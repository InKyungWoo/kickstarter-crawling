"""
Kickstarter HTTP 클라이언트

curl_cffi로 Chrome을 impersonate해서 Cloudflare 봇 탐지를 우회한다
- 요청 간 rate limit
- 지수 백오프 재시도, 403/429 시 세션 재생성
- /graph 요청용 CSRF 토큰 관리 (프로젝트 페이지에서 추출, 실패 시 갱신)
"""
import logging
import re
import time

from curl_cffi import requests as curl_requests

import config

logger = logging.getLogger(__name__)

CSRF_RE = re.compile(r'<meta name="csrf-token" content="([^"]+)"')


class KickstarterClient:
    def __init__(self, delay: float = config.REQUEST_DELAY):
        self.delay = delay
        self._last_request_at = 0.0
        self.csrf_token: str | None = None
        self._new_session()

    def _new_session(self):
        self.session = curl_requests.Session(impersonate=config.IMPERSONATE)
        self.csrf_token = None

    def _throttle(self):
        elapsed = time.monotonic() - self._last_request_at
        if elapsed < self.delay:
            time.sleep(self.delay - elapsed)
        self._last_request_at = time.monotonic()

    def _request(self, method: str, url: str, **kwargs):
        """재시도/백오프를 적용한 단일 요청. 403/429면 세션을 갈아엎는다."""
        last_error = None
        for attempt in range(config.MAX_RETRIES + 1):
            if attempt > 0:
                wait = config.RETRY_BACKOFF * attempt
                logger.warning("재시도 %d/%d (%.0f초 대기): %s", attempt, config.MAX_RETRIES, wait, url)
                time.sleep(wait)
            self._throttle()
            try:
                resp = self.session.request(method, url, timeout=30, **kwargs)
            except Exception as e:
                last_error = e
                continue
            if resp.status_code in (403, 429):
                last_error = RuntimeError(f"HTTP {resp.status_code}: {url}")
                self._new_session()
                continue
            if resp.status_code >= 500:
                last_error = RuntimeError(f"HTTP {resp.status_code}: {url}")
                continue
            return resp
        raise last_error

    def get(self, url: str, **kwargs):
        return self._request("GET", url, **kwargs)

    def get_json(self, url: str, **kwargs) -> dict:
        resp = self._request("GET", url, headers={"Accept": "application/json"}, **kwargs)
        return resp.json()

    def get_project_html(self, url: str) -> str:
        """프로젝트 페이지 HTML을 가져오고 CSRF 토큰을 갱신한다."""
        resp = self._request("GET", url)
        html = resp.text
        m = CSRF_RE.search(html)
        if m:
            self.csrf_token = m.group(1)
        return html

    def graph(self, query: str, variables: dict) -> dict:
        """GraphQL POST. CSRF 토큰이 없으면 먼저 홈페이지에서 확보한다."""
        if not self.csrf_token:
            self.get_project_html(config.BASE_URL)
        for attempt in range(2):
            resp = self._request(
                "POST",
                config.GRAPH_URL,
                json={"query": query, "variables": variables},
                headers={"X-CSRF-Token": self.csrf_token or "", "Content-Type": "application/json"},
            )
            data = resp.json()
            if resp.status_code == 200 and "data" in data:
                return data
            # 쿼리 자체가 잘못된 경우(validation error)는 재시도해도 소용없다
            errors = data.get("errors", [])
            messages = " | ".join(e.get("message", "") for e in errors)
            if errors and "CSRF" not in messages and resp.status_code == 200:
                raise RuntimeError(f"GraphQL 쿼리 오류: {messages}")
            if attempt == 0:
                logger.warning("graph 응답 이상(HTTP %d), CSRF 갱신 후 재시도", resp.status_code)
                self._new_session()
                self.get_project_html(config.BASE_URL)
        raise RuntimeError(f"graph 요청 실패: {data}")
