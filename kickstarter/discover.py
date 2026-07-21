"""
Phase A: discover API로 종료 프로젝트 목록 수집 (최근 3년치).

- 카테고리 × 상태(successful/failed/canceled)별로 수집
- discover는 한 쿼리당 2,400개까지만 노출되므로, 개수가 넘으면
  마감일(deadline) 구간을 재귀적으로 이등분해 각 구간을 2,400 이하로 낮춘 뒤 페이지네이션
- 펀딩 오픈일(launched_at) 하한은 클라이언트 측에서 필터
"""
import json
import logging
from datetime import date, datetime, timezone
from urllib.parse import urlencode

import config
from kickstarter.client import KickstarterClient

logger = logging.getLogger(__name__)


def _params(category_id: int, state: str, deadline_after: str, deadline_before: str, page: int = 1) -> str:
    return urlencode({
        "format": "json",
        "category_id": category_id,
        "state": state,
        "sort": "end_date",
        "deadline_after": deadline_after,
        "deadline_before": deadline_before,
        "page": page,
    })


def count_hits(client: KickstarterClient, category_id: int, state: str, da: str, db: str) -> int:
    url = f"{config.DISCOVER_URL}?{_params(category_id, state, da, db)}"
    return client.get_json(url).get("total_hits", 0)


def fetch_window(client: KickstarterClient, category_id: int, state: str, da: str, db: str) -> list[dict]:
    """한 (카테고리·상태·마감일 구간)을 끝까지 페이지네이션. 구간은 2,400 이하로 가정."""
    projects = []
    for page in range(1, config.MAX_DISCOVER_PAGES + 1):
        url = f"{config.DISCOVER_URL}?{_params(category_id, state, da, db, page)}"
        data = client.get_json(url)
        page_projects = data.get("projects", [])
        if not page_projects:
            break
        # discover는 검색엔진이라 연관 프로젝트(다른 카테고리)를 끼워 보내므로 정확히 일치하는 것만
        projects.extend(p for p in page_projects if (p.get("category") or {}).get("id") == category_id)
        if not data.get("has_more"):
            break
    return projects


def collect_state(client: KickstarterClient, category_id: int, state: str,
                  da: str, db: str, acc: list[dict]) -> None:
    """마감일 구간 [da, db]를 재귀 분할하며 수집. 각 leaf는 2,400 이하로 내려 전량 도달."""
    total = count_hits(client, category_id, state, da, db)
    if total == 0:
        return
    if total <= config.DISCOVER_CAP:
        got = fetch_window(client, category_id, state, da, db)
        acc.extend(got)
        logger.info("    %s [%s ~ %s]: %d개 (수집 %d)", state, da, db, total, len(got))
        return

    d_a, d_b = date.fromisoformat(da), date.fromisoformat(db)
    if (d_b - d_a).days <= 1:
        # 하루 구간인데도 2,400 초과 (혹시모르니) — 가능한 만큼만 수집하고 경고
        logger.warning("    %s [%s ~ %s]: %d개, 캡 초과 구간 더 못 쪼갬 → 최대 %d개만",
                       state, da, db, total, config.DISCOVER_CAP)
        acc.extend(fetch_window(client, category_id, state, da, db))
        return

    mid = (d_a + (d_b - d_a) / 2).isoformat()
    logger.info("    %s [%s ~ %s]: %d개 > %d → %s 기준 분할",
                state, da, db, total, config.DISCOVER_CAP, mid)
    # [da, mid] ∪ [mid, db] (경계 1일 겹침은 id 중복제거로 처리)
    collect_state(client, category_id, state, da, mid, acc)
    collect_state(client, category_id, state, mid, db, acc)


def run_discover(category_keys: list[str], limit: int | None, delay: float) -> None:
    """카테고리별 목록을 수집해 LIST_FILE(jsonl)에 저장. 실행할 때마다 새로 쓴다."""
    config.DATA_DIR.mkdir(exist_ok=True)
    client = KickstarterClient(delay=delay)

    launched_after_ts = int(
        datetime.fromisoformat(config.LAUNCHED_AFTER).replace(tzinfo=timezone.utc).timestamp()
    )
    seen_ids: set[int] = set()
    count = 0
    with open(config.LIST_FILE, "w", encoding="utf-8") as f:
        for key in category_keys:
            cat = config.CATEGORIES[key]
            logger.info("[discover] %s (category_id=%d), 기간 마감 %s~%s, 오픈>=%s, limit=%s",
                        cat["name"], cat["id"], config.DEADLINE_AFTER, config.DEADLINE_BEFORE,
                        config.LAUNCHED_AFTER, limit)
            cat_projects: list[dict] = []
            for state in config.STATES:
                collect_state(client, cat["id"], state,
                              config.DEADLINE_AFTER, config.DEADLINE_BEFORE, cat_projects)

            kept = 0
            for p in cat_projects:
                if p["id"] in seen_ids:
                    continue
                if (p.get("launched_at") or 0) < launched_after_ts:  # 오픈일 하한 (클라이언트 필터)
                    continue
                seen_ids.add(p["id"])
                p["_category_key"] = key
                f.write(json.dumps(p, ensure_ascii=False) + "\n")
                kept += 1
                count += 1
                if limit and kept >= limit:
                    break
            logger.info("[discover] %s: %d개 수집", cat["name"], kept)
    logger.info("[discover] 완료: 총 %d개 → %s", count, config.LIST_FILE)
