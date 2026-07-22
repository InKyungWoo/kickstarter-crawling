"""
Phase A: discover API로 종료 프로젝트 목록 수집

- 카테고리 × 상태(successful/failed/canceled)별로 페이지네이션
- discover API는 쿼리당 200페이지(약 2,400개)까지만 노출되므로 샘플 모드에서는 상태별로 limit을 균등 배분해서 수집
"""
import json
import logging
from urllib.parse import urlencode

import config
from kickstarter.client import KickstarterClient

logger = logging.getLogger(__name__)


def fetch_state_projects(client: KickstarterClient, category_id: int, state: str, limit: int) -> list[dict]:
    """한 카테고리·상태 조합에서 최대 limit개 수집."""
    projects = []
    for page in range(1, config.MAX_DISCOVER_PAGES + 1):
        params = {
            "format": "json",
            "category_id": category_id,
            "state": state,
            "sort": "end_date",
            "page": page,
        }
        url = f"{config.DISCOVER_URL}?{urlencode(params)}"
        data = client.get_json(url)
        page_projects = data.get("projects", [])
        if not page_projects:
            break
        # discover는 검색엔진이라 연관 프로젝트(다른 카테고리)를 끼워 보내므로
        # 요청한 카테고리와 정확히 일치하는 것만 남긴다!!!
        matched = [p for p in page_projects
                   if (p.get("category") or {}).get("id") == category_id]
        dropped = len(page_projects) - len(matched)
        projects.extend(matched)
        logger.info("  %s / page %d: +%d%s (누적 %d, total_hits=%s)",
                    state, page, len(matched),
                    f" (타 카테고리 {dropped}개 제외)" if dropped else "",
                    len(projects), data.get("total_hits"))
        if len(projects) >= limit:
            break
        if not data.get("has_more"):
            break
    return projects[:limit]


def run_discover(category_keys: list[str], limit: int, delay: float) -> None:
    """카테고리별 목록을 수집해 LIST_FILE(jsonl)에 저장. 실행할 때마다 새로 쓴다."""
    config.DATA_DIR.mkdir(exist_ok=True)
    client = KickstarterClient(delay=delay)

    seen_ids: set[int] = set()
    count = 0
    with open(config.LIST_FILE, "w", encoding="utf-8") as f:
        for key in category_keys:
            cat = config.CATEGORIES[key]
            logger.info("[discover] %s (category_id=%d), limit=%d", cat["name"], cat["id"], limit)
            # 상태별 균등 배분 (나머지는 앞 상태에 몰아줌)
            base, rem = divmod(limit, len(config.STATES))
            for i, state in enumerate(config.STATES):
                state_limit = base + (1 if i < rem else 0)
                if state_limit == 0:
                    continue
                for p in fetch_state_projects(client, cat["id"], state, state_limit):
                    if p["id"] in seen_ids:
                        continue
                    seen_ids.add(p["id"])
                    p["_category_key"] = key
                    f.write(json.dumps(p, ensure_ascii=False) + "\n")
                    count += 1
    logger.info("[discover] 완료: 총 %d개 → %s", count, config.LIST_FILE)
