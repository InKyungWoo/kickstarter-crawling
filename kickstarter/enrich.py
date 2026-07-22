"""
Phase B: 프로젝트별 상세 수집 (GraphQL /graph, 프로젝트당 1 요청)

- 배지·카운트가 모두 GraphQL로 제공되므로 프로젝트별 HTML 요청은 불필요
- CSRF 토큰은 세션당 1회만 확보해서 재사용
- 결과는 ENRICHED_FILE에 append (재실행 시 완료된 프로젝트는 자동 스킵)
"""
import json
import logging

import config
from kickstarter.client import KickstarterClient

logger = logging.getLogger(__name__)

PROJECT_QUERY = """
query GetProject($slug: String!) {
  project(slug: $slug) {
    id
    story
    risks
    commentsCount
    backersCount
    percentFunded
    state
    launchedAt
    deadlineAt
    isProjectWeLove
    video { id }
    posts { totalCount }
    creator {
      id
      name
      isSuperbacker
      isBackerFavorite
      badges
      launchedProjects { totalCount }
      backedProjects { totalCount }
      backingsCount
    }
  }
}
"""


def load_done_ids() -> set[int]:
    done = set()
    if config.ENRICHED_FILE.exists():
        with open(config.ENRICHED_FILE, encoding="utf-8") as f:
            for line in f:
                try:
                    done.add(json.loads(line)["id"])
                except (json.JSONDecodeError, KeyError):
                    continue
    return done


def run_enrich(delay: float) -> None:
    with open(config.LIST_FILE, encoding="utf-8") as f:
        projects = [json.loads(line) for line in f]

    done = load_done_ids()
    todo = [p for p in projects if p["id"] not in done]
    logger.info("[enrich] 대상 %d개 (이미 완료 %d개 스킵)", len(todo), len(projects) - len(todo))
    if not todo:
        return

    client = KickstarterClient(delay=delay)
    failures = 0
    with open(config.ENRICHED_FILE, "a", encoding="utf-8") as f:
        for i, p in enumerate(todo, 1):
            try:
                result = client.graph(PROJECT_QUERY, {"slug": p["slug"]})
                p["_graph"] = result["data"]["project"]  # 삭제된 프로젝트면 None
                if p["_graph"] is None:
                    p["_enrich_error"] = "project not found in graph"
            except Exception as e:
                failures += 1
                logger.error("enrich 실패 (%s): %s", p["slug"], e)
                p["_graph"] = None
                p["_enrich_error"] = str(e)
            f.write(json.dumps(p, ensure_ascii=False) + "\n")
            f.flush()
            if i % 25 == 0 or i == len(todo):
                logger.info("[enrich] %d/%d (실패 %d)", i, len(todo), failures)
    logger.info("[enrich] 완료 → %s", config.ENRICHED_FILE)
