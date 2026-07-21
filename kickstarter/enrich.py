"""
Phase B: 프로젝트별 상세 수집 (GraphQL /graph).

- 한 요청에 여러 프로젝트를 별칭(alias)으로 묶어 조회 → 요청 수 1/BATCH_SIZE로 감소
- CSRF 토큰은 세션당 1회만 확보해서 재사용
- 결과는 ENRICHED_FILE에 append (재실행 시 완료된 프로젝트는 자동 스킵)
- 배치 요청이 실패하면 그 배치만 개별 조회로 폴백 (한 프로젝트 오류가 배치 전체를 날리지 않도록)
"""
import json
import logging

import config
from kickstarter.client import KickstarterClient

logger = logging.getLogger(__name__)

BATCH_SIZE = 10

# 프로젝트 하나당 가져올 필드 (별칭마다 재사용)
PROJECT_FIELDS = """
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
    environmentalCommitments { commitmentCategory description }
    aiDisclosure {
      involvesAi involvesGeneration involvesFunding involvesOther
      generatedByAiDetails generatedByAiConsent otherAiDetails
      fundingForAiOption fundingForAiConsent fundingForAiAttribution
    }
    creator {
      id name isSuperbacker isBackerFavorite badges
      launchedProjects { totalCount }
      backedProjects { totalCount }
      backingsCount
    }
"""


def _build_batch_query(n: int) -> str:
    """슬러그 n개를 별칭 p0..p{n-1}로 한 번에 조회하는 쿼리."""
    var_defs = ", ".join(f"$s{i}: String!" for i in range(n))
    aliases = "\n".join(f'  p{i}: project(slug: $s{i}) {{ {PROJECT_FIELDS} }}' for i in range(n))
    return f"query Batch({var_defs}) {{\n{aliases}\n}}"


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


def _fetch_batch(client: KickstarterClient, batch: list[dict]) -> None:
    """배치를 조회해 각 프로젝트에 _graph를 채운다. 실패 시 개별 조회로 폴백."""
    query = _build_batch_query(len(batch))
    variables = {f"s{i}": p["slug"] for i, p in enumerate(batch)}
    try:
        data = client.graph(query, variables)["data"]
        for i, p in enumerate(batch):
            p["_graph"] = data.get(f"p{i}")
            if p["_graph"] is None:
                p["_enrich_error"] = "project not found in graph"
    except Exception as e:
        if len(batch) == 1:
            logger.error("enrich 실패 (%s): %s", batch[0]["slug"], e)
            batch[0]["_graph"] = None
            batch[0]["_enrich_error"] = str(e)
            return
        # 배치 실패 → 절반씩 나눠 재시도 (문제 프로젝트 격리)
        logger.warning("배치 %d개 실패, 분할 재시도: %s", len(batch), str(e)[:80])
        mid = len(batch) // 2
        _fetch_batch(client, batch[:mid])
        _fetch_batch(client, batch[mid:])


def run_enrich(delay: float) -> None:
    with open(config.LIST_FILE, encoding="utf-8") as f:
        projects = [json.loads(line) for line in f]

    done = load_done_ids()
    todo = [p for p in projects if p["id"] not in done]
    logger.info("[enrich] 대상 %d개 (이미 완료 %d개 스킵), 배치 %d개씩",
                len(todo), len(projects) - len(todo), BATCH_SIZE)
    if not todo:
        return

    client = KickstarterClient(delay=delay)
    done_count = fail_count = 0
    with open(config.ENRICHED_FILE, "a", encoding="utf-8") as f:
        for start in range(0, len(todo), BATCH_SIZE):
            batch = todo[start:start + BATCH_SIZE]
            _fetch_batch(client, batch)
            for p in batch:
                if p.get("_graph") is None:
                    fail_count += 1
                f.write(json.dumps(p, ensure_ascii=False) + "\n")
            f.flush()
            done_count += len(batch)
            logger.info("[enrich] %d/%d (실패 %d)", done_count, len(todo), fail_count)
    logger.info("[enrich] 완료 → %s", config.ENRICHED_FILE)
