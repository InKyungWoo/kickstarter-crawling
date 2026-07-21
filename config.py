"""전역 설정: 카테고리, 경로, 요청 정책."""
from pathlib import Path

BASE_URL = "https://www.kickstarter.com"
DISCOVER_URL = f"{BASE_URL}/discover/advanced"
GRAPH_URL = f"{BASE_URL}/graph"

# Kickstarter 카테고리 ID (Games=12의 하위 카테고리)
CATEGORIES = {
    "tabletop": {"id": 34, "name": "Tabletop Games"},
    "video": {"id": 35, "name": "Video Games"},
}

# 종료된 프로젝트 상태
STATES = ["successful", "failed", "canceled"]

# 요청 정책
REQUEST_DELAY = 1.0          # 요청 간 최소 간격(초)
MAX_RETRIES = 6              # 일시적 차단이 잦아 넉넉히 재시도
RETRY_BACKOFF = 5.0          # 네트워크/5xx 재시도 대기 기본값(초)
BLOCK_COOLDOWN = 20.0        # 403/429 차단 시 첫 쿨다운(초), 재시도마다 증가
BLOCK_COOLDOWN_MAX = 90.0    # 쿨다운 상한(초)
IMPERSONATE = "chrome"       # curl_cffi 브라우저 impersonation 타겟

# discover API는 쿼리당 200페이지(약 2,400개)까지만 노출
MAX_DISCOVER_PAGES = 200
DISCOVER_CAP = 2400          # 한 쿼리로 도달 가능한 최대 개수 (이 이하가 되도록 날짜 구간 분할)
DEFAULT_LIMIT = None         # 카테고리당 상한 (None=전체, 테스트 시 정수 지정)

# 수집 대상 기간
LAUNCHED_AFTER = "2023-01-01"   # 이 날짜 이후 오픈한 프로젝트만 (클라이언트 필터)
DEADLINE_AFTER = "2023-01-01"   # discover deadline_after 파라미터
DEADLINE_BEFORE = "2026-06-30"  # discover deadline_before 파라미터

# 경로
PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "output"
LIST_FILE = DATA_DIR / "projects_list.jsonl"
ENRICHED_FILE = DATA_DIR / "projects_enriched.jsonl"
EXCEL_BASENAME = "kickstarter_projects"

EXCEL_CELL_LIMIT = 32767     # Excel 셀 하나의 최대 문자 수
