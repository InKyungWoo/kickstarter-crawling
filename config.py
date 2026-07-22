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
MAX_RETRIES = 3
RETRY_BACKOFF = 5.0          # 재시도 대기 기본값(초), 시도마다 배수 증가
IMPERSONATE = "chrome"       # curl_cffi 브라우저 impersonation 타겟

# discover API는 쿼리당 200페이지(약 2,400개)까지만 노출
MAX_DISCOVER_PAGES = 200
DEFAULT_LIMIT = 200          # 카테고리당 기본 수집 개수 (샘플 모드)

# 경로
PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "output"
LIST_FILE = DATA_DIR / "projects_list.jsonl"
ENRICHED_FILE = DATA_DIR / "projects_enriched.jsonl"
EXCEL_BASENAME = "kickstarter_projects"

EXCEL_CELL_LIMIT = 32767     # Excel 셀 하나의 최대 문자 수
