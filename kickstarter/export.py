"""
Phase C: 수집된 JSONL → Excel(카테고리별 시트) 또는 CSV 변환.
"""
import json
import logging
import re
from datetime import datetime, timezone

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

import config

logger = logging.getLogger(__name__)

# Excel이 허용하지 않는 제어 문자
ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# Environmental commitments의 commitmentCategory enum → 페이지 표시 하위 헤더
ENV_CATEGORY_LABELS = {
    "long_lasting_design": "Long-lasting design",
    "reusability_and_recyclability": "Reusability and recyclability",
    "sustainable_materials": "Sustainable materials",
    "sustainable_distribution": "Sustainable Distribution",
    "environmentally_friendly_factories": "Environmentally friendly factories",
    "something_else": "Something else",
}

# Use of AI 응답 필드 → 페이지에 표시되는 실제 설문 질문(하위 헤더)
# label이 None인 필드(otherAiDetails)는 헤더 없음
AI_FIELD_LABELS = [
    ("generatedByAiDetails",
     "What parts of your project will use AI generated content? Please be as specific as possible."),
    ("generatedByAiConsent",
     "Do you have the consent of owners of the works that were (or will be) used to produce the AI generated portion of your projects? Please explain."),
    ("fundingForAiOption",
     "Regarding the potential impact of your project's AI usage on jobs, please describe how it might affect existing jobs and skills."),
    ("fundingForAiConsent",
     "Do you have the consent of owners of the works that were (or will be) used to develop your project's AI technology? Please explain."),
    ("fundingForAiAttribution",
     "Will you attribute the works used to develop your project's AI technology? Please explain."),
    ("otherAiDetails", None),
]

# 상위 섹션 헤더 라벨
TOP_STORY = "Story"
TOP_RISKS = "Risks and challenges"
TOP_ENV = "Environmental commitments"
TOP_AI = "Use of AI"

# story_text 앞/뒤 컬럼 (story 컬럼은 본문 길이에 따라 동적으로 삽입)
COLUMNS_LEAD = ["project_url", "title", "blurb"]
COLUMNS_TAIL = [
    "state", "category",
    "backers_count", "goal", "pledged", "currency",
    "usd_rate_campaign", "usd_pledged", "usd_rate_current", "usd_pledged_current", "percent_funded",
    "has_photos", "photo_count", "has_video", "story_video_count", "has_main_video",
    "start_date", "end_date", "duration_days",
    "creator_name", "creator_badges", "creator_created_count", "creator_backed_count",
    "is_project_we_love", "updates_count", "comments_count",
]


def parse_story(story_html: str | None) -> tuple[str, int, int]:
    """story HTML → (plain text, 이미지 수, 동영상 수)"""
    if not story_html:
        return "", 0, 0
    soup = BeautifulSoup(story_html, "lxml")
    img_count = len(soup.find_all("img"))
    video_count = len(soup.find_all("video")) + len(soup.find_all("iframe"))
    text = soup.get_text("\n", strip=True)
    return text, img_count, video_count


def build_story(g: dict) -> tuple[list[tuple[str, str | None]], int, int]:
    """캠페인 본문을 서식 조각(runs)과 이미지/영상 수로 반환.

    runs 원소: (텍스트, 종류). 종류 — 'top'(상위 헤더, 볼드) / 'sub'(하위 헤더, 밑줄) / None(본문).
    페이지 FAQ 위 순서 그대로: Story → Risks and challenges → Environmental commitments → Use of AI.
    """
    body, img_count, video_count = parse_story(g.get("story"))
    runs: list[tuple[str, str | None]] = [(TOP_STORY, "top"), ("\n", None), (body, None)]

    risks = (g.get("risks") or "").strip()
    if risks:
        runs += [("\n\n", None), (TOP_RISKS, "top"), ("\n", None), (risks, None)]

    env_items = []
    for e in (g.get("environmentalCommitments") or []):
        desc = (e.get("description") or "").strip()
        if not desc:
            continue
        cat = e.get("commitmentCategory")
        header = ENV_CATEGORY_LABELS.get(cat, (cat or "").replace("_", " ").capitalize())
        env_items.append((header, desc))
    if env_items:
        runs += [("\n\n", None), (TOP_ENV, "top")]
        for i, (header, desc) in enumerate(env_items):
            runs.append(("\n\n" if i else "\n", None))
            runs += [(header, "sub"), ("\n", None), (desc, None)]

    ai = g.get("aiDisclosure") or {}
    if ai.get("involvesAi"):
        ai_items = [(label, str(ai[key]).strip()) for key, label in AI_FIELD_LABELS if ai.get(key)]
        if ai_items:
            runs += [("\n\n", None), (TOP_AI, "top")]
            for i, (label, val) in enumerate(ai_items):
                runs.append(("\n\n" if i else "\n", None))
                if label:  # 설문 질문이 있는 필드는 하위 헤더로, 자유 서술(otherAiDetails)은 본문만
                    runs += [(label, "sub"), ("\n", None)]
                runs.append((val, None))

    return runs, img_count, video_count


def _inline_font(kind: str | None) -> InlineFont | None:
    if kind == "top":
        return InlineFont(b=True)       # 상위 헤더: 볼드
    if kind == "sub":
        return InlineFont(u="single")   # 하위 헤더: 밑줄
    return None


def runs_to_cells(runs: list[tuple[str, str | None]], limit: int) -> list[CellRichText]:
    """runs를 셀당 limit 문자로 흘려 여러 리치 텍스트 셀로 나눈다 (서식 유지, 손실 없음).

    본문이 Excel 셀 한도를 넘으면 다음 셀(story_text_2, ...)로 이어진다.
    """
    cells: list[CellRichText] = []
    blocks: list = []
    used = 0

    def flush():
        nonlocal blocks, used
        cells.append(CellRichText(blocks) if blocks else CellRichText(""))
        blocks, used = [], 0

    for text, kind in runs:
        text = ILLEGAL_CHARS.sub("", text)
        while text:
            if used >= limit:
                flush()
            room = limit - used
            piece, text = text[:room], text[room:]
            font = _inline_font(kind)
            blocks.append(TextBlock(font, piece) if font else piece)
            used += len(piece)
    if blocks:
        flush()
    return cells or [CellRichText("")]


def cell_to_plain(cell: CellRichText) -> str:
    """리치 텍스트 셀 → 평문 (CSV/DataFrame 값용)."""
    return "".join(b.text if isinstance(b, TextBlock) else str(b) for b in cell)


def fmt_date(ts) -> str | None:
    if not ts:
        return None
    return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")


def build_row(p: dict, img_count: int, video_count: int) -> dict:
    """story 이외의 컬럼을 만든다. story 컬럼은 run_export에서 별도로 채운다."""
    g = p.get("_graph") or {}
    creator_g = g.get("creator") or {}

    badges = list(creator_g.get("badges") or [])
    if creator_g.get("isSuperbacker") and "Superbacker" not in badges:
        badges.append("Superbacker")

    # 후원 목록이 비공개면 backedProjects가 null → backingsCount로 폴백
    backed = (creator_g.get("backedProjects") or {}).get("totalCount")
    if backed is None:
        backed = creator_g.get("backingsCount")

    launched, deadline = p.get("launched_at"), p.get("deadline")

    return {
        "project_url": (p.get("urls") or {}).get("web", {}).get("project"),
        "title": p.get("name"),
        "blurb": p.get("blurb"),
        "state": p.get("state"),
        "category": (p.get("category") or {}).get("name"),
        "backers_count": g.get("backersCount", p.get("backers_count")),
        "goal": p.get("goal"),
        "pledged": p.get("pledged"),    ## 원금
        "currency": p.get("currency"),  ## 원금 통화
        "usd_rate_campaign": round(float(p["static_usd_rate"]), 4) if p.get("static_usd_rate") else None,   ## 당시 환율
        "usd_pledged": round(float(p.get("usd_pledged") or 0), 2),  ## 당시 환율 적용값
        "usd_rate_current": round(float(p["fx_rate"]), 4) if p.get("fx_rate") else None, ## 수집 시점 환율
        "usd_pledged_current": round(float(p.get("pledged") or 0) * float(p["fx_rate"]), 2) if p.get("fx_rate") else None,  ## 수집 시점 환율 적용값 (페이지 표시)
        "percent_funded": g.get("percentFunded", p.get("percent_funded")),
        "has_photos": img_count > 0,
        "photo_count": img_count,
        "has_video": bool(g.get("video")) or video_count > 0,
        "story_video_count": video_count,
        "has_main_video": bool(g.get("video") or p.get("video")),
        "start_date": fmt_date(launched),
        "end_date": fmt_date(deadline),
        "duration_days": round((deadline - launched) / 86400, 1) if launched and deadline else None,
        "creator_name": (p.get("creator") or {}).get("name"),
        "creator_badges": ", ".join(badges),
        "creator_created_count": (creator_g.get("launchedProjects") or {}).get("totalCount"),
        "creator_backed_count": backed,
        "is_project_we_love": bool(g.get("isProjectWeLove", p.get("staff_pick"))),
        "updates_count": (g.get("posts") or {}).get("totalCount"),
        "comments_count": g.get("commentsCount"),
    }


def unique_path(path):
    """이미 같은 이름의 파일이 있으면 ' (1)', ' (2)'를 붙여 덮어쓰지 않는 경로를 돌려준다."""
    if not path.exists():
        return path
    n = 1
    while True:
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def _build_sheet(sel_projects: list[dict]):
    """한 카테고리의 (DataFrame, story 셀 목록, story 컬럼명) 구성."""
    rows, cells_per_row = [], []
    for p in sel_projects:
        runs, img_count, video_count = build_story(p.get("_graph") or {})
        rows.append(build_row(p, img_count, video_count))
        cells_per_row.append(runs_to_cells(runs, config.EXCEL_CELL_LIMIT))

    max_cells = max((len(c) for c in cells_per_row), default=1)
    # 본문이 한 셀을 넘으면 story_text_2, story_text_3, ... 로 이어붙임
    story_cols = ["story_text"] + [f"story_text_{i}" for i in range(2, max_cells + 1)]

    for row, cells in zip(rows, cells_per_row):
        for i, name in enumerate(story_cols):
            row[name] = cell_to_plain(cells[i]) if i < len(cells) else ""

    columns = COLUMNS_LEAD + story_cols + COLUMNS_TAIL
    return pd.DataFrame(rows, columns=columns), cells_per_row, story_cols


def run_export(fmt: str = "xlsx") -> None:
    config.OUTPUT_DIR.mkdir(exist_ok=True)
    # 최신 목록 파일을 기준으로 삼고, 상세 수집 파일에서는 _graph만 가져온다
    # (enriched 파일에는 과거 실행의 낡은 분류 정보가 남아 있을 수 있음)
    with open(config.LIST_FILE, encoding="utf-8") as f:
        projects = [json.loads(line) for line in f]
    graphs = {}
    with open(config.ENRICHED_FILE, encoding="utf-8") as f:
        for line in f:
            r = json.loads(line)
            graphs[r["id"]] = r.get("_graph")
    missing = 0
    for p in projects:
        p["_graph"] = graphs.get(p["id"])
        if p["_graph"] is None:
            missing += 1
    if missing:
        logger.warning("[export] 상세 데이터 없는 프로젝트 %d개 (enrich 미실행분)", missing)
    date_str = datetime.now().strftime("%Y-%m-%d")

    # 카테고리별로 (DataFrame, story 셀 목록, story 컬럼명) 구성
    sheets = {}
    for key, cat in config.CATEGORIES.items():
        sel = [p for p in projects
               if p.get("_category_key") == key and (p.get("category") or {}).get("id") == cat["id"]]
        if sel:
            sheets[cat["name"]] = _build_sheet(sel)

    if fmt == "csv":
        for name, (df, _, _) in sheets.items():
            path = unique_path(config.OUTPUT_DIR / f"kickstarter_{name.lower().replace(' ', '_')}_{date_str}.csv")
            df.to_csv(path, index=False, encoding="utf-8-sig")
            logger.info("[export] %s: %d행 → %s", name, len(df), path)
        return

    # 카테고리(tabletop / video)별로 각각 한 파일로 저장
    for name, (df, cells_per_row, story_cols) in sheets.items():
        slug = name.lower().replace(" ", "_")
        path = unique_path(config.OUTPUT_DIR / f"{config.EXCEL_BASENAME}_{slug}_{date_str}.xlsx")
        _write_xlsx(path, name, df, cells_per_row, story_cols)
        logger.info("[export] %s: %d행 → %s", name, len(df), path.name)
    logger.info("[export] 완료 (카테고리 %d개)", len(sheets))


def _write_xlsx(path, sheet_name: str, df: pd.DataFrame, cells_per_row: list, story_cols: list) -> None:
    """한 청크를 스타일 적용해 xlsx 파일로 쓴다."""
    top_align = Alignment(vertical="top")
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(fill_type="solid", start_color="4472C4", end_color="4472C4")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        ws = writer.sheets[sheet_name[:31]]
        ws.freeze_panes = "A2"
        # story 컬럼은 고정 너비 60, 나머지는 내용 길이에 맞춰 자동 조절
        for idx, col_name in enumerate(df.columns, start=1):
            letter = get_column_letter(idx)
            if col_name in story_cols:
                width = 60
            else:
                max_len = max(len(str(v)) for v in [col_name, *df[col_name].fillna("")])
                width = min(max_len + 2, 100)
            ws.column_dimensions[letter].width = width
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = top_align
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
        # story 컬럼들을 리치 텍스트(상위 헤더 볼드/하위 헤더 밑줄)로 채움
        first_story_col = list(df.columns).index(story_cols[0]) + 1
        for r, cells in enumerate(cells_per_row, start=2):
            for i, cell_rt in enumerate(cells):
                ws.cell(row=r, column=first_story_col + i).value = cell_rt
