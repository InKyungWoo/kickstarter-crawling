# Kickstarter Crawler

Kickstarter의 **Tabletop Games / Video Games** 카테고리에서 **최근 3년치 종료 프로젝트**를 수집해 Excel로 정리하는 크롤러입니다.

- 기본 수집 대상: 펀딩 오픈 **2023-01-01 이후**, 마감 **2026-06-30 이전** (기간은 `config.py`에서 조절)
- 캠페인 본문 전체, 펀딩 금액, 크리에이터 정보, updates,comments cnt 등

## 설치 (최초 1회)

Python 3.10 이상이 필요합니다.

```bash
git clone <저장소 주소>
cd kickstarter-crawler
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## 실행

```bash
source .venv/bin/activate   # 터미널 열 때마다 1회 (또는 python 대신 .venv/bin/python 사용)

python main.py                        # 기본: 두 카테고리 3년치 전량 수집 → Excel
python main.py --limit 20             # 소규모 테스트 (카테고리당 20개)
python main.py --category tabletop    # 한 카테고리만 (tabletop / video)
python main.py --delay 1.5            # 요청 간격을 넉넉히 (속도제한 회피)
python main.py --stage export         # 수집 없이 Excel만 다시 생성
python main.py --stage export --format csv   # CSV로 출력
```

| 옵션         | 기본값 | 설명                                         |
| ------------ | ------ | -------------------------------------------- |
| `--category` | 둘 다  | `tabletop`, `video` 중 선택 (복수 가능)      |
| `--limit`    | 전체   | 카테고리당 수집 개수 상한 (테스트용)         |
| `--delay`    | 1.0    | 요청 간격(초) — 너무 줄이면 429/403 차단     |
| `--stage`    | all    | `discover` / `enrich` / `export` 단계별 실행 |
| `--format`   | xlsx   | `xlsx` 또는 `csv`                            |

- 실행 도중 중단해도 다시 실행하면 **수집 완료된 프로젝트는 건너뛰고 이어서** 진행합니다 (enrich 단계).
- 결과 파일명에는 날짜가 붙고, 같은 이름이 있으면 `(1)`, `(2)`가 붙어 덮어쓰지 않습니다.
- 3년치 전량은 약 3만 개 규모라 enrich에 **수십 분~1시간** 정도 걸립니다.

## 파일 구성

```
kickstarter-crawler/
├── .docs/PLAN.md           # 설계 문서 및 구현 노트
├── data/                   # (자동 생성) 수집 원본 — 중단 후 재개를 위한 체크포인트
│   ├── projects_list.jsonl      # 목록 수집 결과 (프로젝트당 1줄)
│   └── projects_enriched.jsonl  # 상세 수집 결과 (본문 HTML 원본 포함)
├── kickstarter/            # 크롤러 핵심 로직
│   ├── client.py           #   HTTP 클라이언트 — curl_cffi로 Chrome 위장(Cloudflare 우회),
│   │                       #   요청 간격 제한, 403/429 시 세션 재생성 + 점증 쿨다운 재시도, CSRF 토큰 관리
│   ├── discover.py         #   1단계(목록) — discover API로 카테고리×상태별 수집,
│   │                       #   마감일 구간을 재귀 분할해 2,400 캡 우회, data/projects_list.jsonl 저장
│   ├── enrich.py           #   2단계(상세) — GraphQL 별칭 배치(10개/요청)로 본문·risks·환경·AI 섹션,
│   │                       #   크리에이터 배지/카운트, updates/comments 수를 받아 저장
│   └── export.py           #   3단계(출력) — 수집 데이터를 Excel/CSV로 변환,
│                           #   본문 섹션 합치기·서식, 환율 컬럼 계산, 카테고리별 파일 저장
├── output/                 # (자동 생성) 최종 결과물 (xlsx/csv)
├── config.py               # 전역 설정 — 카테고리 ID, 수집 기간, 요청 간격/재시도 정책, 경로
├── main.py                 # 실행 진입점 — CLI 옵션을 받아 3단계 파이프라인을 순서대로 실행
└── requirements.txt        # 의존성 목록
```

## 결과물 (Excel)

카테고리별로 **각각 한 파일**(`kickstarter_projects_tabletop_games_<날짜>.xlsx`, `..._video_games_<날짜>.xlsx`), 프로젝트당 1행. 주요 컬럼:

- **캠페인**: `project_url`, `title`, `blurb`(요약), `story_text`(본문 전체 — Story + Risks and challenges + Environmental commitments + Use of AI, 페이지 FAQ 위의 모든 섹션. 상위 헤더 볼드·하위 헤더 밑줄), `state`, `category`
- **금액**: `pledged`+`currency`(원금), `usd_rate_campaign`→`usd_pledged`(캠페인 당시 환율 환산), `usd_rate_current`→`usd_pledged_current`(수집 시점 환율 환산, 페이지 표시 방식), `goal`, `percent_funded`
- **미디어**: `has_photos`/`photo_count`, `has_video`/`story_video_count`/`has_main_video`
- **기간**: `start_date`, `end_date`, `duration_days`
- **크리에이터**: `creator_name`, `creator_badges`(RepeatCreator, BackerFavorite, Superbacker 등), `creator_created_count`, `creator_backed_count`, `is_project_we_love`
- **활동**: `updates_count`, `comments_count`

> 참고: Excel 셀 한도(32,767자)를 넘는 긴 본문은 잘리지 않고 다음 칸(`story_text_2`, `story_text_3`, …)으로 이어집니다.<br/>
> 본문 원본(HTML 포함)은 `data/projects_enriched.jsonl`에 보존됩니다.

## 동작 방식

1. **discover** — `GET /discover/advanced?format=json&category_id=…&state=…&deadline_after=…&deadline_before=…`로 목록 수집. 한 쿼리는 최대 2,400개까지만 노출되므로, (카테고리×상태) 개수가 넘으면 **마감일 구간을 재귀적으로 이등분**해 각 구간을 2,400 이하로 낮춘 뒤 페이지네이션합니다. 검색엔진 특성상 끼어드는 타 카테고리 프로젝트는 `category.id`로 걸러내고, 펀딩 오픈일 하한은 클라이언트에서 필터합니다.
2. **enrich** — 본문·상세 정보는 GraphQL(`POST /graph`)로 수집합니다. **한 요청에 프로젝트 10개를 별칭(alias)으로 묶어** 조회해 요청 수를 크게 줄입니다 (GraphQL 복잡도 한도 500). CSRF 토큰은 세션당 1회만 확보합니다.
3. **export** — 목록과 상세를 합쳐 카테고리별 Excel/CSV를 생성합니다.

> Kickstarter는 Cloudflare 봇 차단을 쓰기 때문에 일반 HTTP 요청은 403이 뜹니다.<br/>
> 이 프로젝트는 `curl_cffi`의 Chrome impersonation으로 우회합니다.<br/>
> 요청 간격(기본 1초)을 지나치게 줄이면 429/403 차단이 발생할 수 있으며 이 경우 client가 세션을 새로 만들고 점점 긴 쿨다운으로 재시도합니다. <br/>
> 차단이 잦으면 `--delay`를 1.5~2초로 올려 실행하세요.
