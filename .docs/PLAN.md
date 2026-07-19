# Kickstarter 프로젝트 크롤러 (Tabletop / Video Games → Excel)

## Context

Kickstarter의 Games 하위 두 카테고리의 **최근 3년치 종료 프로젝트**를 수집해 Excel로 출력한다.

- 기간: 펀딩 오픈(launched_at) >= **2023-01-01**, 마감(deadline) **2023-01-01 ~ 2026-06-30**. (`config.py`에서 조절)
- 사이트는 React SPA라 HTML 파싱 대신 **JSON/GraphQL API 직접 호출**:
  - 목록: `GET /discover/advanced?format=json&...` (환율·상태 등 메타데이터 포함, canceled 상태 지원)
  - 본문 상세: `POST /graph` (GraphQL) — 프로젝트 페이지의 `<meta name="csrf-token">` 토큰 사용
- plain curl은 **403/429 (Cloudflare)** → **curl_cffi** (Chrome impersonation) 필수.
- 결정 사항:
  - **최근 3년치 전량** 수집 (약 3만 개)
  - 본문 전문을 Excel 셀에 저장, 32,767자 초과 시 잘리지 않고 다음 칸으로 이어붙임
  - Excel은 **카테고리별로 각각 한 파일**

## 수집 필드

| 구분     | 필드                                                                                                                                              | 출처                                                  |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------- |
| campaign | 제목, 요약(blurb), 본문 전체(story+risks+환경+AI), state, category, backers, goal, pledged, 환율/USD 환산, 달성률, 시작/종료일·기간, 프로젝트 URL | discover JSON + GraphQL                               |
| campaign | 사진/영상 유무·개수                                                                                                                               | story HTML의 `<img>`/video·iframe 카운트 + 메인 video |
| creator  | 이름, 배지(RepeatCreator/BackerFavorite/Superbacker 등), 만든/후원 프로젝트 수                                                                    | GraphQL `creator`                                     |
| 기타     | updates 수(`posts.totalCount`), comments 수(`commentsCount`), Project We Love                                                                     | GraphQL                                               |

## 3단계 파이프라인

흐름: `main.py` → `discover.py`(목록 → `data/projects_list.jsonl`) → `enrich.py`(상세 → `data/projects_enriched.jsonl`) → `export.py`(`output/`에 Excel). 중간 결과가 파일로 남아 재개 가능.

### Phase A — 목록 수집 (`discover.py`)

- `GET /discover/advanced?format=json&category_id={34|35}&state={successful|failed|canceled}&sort=end_date&deadline_after=…&deadline_before=…`
- **2,400개 캡 우회**: discover는 한 쿼리당 최대 2,400개(200페이지)만 노출. (카테고리×상태) total_hits가 2,400을 넘으면 **마감일 구간을 재귀적으로 이등분**(`collect_state`)해 각 leaf를 2,400 이하로 낮춘 뒤 페이지네이션. 날짜는 연속값이라 어떤 밀도든 캡 이하로 쪼개짐.
- 검색엔진 특성상 끼어드는 타 카테고리(Toys/Apps 등)는 `category.id` 정확히 일치하는 것만 남김.
- 펀딩 오픈일 하한(`LAUNCHED_AFTER`)은 클라이언트에서 필터 (discover는 마감일 필터만 지원, launched 필터 없음).
- 확보 필드: id, slug, name, blurb, goal, pledged, currency, usd_pledged, static_usd_rate, fx_rate, backers_count, launched_at, deadline, state, creator, urls, category.

### Phase B — 상세 수집 (`enrich.py`)

- `POST /graph` GraphQL. **한 요청에 프로젝트 10개를 별칭(p0..p9)으로 묶어** 조회 → 요청 수 1/10 (GraphQL 복잡도 한도 500, 프로젝트당 ~40).
- 필드: `story`, `risks`, `commentsCount`, `backersCount`, `percentFunded`, `posts.totalCount`, `isProjectWeLove`, `video`, `environmentalCommitments{commitmentCategory,description}`, `aiDisclosure{...}`, `creator{name,badges,isSuperbacker,launchedProjects,backedProjects,backingsCount}`.
- 배치 실패 시 절반씩 나눠 재시도(문제 프로젝트 격리), 단일 실패는 `_graph=None` 기록.
- CSRF 토큰은 세션당 1회 확보 후 재사용. 결과를 `projects_enriched.jsonl`에 append, 재실행 시 완료 id 자동 스킵.

### Phase C — Excel 출력 (`export.py`)

- pandas + openpyxl. 카테고리별로 각각 한 파일: `kickstarter_projects_{tabletop|video}_games_<날짜>.xlsx`.
- `story_text`: Story → Risks and challenges → Environmental commitments → Use of AI 순서로 합침. **상위 헤더 볼드, 하위 헤더(환경 항목명·AI 설문 질문) 밑줄** (openpyxl 리치 텍스트). AI 자유서술(`otherAiDetails`)은 헤더 없이 본문만.
- 32,767자 초과 본문은 `story_text_2`, `story_text_3`… 로 이어붙여 손실 없음.
- 환율: `pledged`+`currency`(원금), `usd_rate_campaign`→`usd_pledged`(당시 환율), `usd_rate_current`→`usd_pledged_current`(수집 시점 환율, 페이지 표시값).
- 날짜는 unix → `YYYY-MM-DD`, `duration_days` = deadline − launched. 헤더 파란 배경+흰 볼드, story 컬럼 폭 60·나머지 자동, 셀 위쪽 정렬. `--format csv`도 지원.

## client.py (`client.py`)

- curl_cffi `impersonate="chrome"`로 Cloudflare 우회.
- 요청 간 최소 간격(rate limit). 403/429는 시간 기반 차단이라 **세션 재생성 + 점증 쿨다운(20→40→60초, 상한 90초)으로 최대 6회 재시도**. 5xx/네트워크 오류는 백오프 재시도.
- `/graph`용 CSRF 토큰 추출·갱신. GraphQL validation 오류는 재시도 없이 즉시 실패.

## 주요 API

- **discover REST**: 한 쿼리 2,400개 캡(201페이지는 non-JSON 차단). `deadline_after`/`deadline_before`(YYYY-MM-DD) 날짜 필터 지원. canceled 상태 지원. 환율 필드(static_usd_rate/fx_rate) 제공.
- **GraphQL introspection 차단** — 필드 존재 여부는 잘못된 필드 요청 시 나오는 에러 메시지로 확인. 캠페인 탭 FAQ 위 섹션은 story/risks/environmentalCommitments/aiDisclosure 4종.
- (참고) GraphQL `Query.projects` 커넥션은 캡 10,000·본문 bulk 반환이 가능하나 **canceled 상태 필터 불가**(PublicProjectState에 없음)라 미채택. REST + 날짜 분할이 canceled·환율 모두 커버.

## 검증 방법

- 날짜 구간 분할: 각 leaf ≤ 2,400, 전 기간 누락·중복 없음 (mock 단위 테스트로 확인).
- 소규모 실행 후 Excel에서 알려진 프로젝트를 실제 페이지와 대조 (backers, pledged, 기간, updates/comments, creator 배지, 환율값).
- 본문 빈 행 비율(graph 실패 감지), 세 상태 모두 포함 확인.
- 예상 소요: 3년치 ~~29,000개, enrich 배치 10개/요청 → 수십 분~~1시간. 429 잦으면 `--delay` 1.5~2초.
