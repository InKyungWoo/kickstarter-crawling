"""
Kickstarter 종료 프로젝트 크롤러 실행 진입점.

사용 예:
    python main.py                          # 최근 3년치 두 카테고리 전량 수집 → Excel
    python main.py --limit 20               # 소규모 테스트 (카테고리당 20개)
    python main.py --category tabletop      # 한 카테고리만
    python main.py --stage export --format csv   # 수집된 데이터를 CSV로만 재변환

수집 기간은 config.py의 LAUNCHED_AFTER / DEADLINE_AFTER / DEADLINE_BEFORE로 조절.
"""
import argparse
import logging

import config
from kickstarter.discover import run_discover
from kickstarter.enrich import run_enrich
from kickstarter.export import run_export


def main():
    parser = argparse.ArgumentParser(description="Kickstarter 종료 프로젝트 크롤러")
    parser.add_argument("--category", nargs="+", choices=list(config.CATEGORIES),
                        default=list(config.CATEGORIES), help="수집할 카테고리 (기본: 전부)")
    parser.add_argument("--limit", type=int, default=config.DEFAULT_LIMIT,
                        help="카테고리당 수집 개수 상한 (기본: 전체). 테스트 시 작은 수 지정")
    parser.add_argument("--delay", type=float, default=config.REQUEST_DELAY,
                        help=f"요청 간격 초 (기본 {config.REQUEST_DELAY})")
    parser.add_argument("--stage", choices=["all", "discover", "enrich", "export"],
                        default="all", help="실행할 단계 (기본 all)")
    parser.add_argument("--format", choices=["xlsx", "csv"], default="xlsx",
                        help="출력 형식 (기본 xlsx)")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        datefmt="%H:%M:%S")

    if args.stage in ("all", "discover"):
        run_discover(args.category, args.limit, args.delay)
    if args.stage in ("all", "enrich"):
        run_enrich(args.delay)
    if args.stage in ("all", "export"):
        run_export(args.format)


if __name__ == "__main__":
    main()
