"""기술교범 한글 파일(.hwp/.hwpx)에서 조건에 맞는 표를 선별해 엑셀로 추출.

사용법:
    python -m src.main <입력파일 또는 폴더> [-c 규칙파일] [-o 출력폴더] [--all]

예시:
    python -m src.main data\\교범.hwpx
    python -m src.main data\\ -c config\\target_tables.yaml -o output\\
    python -m src.main data\\교범.hwpx --all     # 규칙 무시하고 모든 표 추출
"""
from __future__ import annotations

import argparse
import logging
import sys
import zipfile
from pathlib import Path

import yaml

from .excel_writer import SHEET_LAYOUT_PER_TABLE, write_workbook
from .hwpx_parser import parse_tables
from .models import MatchResult
from .table_matcher import select_tables
from .visual_merge import apply_visual_merges

logger = logging.getLogger("extractor")


def setup_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_dir / "extract.log", encoding="utf-8"),
        ],
    )


def load_config(config_path: Path, *, require_rules: bool = True) -> tuple[list[dict], dict]:
    """YAML 설정에서 rules 목록과 output 옵션을 읽어 반환."""
    with open(config_path, encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}
    rules = config.get("rules") or []
    if require_rules and not rules:
        raise ValueError(f"규칙 파일에 rules 항목이 없습니다: {config_path}")
    output = config.get("output") or {}
    return rules, output


def collect_input_files(target: Path) -> list[Path]:
    if target.is_file():
        return [target]
    if target.is_dir():
        files = sorted(
            p for p in target.iterdir()
            if p.suffix.lower() in (".hwp", ".hwpx")
        )
        if not files:
            raise FileNotFoundError(f"폴더에 .hwp/.hwpx 파일이 없습니다: {target}")
        return files
    raise FileNotFoundError(f"입력 경로를 찾을 수 없습니다: {target}")


def process_file(
    path: Path,
    rules: list[dict],
    out_dir: Path,
    extract_all: bool,
    *,
    include_title: bool = True,
    include_footer: bool = True,
    visual_merge: bool = True,
    sheet_layout: str = "per_table",
) -> dict:
    """파일 하나 처리 후 요약 정보 반환."""
    logger.info("=" * 70)
    logger.info(f"처리 시작: {path.name}")

    if zipfile.is_zipfile(path):
        # 정상 HWPX → XML 직접 파싱 (한글 불필요).
        # 꼬리말도 구역 XML(+이전 구역 상속)로 채운다.
        # COM Goto 보정은 한글 '찾아가기' 팝업을 띄우므로 사용하지 않는다.
        tables = parse_tables(path)
    else:
        # DRM 암호화 또는 구형 HWP → 한글 COM으로 문서를 열어 표를 읽음
        logger.info("  HWPX(ZIP) 형식이 아님 (DRM/구형 HWP 추정) → 한글로 문서를 열어 표를 읽습니다...")
        from .hwp_com_reader import read_tables_via_com
        tables = read_tables_via_com(path)
    if visual_merge:
        tables = [apply_visual_merges(t) for t in tables]
    logger.info(f"  문서 내 표 {len(tables)}개 발견")

    if extract_all:
        results = [MatchResult(table=t, rule_name="전체", score=0, reasons=["--all 옵션"])
                   for t in tables]
    else:
        results = select_tables(tables, rules, logger=logger)

    out_xlsx = out_dir / f"{path.stem}_tables.xlsx"
    write_workbook(
        results,
        out_xlsx,
        include_title=include_title,
        include_footer=include_footer,
        sheet_layout=sheet_layout,
    )
    logger.info(f"  추출 {len(results)}/{len(tables)}개 표 → {out_xlsx}")

    return {"file": path.name, "total": len(tables), "extracted": len(results),
            "output": str(out_xlsx)}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="한글 파일(.hwp/.hwpx)에서 조건에 맞는 표를 엑셀로 추출")
    parser.add_argument("input", help="입력 .hwp/.hwpx 파일 또는 폴더 경로")
    parser.add_argument("-c", "--config", default="config/target_tables.yaml",
                        help="표 선별 규칙 YAML 파일 (기본: config/target_tables.yaml)")
    parser.add_argument("-o", "--output", default="output",
                        help="출력 폴더 (기본: output)")
    parser.add_argument("--all", action="store_true",
                        help="선별 규칙을 무시하고 문서의 모든 표 추출 (규칙 튜닝용)")
    args = parser.parse_args(argv)

    out_dir = Path(args.output)
    setup_logging(out_dir)

    config_path = Path(args.config)
    if args.all:
        _, output_opts = load_config(config_path, require_rules=False)
        rules = []
    else:
        rules, output_opts = load_config(config_path, require_rules=True)
    include_title = bool(output_opts.get("include_title", True))
    include_footer = bool(output_opts.get("include_footer", True))
    visual_merge = bool(output_opts.get("visual_merge", True))
    sheet_layout = output_opts.get("sheet_layout", SHEET_LAYOUT_PER_TABLE)
    if sheet_layout not in ("per_table", "single_sheet"):
        sheet_layout = SHEET_LAYOUT_PER_TABLE
    files = collect_input_files(Path(args.input))
    layout_label = (
        "표당 시트 1개" if sheet_layout == SHEET_LAYOUT_PER_TABLE else "단일 시트 연속 배치"
    )
    logger.info(
        f"입력 파일 {len(files)}개, 규칙 {len(rules)}개 로드"
        f" (표 제목: {'포함' if include_title else '미포함'},"
        f" 꼬리말: {'포함' if include_footer else '미포함'},"
        f" 숨은 가로선 병합: {'적용' if visual_merge else '미적용'},"
        f" 시트 배치: {layout_label})"
    )

    summary, failed = [], []
    for f in files:
        try:
            summary.append(process_file(
                f, rules, out_dir, args.all,
                include_title=include_title,
                include_footer=include_footer,
                visual_merge=visual_merge,
                sheet_layout=sheet_layout,
            ))
        except Exception:
            logger.exception(f"처리 실패: {f}")
            failed.append(f.name)

    logger.info("=" * 70)
    logger.info("[처리 결과 요약]")
    for s in summary:
        logger.info(f"  {s['file']}: 표 {s['total']}개 중 {s['extracted']}개 추출 → {s['output']}")
    if failed:
        logger.error(f"  실패한 파일: {', '.join(failed)}")

    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
