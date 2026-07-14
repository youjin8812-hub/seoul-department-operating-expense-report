"""
전체 파이프라인 실행 (마스터 스크립트)

사용법:
    python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html docx hwpx
    python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html docx
    python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html

인자 없이 실행하면 사용법을 출력하고 종료한다.

- html : output/dashboard.html
- docx : output/seoul_expense_executive_report.docx
- hwpx : output/seoul_expense_executive_report.hwpx (Windows + 한컴오피스 한글 감지 시에만 생성, DOCX 필요)

hwpx를 요청했지만 docx가 formats에 없으면, hwpx 생성을 위해 docx 단계를 자동으로 함께 실행한다.
"""
import argparse
import sys

USAGE = """\
사용법: python scripts/run_pipeline.py --input <CSV경로> --formats <html> <docx> <hwpx> [조합]

예시:
  python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html docx hwpx
  python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html docx
  python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html

옵션:
  --input     원본 CSV 경로 (필수)
  --formats   생성할 산출물 형식 (필수, 1개 이상): html docx hwpx

산출물:
  html -> output/dashboard.html
  docx -> output/seoul_expense_executive_report.docx
  hwpx -> output/seoul_expense_executive_report.hwpx (Windows + 한컴오피스 한글 필요, 없으면 자동 건너뜀)

참고: hwpx만 지정해도 내부적으로 docx가 먼저 생성됩니다(HWPX는 DOCX 변환 결과이기 때문).
"""


def main():
    if len(sys.argv) == 1:
        print(USAGE)
        sys.exit(0)

    parser = argparse.ArgumentParser(
        description="서울시 본청 부서운영업무추진비 분석 파이프라인",
        add_help=True,
    )
    parser.add_argument("--input", required=True, help="원본 CSV 경로")
    parser.add_argument("--formats", required=True, nargs="+", choices=["html", "docx", "hwpx"],
                         help="생성할 산출물 형식 (html docx hwpx 중 1개 이상)")
    args = parser.parse_args()

    formats = set(args.formats)
    need_docx = "docx" in formats or "hwpx" in formats
    if "hwpx" in formats and "docx" not in formats:
        print("[안내] hwpx 생성을 위해 docx 단계를 함께 실행합니다(HWPX는 DOCX 변환 결과).")

    results = {}

    # 1. 집계·차트·JSON (모든 형식의 공통 선행 단계)
    print("\n========== 1단계: 집계·차트·분석 JSON ==========")
    import generate_charts
    generate_charts.run(args.input)
    results["charts"] = True

    # 2. HTML 대시보드
    if "html" in formats:
        print("\n========== 2단계: 인터랙티브 대시보드(HTML) ==========")
        import generate_dashboard
        generate_dashboard.run(args.input)
        results["html"] = True

    # 3. DOCX 보고서
    if need_docx:
        print("\n========== 3단계: DOCX 상세 보고서 ==========")
        import generate_docx_report
        generate_docx_report.run()
        results["docx"] = True

    # 4. HWPX 변환 (선택적, 조건 불충족 시 자동 건너뜀)
    hwpx_ok = None
    if "hwpx" in formats:
        print("\n========== 4단계: HWPX 변환 (Windows + 한컴오피스 전용) ==========")
        import convert_docx_to_hwpx
        hwpx_ok, hwpx_reason = convert_docx_to_hwpx.run()
        results["hwpx"] = hwpx_ok
        print(f"[HWPX] {'성공' if hwpx_ok else '건너뜀/실패'}: {hwpx_reason}")
        if not hwpx_ok:
            print("[안내] HWPX가 생성되지 않았지만 HTML·DOCX 산출물은 그대로 유지됩니다.")

    # 5. 검증
    print("\n========== 5단계: 검증 ==========")
    if need_docx:
        import validate_data_and_privacy
        validate_data_and_privacy.run(args.input)
        import validate_docx_layout
        validate_docx_layout.run()
    if hwpx_ok:
        import validate_hwpx_output
        validate_hwpx_output.run()

    print("\n========== 파이프라인 완료 ==========")
    print("요청 형식:", sorted(formats), "| 결과:", results)


if __name__ == "__main__":
    main()
