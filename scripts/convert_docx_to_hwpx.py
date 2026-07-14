"""
DOCX -> HWPX 변환 (한컴오피스 한글 COM 자동화)

원칙:
- 포맷 식별자를 임의로 가정하지 않는다. Open()은 빈 포맷 문자열(확장자 자동인식)을 사용하고,
  SaveAs()는 대상 포맷의 실제 이름인 "HWPX"를 사용한다(둘 다 한컴 공식 API 문서에 명시된 표준 경로).
- 각 COM 호출은 개별적으로 try/except로 감싸 정확한 오류 메시지를 로그에 남긴다.
- 결과 파일이 실제로 존재하고, 0바이트보다 크며, 유효한 ZIP(HWPX 내부 구조) 컨테이너인 경우에만 성공으로 판단한다.
  COM 호출이 예외 없이 끝났다는 사실만으로 성공을 단정하지 않는다.
"""
import os
import platform
import sys
import time
import zipfile
import traceback

DOCX_PATH = os.path.abspath("output/seoul_expense_executive_report.docx")
HWPX_PATH = os.path.abspath("output/seoul_expense_executive_report.hwpx")
LOG_PATH = "output/hwpx_conversion_log.txt"


def log(msg, logf):
    print(msg)
    logf.write(msg + "\n")
    logf.flush()


def detect_hancom_installation(logf):
    if platform.system() != "Windows":
        log(f"[건너뜀] Windows가 아닌 환경({platform.system()}) - HWPX 변환은 Windows + 한컴오피스 전용 기능", logf)
        return False
    try:
        import win32com.client
    except ImportError as e:
        log(f"[실패] pywin32(win32com) 임포트 실패: {e}", logf)
        return False
    try:
        import winreg
        winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\HWPFrame.HwpObject\CLSID")
        log("[확인] Windows 환경 + 레지스트리에 HWPFrame.HwpObject CLSID 등록되어 있음", logf)
        return True
    except FileNotFoundError:
        log("[실패] HWPFrame.HwpObject 레지스트리 항목을 찾을 수 없음 - 한컴오피스 미설치 가능성", logf)
        return False
    except Exception as e:
        log(f"[실패] 레지스트리 확인 중 오류: {e}", logf)
        return False


def verify_reopen(hwpx_path, logf):
    """실제로 한컴 한글에서 재오픈되는 경우에만 성공으로 인정한다."""
    import win32com.client
    hwp = None
    try:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        except Exception:
            pass
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        opened = hwp.Open(hwpx_path, "", "")
        log(f"[재오픈검증] Open() 반환값: {opened}", logf)
        if opened:
            try:
                pc = hwp.PageCount
                log(f"[재오픈검증] PageCount = {pc}", logf)
            except Exception:
                pass
        try:
            hwp.Clear(1)
            hwp.Quit()
        except Exception:
            pass
        return bool(opened)
    except Exception as e:
        log(f"[재오픈검증실패] {e}", logf)
        try:
            if hwp:
                hwp.Quit()
        except Exception:
            pass
        return False


def validate_hwpx_zip_structure(path, logf):
    try:
        if not zipfile.is_zipfile(path):
            log(f"[검증실패] {path}는 유효한 ZIP 컨테이너가 아님", logf)
            return False
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            log(f"[검증] HWPX 내부 파일 수: {len(names)}", logf)
            has_content = any("content" in n.lower() or "section" in n.lower() or n.lower() == "mimetype" for n in names)
            has_manifest = any("manifest" in n.lower() or "container.xml" in n.lower() for n in names)
            log(f"[검증] 콘텐츠 관련 파일 존재: {has_content} | 매니페스트 관련 파일 존재: {has_manifest}", logf)
            sample = names[:10]
            log(f"[검증] 내부 파일 샘플: {sample}", logf)
            return has_content or has_manifest
    except Exception as e:
        log(f"[검증실패] ZIP 구조 검사 중 오류: {e}", logf)
        return False


def convert_docx_to_hwpx(docx_path, hwpx_path, logf):
    import win32com.client
    hwp = None
    try:
        log("[진행] HWPFrame.HwpObject COM 객체 생성 시도", logf)
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        log("[성공] COM 객체 생성 완료", logf)
    except Exception as e:
        log(f"[실패] COM 객체 생성 실패: {e}", logf)
        log(traceback.format_exc(), logf)
        return False

    try:
        log("[진행] 보안 모듈 등록 시도 (자동화 보안경고 방지)", logf)
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        log("[성공] 보안 모듈 등록 완료", logf)
    except Exception as e:
        log(f"[경고] 보안 모듈 등록 실패(보안 승인창이 뜰 수 있음): {e}", logf)

    try:
        try:
            hwp.XHwpWindows.Item(0).Visible = False
            log("[정보] 창 숨김 모드 설정", logf)
        except Exception as e:
            log(f"[정보] 창 숨김 설정 불가(무시하고 진행): {e}", logf)

        log(f"[진행] DOCX 열기 시도 (확장자 자동인식, format=''): {docx_path}", logf)
        opened = hwp.Open(docx_path, "", "")
        log(f"[결과] Open() 반환값: {opened}", logf)
        if not opened:
            log("[실패] Open()이 False/0을 반환함 - 문서가 정상적으로 열리지 않았을 가능성", logf)
            return False
    except Exception as e:
        log(f"[실패] DOCX 열기 중 예외 발생: {e}", logf)
        log(traceback.format_exc(), logf)
        return False

    try:
        docx_page_count = hwp.PageCount
        log(f"[확인] DOCX를 한글 엔진으로 열었을 때 PageCount = {docx_page_count}", logf)
    except Exception as e:
        log(f"[정보] PageCount 속성 조회 실패(무시하고 진행): {e}", logf)

    try:
        log(f"[진행] HWPX로 저장 시도 (format='HWPX'): {hwpx_path}", logf)
        saved = hwp.SaveAs(hwpx_path, "HWPX", "")
        log(f"[결과] SaveAs() 반환값: {saved}", logf)
    except Exception as e:
        log(f"[실패] HWPX 저장 중 예외 발생: {e}", logf)
        log(traceback.format_exc(), logf)
        try:
            hwp.Clear(1)
            hwp.Quit()
        except Exception:
            pass
        return False

    try:
        hwp.Clear(1)
        hwp.Quit()
        log("[정보] 한글 프로세스 정상 종료", logf)
    except Exception as e:
        log(f"[경고] 한글 종료 중 오류(프로세스가 남아있을 수 있음): {e}", logf)

    return True


def run():
    """반환값: (성공여부: bool, 사유: str). HWPX 실패·건너뜀 시에도 DOCX/HTML은 손대지 않는다."""
    os.makedirs("output", exist_ok=True)
    with open(LOG_PATH, "w", encoding="utf-8") as logf:
        log("=== DOCX -> HWPX 변환 시작 ===", logf)
        log(f"입력: {DOCX_PATH}", logf)
        log(f"출력: {HWPX_PATH}", logf)

        if not os.path.exists(DOCX_PATH):
            reason = f"입력 DOCX 파일이 존재하지 않음: {DOCX_PATH}"
            log(f"[실패] {reason}", logf)
            return False, reason

        if not detect_hancom_installation(logf):
            reason = "Windows + 한컴오피스 한글이 감지되지 않아 HWPX 변환을 건너뜀 (HTML·DOCX는 정상 유지)"
            log(f"[건너뜀] {reason}", logf)
            return False, reason

        ok = convert_docx_to_hwpx(DOCX_PATH, HWPX_PATH, logf)
        if not ok:
            reason = "변환 과정에서 오류 발생 - HWPX 미생성으로 처리"
            log(f"[최종실패] {reason}", logf)
            return False, reason

        time.sleep(1)
        if not os.path.exists(HWPX_PATH):
            reason = f"SaveAs가 오류 없이 반환됐으나 실제 파일이 생성되지 않음: {HWPX_PATH}"
            log(f"[최종실패] {reason}", logf)
            return False, reason

        size = os.path.getsize(HWPX_PATH)
        log(f"[확인] 생성된 파일 크기: {size} bytes", logf)
        if size == 0:
            reason = "파일 크기 0바이트 - 정상 생성 아님"
            log(f"[최종실패] {reason}", logf)
            return False, reason

        if not validate_hwpx_zip_structure(HWPX_PATH, logf):
            reason = "HWPX 내부 구조가 예상과 다름 - 손상되었거나 정상 저장되지 않았을 가능성"
            log(f"[최종실패] {reason}", logf)
            return False, reason

        if not verify_reopen(HWPX_PATH, logf):
            reason = "생성된 HWPX가 한컴 한글에서 재오픈되지 않음 - 성공으로 처리하지 않음"
            log(f"[최종실패] {reason}", logf)
            return False, reason

        log("[최종성공] DOCX -> HWPX 변환, 구조 검증, 재오픈 검증 완료", logf)
        return True, "성공"


def main():
    ok, reason = run()
    print(f"[convert_docx_to_hwpx] {'성공' if ok else '실패/건너뜀'}: {reason}")
    sys.exit(0 if ok else 2)


if __name__ == "__main__":
    main()
