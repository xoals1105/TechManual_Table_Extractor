"""구형 .hwp 파일을 .hwpx 로 변환 (한글 프로그램 COM 자동화 사용).

주의:
  - 한글(Hancom Office)이 설치된 PC에서만 동작한다.
  - 무인 실행을 위해 보안 승인 모듈(FilePathCheckerModule) 등록이 필요할 수 있다.
    (미등록 시 파일을 열 때마다 접근 승인 팝업이 뜸)
"""
from __future__ import annotations

from pathlib import Path


def convert_hwp_to_hwpx(hwp_path: str | Path, out_dir: str | Path | None = None) -> Path:
    """hwp → hwpx 변환 후 생성된 hwpx 경로 반환."""
    try:
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError(
            "pywin32가 설치되어 있지 않습니다. .hwp 변환에는 pywin32와 "
            "한글(Hancom Office) 설치가 필요합니다."
        ) from e

    hwp_path = Path(hwp_path).resolve()
    out_dir = Path(out_dir).resolve() if out_dir else hwp_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / (hwp_path.stem + ".hwpx")

    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        # 보안 승인 모듈 등록 (레지스트리에 등록되어 있어야 실제 효과 있음)
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass  # 미등록 환경에서는 팝업이 뜰 수 있으나 변환 자체는 가능

        hwp.XHwpWindows.Item(0).Visible = False
        if not hwp.Open(str(hwp_path), "HWP", "forceopen:true"):
            raise RuntimeError(f"한글 파일 열기 실패: {hwp_path}")
        if not hwp.SaveAs(str(out_path), "HWPX", ""):
            raise RuntimeError(f"HWPX 저장 실패: {out_path}")
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass

    return out_path
