"""
건축물대장 열람본(갑/을) 이미지를 업로드했을 때, 공공데이터 API로는 제공되지 않는
위반건축물 여부 · 소유자현황 · 변동사항 등을 참고할 수 있도록 돕는 보조 모듈.

건축HUB Open API(표제부/총괄표제부 등)는 위반건축물 여부·소유자·변동이력 필드
자체를 제공하지 않는다(소유자 오픈API는 2026년 기준 data.go.kr에서 서비스 종료).
이 정보는 정부24/세움터가 발급하는 "열람용" 이미지·PDF에만 존재하므로, 업로드된
이미지를 OCR로 훑어 핵심 신호(위반건축물 여부, 텍스트 검색용 원문)만 뽑아내고,
정확한 확인은 항상 업로드된 원본 이미지를 함께 보여줘서 대조하도록 한다.

Tesseract 한글 OCR은 이런 표 형식 스캔 문서에서 숫자/키워드는 비교적 잘 잡아내지만
사람 이름·주소 등 작은 글씨는 오탈자가 잦다. 그래서 구조화된 표로 완전히 믿지 말고,
"참고용" 텍스트로만 취급한다 — 원본 이미지가 항상 최종 진실이다.
"""
from __future__ import annotations

import io

import numpy as np
from PIL import Image

try:
    import pytesseract
    _TESSERACT_IMPORT_ERROR = None
except ImportError as e:
    pytesseract = None
    _TESSERACT_IMPORT_ERROR = e


VIOLATION_KEYWORDS = ["위반건축물", "위반 건축물"]


def is_ocr_available() -> tuple[bool, str]:
    """(사용 가능 여부, 안내 메시지) 반환. 서버(Streamlit Cloud)에 tesseract 바이너리
    (packages.txt의 tesseract-ocr/tesseract-ocr-kor)가 없으면 여기서 걸러진다."""
    if pytesseract is None:
        return False, f"pytesseract 모듈을 불러오지 못했습니다: {_TESSERACT_IMPORT_ERROR}"
    try:
        pytesseract.get_tesseract_version()
    except Exception as e:
        return False, (
            "Tesseract OCR 실행 파일을 찾을 수 없습니다. "
            "requirements.txt(pytesseract)와 packages.txt(tesseract-ocr, tesseract-ocr-kor)가 "
            f"배포 환경에 반영됐는지 확인하세요. (원본 오류: {e})"
        )
    return True, ""


def preprocess_for_ocr(image: Image.Image) -> Image.Image:
    """그레이스케일 -> 3배 확대 -> 이진화. 열람용 문서 특유의 옅은 회색 워터마크
    ("열람용")를 지우고 작은 표 글씨의 인식률을 끌어올리기 위한 전처리."""
    gray = image.convert("L")
    w, h = gray.size
    upscaled = gray.resize((w * 3, h * 3), Image.LANCZOS)
    arr = np.array(upscaled)
    binary = np.where(arr < 150, 255, 0).astype("uint8")
    return Image.fromarray(255 - binary)


def ocr_image_text(file_bytes: bytes) -> str:
    """이미지 바이트 -> OCR 텍스트. 실패해도 예외를 던지지 않고 빈 문자열을 반환한다
    (OCR은 보조 기능이라, 실패해도 원본 이미지 표시 자체는 방해하면 안 된다)."""
    ok, _ = is_ocr_available()
    if not ok:
        return ""
    try:
        image = Image.open(io.BytesIO(file_bytes))
        processed = preprocess_for_ocr(image)
        return pytesseract.image_to_string(processed, lang="kor+eng", config="--psm 6")
    except Exception:
        return ""


def detect_violation_building(text: str) -> bool:
    return any(kw in text for kw in VIOLATION_KEYWORDS)


def analyze_ledger_image(file_bytes: bytes, filename: str = "") -> dict:
    """업로드된 건축물대장 이미지 1장을 분석. 항상 원본 바이트를 함께 들고 있어서
    화면에 이미지 자체를 보여주고 OCR 결과와 나란히 대조할 수 있게 한다."""
    ocr_ok, ocr_message = is_ocr_available()
    raw_text = ocr_image_text(file_bytes) if ocr_ok else ""
    return {
        "filename": filename,
        "image_bytes": file_bytes,
        "ocr_available": ocr_ok,
        "ocr_message": ocr_message,
        "raw_text": raw_text,
        "is_violation": detect_violation_building(raw_text),
    }


def analyze_ledger_images(files: list[tuple[bytes, str]]) -> list[dict]:
    return [analyze_ledger_image(b, name) for b, name in files]
