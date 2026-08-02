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
import re

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


def ocr_region(image: Image.Image, box: tuple[float, float, float, float], psm: int = 4, scale: int = 2,
               pad: bool = True) -> str:
    """이미지의 일부(상대좌표 0~1 기준 x0,y0,x1,y1)만 잘라 OCR한다.

    전체 페이지를 한 번에 OCR하면(ocr_image_text) 표 전체의 격자선·워터마크·
    옆 칸 글자가 서로 간섭해서 작은 글씨가 뭉개지는데, 이 프로젝트가 실제로
    검증한 바로는 표 하나만 딱 잘라서 적당히 확대하고(2배) 순수 그레이스케일로
    (이진화 없이) --psm 4로 돌리면 인식률이 크게 올라간다 — 예를 들어 소유자
    이름이 "윤명분"으로 정확히 나온 반면, 전체 페이지 OCR에서는 "륜령분"으로
    깨졌었다. 이진화(preprocess_for_ocr)는 워터마크 제거에는 유용하지만, 이미
    좁게 잘라낸 영역에서는 오히려 획을 뭉개서 정확도를 떨어뜨렸다.
    """
    w, h = image.size
    x0, y0, x1, y1 = box
    crop = image.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h))).convert("L")
    crop = crop.resize((crop.width * scale, crop.height * scale), Image.LANCZOS)
    if pad:
        padded = Image.new("L", (crop.width + 40, crop.height + 40), 255)
        padded.paste(crop, (20, 20))
        crop = padded
    try:
        return pytesseract.image_to_string(crop, lang="kor+eng", config=f"--psm {psm}")
    except Exception:
        return ""


def detect_page_kind(text: str) -> str:
    """표지 큰 글씨 "(갑)"/"제1쪽"은 소유자현황이 있는 갑(甲)페이지,
    "변동사항"/"(을)"은 변동사항·인허가정보가 있는 을(乙)페이지의 신호다.
    이 프로젝트가 검증한 스캔본에서는 본문 라벨(예: "소유자현황")은 OCR로 뭉개져도
    이 표지성 큰 글자는 비교적 잘 살아남았다."""
    if "(갑)" in text or "건축물대장(갑" in text or "제1쪽" in text:
        return "갑"
    if "변동사항" in text or "(을)" in text or "건축물대장(을" in text or "제2쪽" in text:
        return "을"
    return "unknown"


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
        "page_kind": detect_page_kind(raw_text) if raw_text else "unknown",
    }


def analyze_ledger_images(files: list[tuple[bytes, str]]) -> list[dict]:
    return [analyze_ledger_image(b, name) for b, name in files]


# 소유자현황 표 한 행: <구조/용도/면적> | <성명> <주소(...동/리/가...)> ... <변동일 YYYY.M.D>
# 이 프로젝트가 실제로 검증한 스캔본에서는 구조가 무너진 OCR 텍스트라도 이 순서(이름 ->
# 동/리/가가 들어간 주소 -> 날짜)는 살아남았다 — 표 셀 경계가 아니라 이 순서 자체로 매칭한다.
_OWNER_ROW_RE = re.compile(
    r"([가-힣]{2,4})\s+.{2,30}?(?:동|리|가)\D{0,20}?(\d{4}\s*[.\-]\s*\d{1,2}\s*[.\-]\s*\d{1,2})"
)
_OWNER_NAME_STOPWORDS = {"성명", "명칭", "소유자", "소유지", "주소", "변동일", "변동원인", "구분"}


def extract_owners(text: str) -> list[dict]:
    """소유자현황 표에서 성명·변동일 후보를 정규식으로 뽑는다.

    표 셀 구분이 무너진 OCR 텍스트에서 뽑는 것이라 이름 등이 오탈자로 나올 수 있다
    (예: '윤명분'이 '륜령분'으로) — 참고용이며, 정확한 값은 원본 이미지 대조가 필요하다.
    """
    owners = []
    seen = set()
    for m in _OWNER_ROW_RE.finditer(text):
        name = m.group(1)
        date = re.sub(r"\s+", "", m.group(2))
        if name in _OWNER_NAME_STOPWORDS:
            continue
        key = (name, date)
        if key in seen:
            continue
        seen.add(key)
        owners.append({"성명": name, "변동일": date})
    return owners


_CHANGE_LINE_RE = re.compile(r"\d{4}\s*[.\-]\s*\d{1,2}\s*[.\-]\s*\d{1,2}")


def extract_change_history(text: str) -> list[str]:
    """'변동사항' 절 이후에서 날짜가 포함된 줄만, 정리하지 않고 원문에 가깝게 뽑는다.

    을(乙)페이지 변동사항은 2단으로 나뉜 표라 OCR이 두 컬럼을 한 줄에 섞어버리는 경우가
    많다 — 깔끔한 표로 재구성하려 하지 않고, 날짜가 포함된 원문 줄을 그대로 보여준다.
    """
    idx = text.find("변동사항")
    if idx == -1:
        return []
    section = text[idx:]
    lines = []
    for line in section.splitlines():
        line = re.sub(r"\s+", " ", line).strip()
        if _CHANGE_LINE_RE.search(line) and len(line) > 8:
            lines.append(line)
    return lines[:12]


# 갑(甲)페이지의 소유자현황(성명/주소/소유권지분/변동일) 칸은 건축물현황 표 오른쪽
# 절반, 세로로는 중간쯤에서 시작한다. 을(乙)페이지의 "구분(건축주/설계자/공사감리자/
# 공사시공자)·성명 또는 명칭·면허(등록)번호" 칸은 표 맨 위 왼쪽 절반이다. 이 프로젝트가
# 검증한 표준 양식 스캔본(가로 997x638 안팎) 기준 상대좌표 — 스캔 해상도가 달라도
# 비율은 같은 표준 서식이라 어느 정도는 유지될 것으로 기대하지만, 문서 스캔 방식에
# 따라 어긋날 수 있다(그 경우 아래 정규식이 아무 것도 못 찾고 조용히 빈 리스트를
# 반환한다 — 잘못된 값을 지어내지는 않는다).
_GAP_OWNER_REGION = (0.50, 0.55, 1.0, 0.72)
_EUL_PARTIES_REGION = (0.01, 0.21, 0.44, 0.40)

_NAME_TOKEN_RE = re.compile(r"[가-힣]{2,4}")
_OWNER_NAME_BLOCKLIST = _OWNER_NAME_STOPWORDS | {
    "서울특별시", "용산구", "이태원동", "이래원동", "주소변경", "즈소변경",
}


def extract_owners_region(image: Image.Image) -> list[dict]:
    """소유자현황 칸만 잘라 OCR한 뒤, 날짜 하나마다 그 앞쪽에서 가장 가까운
    이름다운 토큰을 붙인다. 표 셀 경계가 아니라 "날짜 앞에 이름이 온다"는
    문서 관례에 기대는 것이라 완벽하지 않지만(예: 두 번째 소유자는 이름이
    아예 인식 안 될 수 있음), 전체 페이지 OCR보다 정확도가 뚜렷이 높다."""
    text = ocr_region(image, _GAP_OWNER_REGION)
    owners = []
    for m in _CHANGE_LINE_RE.finditer(text):
        date = re.sub(r"\s+", "", m.group(0))
        before = text[max(0, m.start() - 80):m.start()]
        candidates = [
            n for n in _NAME_TOKEN_RE.findall(before)
            if n not in _OWNER_NAME_BLOCKLIST and not any(c in n for c in "동구시군읍면리등")
        ]
        name = candidates[-1] if candidates else None
        owners.append({"성명": name, "변동일": date})
    return owners


# 회사/사무소 형태 관계자(설계자·공사감리자)는 이름 끝의 "사무소/건축사/공사" 같은
# 접미사가 OCR에서도 비교적 잘 살아남는다 — 반면 "함석근" 같은 3글자 개인 이름(건축주·
# 공사시공자)은 표 셀이 무너지면서 사라지기 쉽다(재현 테스트 확인). 이 칸은 소유자
# 칸과 최적 설정이 달랐다 — 실측 결과 스케일 5배 + 패딩 없이 + --psm 11(문자 단위
# 인식)일 때만 "사무소" 마지막 음절까지 정확히 살아남았고(스케일 2배·패딩 포함으로는
# 매번 "사무소"→"사루소/사두소"로 깨졌다), 그 설정에서는 "건축주" 바로 뒤에 "할소근"
# (함석근의 오독)처럼 사람 이름도 간간이 붙어 나와서 시도할 가치가 있었다.
_FIRM_NAME_RE = re.compile(r"[가-힣]{2,10}(?:건축사사무소|사무소|합동사무소|공사(?!감리))")
_BUILDER_NAME_RE = re.compile(r"건축주\s*\|?\s*([가-힣]{2,4})")
_PARTY_LABELS = {"건축주", "설계자", "공사감리자", "공사시공자", "구분", "성명", "명칭", "현장관리인"}


def extract_parties_region(image: Image.Image) -> dict:
    """건축주/설계자/공사감리자/공사시공자 칸에서 회사(사무소)명 + (있으면) 건축주 이름을 뽑는다."""
    text = ocr_region(image, _EUL_PARTIES_REGION, psm=11, scale=5, pad=False)
    seen = set()
    firms = []
    for m in _FIRM_NAME_RE.finditer(text):
        firm = m.group(0)
        if firm not in seen:
            seen.add(firm)
            firms.append(firm)
    builder_m = _BUILDER_NAME_RE.search(text)
    builder = builder_m.group(1) if builder_m else None
    if builder in _PARTY_LABELS:
        builder = None
    return {"firms": firms, "builder": builder}


def extract_ledger_content(docs: list[dict]) -> dict:
    """analyze_ledger_image()로 얻은 문서(들)에서 소유자현황·변동사항·관계 업체를 모은다.

    각 문서를 갑/을 페이지로 구분해, 갑페이지는 표 영역만 다시 잘라 OCR한
    고정밀 결과로 소유자를 뽑고(전체 페이지 정규식 결과는 이름을 못 찾은
    자리를 메우는 보조용으로만 씀), 을페이지는 변동사항 + 관계 업체를 뽑는다.
    """
    owners, changes, firms = [], [], []
    builder = None
    for doc in docs:
        text = doc.get("raw_text") or ""
        if not text or not doc.get("ocr_available"):
            continue
        page_kind = doc.get("page_kind") or detect_page_kind(text)
        image = None
        if doc.get("image_bytes"):
            try:
                image = Image.open(io.BytesIO(doc["image_bytes"]))
            except Exception:
                image = None

        fallback_owners = extract_owners(text)
        if page_kind == "갑" and image is not None:
            region_owners = extract_owners_region(image)
            for i, owner in enumerate(region_owners):
                if owner["성명"] is None and i < len(fallback_owners):
                    owner["성명"] = fallback_owners[i]["성명"] + "(정확도 낮음)"
                owner["성명"] = owner["성명"] or "성명 미상(OCR)"
            owners.extend(region_owners or fallback_owners)
        else:
            # 을/unknown 페이지에는 소유자 행 패턴이 원래 없으므로 이 호출은 보통 빈 리스트만
            # 돌려준다 — 이미지가 없어 정밀 추출을 못 하는 갑페이지에 대한 대비용이다.
            owners.extend(fallback_owners)

        if page_kind == "을" and image is not None:
            parties = extract_parties_region(image)
            firms.extend(parties["firms"])
            builder = builder or parties["builder"]
        changes.extend(extract_change_history(text))

    # 중복 제거 (성명+변동일 기준)
    uniq_owners, seen_owner_keys = [], set()
    for owner in owners:
        key = (owner["성명"], owner["변동일"])
        if key not in seen_owner_keys:
            seen_owner_keys.add(key)
            uniq_owners.append(owner)

    uniq_firms = list(dict.fromkeys(firms))
    return {"owners": uniq_owners, "changes": changes, "firms": uniq_firms, "builder": builder}
