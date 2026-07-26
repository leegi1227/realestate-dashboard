"""
국토교통부 건축물대장정보(BuildingLedger) 조회 - 간편 스크립트
================================================================

사용법
------
1) 아래 "설정" 부분만 수정합니다.
   - ADDRESS_KEYWORD 에 "시/구 + 동" 이름을 넣으면 시군구코드/법정동코드를 자동으로 찾아줍니다.
     (코드를 직접 알고 있다면 SIGUNGU_CODE / BDONG_CODE에 바로 입력해도 됩니다.)
   - BUN, JI 에 번지를 입력합니다.
   - LEDGER_TYPE 으로 어떤 건축물대장 정보를 볼지 고릅니다.
2) 실행: python building_example.py
3) 결과가 화면에 출력되고, SAVE_EXCEL=True면 building_result.xlsx로 자동 저장됩니다.

사전 준비
---------
- https://www.data.go.kr 에서 "건축HUB_건축물대장정보 서비스"를 신청하고 서비스키를 발급받습니다.
- pip install PublicDataReader openpyxl --upgrade

참고: 라이브러리 버그 우회
--------------------------
PublicDataReader(pip 배포판)의 BuildingLedger.get_data()는 페이지 수를
"총건수 // numOfRows + 1" 로 계산합니다. 공식 가이드(OpenAPI활용가이드-
건축HUB_건축물대장_1.0)에는 올림(ceiling) 나눗셈으로 계산하라고 되어
있는데, 나머지가 0일 때(예: 총건수 1, numOfRows 1) 라이브러리 공식은
1페이지를 더 요청하게 되고, 그 응답이 비면 이미 모은 데이터를 통째로
버리는 버그가 있습니다. 아래 get_building_ledger()는 이를 올바르게
(len(rows) >= total_count 되면 종료) 우회 처리합니다.
"""

import datetime
import os
import re
import time
import requests
import xmltodict
import pandas as pd
from PublicDataReader import TransactionPrice, BuildingLedger, code_bdong

requests.packages.urllib3.disable_warnings()

# ========================= 설정 (여기만 수정) =========================
# 서비스키는 소스에 직접 적지 말고 환경변수로 주입한다 (공개 저장소 유출 방지).
# Windows: set BUILDING_LEDGER_SERVICE_KEY=발급받은_키
# Mac/Linux: export BUILDING_LEDGER_SERVICE_KEY=발급받은_키
SERVICE_KEY = os.environ.get("BUILDING_LEDGER_SERVICE_KEY", "")

ADDRESS_KEYWORD = "성남시 분당구 백현동"  # 시/군/구 + 동 이름 (모르면 SIGUNGU_CODE/BDONG_CODE를 직접 입력)
SIGUNGU_CODE = None                       # 직접 알고 있으면 "41135" 처럼 입력 (있으면 ADDRESS_KEYWORD보다 우선)
BDONG_CODE = None                         # 직접 알고 있으면 "11000" 처럼 입력

BUN = 237       # 번지 본번
JI = 0          # 번지 부번 (없으면 0)

LEDGER_TYPE = "표제부"
# 선택 가능: 기본개요, 총괄표제부, 표제부, 층별개요, 부속지번,
#           전유공용면적, 오수정화시설, 주택가격, 전유부, 지역지구구역, 소유자

START_DATE = None   # 예: "20200101" (생성일자 검색 시작일, 선택)
END_DATE = None     # 예: "20241231" (생성일자 검색 종료일, 선택)

SAVE_EXCEL = True
EXCEL_PATH = "building_result.xlsx"
VERBOSE = True
# ======================================================================


def resolve_dong_code(keyword: str):
    """'시/구 + 동' 이름으로 시군구코드(5자리)/법정동코드(5자리)를 자동 검색"""
    df = code_bdong()
    # 아직 폐지되지 않은(현재 유효한) 법정동만 대상으로 함
    active = df[df["말소일자"] == ""].copy()
    # 도시의 '동'은 읍면동명 컬럼에, 농어촌의 '리'는 동리명 컬럼에 들어있음 -> 합쳐서 검색
    active["동명"] = active["읍면동명"].where(active["읍면동명"] != "", active["동리명"])

    tokens = keyword.split()
    dong_name = tokens[-1]
    region_hint = " ".join(tokens[:-1])

    matched = active[active["동명"] == dong_name]
    if region_hint:
        region_filtered = matched[
            matched["시군구명"].str.contains(region_hint.split()[-1], na=False)
            | matched["시도명"].str.contains(region_hint.split()[0], na=False)
        ]
        if len(region_filtered) > 0:
            matched = region_filtered

    if len(matched) == 0:
        raise ValueError(f"'{keyword}'에 해당하는 법정동을 찾지 못했습니다. SIGUNGU_CODE/BDONG_CODE를 직접 입력해주세요.")
    if len(matched) > 1:
        print(f"'{keyword}' 검색 결과가 {len(matched)}건입니다. 첫 번째 결과를 사용합니다:")
        print(matched[["시도명", "시군구명", "동명", "법정동코드"]].to_string(index=False))

    row = matched.iloc[0]
    code10 = row["법정동코드"]
    return code10[:5], code10[5:], row


def get_sigungu_list() -> pd.DataFrame:
    """전국 시군구 목록을 중복 없이 반환. 컬럼: 시도명, 시군구명, 시군구코드

    code_bdong()은 법정동(읍면동) 단위라 같은 시군구가 동 개수만큼(수십 번)
    반복되므로, 실거래가 조회용 지역 콤보박스에 쓰려면 시군구코드 기준으로
    중복 제거해야 한다. 시/도 전체를 가리키는 상위 코드(예: 서울 11000,
    끝자리 000이고 시군구명이 비어있는 행)는 실거래가 조회 단위가 아니므로 제외한다.
    """
    df = code_bdong()
    active = df[df["말소일자"] == ""]
    sigungu = active[["시도명", "시군구명", "시군구코드"]].drop_duplicates(subset=["시군구코드"]).copy()
    sigungu = sigungu[~((sigungu["시군구명"] == "") & sigungu["시군구코드"].str.endswith("000"))]
    # 세종특별자치시처럼 구가 없어 시군구명이 비어있는 경우 시도명으로 채움
    sigungu["시군구명"] = sigungu.apply(lambda r: r["시군구명"] or r["시도명"], axis=1)
    return sigungu.sort_values(["시도명", "시군구명"]).reset_index(drop=True)


def get_dong_list(sigungu_code: str) -> list:
    """해당 시군구 안의 동(읍면동/리) 이름 목록을 중복 없이 정렬해서 반환.

    TransactionPrice API는 시군구 단위로만 조회되므로(동을 파라미터로 못
    받음), 동 콤보박스는 조회 후 결과를 '법정동' 컬럼으로 걸러내는 용도로 쓴다.
    """
    df = code_bdong()
    active = df[(df["말소일자"] == "") & (df["시군구코드"] == sigungu_code)].copy()
    active["동명"] = active["읍면동명"].where(active["읍면동명"] != "", active["동리명"])
    names = sorted({n for n in active["동명"] if n})
    return names


def get_bdong_code_map(sigungu_code: str) -> dict:
    """해당 시군구 안의 동명 -> 법정동코드(5자리) 매핑. 시가표준액 조회 시

    실거래가 결과의 '법정동'(동 이름 텍스트)을 건축물대장 API가 요구하는
    법정동코드(숫자)로 되돌리는 데 쓴다.
    """
    df = code_bdong()
    active = df[(df["말소일자"] == "") & (df["시군구코드"] == sigungu_code)].copy()
    active["동명"] = active["읍면동명"].where(active["읍면동명"] != "", active["동리명"])
    active = active[active["동명"] != ""]
    return dict(zip(active["동명"], active["법정동코드"].str[5:]))


def add_standard_price_column(
    df: pd.DataFrame,
    api: BuildingLedger,
    sigungu_code: str,
    progress_callback=None,
    wait_time: float = 0.3,
) -> pd.DataFrame:
    """실거래가 결과에 각 지번의 최신 시가표준액(개별/공동주택가격)을 컬럼으로 추가.

    지번이 마스킹(별표 포함)된 상업업무용 거래, 토지/상업용처럼 건축HUB
    주택가격 API 범위 밖인 유형은 값이 비어있게 된다. 같은 지번이 여러
    행(호실별 거래)에 걸쳐 반복되므로, 지번 단위로만 조회해 API 호출을
    최소화한다.
    """
    if df is None or df.empty or "지번" not in df.columns or "법정동" not in df.columns:
        return df

    out = df.copy()
    bdong_map = get_bdong_code_map(sigungu_code)

    unique_pairs = out[["법정동", "지번"]].drop_duplicates()
    unique_pairs = unique_pairs[~unique_pairs["지번"].astype(str).str.contains(r"\*", na=False)]

    price_lookup = {}
    total = len(unique_pairs)
    for i, (_, row) in enumerate(unique_pairs.iterrows(), start=1):
        dong, jibun = row["법정동"], str(row["지번"])
        bdong_code = bdong_map.get(dong)
        if bdong_code:
            if "-" in jibun:
                bun_s, ji_s = jibun.split("-", 1)
            else:
                bun_s, ji_s = jibun, "0"
            try:
                bun_i, ji_i = int(bun_s), int(ji_s)
            except ValueError:
                bun_i = ji_i = None

            if bun_i is not None:
                try:
                    price_df = get_building_ledger(
                        api, ledger_type="주택가격", sigungu_code=sigungu_code, bdong_code=bdong_code,
                        bun=bun_i, ji=ji_i, max_rows=200, wait_time=wait_time,
                    )
                except Exception:
                    price_df = None

                if price_df is not None and not price_df.empty \
                        and "stdDay" in price_df.columns and "주택가격" in price_df.columns:
                    price_df = price_df.copy()
                    price_df["_연도"] = price_df["stdDay"].apply(extract_year)
                    price_df = price_df.dropna(subset=["_연도"])
                    if not price_df.empty:
                        latest = price_df.loc[price_df["_연도"].idxmax()]
                        price_lookup[(dong, jibun)] = pd.to_numeric(latest["주택가격"], errors="coerce")

        if progress_callback:
            progress_callback(i, total)

    out["시가표준액(억)"] = [
        round(price_lookup[(d, str(j))] / 1e8, 2)
        if (d, str(j)) in price_lookup and pd.notna(price_lookup[(d, str(j))]) else None
        for d, j in zip(out["법정동"], out["지번"])
    ]
    return out


def parse_masked_jibun(jibun: str):
    """마스킹된 지번(예: '8*' -> 2자리 중 '8'만 보임, '1**' -> 3자리 중 '1'만 보임)에서

    (보이는 접두부, 전체 자릿수)를 추출한다. 마스킹이 아니면(별표 없음) None.
    """
    if not jibun:
        return None
    m = re.match(r"^(\d*)(\*+)$", str(jibun).strip())
    if not m:
        return None
    prefix, stars = m.groups()
    return prefix, len(prefix) + len(stars)


def reverse_match_transactions(tx_df: pd.DataFrame, title_df: pd.DataFrame) -> pd.DataFrame:
    """마스킹된 지번 거래를 같은 동의 표제부 전체와 대조해 필지를 역매칭한다.

    국토부는 상업업무용 실거래가의 지번을 '소격동 8*'처럼 마스킹해 공개한다.
    표제부의 연면적·대지면적·건축년도 3값이 거래행의 값과 정확히 일치하는
    필지는 사실상 그 건물 하나뿐이라는 점(deal-locator MCP와 동일한 원리)을
    이용해 역으로 특정한다. 결과에 '역매칭주소'/'매칭신뢰도' 컬럼을 추가하며,
    마스킹되지 않은 일반 거래는 건드리지 않는다.

    신뢰도 등급
    -----------
    확정매칭 : 연면적·건축년도·대지면적 3값이 모두 정밀 일치하는 필지가 단 하나
    추정매칭 : 연면적·건축년도가 일치하는 필지가 단 하나(대지면적은 근사치)
    인접후보 : 조건을 만족하는 필지가 여러 개 — 확인 없이 인용 금지
    미확정   : 조건을 만족하는 필지를 찾지 못함
    """
    if tx_df is None or tx_df.empty or "지번" not in tx_df.columns:
        return tx_df
    if title_df is None or title_df.empty:
        return tx_df

    out = tx_df.copy()
    out["역매칭주소"] = None
    out["매칭신뢰도"] = None

    td = title_df.copy()
    td["_연면적"] = pd.to_numeric(td.get("연면적"), errors="coerce")
    td["_대지면적"] = pd.to_numeric(td.get("대지면적"), errors="coerce")
    td["_건축년도"] = td["사용승인일"].apply(extract_year) if "사용승인일" in td.columns else None
    if "번" in td.columns:
        bun_digits = td["번"].astype(str).str.lstrip("0")
        td["_번"] = bun_digits.where(bun_digits != "", "0")
    else:
        td["_번"] = None

    for idx, row in out.iterrows():
        parsed = parse_masked_jibun(row.get("지번"))
        if not parsed:
            continue
        prefix, length = parsed

        area = pd.to_numeric(row.get("건물면적"), errors="coerce")
        land = pd.to_numeric(row.get("대지면적"), errors="coerce")
        try:
            year = int(row.get("건축년도"))
        except (TypeError, ValueError):
            year = None
        if pd.isna(area) or year is None:
            out.at[idx, "매칭신뢰도"] = "미확정(거래 정보 부족)"
            continue

        candidates = td[(td["_번"].str.len() == length) & (td["_번"].str.startswith(prefix))] \
            if td["_번"].notna().any() else td

        exact = candidates[
            candidates["_연면적"].sub(area).abs().le(0.5)
            & candidates["_건축년도"].eq(year)
            & (candidates["_대지면적"].sub(land).abs().le(2) if pd.notna(land) else True)
        ]
        if len(exact) == 1:
            r = exact.iloc[0]
            out.at[idx, "역매칭주소"] = r.get("도로명대지위치") or r.get("대지위치")
            out.at[idx, "매칭신뢰도"] = "확정매칭"
            continue

        loose = candidates[candidates["_연면적"].sub(area).abs().le(3) & candidates["_건축년도"].eq(year)]
        if len(loose) == 1:
            r = loose.iloc[0]
            out.at[idx, "역매칭주소"] = r.get("도로명대지위치") or r.get("대지위치")
            out.at[idx, "매칭신뢰도"] = "추정매칭"
        elif len(loose) > 1:
            out.at[idx, "역매칭주소"] = "; ".join(str(a) for a in loose["대지위치"].head(3))
            out.at[idx, "매칭신뢰도"] = "인접후보(다수)"
        else:
            out.at[idx, "매칭신뢰도"] = "미확정"

    return out


# 실거래가 결과에서 '지번'(또는 시군구/법정동 텍스트)과 같은 정보를 중복해서
# 담고 있는 원시 코드 컬럼들 — 주소 컬럼을 만든 뒤에는 불필요하므로 제거한다.
_ADDRESS_REDUNDANT_COLUMNS = [
    "법정동시군구코드", "법정동읍면동코드", "법정동지번코드",
    "법정동본번코드", "법정동부번코드",
    "도로명시군구코드", "도로명코드", "도로명일련번호코드", "도로명지상지하코드",
    "도로명건물본번호코드", "도로명건물부번호코드",
]


def add_address_column(df: pd.DataFrame, sido: str = "", sigungu_name: str = "") -> pd.DataFrame:
    """실거래가 조회 결과에 '주소'(시도+시군구+법정동+지번) 컬럼을 맨 앞에 추가하고,

    같은 정보를 다른 형태(법정동시군구코드/본번코드/부번코드 등)로 중복해서
    담고 있던 원시 코드 컬럼은 제거한다. 또한 완전히 동일한 거래가 중복
    보고된 행(같은 주소·지번·계약일·거래금액·층)이 있으면 하나만 남긴다.
    (같은 지번의 아파트 단지에서 호실만 다른 서로 다른 거래는 중복이
    아니므로 그대로 유지된다.)
    """
    if df is None or df.empty:
        return df

    out = df.copy()
    sigungu_series = out["시군구"].astype(str) if "시군구" in out.columns else pd.Series([sigungu_name] * len(out), index=out.index)
    dong_series = out["법정동"].astype(str) if "법정동" in out.columns else pd.Series([""] * len(out), index=out.index)
    jibun_series = out["지번"].astype(str) if "지번" in out.columns else pd.Series([""] * len(out), index=out.index)

    def build_address(sigungu, dong, jibun):
        parts = [p for p in [sido, sigungu, dong] if p and p.lower() != "nan"]
        addr = " ".join(parts)
        # 상업업무용처럼 지번이 마스킹(예: '1**', '3*')된 경우, 별표 섞인
        # 부정확한 번지를 붙이지 않고 법정동까지만 정확하게 표시한다.
        if jibun and jibun.lower() != "nan" and "*" not in jibun:
            addr = f"{addr} {jibun}번지" if addr else f"{jibun}번지"
        return addr

    out.insert(0, "주소", [
        build_address(s, d, j) for s, d, j in zip(sigungu_series, dong_series, jibun_series)
    ])

    drop_cols = [c for c in _ADDRESS_REDUNDANT_COLUMNS if c in out.columns]
    if drop_cols:
        out = out.drop(columns=drop_cols)

    dedup_cols = [c for c in ["주소", "지번", "계약년도", "계약월", "계약일", "거래금액", "층"] if c in out.columns]
    if dedup_cols:
        out = out.drop_duplicates(subset=dedup_cols).reset_index(drop=True)

    return out


_PYEONG_PER_SQM = 3.305785
_UNIT_PRICE_KEYWORDS = ("가격", "단가", "금액", "시세")


def add_pyeong_columns(df: pd.DataFrame) -> pd.DataFrame:
    """'㎡'(제곱미터)가 들어간 컬럼 바로 옆에 평 환산 컬럼을 추가한다.

    일반 면적 컬럼(예: '연면적(㎡)')은 평 = ㎡ ÷ 3.305785로 변환하지만,
    '㎡당가격'처럼 단위가격을 담은 컬럼은 반대 방향(평당가격 = ㎡당가격 × 3.305785)
    으로 계산해야 한다 — 평 하나가 3.305785㎡이므로 더 큰 단위(평)당 가격은
    더 커야 한다. 컬럼명에 가격 관련 키워드가 있으면 단위가격으로 판단해서
    곱셈으로, 아니면 순수 면적으로 보고 나눗셈으로 계산한다.
    """
    if df is None or df.empty:
        return df

    out = df.copy()
    sqm_cols = [c for c in out.columns if "㎡" in str(c) or "제곱미터" in str(c)]
    for col in sqm_cols:
        pyeong_col = str(col).replace("㎡", "평").replace("제곱미터", "평")
        if pyeong_col == col or pyeong_col in out.columns:
            pyeong_col = f"{col}(평)"
        vals = pd.to_numeric(out[col], errors="coerce")
        is_unit_price = any(k in str(col) for k in _UNIT_PRICE_KEYWORDS)
        converted = vals * _PYEONG_PER_SQM if is_unit_price else vals / _PYEONG_PER_SQM
        insert_at = out.columns.get_loc(col) + 1
        out.insert(insert_at, pyeong_col, converted.round(2))

    return out


def get_building_ledger(
    api: BuildingLedger,
    ledger_type: str,
    sigungu_code: str,
    bdong_code: str,
    bun: str = None,
    ji: str = None,
    plat_code: str = None,
    start_date: str = None,
    end_date: str = None,
    translate: bool = True,
    verbose: bool = False,
    wait_time: float = 1,
    max_rows: int = None,
    progress_callback=None,
    **kwargs,
) -> pd.DataFrame:
    """BuildingLedger.get_data()의 페이지네이션 버그를 우회한 버전

    Parameters
    ----------
    max_rows : int, optional
        지정하면 이 건수에 도달하는 즉시 조회를 멈춘다 (동 전체처럼 큰
        조회의 안전장치). None이면 totalCount까지 전부 수집.
    progress_callback : callable, optional
        페이지를 하나 받을 때마다 progress_callback(수집건수, 총건수)로
        호출한다. Streamlit 진행바 등 UI 연동용.
    """
    url = api.meta_dict[ledger_type]["url"]
    columns = api.meta_dict[ledger_type]["columns"]

    if ledger_type != "소유자":
        params = {
            "serviceKey": requests.utils.unquote(api.service_key),
            "numOfRows": 100,  # 공식 가이드 기준 1회 요청 최대 100건
            "sigunguCd": sigungu_code,
            "bjdongCd": bdong_code,
        }
        if bun:
            params["bun"] = str(bun).zfill(4)
        if ji:
            params["ji"] = str(ji).zfill(4)
        if plat_code:
            params["platGbCd"] = plat_code
        if start_date:
            params["startDate"] = str(start_date)
        if end_date:
            params["endDate"] = str(end_date)
    else:
        params = {
            "serviceKey": requests.utils.unquote(api.service_key),
            "numOfRows": 100,
            "sigungu_cd": sigungu_code,
            "bjdong_cd": bdong_code,
        }
        if bun:
            params["bun"] = str(bun).zfill(4)
        if ji:
            params["ji"] = str(ji).zfill(4)
        if plat_code:
            params["plat_gb_cd"] = plat_code
    params.update(kwargs)

    rows = []
    page = 1
    total_count = None
    while True:
        params["pageNo"] = page
        res = requests.get(url, params=params, verify=False, timeout=15)
        try:
            res_json = xmltodict.parse(res.text)
        except Exception:
            raise Exception(
                f"API가 XML이 아닌 응답을 반환했습니다 (HTTP {res.status_code}). "
                f"응답 본문 일부: {res.text[:300]!r}"
            )
        if not res_json.get("response"):
            raise Exception(f"API 요청이 실패했습니다: {res_json}")
        if res_json["response"]["header"]["resultCode"] != "00":
            raise Exception(res_json["response"]["header"]["resultMsg"])

        body = res_json["response"]["body"]
        total_count = int(body["totalCount"])
        items = body.get("items")
        if not items:
            break

        data = items["item"]
        data = data if isinstance(data, list) else [data]
        rows.extend(data)

        if verbose:
            print(f"page {page} 요청 - 누적 {len(rows)} / {total_count}건")
        if progress_callback:
            progress_callback(len(rows), total_count)

        if max_rows and len(rows) >= max_rows:
            rows = rows[:max_rows]
            break
        if len(rows) >= total_count:
            break
        page += 1
        time.sleep(wait_time)

    if rows:
        # 선언된 columns를 앞세우되, API가 실제로 내려준 미선언 필드
        # (예: 주택가격의 stdDay)도 뒤에 붙여서 보존한다.
        df = pd.DataFrame(rows)
        ordered_cols = [c for c in columns if c in df.columns] + \
                       [c for c in df.columns if c not in columns]
        df = df[ordered_cols]
    else:
        df = pd.DataFrame(columns=columns)

    if translate:
        df = api.translate_columns(df)
    return df


ALL_LEDGER_TYPES = [
    "기본개요", "총괄표제부", "표제부", "층별개요", "부속지번",
    "전유공용면적", "오수정화시설", "주택가격", "지역지구구역", "전유부", "소유자",
]
# 사용자가 편하게 부르는 이름 -> meta_dict의 실제 키
LEDGER_TYPE_ALIASES = {"가격": "주택가격"}

# data.go.kr 고유번호 15021136(건축물소유정보 서비스)가 폐지되어 더 이상 조회 불가.
# (2026-07 확인: 신청 페이지 자체가 삭제됨. 개인정보 보안 사유로 추정)
DISCONTINUED_LEDGER_TYPES = {"소유자": "건축물 소유자 정보 오픈API(고유번호 15021136)는 서비스가 종료되었습니다."}


def get_full_building_report(
    api: BuildingLedger,
    sigungu_code: str,
    bdong_code: str,
    bun: str = None,
    ji: str = None,
    plat_code: str = None,
    ledger_types=None,
    verbose: bool = False,
    wait_time: float = 1,
) -> dict:
    """특정 번지의 건축물대장 전체 종류(기본개요~소유자)를 한 번에 조회.

    반환값은 {ledger_type: DataFrame} 딕셔너리. 개별 항목 조회가 실패해도
    (예: 소유자 정보는 별도 서비스 승인이 필요해서 403이 날 수 있음)
    전체가 중단되지 않고, 실패한 항목만 오류 메시지가 담긴 DataFrame으로 표시됩니다.
    """
    ledger_types = ledger_types or ALL_LEDGER_TYPES
    report = {}
    for ledger_type in ledger_types:
        ledger_type = LEDGER_TYPE_ALIASES.get(ledger_type, ledger_type)
        if verbose:
            print(f"=== {ledger_type} 조회 중 ===")
        if ledger_type in DISCONTINUED_LEDGER_TYPES:
            report[ledger_type] = pd.DataFrame({"오류": [DISCONTINUED_LEDGER_TYPES[ledger_type]]})
            continue
        try:
            df = get_building_ledger(
                api,
                ledger_type=ledger_type,
                sigungu_code=sigungu_code,
                bdong_code=bdong_code,
                bun=bun,
                ji=ji,
                plat_code=plat_code,
                verbose=verbose,
                wait_time=wait_time,
            )
        except Exception as e:
            df = pd.DataFrame({"오류": [str(e)]})
        report[ledger_type] = df
    return report


def save_report_to_excel(report: dict, path: str):
    """get_full_building_report() 결과를 대장 종류별 시트로 나눠 엑셀 하나로 저장"""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for ledger_type, df in report.items():
            df.to_excel(writer, sheet_name=ledger_type[:31], index=False)


def extract_year(value) -> int:
    """'19810204', '1947' 등에서 4자리 연도를 안전하게 추출. 실패 시 None."""
    if value is None:
        return None
    text = str(value).strip()
    if len(text) < 4 or not text[:4].isdigit():
        return None
    year = int(text[:4])
    if year < 1800 or year > 2100:
        return None
    return year


def _to_numeric(series):
    if series is None:
        return pd.Series(dtype=float)
    return pd.to_numeric(series, errors="coerce")


def analyze_district_stats(title_df: pd.DataFrame) -> dict:
    """표제부 전체(동 단위) DataFrame에서 총괄/주용도별/연대별/노후도/규모벤치마크 통계 산출"""
    empty = {"총괄": {}, "주용도별": pd.DataFrame(), "연대별": pd.DataFrame(),
             "노후도분포": pd.DataFrame(), "규모벤치마크": pd.DataFrame()}
    if title_df is None or title_df.empty:
        return empty

    df = title_df.copy()
    total_count = len(df)
    approval_year = df["사용승인일"].apply(extract_year) if "사용승인일" in df.columns else pd.Series([None] * total_count)
    this_year = datetime.date.today().year
    age_years = approval_year.apply(lambda y: (this_year - y) if y else None)

    total_area = _to_numeric(df.get("연면적")).sum()
    avg_floors = _to_numeric(df.get("지상층수")).mean()
    avg_age = age_years.dropna().mean()

    summary = {
        "총동수": total_count,
        "총연면적(㎡)": round(float(total_area), 2),
        "평균층수": round(float(avg_floors), 1) if pd.notna(avg_floors) else None,
        "평균경과연수": round(float(avg_age), 1) if pd.notna(avg_age) else None,
    }

    by_purpose = pd.DataFrame()
    if "주용도코드명" in df.columns:
        tmp = df.assign(_area=_to_numeric(df.get("연면적")))
        grouped = tmp.groupby("주용도코드명").agg(건수=("주용도코드명", "size"), 연면적=("_area", "sum"))
        grouped["비율(%)"] = (grouped["건수"] / total_count * 100).round(1)
        by_purpose = grouped.sort_values("건수", ascending=False).head(10).reset_index()
        by_purpose = by_purpose.rename(columns={"주용도코드명": "주용도"})

    by_decade = pd.DataFrame()
    decades = (approval_year.dropna() // 10 * 10).astype(int)
    if len(decades) > 0:
        counts = decades.value_counts().sort_index()
        by_decade = pd.DataFrame({
            "연대": [f"{d}년대" for d in counts.index],
            "건수": counts.values,
            "비율(%)": (counts.values / total_count * 100).round(1),
        })

    age_bins = [0, 10, 20, 30, 40, 9999]
    age_labels = ["10년 미만", "10~20년", "20~30년", "30~40년", "40년 이상"]
    bucket = pd.cut(age_years, bins=age_bins, labels=age_labels, right=False)
    bucket_counts = bucket.value_counts().reindex(age_labels).fillna(0).astype(int)
    old_age_dist = pd.DataFrame({
        "구간": age_labels,
        "건수": bucket_counts.values,
        "비율(%)": (bucket_counts.values / total_count * 100).round(1),
    })

    benchmark = pd.DataFrame()
    if not by_purpose.empty:
        rows = []
        for purpose in by_purpose["주용도"].head(5):
            sub = df[df["주용도코드명"] == purpose]
            rows.append({
                "주용도": purpose,
                "층수(중앙값)": _to_numeric(sub.get("지상층수")).median(),
                "용적률(중앙값,%)": _to_numeric(sub.get("용적률")).median(),
                "높이(중앙값,m)": _to_numeric(sub.get("높이")).median(),
            })
        benchmark = pd.DataFrame(rows)

    return {
        "총괄": summary,
        "주용도별": by_purpose,
        "연대별": by_decade,
        "노후도분포": old_age_dist,
        "규모벤치마크": benchmark,
    }


def analyze_old_buildings(title_df: pd.DataFrame, min_age_years: int = 30) -> pd.DataFrame:
    """표제부 전체(동 단위)에서 경과연수 min_age_years 이상인 건물만 내림차순 정렬"""
    if title_df is None or title_df.empty:
        return pd.DataFrame()

    df = title_df.copy()
    approval_year = df["사용승인일"].apply(extract_year) if "사용승인일" in df.columns else None
    this_year = datetime.date.today().year
    df["경과연수"] = approval_year.apply(lambda y: (this_year - y) if y else None) if approval_year is not None else None

    df = df[df["경과연수"].notna() & (df["경과연수"] >= min_age_years)]
    df = df.sort_values("경과연수", ascending=False)

    keep_cols = ["건물명", "도로명대지위치", "대지위치", "주용도코드명", "사용승인일",
                 "경과연수", "연면적", "지상층수", "세대수"]
    keep_cols = [c for c in keep_cols if c in df.columns]
    return df[keep_cols].reset_index(drop=True)


_NON_REINFORCED_STRUCTURES = ["벽돌구조", "블록구조", "석구조", "목구조"]


def analyze_seismic_risk(title_df: pd.DataFrame, top_n: int = 30) -> dict:
    """내진설계 적용 여부 필드 + 사용승인연도 기준 의무화 연혁으로 미적용 추정 분류

    내진설계 의무화 연혁: 1988년 도입(6층/10만㎡ 이상) -> 이후 단계적 강화 ->
    2017.12 전면 의무화(2층 또는 200㎡ 이상). `내진 설계 적용 여부` 필드가
    명시돼 있으면 그 값을 우선하고, 없으면 준공연도로 추정한다.
    """
    empty = {"분류별집계": pd.DataFrame(), "취약우선목록": pd.DataFrame()}
    if title_df is None or title_df.empty:
        return empty

    df = title_df.copy()
    df["_사용승인연도"] = df["사용승인일"].apply(extract_year) if "사용승인일" in df.columns else None

    def classify(row):
        flag = str(row.get("내진 설계 적용 여부", "") or "").strip()
        year = row["_사용승인연도"]
        if flag == "1":
            return "내진설계 적용(명시)"
        if flag == "0":
            return "내진설계 미적용(명시)"
        if year is None:
            return "준공년도 미상"
        if year < 1988:
            return "1988년 이전 준공 (의무화 전, 미적용 추정)"
        if year < 2018:
            return "1988~2017 준공 (기준 완화기간, 확인 필요)"
        return "2018년 이후 준공 (전면의무화, 적용 추정)"

    df["내진분류"] = df.apply(classify, axis=1)

    total = len(df)
    dist = df["내진분류"].value_counts().reset_index()
    dist.columns = ["분류", "건수"]
    dist["비율(%)"] = (dist["건수"] / total * 100).round(1)

    priority_order = {
        "내진설계 미적용(명시)": 0,
        "1988년 이전 준공 (의무화 전, 미적용 추정)": 1,
        "1988~2017 준공 (기준 완화기간, 확인 필요)": 2,
        "준공년도 미상": 3,
        "내진설계 적용(명시)": 4,
        "2018년 이후 준공 (전면의무화, 적용 추정)": 4,
    }
    df["_우선순위"] = df["내진분류"].map(priority_order).fillna(9)
    if "구조코드명" in df.columns:
        df["_비보강구조"] = df["구조코드명"].isin(_NON_REINFORCED_STRUCTURES)
    else:
        df["_비보강구조"] = False
    df = df.sort_values(["_우선순위", "_비보강구조", "_사용승인연도"], ascending=[True, False, True])

    keep_cols = ["건물명", "도로명대지위치", "대지위치", "주용도코드명", "구조코드명",
                 "사용승인일", "지상층수", "내진분류"]
    keep_cols = [c for c in keep_cols if c in df.columns]
    priority_list = df.head(top_n)[keep_cols].reset_index(drop=True)

    return {"분류별집계": dist, "취약우선목록": priority_list}


PRICE_HISTORY_WARNING = (
    "공시가격은 시세가 아닙니다(통상 시세의 60~70% 수준, 연도별 현실화율 정책 "
    "변동 포함). 증감률·CAGR을 시세·실거래가 상승률로 해석하지 마세요."
)


def analyze_price_history(price_df: pd.DataFrame, top_units: int = 10) -> dict:
    """주택가격(공시가격) DataFrame을 관리건축물대장PK(호)별로 묶어 연도별 추이·증감률·CAGR 계산"""
    if (price_df is None or price_df.empty
            or "관리건축물대장PK" not in price_df.columns or "stdDay" not in price_df.columns):
        return {"단위목록": [], "경고": PRICE_HISTORY_WARNING}

    df = price_df.copy()
    df["_연도"] = df["stdDay"].apply(extract_year)
    df["_가격"] = pd.to_numeric(df.get("주택가격"), errors="coerce")
    df = df.dropna(subset=["_연도", "_가격"])

    units = []
    for pk, group in df.groupby("관리건축물대장PK"):
        g = group.sort_values("_연도")
        if g.empty:
            continue
        first_year, first_price = int(g.iloc[0]["_연도"]), float(g.iloc[0]["_가격"])
        last_year, last_price = int(g.iloc[-1]["_연도"]), float(g.iloc[-1]["_가격"])
        years_span = last_year - first_year
        total_change_pct = ((last_price / first_price) - 1) * 100 if first_price else None
        cagr_pct = (((last_price / first_price) ** (1 / years_span)) - 1) * 100 \
            if first_price and years_span > 0 else None

        timeline = g[["_연도", "_가격"]].rename(columns={"_연도": "연도", "_가격": "주택가격"})
        units.append({
            "관리건축물대장PK": pk,
            "최초연도": first_year, "최초가격": first_price,
            "최신연도": last_year, "최신가격": last_price,
            "총증감률(%)": round(total_change_pct, 1) if total_change_pct is not None else None,
            "연평균상승률CAGR(%)": round(cagr_pct, 1) if cagr_pct is not None else None,
            "추이": timeline,
        })

    units.sort(key=lambda u: u["최신가격"], reverse=True)
    return {"단위목록": units[:top_units], "경고": PRICE_HISTORY_WARNING}


def summarize_zoning(zoning_df: pd.DataFrame) -> list:
    """지역지구구역 표에서 중복을 제거한 용도지역/지구/구역 목록을 문자열 리스트로 반환.

    같은 필지라도 세대(관리건축물대장PK)마다 행이 반복 등록돼 있어(예: 아파트 한
    단지가 59행) 그대로 보여주면 의미 없이 길어지므로, 구분·코드명 조합으로
    중복 제거한다.
    """
    if zoning_df is None or zoning_df.empty:
        return []
    cols = [c for c in ["지역지구구역구분코드명", "지역지구구역코드명"] if c in zoning_df.columns]
    if not cols:
        return []
    unique = zoning_df[cols].drop_duplicates()
    results = []
    for _, row in unique.iterrows():
        name = str(row.get("지역지구구역코드명", "") or "").strip()
        if not name or name.lower() == "none":
            continue
        results.append(name)
    return results


# 브이월드 데이터 API의 용도지역/지구/구역 관련 레이어 (레이어코드, 표시명)
_VWORLD_ZONING_LAYERS = [
    ("LT_C_UQ111", "도시지역"),
    ("LT_C_UQ112", "관리지역"),
    ("LT_C_UQ113", "농림지역"),
    ("LT_C_UQ114", "자연환경보전지역"),
    ("LT_C_UQ121", "경관지구"),
    ("LT_C_UQ123", "고도지구"),
    ("LT_C_UQ124", "방화지구"),
    ("LT_C_UQ125", "방재지구"),
    ("LT_C_UQ126", "보호지구"),
    ("LT_C_UQ127", "시설보호지구"),
    ("LT_C_UQ128", "취락지구"),
    ("LT_C_UQ129", "개발진흥지구"),
    ("LT_C_UQ130", "특정용도제한지구"),
    ("LT_C_UQ141", "국토계획구역"),
    ("LT_C_UQ162", "도시자연공원구역"),
    ("LT_C_UD801", "개발제한구역(그린벨트)"),
]


def geocode_address_vworld(vworld_key: str, address: str):
    """도로명주소를 브이월드 주소검색 API로 좌표(경도, 위도)로 변환. 실패 시 None."""
    url = "https://api.vworld.kr/req/address"
    params = {
        "service": "address", "request": "getcoord", "version": "2.0",
        "crs": "epsg:4326", "key": vworld_key, "type": "ROAD", "address": address,
    }
    try:
        res = requests.get(url, params=params, verify=False, timeout=10)
        data = res.json()
        point = data["response"]["result"]["point"]
        return float(point["x"]), float(point["y"])
    except Exception:
        return None


def get_vworld_zoning_detail(vworld_key: str, lon: float, lat: float, retries: int = 2) -> dict:
    """좌표 지점이 속한 용도지역/지구/구역을 브이월드 데이터 API로 레이어별 조회.

    반환값: {레이어명: [그 지점이 속한 구역명, ...]}. 레이어에 해당 사항이
    없거나 (아직 승인이 덜 퍼졌거나 일시적 오류로) API 호출이 실패하면 그
    레이어만 결과에서 조용히 생략된다 — 나머지 레이어 조회는 계속 진행된다.
    """
    url = "https://api.vworld.kr/req/data"
    result = {}
    for layer_code, label in _VWORLD_ZONING_LAYERS:
        params = {
            "key": vworld_key, "service": "data", "request": "GetFeature",
            "data": layer_code, "geomFilter": f"POINT({lon} {lat})",
            "size": 10, "page": 1,
        }
        data = None
        for _ in range(retries):
            try:
                res = requests.get(url, params=params, verify=False, timeout=10)
                candidate = res.json()
                if candidate.get("response", {}).get("status") == "OK":
                    data = candidate
                    break
            except Exception:
                continue
        if data is None:
            continue

        features = data["response"].get("result", {}).get("featureCollection", {}).get("features", [])
        names = []
        for feat in features:
            props = feat.get("properties", {})
            name = next((v for k, v in props.items() if "nm" in k.lower() and v), None)
            if not name:
                name = next((v for v in props.values() if isinstance(v, str) and v.strip()), None)
            if name:
                names.append(name)
        if names:
            result[label] = sorted(set(names))
    return result


def combine_zoning_sources(master: dict) -> list:
    """건축HUB 표제부의 지역지구구역 + 브이월드 레이어 조회 결과를 하나의 중복 없는 목록으로 합침"""
    combined = list(master.get("지역지구") or [])
    vworld_detail = master.get("브이월드용도지역") or {}
    for values in vworld_detail.values():
        for v in values:
            if v not in combined:
                combined.append(v)
    return combined


# 표제부 주용도코드명 -> TransactionPrice property_type 추정 매핑
_PROPERTY_TYPE_KEYWORDS = [
    (["아파트"], "아파트"),
    (["다세대주택", "연립주택"], "연립다세대"),
    (["단독주택", "다가구주택"], "단독다가구"),
    (["오피스텔"], "오피스텔"),
]
_COMMERCIAL_KEYWORDS = [
    "근린생활시설", "업무시설", "판매시설", "숙박시설", "위락시설",
    "운동시설", "교육연구시설", "의료시설", "제1종근린생활시설", "제2종근린생활시설",
]


def guess_property_type(purpose_name: str):
    """표제부 주용도코드명에서 TransactionPrice property_type을 추정. 모르면 None."""
    if not purpose_name:
        return None
    for keywords, ptype in _PROPERTY_TYPE_KEYWORDS:
        if any(k in purpose_name for k in keywords):
            return ptype
    if any(k in purpose_name for k in _COMMERCIAL_KEYWORDS):
        return "상업업무용"
    if "공장" in purpose_name or "창고" in purpose_name:
        return "공장창고등"
    return None


def match_transactions_to_parcel(
    tp_api: TransactionPrice,
    property_type: str,
    trade_type: str,
    sigungu_code: str,
    bun,
    ji,
    months_lookback: int = 12,
    sido: str = "",
    sigungu_name: str = "",
) -> dict:
    """추정된 property_type으로 최근 N개월 실거래가를 받아 이 번지(bun/ji)와 정확히 일치하는 건만 추림.

    반환되는 DataFrame에는 시도+시군구+법정동(동)+지번을 합친 '주소' 컬럼이
    포함된다 — 지번만으로는 어느 동인지 알 수 없어 부정확하므로 항상 동
    단위까지 표시한다.

    상업업무용은 지번이 마스킹(예: '소격동 8*')돼 공개되므로 완전한 특정이
    불가능하다 — 이 경우 deal-locator MCP의 역매칭 기능을 안내한다.
    """
    if property_type is None:
        return {"status": "no_type", "df": None, "note": "주용도로 실거래가 유형을 추정할 수 없습니다."}

    end = datetime.date.today()
    end_ym = end.strftime("%Y%m")
    start_ym = (end - datetime.timedelta(days=30 * months_lookback)).strftime("%Y%m")

    try:
        df = tp_api.get_data(
            property_type=property_type, trade_type=trade_type,
            sigungu_code=sigungu_code, start_year_month=start_ym, end_year_month=end_ym,
        )
    except Exception as e:
        return {"status": "error", "df": None, "note": str(e)}

    if df is None or df.empty or "지번" not in df.columns:
        return {
            "status": "no_data", "df": pd.DataFrame(),
            "note": f"최근 {months_lookback}개월 내 {property_type} {trade_type} 거래가 없습니다.",
        }

    df = add_address_column(df, sido=sido, sigungu_name=sigungu_name)

    bun_i = int(bun) if bun else None
    ji_i = int(ji) if ji else 0
    target_jibun = f"{bun_i}-{ji_i}" if ji_i else str(bun_i)

    matched = df[df["지번"].astype(str).str.strip() == target_jibun] if bun_i else pd.DataFrame()

    masked = property_type == "상업업무용"
    if matched.empty:
        note = f"최근 {months_lookback}개월 내 이 번지({target_jibun})의 {property_type} {trade_type} 거래를 찾지 못했습니다."
        if masked:
            note += " 상업업무용은 지번이 마스킹되어 공개되어 완전한 특정이 어렵습니다 — deal-locator MCP의 역매칭 기능을 이용해보세요."
        return {"status": "no_match", "df": df, "note": note}

    return {"status": "matched", "df": matched, "note": f"{len(matched)}건 매칭됨"}


def build_master_report(
    api: BuildingLedger,
    tp_api: TransactionPrice,
    sigungu_code: str,
    bdong_code: str,
    bun,
    ji,
    months_lookback: int = 12,
    district_title_df: pd.DataFrame = None,
    sido: str = "",
    sigungu_name: str = "",
    vworld_key: str = None,
) -> dict:
    """지번 하나에 대해 단일조회·실거래가·노후도·내진·공시가격(+선택적 동단위 비교)을 한 번에 모은 종합 리포트

    vworld_key를 넘기면 브이월드 데이터 API로 용도지역/지구/구역 상세(도시지역,
    경관지구, 고도지구, 개발제한구역 등 15종 레이어)를 추가로 조회한다.
    """
    result = {}

    title_df = get_building_ledger(api, ledger_type="표제부", sigungu_code=sigungu_code,
                                    bdong_code=bdong_code, bun=bun, ji=ji)
    result["표제부"] = title_df

    if title_df is not None and not title_df.empty:
        purpose = title_df.iloc[0].get("주용도코드명")
        ptype = guess_property_type(purpose)
        result["추정부동산유형"] = ptype
        result["실거래가"] = match_transactions_to_parcel(
            tp_api, ptype, "매매", sigungu_code, bun, ji, months_lookback=months_lookback,
            sido=sido, sigungu_name=sigungu_name,
        )
        result["노후도"] = analyze_old_buildings(title_df, min_age_years=0)
        result["내진분석"] = analyze_seismic_risk(title_df, top_n=1)
    else:
        result["추정부동산유형"] = None
        result["실거래가"] = {"status": "no_building", "df": None, "note": "표제부 조회 결과가 없습니다."}
        result["노후도"] = pd.DataFrame()
        result["내진분석"] = {"분류별집계": pd.DataFrame(), "취약우선목록": pd.DataFrame()}

    price_df = get_building_ledger(
        api, ledger_type="주택가격", sigungu_code=sigungu_code, bdong_code=bdong_code,
        bun=bun, ji=ji, max_rows=5000, wait_time=0.3,
    )
    result["공시가격"] = analyze_price_history(price_df, top_units=10)

    zoning_df = get_building_ledger(
        api, ledger_type="지역지구구역", sigungu_code=sigungu_code, bdong_code=bdong_code,
        bun=bun, ji=ji,
    )
    result["지역지구"] = summarize_zoning(zoning_df)

    result["브이월드용도지역"] = {}
    result["좌표"] = None
    if vworld_key and title_df is not None and not title_df.empty:
        addr_for_geocode = title_df.iloc[0].get("도로명대지위치") or title_df.iloc[0].get("대지위치")
        if addr_for_geocode:
            coord = geocode_address_vworld(vworld_key, addr_for_geocode)
            if coord:
                result["좌표"] = coord
                result["브이월드용도지역"] = get_vworld_zoning_detail(vworld_key, coord[0], coord[1])

    if district_title_df is not None and not district_title_df.empty:
        result["동단위통계"] = analyze_district_stats(district_title_df)
    else:
        result["동단위통계"] = None

    return result


# (일반체, 굵은체) 후보들. Windows(로컬)와 Linux(Streamlit Cloud 등 클라우드 배포)를
# 둘 다 지원하기 위해 여러 경로를 순서대로 시도한다. 클라우드에 배포할 때는
# packages.txt에 `fonts-nanum`을 추가하면 두 번째 경로가 설치되어 잡힌다.
_KOREAN_FONT_CANDIDATE_SETS = [
    (r"C:\Windows\Fonts\malgun.ttf", r"C:\Windows\Fonts\malgunbd.ttf"),
    ("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
    ("/usr/share/fonts/truetype/nanum/NanumGothic-Bold.ttf", "/usr/share/fonts/truetype/nanum/NanumGothic-Bold.ttf"),
]


def _find_korean_font():
    """사용 가능한 (일반체, 굵은체) 한글 폰트 경로 쌍을 찾는다. 없으면 (None, None)."""
    for regular, bold in _KOREAN_FONT_CANDIDATE_SETS:
        if os.path.exists(regular) and os.path.exists(bold):
            return regular, bold
    return None, None


def split_common_and_varying(df: pd.DataFrame):
    """여러 행을 가진 표(예: 층별개요의 각 층, 전유공용면적의 각 호실)에서

    건물 전체에 공통인(모든 행에서 값이 똑같은) 컬럼과, 층/호실마다
    달라지는 핵심 컬럼을 분리한다. 공통 컬럼은 한 번만 보여주고,
    달라지는 컬럼만 표로 모아서 층/호실 전체를 한 번에 볼 수 있게 하기 위함.

    반환값: (공통정보 dict, 달라지는 컬럼만 남긴 DataFrame)
    """
    if len(df) <= 1:
        return {}, df

    common = {}
    varying_cols = []
    for col in df.columns:
        values = df[col].astype(str).fillna("")
        uniq = values.unique()
        if len(uniq) <= 1:
            val = uniq[0] if len(uniq) else ""
            if val not in ("", "None", "nan"):
                common[col] = val
        else:
            varying_cols.append(col)

    varying_df = df[varying_cols] if varying_cols else df
    return common, varying_df


# 핵심정보 카드에 쓸 (원본컬럼, 표시라벨) 순서. 표제부/총괄표제부에서 값이 있는 것만 뽑아 쓴다.
_CORE_FIELD_LABELS = [
    ("건물명", "건물명"), ("주용도코드명", "주용도"), ("구조코드명", "구조"), ("지붕코드명", "지붕"),
    ("지상층수", "지상층수"), ("지하층수", "지하층수"), ("세대수", "세대수"), ("가구수", "가구수"),
    ("호수", "호수"), ("대지면적", "대지면적(㎡)"), ("건축면적", "건축면적(㎡)"), ("연면적", "연면적(㎡)"),
    ("건폐율", "건폐율(%)"), ("용적률", "용적률(%)"), ("사용승인일", "사용승인일"), ("허가일", "허가일"),
    ("착공일", "착공일"), ("내진 설계 적용 여부", "내진설계"),
]

# 여러 행(층/호실 등)을 갖는 대장 종류에서, 한눈에 보기용으로 남길 핵심 컬럼
_MULTIROW_KEEP_COLS = {
    "층별개요": ["층번호명", "면적", "주용도코드명", "기타용도"],
    "전유공용면적": ["호명칭", "층번호명", "전유공용구분코드명", "면적"],
    "부속지번": ["지번주소", "부속대장구분코드명"],
    "지역지구구역": ["지역지구구역구분코드명", "지역지구구역코드명"],
}

# 한 줄 요약으로만 보여줄 소규모 항목: {대장종류: [원본컬럼, ...]}
_ONE_LINE_FIELDS = {
    "오수정화시설": ["형식코드명", "용량(루베)", "용량(인용)"],
    "주택가격": ["주택가격", "생성일자"],
}


def _clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in ("none", "nan"):
        return ""
    return text


def generate_pdf_report(report: dict, address_label: str = "", title: str = "건축물대장 종합 리포트") -> bytes:
    """get_full_building_report() 결과를 1페이지 분량의 깔끔한 요약 PDF로 변환.

    11개 대장 종류를 전부 원본 컬럼 그대로 나열하면 항목이 너무 많고
    난잡해지므로, 실제로 의미 있는 핵심 필드만 골라 하나의 요약 카드로
    합치고, 값이 없는 항목/대장 종류는 아예 표시하지 않는다.
    """
    from fpdf import FPDF
    from fpdf.fonts import FontFace

    label_style = FontFace(emphasis="BOLD")

    regular, bold = _find_korean_font()
    if not regular:
        raise FileNotFoundError(
            "한글 PDF 출력을 위한 폰트를 찾을 수 없습니다. Windows에서는 맑은 고딕이 "
            "기본 설치돼 있어야 하고, 리눅스 서버(예: Streamlit Cloud)에서는 "
            "packages.txt에 `fonts-nanum`을 추가해야 합니다."
        )

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.add_font("Malgun", "", regular)
    pdf.add_font("Malgun", "B", bold)

    pdf.set_font("Malgun", "B", 18)
    pdf.cell(0, 11, title, new_x="LMARGIN", new_y="NEXT")

    # 대지위치는 표제부 -> 총괄표제부 -> 기본개요 순으로 찾아서 부제목에 표시
    core_src = None
    for key in ("표제부", "총괄표제부", "기본개요"):
        candidate = report.get(key)
        if candidate is not None and not candidate.empty and "오류" not in candidate.columns:
            core_src = candidate.iloc[0]
            break

    subtitle = address_label or (core_src.get("도로명대지위치") or core_src.get("대지위치", "") if core_src is not None else "")
    pdf.set_font("Malgun", "", 11)
    pdf.set_text_color(90, 90, 90)
    pdf.cell(0, 7, _clean(subtitle), new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(2)

    # --- 핵심 정보 카드 (2열 그리드) ---
    if core_src is not None:
        pairs = [(label, _clean(core_src.get(col, ""))) for col, label in _CORE_FIELD_LABELS]
        pairs = [(label, val) for label, val in pairs if val]
        if pairs:
            pdf.set_font("Malgun", "B", 12)
            pdf.set_fill_color(235, 235, 235)
            pdf.cell(0, 9, "핵심 정보", new_x="LMARGIN", new_y="NEXT", fill=True)
            pdf.set_font("Malgun", "", 10)
            rows = [pairs[i:i + 2] for i in range(0, len(pairs), 2)]
            with pdf.table(text_align="LEFT", line_height=6, borders_layout="MINIMAL",
                            col_widths=(28, 62, 28, 62)) as table:
                for row_pairs in rows:
                    row_obj = table.row()
                    for label, val in row_pairs:
                        row_obj.cell(label, style=label_style)
                        row_obj.cell(val)
                    if len(row_pairs) == 1:
                        row_obj.cell("")
                        row_obj.cell("")
            pdf.ln(3)

    # --- 층별개요 / 전유공용면적 등 다건 항목: 핵심 컬럼만 표로 ---
    for ledger_type, keep_cols in _MULTIROW_KEEP_COLS.items():
        df = report.get(ledger_type)
        if df is None or df.empty or "오류" in df.columns:
            continue
        cols = [c for c in keep_cols if c in df.columns]
        if not cols:
            continue
        pdf.set_font("Malgun", "B", 12)
        pdf.set_fill_color(235, 235, 235)
        pdf.cell(0, 9, ledger_type, new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.set_font("Malgun", "", 9)
        table_data = [cols] + [[_clean(v) for v in row] for row in df[cols].itertuples(index=False)]
        with pdf.table(text_align="LEFT", line_height=5, borders_layout="MINIMAL") as table:
            for data_row in table_data:
                row_obj = table.row()
                for datum in data_row:
                    row_obj.cell(datum)
        pdf.ln(3)

    # --- 오수정화시설 / 주택가격 등 소규모 항목: 한 줄 요약 ---
    one_liners = []
    for ledger_type, cols in _ONE_LINE_FIELDS.items():
        df = report.get(ledger_type)
        if df is None or df.empty or "오류" in df.columns:
            continue
        row = df.iloc[0]
        parts = [f"{c}: {_clean(row[c])}" for c in cols if c in df.columns and _clean(row[c])]
        if parts:
            one_liners.append(f"{ledger_type} — " + " · ".join(parts))
    if one_liners:
        pdf.set_font("Malgun", "B", 12)
        pdf.set_fill_color(235, 235, 235)
        pdf.cell(0, 9, "기타 정보", new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.set_font("Malgun", "", 10)
        for line in one_liners:
            pdf.multi_cell(0, 6, line, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

    # --- 조회되지 않은/미제공 항목은 한 줄로만 안내 ---
    unavailable = []
    for ledger_type, df in report.items():
        if ledger_type == "소유자":
            unavailable.append("소유자(공공API 미제공)")
        elif df is None or df.empty or "오류" in df.columns:
            if ledger_type not in ("표제부", "총괄표제부", "기본개요"):
                unavailable.append(ledger_type)
    if unavailable:
        pdf.set_font("Malgun", "", 9)
        pdf.set_text_color(130, 130, 130)
        pdf.multi_cell(0, 5, "해당 없음 / 조회 불가: " + ", ".join(unavailable), new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)

    return bytes(pdf.output())


_TRANSACTION_COLS_FOR_PDF = ["주소", "계약년도", "계약월", "계약일", "거래금액", "층", "전용면적", "건축년도"]


def build_executive_summary(master: dict) -> str:
    """통합 리포트의 핵심 내용을 자연어 한 단락으로 요약 (보고서 서두 '핵심 요약'용)"""
    title_df = master.get("표제부")
    core = title_df.iloc[0] if title_df is not None and not title_df.empty else None
    if core is None:
        return "표제부 조회 결과가 없어 요약을 생성할 수 없습니다."

    purpose = core.get("주용도코드명") or "용도 미상"
    structure = core.get("구조코드명") or "구조 미상"
    approval_year = extract_year(core.get("사용승인일"))

    old_df = master.get("노후도")
    age = None
    if old_df is not None and not old_df.empty and "경과연수" in old_df.columns:
        age = int(old_df.iloc[0]["경과연수"])

    seismic = master.get("내진분석") or {}
    seismic_list = seismic.get("취약우선목록")
    seismic_label = seismic_list.iloc[0]["내진분류"] if seismic_list is not None and not seismic_list.empty else "내진 정보 미상"

    zoning = combine_zoning_sources(master)
    zoning_clause = f" 용도지역/지구는 {', '.join(zoning)}입니다." if zoning else ""

    sentences = [
        f"이 건물은 {purpose}(구조: {structure})로, "
        + (f"{approval_year}년 준공되어 약 {age}년이 경과했습니다." if approval_year and age is not None
           else "준공연도가 명확하지 않습니다.")
        + f" 내진 분류는 '{seismic_label}'입니다."
        + zoning_clause
    ]

    tx = master.get("실거래가") or {}
    tx_df = tx.get("df")
    if tx.get("status") == "matched":
        sentences.append(f"최근 조회 기간 내 이 번지에서 {len(tx_df)}건의 실거래가 확인됐습니다.")
    elif tx_df is not None and not tx_df.empty and "거래금액" in tx_df.columns:
        avg_price = pd.to_numeric(tx_df["거래금액"], errors="coerce").mean()
        if pd.notna(avg_price):
            sentences.append(
                f"이 번지의 직접 거래는 확인되지 않았으나, 같은 유형·지역의 최근 평균 거래가는 "
                f"{avg_price / 1e4:.1f}억원 수준입니다({len(tx_df)}건 기준)."
            )
    else:
        sentences.append("실거래가 매칭 데이터가 충분하지 않습니다.")

    ph = master.get("공시가격") or {}
    units = ph.get("단위목록") or []
    if units:
        u = units[0]
        sentences.append(
            f"공시가격은 {u['최초연도']}년 {u['최초가격'] / 1e8:.2f}억원에서 {u['최신연도']}년 "
            f"{u['최신가격'] / 1e8:.2f}억원으로 연평균 {u['연평균상승률CAGR(%)']}% 상승했습니다."
        )

    district = master.get("동단위통계")
    if district and district.get("총괄") and age is not None and district["총괄"].get("평균경과연수"):
        diff = age - district["총괄"]["평균경과연수"]
        sentences.append(
            f"동일 법정동 평균 경과연수({district['총괄']['평균경과연수']}년) 대비 이 건물은 "
            f"{abs(diff):.1f}년 {'더 오래되었습니다' if diff > 0 else '더 신축입니다'}."
        )

    return " ".join(s for s in sentences if s)


def generate_master_pdf_report(master: dict, address_label: str = "", title: str = "부동산 통합 리포트") -> bytes:
    """build_master_report() 결과(단일조회+실거래가+노후도+내진+공시가격+동단위통계)를 하나의 PDF로."""
    from fpdf import FPDF
    from fpdf.fonts import FontFace

    label_style = FontFace(emphasis="BOLD")

    regular, bold = _find_korean_font()
    if not regular:
        raise FileNotFoundError(
            "한글 PDF 출력을 위한 폰트를 찾을 수 없습니다. 리눅스 서버(예: Streamlit Cloud)에서는 "
            "packages.txt에 `fonts-nanum`을 추가해야 합니다."
        )

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.add_font("Malgun", "", regular)
    pdf.add_font("Malgun", "B", bold)

    pdf.set_font("Malgun", "B", 18)
    pdf.cell(0, 11, title, new_x="LMARGIN", new_y="NEXT")

    title_df = master.get("표제부")
    core_src = title_df.iloc[0] if title_df is not None and not title_df.empty else None
    subtitle = address_label or (core_src.get("도로명대지위치") or core_src.get("대지위치", "") if core_src is not None else "")
    pdf.set_font("Malgun", "", 11)
    pdf.set_text_color(90, 90, 90)
    pdf.cell(0, 7, _clean(subtitle), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Malgun", "", 9)
    pdf.cell(0, 6, f"생성일시: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(2)

    def section_title(text):
        pdf.set_font("Malgun", "B", 12)
        pdf.set_fill_color(235, 235, 235)
        pdf.cell(0, 9, text, new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.set_font("Malgun", "", 10)

    # --- 0. 핵심 요약 ---
    section_title("핵심 요약")
    pdf.multi_cell(0, 6, build_executive_summary(master), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # --- 1. 단일조회: 핵심 정보 카드 ---
    if core_src is not None:
        pairs = [(label, _clean(core_src.get(col, ""))) for col, label in _CORE_FIELD_LABELS]
        pairs = [(label, val) for label, val in pairs if val]
        zoning = combine_zoning_sources(master)
        if zoning:
            pairs.insert(0, ("용도지역/지구", ", ".join(zoning)))
        if pairs:
            section_title("① 단일 조회 — 핵심 정보")
            rows = [pairs[i:i + 2] for i in range(0, len(pairs), 2)]
            with pdf.table(text_align="LEFT", line_height=6, borders_layout="MINIMAL",
                            col_widths=(28, 62, 28, 62)) as table:
                for row_pairs in rows:
                    row_obj = table.row()
                    for label, val in row_pairs:
                        row_obj.cell(label, style=label_style)
                        row_obj.cell(val)
                    if len(row_pairs) == 1:
                        row_obj.cell("")
                        row_obj.cell("")
            pdf.ln(3)

    # --- 2. 실거래가 매칭 ---
    section_title("② 실거래가")
    tx = master.get("실거래가") or {}
    tx_df = tx.get("df")
    ptype = master.get("추정부동산유형")
    pdf.multi_cell(0, 6, f"추정 부동산 유형: {ptype or '판별 불가'}", new_x="LMARGIN", new_y="NEXT")
    pdf.multi_cell(0, 6, tx.get("note", ""), new_x="LMARGIN", new_y="NEXT")

    if tx_df is not None and not tx_df.empty and "거래금액" in tx_df.columns:
        if tx.get("status") != "matched":
            prices = pd.to_numeric(tx_df["거래금액"], errors="coerce").dropna()
            if not prices.empty:
                pdf.set_font("Malgun", "B", 10)
                pdf.multi_cell(
                    0, 6,
                    f"[참고] 같은 유형·지역 최근 거래 {len(prices)}건 — "
                    f"평균 {prices.mean() / 1e4:.1f}억 · 중앙값 {prices.median() / 1e4:.1f}억 "
                    f"(최저 {prices.min() / 1e4:.1f}억 ~ 최고 {prices.max() / 1e4:.1f}억)",
                    new_x="LMARGIN", new_y="NEXT",
                )
                pdf.set_font("Malgun", "", 10)
            pdf.set_font("Malgun", "", 9)
            pdf.multi_cell(0, 5, "아래는 이 번지와 정확히 일치하지 않는, 인근 참고 거래 목록입니다.",
                           new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Malgun", "", 10)

        cols = [c for c in _TRANSACTION_COLS_FOR_PDF if c in tx_df.columns]
        show_df = tx_df if tx.get("status") == "matched" else tx_df.head(10)
        table_data = [cols] + [[_clean(v) for v in row] for row in show_df[cols].itertuples(index=False)]
        pdf.set_font("Malgun", "", 8)
        with pdf.table(text_align="LEFT", line_height=5, borders_layout="MINIMAL") as table:
            for data_row in table_data:
                row_obj = table.row()
                for datum in data_row:
                    row_obj.cell(datum)
        pdf.set_font("Malgun", "", 10)
    pdf.ln(3)

    # --- 3. 노후도 / 내진 ---
    section_title("③ 노후도 · 내진")
    old_df = master.get("노후도")
    if old_df is not None and not old_df.empty and "경과연수" in old_df.columns:
        pdf.multi_cell(0, 6, f"경과연수: {old_df.iloc[0]['경과연수']:.0f}년 "
                              f"(사용승인일 {old_df.iloc[0].get('사용승인일', '-')})",
                       new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.multi_cell(0, 6, "경과연수 산출 불가 (사용승인일 정보 없음)", new_x="LMARGIN", new_y="NEXT")

    seismic = master.get("내진분석") or {}
    seismic_list = seismic.get("취약우선목록")
    if seismic_list is not None and not seismic_list.empty:
        pdf.multi_cell(0, 6, f"내진 분류: {seismic_list.iloc[0].get('내진분류', '-')}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # --- 4. 공시가격(시가표준액) 시계열 ---
    section_title("④ 공시가격(시가표준액) 시계열")
    ph = master.get("공시가격") or {}
    units = ph.get("단위목록") or []
    if not units:
        pdf.multi_cell(0, 6, "이 번지의 공시가격(주택가격) 데이터가 없습니다.", new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_font("Malgun", "", 9)
        pdf.multi_cell(0, 5, ph.get("경고", ""), new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1)

        summary_cols = ["호(관리건축물대장PK)", "최초연도", "최초가격(억)", "최신연도", "최신가격(억)", "총증감(%)", "CAGR(%)"]
        summary_rows = [summary_cols] + [
            [
                str(u["관리건축물대장PK"])[-6:], str(u["최초연도"]), f"{u['최초가격'] / 1e8:.2f}",
                str(u["최신연도"]), f"{u['최신가격'] / 1e8:.2f}",
                str(u["총증감률(%)"]), str(u["연평균상승률CAGR(%)"]),
            ]
            for u in units[:10]
        ]
        pdf.set_font("Malgun", "", 9)
        with pdf.table(text_align="LEFT", line_height=5, borders_layout="MINIMAL") as table:
            for data_row in summary_rows:
                row_obj = table.row()
                for datum in data_row:
                    row_obj.cell(datum)
        pdf.ln(2)

        pdf.set_font("Malgun", "B", 9)
        pdf.multi_cell(0, 5, f"최고가 호(…{str(units[0]['관리건축물대장PK'])[-6:]}) 연도별 추이", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Malgun", "", 9)
        trend = units[0]["추이"]
        trend_rows = [["연도", "주택가격(억)"]] + [
            [str(int(r["연도"])), f"{r['주택가격'] / 1e8:.2f}"] for _, r in trend.iterrows()
        ]
        with pdf.table(text_align="LEFT", line_height=5, borders_layout="MINIMAL", col_widths=(30, 30)) as table:
            for data_row in trend_rows:
                row_obj = table.row()
                for datum in data_row:
                    row_obj.cell(datum)
        pdf.set_font("Malgun", "", 10)
    pdf.ln(3)

    # --- 5. 동단위 통계 비교 ---
    district = master.get("동단위통계")
    if district and district.get("총괄"):
        section_title("⑤ 동단위 통계 비교")
        s = district["총괄"]
        pdf.multi_cell(
            0, 6,
            f"동 전체 총동수 {s['총동수']:,} · 총연면적 {s['총연면적(㎡)']:,.0f}㎡ · "
            f"평균층수 {s['평균층수']} · 평균경과연수 {s['평균경과연수']}년",
            new_x="LMARGIN", new_y="NEXT",
        )
        if old_df is not None and not old_df.empty and "경과연수" in old_df.columns and s.get("평균경과연수"):
            diff = old_df.iloc[0]["경과연수"] - s["평균경과연수"]
            pdf.multi_cell(0, 6, f"이 건물은 동 평균보다 {abs(diff):.1f}년 {'더 오래됨' if diff > 0 else '더 신축'}",
                           new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

        by_purpose = district.get("주용도별")
        if by_purpose is not None and not by_purpose.empty:
            pdf.set_font("Malgun", "B", 9)
            pdf.multi_cell(0, 5, "주용도별 분포 (상위 5)", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Malgun", "", 9)
            cols = [c for c in ["주용도", "건수", "비율(%)"] if c in by_purpose.columns]
            rows = [cols] + [[_clean(v) for v in row] for row in by_purpose[cols].head(5).itertuples(index=False)]
            with pdf.table(text_align="LEFT", line_height=5, borders_layout="MINIMAL") as table:
                for data_row in rows:
                    row_obj = table.row()
                    for datum in data_row:
                        row_obj.cell(datum)
            pdf.ln(2)

        age_dist = district.get("노후도분포")
        if age_dist is not None and not age_dist.empty:
            pdf.set_font("Malgun", "B", 9)
            pdf.multi_cell(0, 5, "노후도 분포", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Malgun", "", 9)
            rows = [["구간", "건수", "비율(%)"]] + [
                [_clean(v) for v in row] for row in age_dist[["구간", "건수", "비율(%)"]].itertuples(index=False)
            ]
            with pdf.table(text_align="LEFT", line_height=5, borders_layout="MINIMAL") as table:
                for data_row in rows:
                    row_obj = table.row()
                    for datum in data_row:
                        row_obj.cell(datum)
        pdf.set_font("Malgun", "", 10)
        pdf.ln(3)

    # --- 하단 고지 ---
    pdf.set_font("Malgun", "", 8)
    pdf.set_text_color(140, 140, 140)
    pdf.multi_cell(
        0, 4.5,
        "본 리포트는 공공데이터포털(data.go.kr) 국토교통부 실거래가 공개시스템·건축HUB 건축물대장 "
        "공식 API 실측값을 자동 종합해 생성되었습니다. 법적 효력이 있는 공부(公簿)가 아니며 참고 "
        "자료입니다. 계약·감정·고객 제시 등 실제 의사결정 전에는 반드시 원문(건축물대장, 실거래가 "
        "공개시스템)을 직접 확인하시기 바랍니다.",
        new_x="LMARGIN", new_y="NEXT",
    )
    pdf.set_text_color(0, 0, 0)

    return bytes(pdf.output())


def main():
    api = BuildingLedger(SERVICE_KEY)

    sigungu_code, bdong_code = SIGUNGU_CODE, BDONG_CODE
    if not (sigungu_code and bdong_code):
        sigungu_code, bdong_code, row = resolve_dong_code(ADDRESS_KEYWORD)
        print(f"주소 인식 결과: {row['시도명']} {row['시군구명']} {row['읍면동명']}{row['동리명']} "
              f"(시군구코드={sigungu_code}, 법정동코드={bdong_code})")

    df = get_building_ledger(
        api,
        ledger_type=LEDGER_TYPE,
        sigungu_code=sigungu_code,
        bdong_code=bdong_code,
        bun=BUN,
        ji=JI,
        start_date=START_DATE,
        end_date=END_DATE,
        verbose=VERBOSE,
    )

    print(df)

    if SAVE_EXCEL:
        df.to_excel(EXCEL_PATH, index=False)
        print(f"엑셀로 저장했습니다: {EXCEL_PATH}")

    return df


if __name__ == "__main__":
    main()

# ---------------------------------------------------------------
# 참고: 특정 건물(단지)의 아파트 실거래가를 함께 보고 싶다면 TransactionPrice를 사용합니다.
# tp_api = TransactionPrice(SERVICE_KEY)
# tp_df = tp_api.get_data(
#     property_type="아파트",
#     trade_type="매매",
#     sigungu_code="41135",
#     year_month="202506",
# )
# print(tp_df.head())
