"""
국토교통부 건축물대장정보 조회 대시보드 (Streamlit)
====================================================

실행 방법
---------
streamlit run dashboard.py

브라우저에서 자동으로 열리며, 안 열리면 터미널에 표시되는
http://localhost:8501 주소를 직접 열면 됩니다.
"""

import datetime
import io
import math
import os

import pandas as pd
import streamlit as st
from PublicDataReader import BuildingLedger, TransactionPrice

from building_example import (
    ALL_LEDGER_TYPES,
    add_address_column,
    add_coordinates_column,
    add_pyeong_columns,
    add_standard_price_column,
    analyze_district_stats,
    analyze_old_buildings,
    analyze_seoul_trade_area_detail,
    analyze_transaction_stats,
    generate_pdf_report,
    geocode_address_kakao,
    get_bdong_code_map,
    get_building_ledger,
    get_dong_list,
    get_full_building_report,
    get_sigungu_list,
    REB_COMMERCIAL_VACANCY_STATBL_IDS,
    reb_current_quarter_id,
    reb_quarter_ids_desc,
    get_reb_vacancy_snapshot,
    get_reb_vacancy_trend,
    get_nearby_stores,
    load_sangkwon_upjong_codes,
    get_seoul_trade_area_locations,
    get_seoul_trade_area_quarter_dataset,
    find_nearest_seoul_trade_area,
    seoul_current_quarter_id,
    SEOUL_TRDAR_SALES_SERVICE,
    SEOUL_TRDAR_STORE_SERVICE,
    SEOUL_TRDAR_FLPOP_SERVICE,
    SEOUL_TRDAR_WRC_POPLTN_SERVICE,
    reverse_match_transactions,
    split_common_and_varying,
)
from report_generator import fetch_report_data, generate_ledger_pptx, generate_pptx

try:
    from ledger_ocr import analyze_ledger_image
except ImportError:
    analyze_ledger_image = None


# Leaflet + OpenStreetMap을 삽입하는 Custom Components v2 컴포넌트 (기존 pydeck
# 렌더링을 대체). 처음에는 카카오맵 JS SDK로 구현했으나, 배포 환경에서 SDK 요청이
# 브라우저의 ORB(Opaque Response Blocking)에 막혀 키/도메인 설정이 맞아도 계속
# 실패했다(카카오 서버가 401을 JSON으로 응답 → 브라우저가 script 태그로 받은
# JSON 응답을 차단). API 키나 도메인 등록이 아예 필요 없는 Leaflet+OSM으로 교체해
# 이 문제 자체를 없앴다. (컴포넌트 프로토콜 자체는 v1의 iframe+postMessage 방식이
# 이 Streamlit 버전에서 먹통이라 이미 v2로 전환해 둔 상태 — 그 부분은 그대로 재사용)
_address_map_dir = os.path.join(os.path.dirname(__file__), "address_map_component")
with open(os.path.join(_address_map_dir, "template.html"), encoding="utf-8") as _f:
    _address_map_html = _f.read()
with open(os.path.join(_address_map_dir, "style.css"), encoding="utf-8") as _f:
    _address_map_css = _f.read()
with open(os.path.join(_address_map_dir, "script.js"), encoding="utf-8") as _f:
    _address_map_js = _f.read()
with open(os.path.join(_address_map_dir, "subway_stations.json"), encoding="utf-8") as _f:
    _subway_stations_json = _f.read()
with open(os.path.join(_address_map_dir, "subway_lines.json"), encoding="utf-8") as _f:
    _subway_lines_json = _f.read()
with open(os.path.join(_address_map_dir, "subway_exits.json"), encoding="utf-8") as _f:
    _subway_exits_json = _f.read()

# 서울 지하철역/노선/출구 데이터를 컴포넌트 JS에 정적으로 박아 넣는다 — 매 rerun마다
# data=로 다시 보내면 (역+노선+출구 합쳐 300KB 안팎을) 매번 소켓에 실어야 하지만, js=
# 문자열 자체에 넣으면 컴포넌트 등록 시 한 번만 전달된다. 이 데이터는 절대 안 바뀌므로
# 안전하다.
_address_map_js = (
    _address_map_js.replace('"__SUBWAY_STATIONS_JSON__"', _subway_stations_json)
    .replace('"__SUBWAY_LINES_JSON__"', _subway_lines_json)
    .replace('"__SUBWAY_EXITS_JSON__"', _subway_exits_json)
)

_address_map_component = st.components.v2.component(
    "address_map", html=_address_map_html, css=_address_map_css, js=_address_map_js,
)


def _mercator_pixel(lat: float, lon: float, zoom: float):
    """위경도를 주어진 줌 레벨의 웹 메르카토르 화면 픽셀 좌표(256px 타일 기준)로 변환."""
    lat_rad = math.radians(lat)
    n = 2 ** zoom
    x = (lon + 180.0) / 360.0 * n * 256
    y = (1.0 - math.log(math.tan(lat_rad) + 1 / math.cos(lat_rad)) / math.pi) / 2.0 * n * 256
    return x, y


def _fit_zoom(lat_span: float, lon_span: float, width_px=850, height_px=420, min_zoom=13, max_zoom=18):
    """지점들의 위경도 범위가 뷰포트에 적당히 들어차도록 줌 레벨을 역산한다.

    지점이 서로 아주 가까우면(좁은 범위) 줌을 더 올려서 화면상 픽셀 간격을
    벌려주고, 그만큼 라벨 겹침 해소(_declutter_label_levels)의 부담도 줄어든다.
    """
    lat_span = max(lat_span, 1e-5)
    lon_span = max(lon_span, 1e-5)
    zoom_lon = math.log2(width_px * 360.0 / (lon_span * 256.0))
    zoom_lat = math.log2(height_px * 360.0 / (lat_span * 256.0))
    return max(min_zoom, min(max_zoom, zoom_lon, zoom_lat))


def _declutter_label_levels(lats, lons, zoom, label_width_px=260, row_height_px=20):
    """겹치는 라벨을 세로로 쌓기 위한 단계(level)를 각 점마다 계산한다.

    deck.gl의 CollisionFilterExtension은 Streamlit이 JSON API에 등록해 두지
    않아 사용할 수 없었다(콘솔 에러로 확인) — 대신 렌더링에 쓰는 고정 줌
    레벨 기준으로 각 점의 실제 화면 픽셀 위치를 계산해서, 라벨 영역이 겹치는
    점들을 서로 다른 높이(level)에 배치하는 방식으로 정적으로 해결한다.
    """
    points = [_mercator_pixel(lat, lon, zoom) for lat, lon in zip(lats, lons)]
    placed = []  # (x, y, level)
    levels = []
    for x, y in points:
        level = 0
        while any(lvl == level and abs(x - px) < label_width_px and abs(y - py) < row_height_px
                   for px, py, lvl in placed):
            level += 1
        placed.append((x, y, level))
        levels.append(level)
    return levels


def render_address_map(
    df: pd.DataFrame,
    lat_col: str = "lat",
    lon_col: str = "lon",
    label_col: str = None,
    show_labels: bool = True,
    vworld_key: str = None,
    enable_selection: bool = False,
    selection_key: str = None,
    highlight_lat: float = None,
    highlight_lon: float = None,
):
    """지점을 구글 지도 스타일 핀으로 지도 위에 표시하고, show_labels=True면 핀 옆에

    주소도 표시한다. show_labels=False면 마커만 표시해 지점이 아주 많을 때 더 깔끔하게
    볼 수 있다.

    실제 렌더링은 address_map_component/(Leaflet을 삽입한 커스텀 Streamlit
    컴포넌트)가 담당한다 — 기본 배경지도(OpenStreetMap)는 키가 필요 없고,
    vworld_key를 넘기면 배경지도를 브이월드(국토지리정보원) 공식 지도로 바꾼다.
    이 함수는 파이썬 쪽에서 라벨 중복 제거, 겹침 방지용 세로 스택 단계 계산까지만
    하고 나머지는 컴포넌트에 데이터로 넘긴다.

    enable_selection=True면 마커 클릭을 감지해서, 클릭된 지점의 데이터(주소/lat/lon이
    담긴 dict)를 반환한다(클릭이 없으면 None).

    highlight_lat/highlight_lon을 넘기면(예: 표에서 행을 선택했을 때) 그 좌표에 노란
    테두리 원(halo)을 표시해 "이 마커가 지금 선택된 지점"임을 시각적으로 나타낸다.
    표 → 지도 방향 강조(이 옵션)와 지도 → 표 방향 강조(enable_selection)는 서로
    독립적으로 켤 수 있다.
    """
    plot_df = df.copy()
    if label_col and label_col in plot_df.columns:
        plot_df["주소"] = plot_df[label_col].astype(str)
    else:
        plot_df["주소"] = "(주소 없음)"

    # 완전히 같은 좌표(같은 지번의 서로 다른 거래 등)에 여러 행이 있으면 같은
    # 주소 라벨이 여러 개 중복 생성되어, 겹침 해소 로직이 이들을 서로 멀리
    # 떼어놓으면서 마커에서 동떨어져 보이는 원인이 된다. 좌표 단위로 하나만
    # 남기고, 나머지 건수는 "외 N건"으로 같은 라벨에 표시한다.
    dedup_key = list(zip(plot_df[lat_col].round(7), plot_df[lon_col].round(7)))
    plot_df["_dedup_key"] = dedup_key
    dup_counts = plot_df["_dedup_key"].value_counts()
    plot_df = plot_df.drop_duplicates(subset="_dedup_key", keep="first").reset_index(drop=True)
    plot_df["주소"] = [
        f"{addr} 외 {dup_counts[key] - 1}건" if dup_counts[key] > 1 else addr
        for addr, key in zip(plot_df["주소"], plot_df["_dedup_key"])
    ]
    plot_df = plot_df.drop(columns=["_dedup_key"])

    if len(plot_df) <= 1:
        zoom = 17
    else:
        lat_span = plot_df[lat_col].max() - plot_df[lat_col].min()
        lon_span = plot_df[lon_col].max() - plot_df[lon_col].min()
        # 여백 없이 딱 맞추면 가장자리 지점이 화면 밖으로 잘릴 수 있어 80% 여유를 둔다.
        # (이 zoom 값은 실제 카카오맵 줌과 정확히 같을 필요 없는, 라벨 간 픽셀 거리를
        # 추정하기 위한 내부 근사치일 뿐 — 실제 화면 범위는 컴포넌트가 LatLngBounds로
        # 자동으로 맞춘다.)
        zoom = _fit_zoom(lat_span * 1.8, lon_span * 1.8)
    levels = _declutter_label_levels(plot_df[lat_col].tolist(), plot_df[lon_col].tolist(), zoom)

    markers = [
        {"lat": float(row[lat_col]), "lon": float(row[lon_col]), "label": row["주소"], "offsetLevel": level}
        for (_, row), level in zip(plot_df.iterrows(), levels)
    ]
    highlight = None
    if highlight_lat is not None and highlight_lon is not None:
        highlight = {"lat": float(highlight_lat), "lon": float(highlight_lon)}

    component_kwargs = {}
    if enable_selection:
        # 트리거 값("selected")은 대응하는 on_selected_change 콜백을 넘겨야만
        # 결과 객체에 노출된다 — 실제로 콜백을 쓸 필요는 없어 빈 함수를 넘긴다.
        component_kwargs["on_selected_change"] = lambda: None

    result = _address_map_component(
        key=selection_key,
        data={
            "markers": markers,
            "highlight": highlight,
            "showLabels": show_labels,
            "enableSelection": enable_selection,
            "singleZoom": zoom if len(plot_df) <= 1 else 17,
            "vworldKey": vworld_key,
        },
        height=460,
        **component_kwargs,
    )
    return result.selected if enable_selection else None


@st.cache_data(ttl=3600, show_spinner=False)
def _load_dong_list(sigungu_code: str):
    """동 콤보박스용 목록을 1시간 캐시."""
    return get_dong_list(sigungu_code)


@st.cache_data(ttl=3600, show_spinner=False)
def _load_sigungu_list():
    """시군구 콤보박스용 목록(중복 제거)을 1시간 캐시."""
    return get_sigungu_list()


@st.cache_data(ttl=900, show_spinner=False)
def _load_district_titles(service_key: str, sigungu_code: str, bdong_code: str):
    """동 전체 표제부를 15분 캐시. 동단위 통계/노후건축물/내진 스캔 탭이 공유한다.

    wait_time=0.15는 밀집 지역(예: 이태원동)처럼 페이지 수가 많은 동에서 건축HUB
    게이트웨이가 HTTP 502(HTTP_ERROR)를 내는 게 실제 관찰됐다 — 단일 지번 조회는
    멀쩡히 되는데 대량 페이지네이션에서만 나는 것으로 보아 순간 요청 속도 문제로
    보인다. 요청 간격을 넉넉히 둬서 이를 줄인다(개별 요청 자체는 이미 재시도됨).
    """
    api = BuildingLedger(service_key)
    return get_building_ledger(
        api,
        ledger_type="표제부",
        sigungu_code=sigungu_code,
        bdong_code=bdong_code,
        max_rows=10000,
        wait_time=0.5,
    )


@st.cache_data(ttl=900, show_spinner=False)
def _load_district_prices(service_key: str, sigungu_code: str, bdong_code: str):
    """동 전체 주택가격(공시가격)을 15분 캐시. bun/ji 없이 호출하면 표제부 전체 수집과

    동일한 방식(동 단위 페이지네이션)으로 동 전체 호(관리건축물대장PK)의 공시가격 이력을
    한 번에 받아온다 — 필지별로 API를 반복 호출하지 않는다. 아파트 등은 건물 1개에 호가
    수십~수백 개라 표제부보다 행 수가 훨씬 많을 수 있어 max_rows를 더 넉넉히 둔다.

    wait_time은 _load_district_titles와 같은 이유(대량 페이지네이션에서 관찰된 HTTP 502)로
    0.5초로 늘렸다.
    """
    api = BuildingLedger(service_key)
    return get_building_ledger(
        api,
        ledger_type="주택가격",
        sigungu_code=sigungu_code,
        bdong_code=bdong_code,
        max_rows=20000,
        wait_time=0.5,
    )


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _load_seoul_trade_area_locations(seoul_key: str):
    """서울시 전체 상권(약 1,650개) 위치 정보를 6시간 캐시 (매 요청마다 다시 받기엔 무겁다)."""
    return get_seoul_trade_area_locations(seoul_key)


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _load_seoul_quarter_dataset(seoul_key: str, service: str):
    """서울시 상권분석서비스 분기 데이터(전역, 수만 건)를 6시간 캐시."""
    return get_seoul_trade_area_quarter_dataset(seoul_key, service)


st.set_page_config(page_title="건축물대장 조회", page_icon="🏢", layout="wide")


def _get_secret(key: str):
    try:
        return st.secrets.get(key)
    except Exception:
        return None


def _check_auth() -> bool:
    """st.secrets에 저장된 AUTH_ID/AUTH_PW로 로그인 게이트를 건다.

    비밀번호 자체는 코드에 없고 Streamlit Cloud의 'Secrets' 설정(배포본) 또는
    로컬 .streamlit/secrets.toml(둘 다 git에 커밋되지 않음)에서만 읽는다.
    """
    if st.session_state.get("authenticated"):
        return True

    st.title("🏢 국토교통부 건축물대장 조회")
    st.info("로그인 후 이용할 수 있습니다.")

    valid_id, valid_pw = _get_secret("AUTH_ID"), _get_secret("AUTH_PW")
    if not valid_id or not valid_pw:
        st.error(
            "서버에 로그인 정보(AUTH_ID/AUTH_PW)가 설정되어 있지 않습니다. "
            "Streamlit Cloud 앱 설정 > Secrets에 추가해주세요."
        )
        return False

    with st.form("login_form"):
        user_id = st.text_input("아이디")
        user_pw = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("로그인", type="primary")

    if submitted:
        if user_id == valid_id and user_pw == valid_pw:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
    return False


if not _check_auth():
    st.stop()

with st.sidebar:
    if st.button("로그아웃", key="logout_button"):
        st.session_state.authenticated = False
        st.rerun()

st.title("🏢 국토교통부 건축물대장 조회")
st.caption("PublicDataReader 기반 · 페이지네이션 버그 우회 적용")

with st.sidebar:
    st.header("공통 조회 조건")

    service_key = st.text_input(
        "공공데이터포털 서비스키",
        value="",
        type="password",
        placeholder="data.go.kr에서 발급받은 일반 인증키(Decoding)를 입력하세요",
        help="https://www.data.go.kr 에서 '건축HUB_건축물대장정보 서비스' 등을 활용신청하면 발급됩니다. "
             "이 키는 서버에 저장되지 않고 이 브라우저 세션에서만 사용됩니다.",
    )
    if service_key:
        _key_stripped = service_key.strip()
        _len_note = f"입력된 키 길이: {len(service_key)}자"
        if service_key != _key_stripped:
            _len_note += " ⚠️ 앞/뒤에 공백이나 줄바꿈이 포함되어 있습니다!"
        elif len(service_key) != 64:
            _len_note += " ⚠️ 일반적인 키는 64자입니다. 다시 복사해보세요."
        st.caption(_len_note)
        service_key = _key_stripped

    kakao_key = st.text_input(
        "카카오맵 REST API 키 (선택)",
        value="",
        type="password",
        placeholder="주소 좌표(위도/경도) 조회에 사용",
        help="https://developers.kakao.com 에서 앱 생성 후 'REST API 키'를 복사 (JavaScript 키 아님). "
             "브이월드 지오코딩은 해외 클라우드 배포 환경에서 정책적으로 차단돼(공간정보관리법 제16조) "
             "이 앱의 좌표 조회 기능은 전부 카카오로 전환했습니다.",
    )
    kakao_key = kakao_key.strip() if kakao_key else None

    vworld_key = st.text_input(
        "브이월드(V-World) 인증키 (선택)",
        value="",
        type="password",
        placeholder="지도 배경을 브이월드 지도로 쓰려면 입력 (없어도 기본 배경지도로 정상 동작)",
        help="https://www.vworld.kr 에서 발급. 지도의 배경지도를 브이월드 공식 지도로 바꾸는 데만 쓰입니다.",
    )
    vworld_key = vworld_key.strip() if vworld_key else None

    reb_key = st.text_input(
        "한국부동산원 인증키 (선택)",
        value="",
        type="password",
        placeholder="상업용부동산 공실률 조회에 사용",
        help="https://www.reb.or.kr/r-one (R-ONE 부동산통계정보시스템)에 로그인 후 "
             "'Open API > 인증키 발급내역'에서 발급. '🏬 상업용부동산 공실률' 탭에서만 쓰입니다.",
    )
    reb_key = reb_key.strip() if reb_key else None

    sangkwon_key = st.text_input(
        "소상공인시장진흥공단 인증키 (선택)",
        value="",
        type="password",
        placeholder="주변 상가업소 조회에 사용. 미입력 시 위 공공데이터포털 서비스키로 자동 시도합니다.",
        help="https://www.data.go.kr 에서 '소상공인시장진흥공단_상가(상권)정보_API'를 별도로 활용신청 후 승인되면, "
             "위에 입력한 공공데이터포털 서비스키와 보통 같은 값입니다. '🏪 주변 상가업소' 탭에서만 쓰입니다.",
    )
    sangkwon_key = sangkwon_key.strip() if sangkwon_key else (service_key or None)

    seoul_key = st.text_input(
        "서울 열린데이터광장 인증키 (선택)",
        value="",
        type="password",
        placeholder="서울시 우리마을가게 상권분석서비스(추정매출/생활인구 등)에 사용",
        help="https://data.seoul.go.kr 에 로그인 후 '마이페이지 > 인증키 신청현황'에서 발급/확인. "
             "서울시 소재 상권만 조회 가능합니다. '🏙️ 서울 상권분석' 탭에서만 쓰입니다.",
    )
    seoul_key = seoul_key.strip() if seoul_key else None

    st.divider()
    st.caption(
        "📎 건축물대장 열람본 이미지 업로드 (선택) — 위반건축물 여부·소유자현황·변동사항은 "
        "공공데이터 API에 없어서, 정부24 등에서 발급받은 갑/을 이미지를 올리면 여기서 함께 확인합니다."
    )
    _ledger_files = st.file_uploader(
        "건축물대장 갑/을 이미지 (여러 장 가능)",
        type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="ledger_upload",
        label_visibility="collapsed",
    )
    ledger_docs = []
    if _ledger_files:
        if analyze_ledger_image is None:
            st.caption("⚠️ 이 기능에 필요한 패키지(Pillow/numpy/pytesseract)가 설치되지 않았습니다.")
        else:
            _ledger_cache = st.session_state.setdefault("_ledger_ocr_cache", {})
            for _f in _ledger_files:
                _file_bytes = _f.getvalue()
                _cache_key = (_f.name, len(_file_bytes))
                if _cache_key not in _ledger_cache:
                    with st.spinner(f"'{_f.name}' 텍스트 인식 중..."):
                        _ledger_cache[_cache_key] = analyze_ledger_image(_file_bytes, _f.name)
                ledger_docs.append(_ledger_cache[_cache_key])
            _n_violation = sum(1 for d in ledger_docs if d["is_violation"])
            if _n_violation:
                st.warning(f"🚩 업로드한 {len(ledger_docs)}장 중 {_n_violation}장에서 '위반건축물' 표기가 감지됐습니다.")
            else:
                st.caption(f"{len(ledger_docs)}장 업로드됨 (위반건축물 표기는 감지되지 않음)")
            if not ledger_docs[0]["ocr_available"]:
                st.caption(f"⚠️ {ledger_docs[0]['ocr_message']}")
    st.session_state.ledger_docs = ledger_docs

    st.caption("주소 (서울특별시 + 구 + 동)")
    _addr_sigungu_df = _load_sigungu_list()
    _addr_seoul_gu_df = _addr_sigungu_df[_addr_sigungu_df["시도명"] == "서울특별시"].reset_index(drop=True)
    _addr_gu_options = _addr_seoul_gu_df["시군구명"].tolist()
    addr_gu = st.selectbox(
        "구", _addr_gu_options,
        index=_addr_gu_options.index("강남구") if "강남구" in _addr_gu_options else 0,
        key="addr_gu",
    )
    addr_sigungu_code = _addr_seoul_gu_df.loc[_addr_seoul_gu_df["시군구명"] == addr_gu, "시군구코드"].iloc[0]
    _addr_dong_options = _load_dong_list(addr_sigungu_code)
    addr_dong = st.selectbox(
        "동", _addr_dong_options,
        index=_addr_dong_options.index("역삼동") if "역삼동" in _addr_dong_options else 0,
        key="addr_dong",
    )
    st.caption(f"📍 서울특별시 {addr_gu} {addr_dong}")

    col1, col2 = st.columns(2)
    bun = col1.text_input("번(본번)", value="237")
    ji = col2.text_input("지(부번)", value="0", help="0이면 생략됩니다.")

    st.caption("※ 두 탭(단일 조회 / 종합 리포트)이 위 조건을 공통으로 사용합니다.")


def _resolve_codes():
    bdong_map = get_bdong_code_map(addr_sigungu_code)
    bdong_code = bdong_map.get(addr_dong)
    if not bdong_code:
        raise ValueError(f"'서울특별시 {addr_gu} {addr_dong}'의 법정동코드를 찾지 못했습니다.")
    st.success(
        f"인식된 주소: 서울특별시 {addr_gu} {addr_dong} "
        f"(시군구코드={addr_sigungu_code}, 법정동코드={bdong_code})"
    )
    return addr_sigungu_code, bdong_code, "서울특별시", addr_gu


def _render_ledger_docs_section(ledger_docs):
    """사이드바에서 업로드한 건축물대장 열람본 이미지를, 공공API로는 안 나오는
    위반건축물 여부/소유자현황/변동사항을 확인하는 용도로 각 탭에 이어서 보여준다."""
    if not ledger_docs:
        return
    st.divider()
    st.subheader("📎 업로드한 건축물대장 열람본")
    st.caption(
        "위반건축물 여부·소유자현황·변동사항은 공공데이터 API에 없는 정보라 "
        "업로드한 이미지로 직접 확인합니다. OCR 텍스트는 오탈자가 있을 수 있는 참고용이며, "
        "정확한 값은 옆의 원본 이미지로 대조하세요."
    )
    for doc in ledger_docs:
        col_img, col_info = st.columns([1, 1])
        col_img.image(doc["image_bytes"], caption=doc["filename"], width='stretch')
        with col_info:
            if doc["is_violation"]:
                st.error("🚩 위반건축물 표기 감지됨")
            else:
                st.success("위반건축물 표기 감지되지 않음")
            if not doc["ocr_available"]:
                st.caption(f"⚠️ {doc['ocr_message']}")
            elif doc["raw_text"].strip():
                with st.expander("OCR 원문 텍스트 (검색/복사용)"):
                    st.text(doc["raw_text"])
            else:
                st.caption("텍스트를 인식하지 못했습니다 — 위 이미지로 직접 확인하세요.")


tab_single, tab_report, tab_price, tab_district, tab_old, tab_map, tab_geocode, tab_commercial, tab_sangkwon, tab_seoul, tab_autopptx = st.tabs([
    "🔍 단일 조회", "📋 종합 리포트", "💰 실거래가",
    "📊 동단위 통계", "🏚️ 노후건축물", "🗺️ 지도 업로드",
    "📍 지오코딩", "🏬 상업용부동산 공실률", "🏪 주변 상가업소", "🏙️ 서울 상권분석", "📑 자동 pptx 리포트",
])

# ------------------------------------------------------------------
# 탭 1: 단일 대장 종류 조회
# ------------------------------------------------------------------
with tab_single:
    ledger_type = st.selectbox("건축물대장 종류", ALL_LEDGER_TYPES, index=2, key="single_ledger_type")

    with st.expander("생성일자 필터 (선택)"):
        start_date = st.text_input("검색 시작일 (YYYYMMDD)", value="", key="single_start")
        end_date = st.text_input("검색 종료일 (YYYYMMDD)", value="", key="single_end")

    if st.button("조회하기", type="primary", key="single_submit"):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            api = BuildingLedger(service_key)
            try:
                with st.spinner("주소를 코드로 변환하는 중..."):
                    sigungu_code, bdong_code, addr_sido, addr_sigungu_name = _resolve_codes()
                with st.spinner("건축물대장 조회 중..."):
                    df = get_building_ledger(
                        api,
                        ledger_type=ledger_type,
                        sigungu_code=sigungu_code,
                        bdong_code=bdong_code,
                        bun=bun or None,
                        ji=ji if ji and ji != "0" else None,
                        start_date=start_date or None,
                        end_date=end_date or None,
                        verbose=False,
                    )
                st.session_state.single_df = df
            except Exception as e:
                st.error(f"조회 실패: {e}")
                st.session_state.single_df = None

    single_df = st.session_state.get("single_df")
    if single_df is None:
        st.info("조건을 입력하고 **조회하기**를 눌러주세요.")
    elif single_df.empty:
        st.warning("조회된 데이터가 없습니다. 주소/번지/대장 종류를 확인해주세요.")
    else:
        st.success(f"{len(single_df)}건 조회됨")
        st.dataframe(single_df, width='stretch')

        col_a, col_b = st.columns(2)
        csv_bytes = single_df.to_csv(index=False).encode("utf-8-sig")
        col_a.download_button("📄 CSV 다운로드", csv_bytes, "building_result.csv", "text/csv",
                               width='stretch')

        excel_buf = io.BytesIO()
        single_df.to_excel(excel_buf, index=False)
        col_b.download_button(
            "📊 엑셀 다운로드", excel_buf.getvalue(), "building_result.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
        )

    _render_ledger_docs_section(st.session_state.get("ledger_docs"))

# ------------------------------------------------------------------
# 탭 2: 번지 하나 -> 11종 건축물대장 전체 종합 리포트
# ------------------------------------------------------------------
with tab_report:
    st.write("왼쪽에 입력한 번지 하나에 대해 아래 11종 건축물대장 정보를 한 번에 조회합니다.")
    st.caption(" · ".join(ALL_LEDGER_TYPES))

    if st.button("전체 조회 (리포트 생성)", type="primary", key="report_submit"):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            api = BuildingLedger(service_key)
            try:
                with st.spinner("주소를 코드로 변환하는 중..."):
                    sigungu_code, bdong_code, addr_sido, addr_sigungu_name = _resolve_codes()

                progress = st.progress(0.0, text="조회 준비 중...")
                status = st.empty()
                report = {}
                total = len(ALL_LEDGER_TYPES)
                for i, lt in enumerate(ALL_LEDGER_TYPES, start=1):
                    status.write(f"({i}/{total}) **{lt}** 조회 중...")
                    single_report = get_full_building_report(
                        api, sigungu_code, bdong_code,
                        bun=bun or None,
                        ji=ji if ji and ji != "0" else None,
                        ledger_types=[lt],
                        verbose=False,
                    )
                    report.update(single_report)
                    progress.progress(i / total, text=f"({i}/{total}) {lt} 완료")
                status.write("완료!")

                st.session_state.report = report
                # 자동 pptx 리포트 탭이 "같은 번지로 조회된 결과인지" 확인할 수 있게
                # 조회에 실제로 쓰인 조건을 함께 저장해둔다.
                st.session_state.report_address_key = (
                    sigungu_code, bdong_code, bun or None, ji if ji and ji != "0" else None,
                )
            except Exception as e:
                st.error(f"리포트 생성 실패: {e}")
                st.session_state.report = None
                st.session_state.report_address_key = None

    report = st.session_state.get("report")
    if not report:
        st.info("**전체 조회 (리포트 생성)** 버튼을 눌러주세요.")
    else:
        # --- 요약 카드: 표제부 우선, 없으면 총괄표제부 사용 ---
        summary_src = None
        for key in ("표제부", "총괄표제부"):
            candidate = report.get(key)
            if candidate is not None and not candidate.empty and "오류" not in candidate.columns:
                summary_src = candidate.iloc[0]
                break

        if summary_src is not None:
            st.subheader("📌 요약")
            addr = summary_src.get("도로명대지위치") or summary_src.get("대지위치", "")
            st.markdown(f"**{addr}** · {summary_src.get('건물명', '') or '(건물명 없음)'}")
            _ledger_docs_for_summary = st.session_state.get("ledger_docs") or []
            if any(d["is_violation"] for d in _ledger_docs_for_summary):
                st.error("🚩 업로드한 건축물대장 열람본에서 '위반건축물' 표기가 감지됐습니다. (아래 첨부 이미지 참고)")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("주용도", summary_src.get("주용도코드명", "-"))
            m2.metric("구조", summary_src.get("구조코드명", "-"))
            m3.metric("지상/지하층수", f"{summary_src.get('지상층수', '-')} / {summary_src.get('지하층수', '-')}")
            m4.metric("사용승인일", str(summary_src.get("사용승인일", "-")))
            m5, m6, m7, m8 = st.columns(4)
            m5.metric("대지면적(㎡)", summary_src.get("대지면적", "-"))
            m6.metric("연면적(㎡)", summary_src.get("연면적", "-"))
            m7.metric("건폐율(%)", summary_src.get("건폐율", "-"))
            m8.metric("용적률(%)", summary_src.get("용적률", "-"))
            st.divider()

        # --- 다운로드: 종합 PDF 리포트 (한 문서로 이어짐) / 엑셀(시트별) ---
        addr_label = ""
        if summary_src is not None:
            addr_label = summary_src.get("도로명대지위치") or summary_src.get("대지위치", "") or ""

        col_pptx, col_pdf, col_xlsx = st.columns(3)
        try:
            pptx_bytes = generate_ledger_pptx(
                report, address_label=addr_label, ledger_docs=st.session_state.get("ledger_docs"),
            )
            col_pptx.download_button(
                "📑 종합 리포트 PPTX 다운로드 (표지·차트 포함)",
                pptx_bytes,
                "building_full_report.pptx",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                width='stretch',
            )
        except Exception as e:
            col_pptx.error(f"PPTX 생성 실패: {e}")

        try:
            pdf_bytes = generate_pdf_report(report, address_label=addr_label)
            col_pdf.download_button(
                "📄 간단 요약 PDF 다운로드",
                pdf_bytes,
                "building_full_report.pdf",
                "application/pdf",
                width='stretch',
            )
        except Exception as e:
            col_pdf.error(f"PDF 생성 실패: {e}")

        report_excel_buf = io.BytesIO()
        with pd.ExcelWriter(report_excel_buf, engine="openpyxl") as writer:
            for lt, ldf in report.items():
                ldf.to_excel(writer, sheet_name=lt[:31], index=False)
        col_xlsx.download_button(
            "📊 엑셀 다운로드 (대장 종류별 시트 분리)",
            report_excel_buf.getvalue(),
            "building_full_report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
        )

        st.divider()

        # --- 대장 종류별 상세: 클릭 없이 전부 이어서 표시 ---
        st.subheader("📂 전체 상세 (한 화면에 이어서 표시)")
        for lt in ALL_LEDGER_TYPES:
            ldf = report.get(lt)
            if ldf is None:
                continue
            if lt == "소유자":
                st.markdown(f"#### 🚫 {lt} (공공API 미제공)")
                st.info(
                    "건축물 소유자 정보 오픈API(고유번호 15021136)는 현재 data.go.kr에서 "
                    "서비스가 종료된 상태입니다. 소유자 개인정보 보안 사유로 더 이상 "
                    "공개 API로 제공되지 않는 것으로 보이며, 필요 시 정부24 등기부등본/"
                    "건축물대장 열람 등 별도 인증 경로를 이용해야 합니다. "
                    "사이드바에 열람본 이미지를 업로드했다면 이 탭 맨 아래에서 위반건축물 여부와 "
                    "함께 확인할 수 있습니다."
                )
            elif "오류" in ldf.columns:
                st.markdown(f"#### ⚠️ {lt} (조회 실패)")
                st.error(ldf.iloc[0]["오류"])
            elif ldf.empty:
                st.markdown(f"#### – {lt} (데이터 없음)")
            elif len(ldf) > 1:
                st.markdown(f"#### ✅ {lt} ({len(ldf)}건 — 표 하나로 한눈에 표시)")
                common, varying_df = split_common_and_varying(ldf)
                if common:
                    st.caption(" · ".join(f"{k}: {v}" for k, v in common.items()))
                st.dataframe(varying_df if len(varying_df.columns) > 0 else ldf, width='stretch')
            else:
                st.markdown(f"#### ✅ {lt} ({len(ldf)}건)")
                st.dataframe(ldf, width='stretch')

        _render_ledger_docs_section(st.session_state.get("ledger_docs"))

# ------------------------------------------------------------------
# 탭 3: 실거래가 (전체 부동산 유형)
# ------------------------------------------------------------------
with tab_price:
    _tp_meta = TransactionPrice().meta_dict
    tp_property_type = st.selectbox("부동산 유형", list(_tp_meta.keys()), key="tp_ptype")
    tp_trade_type = st.selectbox("거래 유형", list(_tp_meta[tp_property_type].keys()), key="tp_ttype")

    _sigungu_df = _load_sigungu_list()
    col_sido, col_sigungu = st.columns(2)
    sido_options = sorted(_sigungu_df["시도명"].unique())
    tp_sido = col_sido.selectbox(
        "시/도", sido_options,
        index=sido_options.index("경기도") if "경기도" in sido_options else 0,
        key="tp_sido",
    )
    _sigungu_options_df = _sigungu_df[_sigungu_df["시도명"] == tp_sido].reset_index(drop=True)
    sigungu_names = _sigungu_options_df["시군구명"].tolist()
    tp_sigungu_name = col_sigungu.selectbox(
        "시/군/구", sigungu_names,
        index=sigungu_names.index("성남시 분당구") if "성남시 분당구" in sigungu_names else 0,
        key="tp_sigungu_name",
    )
    tp_sigungu = _sigungu_options_df.loc[_sigungu_options_df["시군구명"] == tp_sigungu_name, "시군구코드"].iloc[0]
    st.caption(f"선택된 시군구코드: {tp_sigungu}")

    dong_options = ["(전체)"] + _load_dong_list(tp_sigungu)
    tp_dong = st.selectbox(
        "동", dong_options, key="tp_dong",
        help="실거래가 API는 시군구 단위로만 조회되므로, 여기서 특정 동을 고르면 조회 후 결과를 그 동으로 걸러서 보여줍니다.",
    )

    tp_with_price = st.checkbox(
        "시가표준액(공시가격) 함께 조회 (지번당 API 1회 추가 호출, 다소 느려질 수 있음)",
        value=False, key="tp_with_price",
        help="지번이 마스킹된 상업업무용 거래, 토지 등은 건축HUB 주택가격 API 범위 밖이라 비어있게 됩니다.",
    )

    tp_with_reverse_match = st.checkbox(
        "마스킹된 지번 역매칭 시도 (표제부 대조 — 특정 동을 골라야 사용 가능, 최초 20~40초)",
        value=False, key="tp_with_reverse_match", disabled=(tp_dong == "(전체)"),
        help="상업업무용처럼 지번이 '8*'같이 마스킹된 거래를, 같은 동의 표제부 전체(연면적·"
             "대지면적·건축년도)와 대조해 실제 필지를 추정합니다. deal-locator MCP와 같은 원리입니다.",
    )
    if tp_dong == "(전체)" and tp_with_reverse_match:
        st.caption("⚠️ 역매칭을 쓰려면 위에서 '동'을 특정 동으로 선택해주세요.")

    tp_with_coords = st.checkbox(
        "주소 좌표(위도/경도) 함께 조회 — 지도 표시용 (카카오 REST API 키 필요, 고유 주소당 1회 호출)",
        value=False, key="tp_with_coords", disabled=not kakao_key,
        help="사이드바에 카카오맵 REST API 키를 입력해야 사용할 수 있습니다.",
    )
    if not kakao_key and tp_with_coords:
        st.caption("⚠️ 좌표 조회를 쓰려면 사이드바에 카카오맵 REST API 키를 입력해주세요.")

    tp_period_mode = st.checkbox("기간으로 조회 (여러 달)", value=False, key="tp_period_mode")
    tp_start = tp_end = tp_year_month = None
    if tp_period_mode:
        colp1, colp2 = st.columns(2)
        tp_start = colp1.text_input("시작 연월 (YYYYMM)", value="202401", key="tp_start")
        tp_end = colp2.text_input("종료 연월 (YYYYMM)", value="202506", key="tp_end")
    else:
        tp_year_month = st.text_input("조회 연월 (YYYYMM)", value="202506", key="tp_ym")

    if st.button("실거래가 조회", type="primary", key="tp_submit"):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            tp_api = TransactionPrice(service_key)
            try:
                with st.spinner("조회 중..."):
                    if tp_period_mode:
                        tp_df = tp_api.get_data(
                            property_type=tp_property_type, trade_type=tp_trade_type,
                            sigungu_code=tp_sigungu,
                            start_year_month=tp_start, end_year_month=tp_end,
                        )
                    else:
                        tp_df = tp_api.get_data(
                            property_type=tp_property_type, trade_type=tp_trade_type,
                            sigungu_code=tp_sigungu, year_month=tp_year_month,
                        )
                    tp_df = add_address_column(tp_df, sido=tp_sido, sigungu_name=tp_sigungu_name)
                    if tp_dong != "(전체)" and "법정동" in tp_df.columns:
                        tp_df = tp_df[tp_df["법정동"] == tp_dong].reset_index(drop=True)

                # 아래 세 항목(시가표준액/역매칭/좌표)은 모두 선택 부가기능이다 — 정부 API가
                # 일시적으로 오류(게이트웨이 502 등)를 내더라도, 이미 성공적으로 받아온 기본
                # 실거래가 결과(tp_df)까지 통째로 날리지 않도록 각각 독립적으로 감싼다.
                if tp_with_price and not tp_df.empty:
                    try:
                        ledger_api = BuildingLedger(service_key)
                        price_progress = st.progress(0.0, text="시가표준액 조회 준비 중...")

                        def _on_price_progress(i, total):
                            price_progress.progress(i / total, text=f"시가표준액 조회 중... ({i}/{total} 지번)")

                        tp_df = add_standard_price_column(
                            tp_df, ledger_api, tp_sigungu, progress_callback=_on_price_progress,
                        )
                        price_progress.empty()
                    except Exception as e:
                        st.warning(f"시가표준액 조회에 실패해 이 항목만 건너뜁니다: {e}")

                if tp_with_reverse_match and tp_dong != "(전체)" and not tp_df.empty:
                    has_masked = tp_df["지번"].astype(str).str.contains(r"\*").any() if "지번" in tp_df.columns else False
                    if has_masked:
                        try:
                            ledger_api = BuildingLedger(service_key)
                            bdong_map = get_bdong_code_map(tp_sigungu)
                            bdong_code = bdong_map.get(tp_dong)
                            if bdong_code:
                                with st.spinner(f"'{tp_dong}' 표제부 전체 수집 중... (역매칭용, 캐시되어 있으면 즉시 완료)"):
                                    dong_title_df = _load_district_titles(service_key, tp_sigungu, bdong_code)
                                tp_df = reverse_match_transactions(tp_df, dong_title_df)
                            else:
                                st.warning(f"'{tp_dong}'의 법정동코드를 찾지 못해 역매칭을 건너뜁니다.")
                        except Exception as e:
                            st.warning(f"역매칭에 실패해 이 항목만 건너뜁니다: {e}")
                    else:
                        st.caption("마스킹된 지번이 없어 역매칭을 건너뜁니다.")

                if tp_with_coords and kakao_key and not tp_df.empty and "주소" in tp_df.columns:
                    try:
                        coord_progress = st.progress(0.0, text="좌표 조회 준비 중...")

                        def _on_coord_progress(i, total):
                            coord_progress.progress(i / total, text=f"주소 좌표 조회 중... ({i}/{total} 고유 주소)")

                        tp_df = add_coordinates_column(
                            tp_df, kakao_key, progress_callback=_on_coord_progress,
                        )
                        coord_progress.empty()
                    except Exception as e:
                        st.warning(f"좌표 조회에 실패해 이 항목만 건너뜁니다: {e}")

                st.session_state.tp_df = tp_df
            except Exception as e:
                st.error(f"조회 실패: {e}")
                st.session_state.tp_df = None

    tp_df = st.session_state.get("tp_df")
    if tp_df is None:
        st.info("조건을 입력하고 **실거래가 조회**를 눌러주세요.")
    elif tp_df.empty:
        st.warning("조회된 거래가 없습니다.")
    else:
        st.success(f"{len(tp_df)}건 조회됨")
        st.dataframe(tp_df, width='stretch')
        col_a, col_b = st.columns(2)
        csv_bytes = tp_df.to_csv(index=False).encode("utf-8-sig")
        col_a.download_button("📄 CSV 다운로드", csv_bytes, "transaction_price.csv", "text/csv",
                               width='stretch', key="tp_csv")
        excel_buf = io.BytesIO()
        tp_df.to_excel(excel_buf, index=False)
        col_b.download_button(
            "📊 엑셀 다운로드", excel_buf.getvalue(), "transaction_price.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch', key="tp_xlsx",
        )

        if "위도" in tp_df.columns and "경도" in tp_df.columns:
            map_df = tp_df.dropna(subset=["위도", "경도"])
            if not map_df.empty:
                st.subheader("🗺️ 거래 위치 지도")
                st.caption(f"좌표가 확인된 {len(map_df)}/{len(tp_df)}건을 표시합니다.")
                render_address_map(map_df, lat_col="위도", lon_col="경도", label_col="주소", vworld_key=vworld_key)

        tx_stats = analyze_transaction_stats(tp_df)
        if tx_stats["총괄"]:
            scope_label = f"{tp_sigungu_name} {tp_dong}" if tp_dong != "(전체)" else tp_sigungu_name
            st.subheader(f"📊 {scope_label} 주변통계")
            st.caption("이미 조회된 거래 내역만 집계한 것으로, 추가 API 호출 없이 즉시 계산됩니다.")
            s = tx_stats["총괄"]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("건수", f"{s['건수']:,}")
            c2.metric("평균 거래가", f"{s['평균거래가(억)']:.2f}억")
            c3.metric("중앙값 거래가", f"{s['중앙값거래가(억)']:.2f}억")
            c4.metric("최고가 / 최저가", f"{s['최고가(억)']:.2f}억 / {s['최저가(억)']:.2f}억")
            if "평균평당가(만원)" in s:
                c5, c6 = st.columns(2)
                c5.metric("평균 평당가", f"{s['평균평당가(만원)']:,.0f}만원" if s["평균평당가(만원)"] is not None else "-")
                c6.metric("평균 면적", f"{s['평균면적(평)']}평")
            if not tx_stats["유형별"].empty:
                st.markdown("**구분별 평균 거래가**")
                st.dataframe(tx_stats["유형별"], width='stretch')
            if not tx_stats["연월별추이"].empty:
                st.markdown("**연월별 평균 거래가 추이**")
                st.line_chart(tx_stats["연월별추이"].set_index("연월")["평균거래가(억)"])

# ------------------------------------------------------------------
# 탭 4: 동단위 통계
# ------------------------------------------------------------------
with tab_district:
    st.caption("동 전체 표제부를 기준으로 집계합니다. 첫 조회는 동 규모에 따라 20~40초 걸릴 수 있고, "
               "이후 15분간 캐시되어 즉시 응답합니다. (사이드바의 번지는 무시하고 동 전체를 봅니다)")
    if st.button("동단위 통계 조회", type="primary", key="district_submit"):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            try:
                with st.spinner("주소를 코드로 변환하는 중..."):
                    sigungu_code, bdong_code, addr_sido, addr_sigungu_name = _resolve_codes()
                with st.spinner("표제부 전체 수집 중... (캐시되어 있으면 즉시 완료)"):
                    title_df = _load_district_titles(service_key, sigungu_code, bdong_code)
                st.session_state.district_stats = analyze_district_stats(title_df)
            except Exception as e:
                st.error(f"조회 실패: {e}")
                st.session_state.district_stats = None

    stats = st.session_state.get("district_stats")
    if not stats or not stats.get("총괄"):
        st.info("**동단위 통계 조회** 버튼을 눌러주세요.")
    else:
        s = stats["총괄"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("총 동수", f"{s['총동수']:,}")
        c2.metric("총 연면적(㎡)", f"{s['총연면적(㎡)']:,.0f}")
        c3.metric("평균 층수", s["평균층수"])
        c4.metric("평균 경과연수", s["평균경과연수"])

        st.subheader("주용도별 (상위 10)")
        st.dataframe(stats["주용도별"], width='stretch')
        st.subheader("사용승인 연대별")
        st.dataframe(stats["연대별"], width='stretch')
        st.subheader("노후도 분포")
        st.dataframe(stats["노후도분포"], width='stretch')
        st.subheader("주용도별 규모 벤치마크 (중앙값)")
        st.dataframe(stats["규모벤치마크"], width='stretch')

# ------------------------------------------------------------------
# 탭 5: 노후건축물
# ------------------------------------------------------------------
with tab_old:
    st.caption("동 전체 표제부에서 경과연수 기준으로 노후 건물을 선별합니다. (동단위 통계와 데이터 공유·캐시)")
    min_age = st.slider("최소 경과연수", 10, 80, 30, key="old_min_age")
    if st.button("노후건축물 조회", type="primary", key="old_submit"):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            try:
                with st.spinner("주소를 코드로 변환하는 중..."):
                    sigungu_code, bdong_code, addr_sido, addr_sigungu_name = _resolve_codes()
                with st.spinner("표제부 전체 수집 중... (캐시되어 있으면 즉시 완료)"):
                    title_df = _load_district_titles(service_key, sigungu_code, bdong_code)
                st.session_state.old_title_df = title_df
            except Exception as e:
                st.error(f"조회 실패: {e}")
                st.session_state.old_title_df = None

    old_title_df = st.session_state.get("old_title_df")
    if old_title_df is None:
        st.info("**노후건축물 조회** 버튼을 눌러주세요.")
    else:
        old_df = analyze_old_buildings(old_title_df, min_age_years=min_age)
        st.success(f"전체 {len(old_title_df)}건 중 경과 {min_age}년↑ {len(old_df)}건")
        st.dataframe(old_df, width='stretch')
        csv_bytes = old_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("📄 CSV 다운로드", csv_bytes, "old_buildings.csv", "text/csv", key="old_csv")

# ------------------------------------------------------------------
# 탭 6: 지도 업로드 — 주소·위도·경도가 담긴 파일을 올리면 지도에 표시
# ------------------------------------------------------------------
with tab_map:
    st.write("**주소·위도·경도** 컬럼이 포함된 CSV 또는 엑셀 파일을 업로드하면 모든 지점을 지도에 표시합니다.")
    st.caption("위도 컬럼명 예시: 위도, lat, latitude, y  ·  경도 컬럼명 예시: 경도, lon, lng, longitude, x")

    uploaded = st.file_uploader("파일 업로드", type=["csv", "xlsx", "xls"], key="map_upload")

    if uploaded is not None:
        try:
            if uploaded.name.lower().endswith(".csv"):
                map_df = pd.read_csv(uploaded)
            else:
                map_df = pd.read_excel(uploaded)
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            map_df = None

        if map_df is not None:
            _lat_names = {"위도", "lat", "latitude", "y"}
            _lon_names = {"경도", "lon", "lng", "longitude", "x"}
            _addr_names = {"주소", "address", "지번주소", "도로명주소"}
            lat_col = next((c for c in map_df.columns if str(c).strip().lower() in _lat_names), None)
            lon_col = next((c for c in map_df.columns if str(c).strip().lower() in _lon_names), None)
            addr_col = next((c for c in map_df.columns if str(c).strip().lower() in _addr_names), None)

            if not lat_col or not lon_col:
                st.error("위도/경도로 보이는 컬럼을 찾지 못했습니다. 아래 미리보기에서 컬럼명을 확인해주세요.")
                st.dataframe(map_df.head(20), width='stretch')
            else:
                cols = [lat_col, lon_col] + ([addr_col] if addr_col else [])
                plot_df = map_df[cols].copy()
                plot_df.columns = ["lat", "lon"] + (["주소"] if addr_col else [])
                plot_df["lat"] = pd.to_numeric(plot_df["lat"], errors="coerce")
                plot_df["lon"] = pd.to_numeric(plot_df["lon"], errors="coerce")
                before = len(plot_df)
                plot_df = plot_df.dropna(subset=["lat", "lon"])
                dropped = before - len(plot_df)

                st.success(f"{len(plot_df)}개 지점을 지도에 표시합니다." + (f" ({dropped}건은 좌표 형식이 올바르지 않아 제외)" if dropped else ""))
                if not addr_col:
                    st.caption("'주소' 컬럼이 없어 마커에 좌표만 표시됩니다. 주소를 툴팁으로 보려면 '주소' 컬럼을 포함해주세요.")

                display_mode = st.selectbox(
                    "지도 표시 방식",
                    ["마커 + 주소 표시", "마커만 표시"],
                    key="map_display_mode",
                    disabled=not addr_col,
                    help="지점이 많아 라벨이 복잡해 보이면 '마커만 표시'로 바꿔보세요.",
                )
                show_labels = bool(addr_col) and display_mode == "마커 + 주소 표시"

                # 표에서 행을 선택한 적이 있으면(이전 rerun에서 위젯이 이미 등록돼
                # session_state에 값이 있으면), 그 지점을 지도에 노란 원으로 표시한다.
                # 위젯 키의 값은 이 rerun이 시작되기 전에 이미 갱신돼 있으므로, 아래
                # st.dataframe 호출보다 먼저 읽어도 방금 클릭한 행을 정확히 반영한다.
                highlight_lat = highlight_lon = highlight_addr = None
                _prev_table_sel = st.session_state.get("map_upload_table_select")
                if _prev_table_sel:
                    _rows = _prev_table_sel.get("selection", {}).get("rows", [])
                    if _rows and 0 <= _rows[0] < len(map_df):
                        _hrow = map_df.iloc[_rows[0]]
                        try:
                            highlight_lat = float(_hrow[lat_col])
                            highlight_lon = float(_hrow[lon_col])
                        except (TypeError, ValueError):
                            highlight_lat = highlight_lon = None
                        highlight_addr = str(_hrow[addr_col]) if addr_col else f"{_rows[0]}번 행"

                selected = render_address_map(
                    plot_df, label_col="주소" if addr_col else None,
                    show_labels=show_labels, vworld_key=vworld_key,
                    enable_selection=True, selection_key="map_upload_selection",
                    highlight_lat=highlight_lat, highlight_lon=highlight_lon,
                )
                if highlight_lat is not None:
                    st.caption(f"📍 표에서 선택한 **{highlight_addr}**의 위치를 지도에 노란 원으로 표시했습니다.")

                st.subheader("업로드한 데이터")
                st.caption("행 번호(맨 왼쪽)를 클릭하면 그 지점의 마커 위치가 지도에 표시됩니다.")
                display_df = add_pyeong_columns(map_df)

                match_mask = None
                if selected:
                    sel_lat = round(float(selected.get("lat")), 7)
                    sel_lon = round(float(selected.get("lon")), 7)
                    lat_num = pd.to_numeric(map_df[lat_col], errors="coerce").round(7)
                    lon_num = pd.to_numeric(map_df[lon_col], errors="coerce").round(7)
                    match_mask = (lat_num == sel_lat) & (lon_num == sel_lon)

                table_kwargs = dict(
                    on_select="rerun", selection_mode="single-row",
                    key="map_upload_table_select", width='stretch',
                )
                if match_mask is not None and match_mask.any():
                    st.caption(
                        f"🔎 지도에서 **{selected.get('주소', '선택한 지점')}** 을(를) 클릭했습니다 — "
                        "아래 표에서 강조 표시된 행입니다. (Streamlit 표는 특정 행으로 자동 스크롤은 "
                        "지원하지 않아, 색상 강조로 대신합니다.)"
                    )

                    def _highlight_selected(row):
                        return ["background-color: #FFE08A"] * len(row) if match_mask.loc[row.name] else [""] * len(row)

                    st.dataframe(display_df.style.apply(_highlight_selected, axis=1), **table_kwargs)
                else:
                    st.dataframe(display_df, **table_kwargs)
    else:
        st.info("파일을 업로드해주세요.")

# ------------------------------------------------------------------
# 탭 7: 지오코딩 — 주소만 넣으면 경도·위도만 반환
# ------------------------------------------------------------------
with tab_geocode:
    st.write("**주소**를 입력하면 경도·위도 좌표만 조회합니다. 도로명·지번 주소 모두 지원합니다.")
    if not kakao_key:
        st.warning("사이드바에 카카오맵 REST API 키를 입력해야 사용할 수 있습니다.")
    else:
        st.subheader("🔍 단일 주소 검색")
        single_addr = st.text_input(
            "주소", placeholder="예: 서울특별시 강남구 테헤란로 152  또는  서울특별시 강남구 역삼동 737",
            key="geo_single_addr",
        )
        if st.button("좌표 찾기", type="primary", key="geo_single_submit"):
            addr = single_addr.strip()
            if not addr:
                st.warning("주소를 입력해주세요.")
            else:
                with st.spinner("조회 중..."):
                    coord, reason = geocode_address_kakao(kakao_key, addr)
                st.session_state.geo_single_result = (addr, coord, reason)

        single_result = st.session_state.get("geo_single_result")
        if single_result:
            found_addr, coord, reason = single_result
            if coord:
                lon, lat = coord
                c1, c2 = st.columns(2)
                c1.metric("경도 (lon)", f"{lon:.6f}")
                c2.metric("위도 (lat)", f"{lat:.6f}")
                render_address_map(
                    pd.DataFrame({"lat": [lat], "lon": [lon], "주소": [found_addr]}),
                    label_col="주소", vworld_key=vworld_key,
                )
            else:
                st.error(f"좌표를 찾지 못했습니다 ({reason}). 주소를 다시 확인해주세요.")

        st.divider()
        st.subheader("📄 파일 일괄 변환")
        st.caption("'주소' 컬럼이 포함된 CSV/엑셀 파일을 올리면 각 행의 경도·위도를 찾아 추가합니다.")
        geo_file = st.file_uploader("파일 업로드", type=["csv", "xlsx", "xls"], key="geo_file_upload")

        if geo_file is not None:
            try:
                if geo_file.name.lower().endswith(".csv"):
                    geo_df = pd.read_csv(geo_file)
                else:
                    geo_df = pd.read_excel(geo_file)
            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
                geo_df = None

            if geo_df is not None:
                _addr_names = {"주소", "address", "지번주소", "도로명주소"}
                addr_col = next((c for c in geo_df.columns if str(c).strip().lower() in _addr_names), None)
                if not addr_col:
                    st.error("주소로 보이는 컬럼을 찾지 못했습니다. 컬럼명을 '주소'로 바꿔서 다시 올려주세요.")
                    st.dataframe(geo_df.head(20), width='stretch')
                else:
                    if st.button("전체 좌표 조회", type="primary", key="geo_batch_submit"):
                        work_df = geo_df[[addr_col]].rename(columns={addr_col: "주소"})
                        geo_progress = st.progress(0.0, text="지오코딩 준비 중...")

                        def _on_geo_progress(i, total):
                            geo_progress.progress(i / total, text=f"좌표 조회 중... ({i}/{total} 고유 주소)")

                        st.session_state.geo_batch_result = add_coordinates_column(
                            work_df, kakao_key, progress_callback=_on_geo_progress,
                        )
                        geo_progress.empty()

        batch_result = st.session_state.get("geo_batch_result")
        if batch_result is not None:
            found = int(batch_result["위도"].notna().sum())
            st.success(f"{found}/{len(batch_result)}건 좌표를 찾았습니다.")
            if found == 0 and "좌표조회실패사유" in batch_result.columns:
                top_reason = batch_result["좌표조회실패사유"].dropna().mode()
                if not top_reason.empty:
                    st.error(f"공통 실패 사유: {top_reason.iloc[0]}")
            st.dataframe(batch_result, width='stretch')
            csv_bytes = batch_result.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "📄 CSV 다운로드", csv_bytes, "geocoded_addresses.csv", "text/csv",
                width='stretch', key="geo_batch_csv",
            )
            map_df = batch_result.dropna(subset=["위도", "경도"])
            if not map_df.empty:
                render_address_map(map_df, lat_col="위도", lon_col="경도", label_col="주소", vworld_key=vworld_key)

# ------------------------------------------------------------------
# 탭 8: 상업용부동산 공실률 (한국부동산원 R-ONE Open API)
# ------------------------------------------------------------------
with tab_commercial:
    st.write("한국부동산원 상업용부동산 임대동향조사의 상권별 공실률을 조회합니다. "
             "(상권 재구획 이후인 2024년 3분기~ 데이터만 제공합니다.)")
    if not reb_key:
        st.warning("사이드바에 한국부동산원 인증키를 입력해야 사용할 수 있습니다.")
    else:
        def _fmt_quarter(q):
            return f"{q[:4]}년 {int(q[4:6])}분기"

        c1, c2 = st.columns(2)
        commercial_type = c1.selectbox("부동산 유형", list(REB_COMMERCIAL_VACANCY_STATBL_IDS.keys()))
        latest_quarter = reb_current_quarter_id()
        quarter_options = reb_quarter_ids_desc("202403", latest_quarter)
        selected_quarter = c2.selectbox("분기", quarter_options, format_func=_fmt_quarter)

        if st.button("공실률 조회", type="primary", key="commercial_submit"):
            statbl_id = REB_COMMERCIAL_VACANCY_STATBL_IDS[commercial_type]
            try:
                with st.spinner("조회 중..."):
                    snap_df, used_quarter, snap_message = get_reb_vacancy_snapshot(reb_key, statbl_id, selected_quarter)
            except RuntimeError as e:
                st.error(str(e))
                snap_df, used_quarter, snap_message = pd.DataFrame(), selected_quarter, None

            st.session_state.commercial_snapshot = snap_df
            st.session_state.commercial_snapshot_meta = (commercial_type, statbl_id, used_quarter, snap_message)

        snap_df = st.session_state.get("commercial_snapshot")
        meta = st.session_state.get("commercial_snapshot_meta")

        if snap_df is not None and meta is not None:
            commercial_type, statbl_id, used_quarter, snap_message = meta
            if snap_df.empty:
                st.info(snap_message or f"{_fmt_quarter(used_quarter)} 데이터가 없습니다. 다른 분기를 선택해보세요.")
            else:
                if used_quarter != selected_quarter:
                    st.caption(f"선택하신 분기는 아직 데이터가 없어 {_fmt_quarter(used_quarter)} 결과를 표시합니다.")

                display_df = snap_df[["CLS_FULLNM", "DTA_VAL"]].rename(
                    columns={"CLS_FULLNM": "상권", "DTA_VAL": "공실률(%)"}
                )
                display_df["공실률(%)"] = pd.to_numeric(display_df["공실률(%)"], errors="coerce").round(2)
                display_df = display_df.sort_values("공실률(%)", ascending=False).reset_index(drop=True)

                search = st.text_input("상권 검색 (예: 연남, 강남, 판교)", key="commercial_search")
                if search.strip():
                    display_df = display_df[display_df["상권"].str.contains(search.strip(), case=False, na=False)]

                st.caption(f"{commercial_type} · {_fmt_quarter(used_quarter)} · {len(display_df)}개 상권 "
                           "— 행을 클릭하면 아래에 분기별 추이가 표시됩니다.")
                event = st.dataframe(
                    display_df, width='stretch', hide_index=True,
                    on_select="rerun", selection_mode="single-row", key="commercial_table_select",
                )
                selected_rows = event.selection.get("rows", []) if event else []
                if selected_rows:
                    picked_addr = display_df.iloc[selected_rows[0]]["상권"]
                    cls_id = snap_df.loc[snap_df["CLS_FULLNM"] == picked_addr, "CLS_ID"].iloc[0]

                    try:
                        with st.spinner("추이 조회 중..."):
                            trend_df, trend_message = get_reb_vacancy_trend(
                                reb_key, statbl_id, cls_id, "202403", latest_quarter,
                            )
                    except RuntimeError as e:
                        st.error(str(e))
                        trend_df = pd.DataFrame()

                    st.subheader(f"📈 {picked_addr} 공실률 추이")
                    if trend_df.empty:
                        st.info(trend_message or "추이 데이터가 없습니다.")
                    else:
                        trend_df = trend_df.copy()
                        trend_df["공실률(%)"] = pd.to_numeric(trend_df["DTA_VAL"], errors="coerce")
                        trend_df = trend_df.sort_values("WRTTIME_IDTFR_ID")
                        x_label_col = "WRTTIME_DESC" if "WRTTIME_DESC" in trend_df.columns else "WRTTIME_IDTFR_ID"
                        if x_label_col == "WRTTIME_IDTFR_ID":
                            trend_df[x_label_col] = trend_df[x_label_col].map(_fmt_quarter)
                        chart_df = trend_df.set_index(x_label_col)[["공실률(%)"]]
                        st.line_chart(chart_df)

# ------------------------------------------------------------------
# 탭 9: 주변 상가업소 (소상공인시장진흥공단 상가(상권)정보 Open API)
# ------------------------------------------------------------------
with tab_sangkwon:
    st.write("**소상공인시장진흥공단 상가(상권)정보 Open API**로 특정 위치 반경 내 실제 점포(상가업소) 목록을 조회합니다. "
             "임장·상권분석 시 주변 업종 구성을 파악하는 데 활용할 수 있습니다. (반경 최대 2,000m)")
    if not sangkwon_key:
        st.warning("사이드바에 소상공인시장진흥공단(또는 공공데이터포털) 인증키를 입력해야 사용할 수 있습니다.")
    elif not kakao_key:
        st.warning("사이드바에 카카오맵 REST API 키를 입력해야 주소를 좌표로 변환할 수 있습니다.")
    else:
        sk_addr = st.text_input(
            "주소", placeholder="예: 서울특별시 마포구 연남동 227-1", key="sangkwon_addr",
        )

        upjong_df = load_sangkwon_upjong_codes()
        c1, c2, c3, c4 = st.columns([2, 2, 2, 1])

        lcls_pairs = upjong_df[["대분류코드", "대분류명"]].drop_duplicates().sort_values("대분류코드")
        lcls_name = c1.selectbox("업종 대분류", ["(전체)"] + lcls_pairs["대분류명"].tolist(), key="sk_lcls")
        lcls_cd = mcls_cd = scls_cd = None

        if lcls_name != "(전체)":
            lcls_cd = lcls_pairs.loc[lcls_pairs["대분류명"] == lcls_name, "대분류코드"].iloc[0]
            mcls_pairs = (
                upjong_df[upjong_df["대분류코드"] == lcls_cd][["중분류코드", "중분류명"]]
                .drop_duplicates().sort_values("중분류코드")
            )
            mcls_name = c2.selectbox("업종 중분류", ["(전체)"] + mcls_pairs["중분류명"].tolist(), key="sk_mcls")
            if mcls_name != "(전체)":
                mcls_cd = mcls_pairs.loc[mcls_pairs["중분류명"] == mcls_name, "중분류코드"].iloc[0]
                scls_pairs = (
                    upjong_df[upjong_df["중분류코드"] == mcls_cd][["소분류코드", "소분류명"]]
                    .drop_duplicates().sort_values("소분류코드")
                )
                scls_name = c3.selectbox("업종 소분류", ["(전체)"] + scls_pairs["소분류명"].tolist(), key="sk_scls")
                if scls_name != "(전체)":
                    scls_cd = scls_pairs.loc[scls_pairs["소분류명"] == scls_name, "소분류코드"].iloc[0]
            else:
                c3.selectbox("업종 소분류", ["(전체)"], key="sk_scls_disabled", disabled=True)
        else:
            c2.selectbox("업종 중분류", ["(전체)"], key="sk_mcls_disabled", disabled=True)
            c3.selectbox("업종 소분류", ["(전체)"], key="sk_scls_disabled", disabled=True)

        radius = c4.number_input("반경(m)", min_value=100, max_value=2000, value=500, step=100, key="sk_radius")

        if st.button("주변 상가업소 조회", type="primary", key="sangkwon_submit"):
            addr = sk_addr.strip()
            if not addr:
                st.warning("주소를 입력해주세요.")
            else:
                with st.spinner("주소 좌표 변환 중..."):
                    coord, reason = geocode_address_kakao(kakao_key, addr)
                if not coord:
                    st.error(f"좌표를 찾지 못했습니다 ({reason}).")
                    st.session_state.sangkwon_result = None
                else:
                    lon, lat = coord
                    try:
                        with st.spinner("주변 상가업소 조회 중..."):
                            stores_df = get_nearby_stores(
                                sangkwon_key, lon, lat, radius=int(radius),
                                inds_lcls_cd=lcls_cd, inds_mcls_cd=mcls_cd, inds_scls_cd=scls_cd,
                            )
                    except RuntimeError as e:
                        st.error(str(e))
                        stores_df = pd.DataFrame()
                    st.session_state.sangkwon_result = (addr, lat, lon, int(radius), stores_df)

        result = st.session_state.get("sangkwon_result")
        if result:
            addr, lat, lon, used_radius, stores_df = result
            if stores_df.empty:
                st.info("반경 내 조회된 상가업소가 없습니다.")
            else:
                st.success(f"'{addr}' 주변 반경 {used_radius}m 내 상가업소 {len(stores_df)}건")

                st.caption("업종 대분류별 분포")
                st.bar_chart(stores_df["indsLclsNm"].value_counts())

                display_cols = {
                    "bizesNm": "상호명", "indsLclsNm": "대분류", "indsMclsNm": "중분류", "indsSclsNm": "소분류",
                    "rdnmAdr": "도로명주소", "lnoAdr": "지번주소", "거리(m)": "거리(m)",
                }
                show_df = stores_df[[c for c in display_cols if c in stores_df.columns]].rename(columns=display_cols)
                st.dataframe(show_df, width='stretch', hide_index=True)

                csv_bytes = show_df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "📄 CSV 다운로드", csv_bytes, "nearby_stores.csv", "text/csv",
                    key="sangkwon_csv",
                )

                map_df = stores_df.copy()
                map_df["표시"] = map_df["bizesNm"].astype(str) + " (" + map_df["indsSclsNm"].astype(str) + ")"
                map_df = pd.concat([
                    pd.DataFrame({"lat": [lat], "lon": [lon], "표시": ["📍 기준 위치"]}),
                    map_df[["lat", "lon", "표시"]],
                ], ignore_index=True)
                render_address_map(map_df, label_col="표시", vworld_key=vworld_key, highlight_lat=lat, highlight_lon=lon)

# ------------------------------------------------------------------
# 탭 10: 서울 상권분석 (서울 열린데이터광장 우리마을가게 상권분석서비스 Open API)
# ------------------------------------------------------------------
with tab_seoul:
    st.write("**서울 열린데이터광장 우리마을가게 상권분석서비스**로 서울시 상권의 추정매출·점포 현황·생활인구·직장인구를 조회합니다. "
             "(서울시 소재 상권만 지원, 데이터는 분기 단위로 갱신됩니다.)")
    if not seoul_key:
        st.warning("사이드바에 서울 열린데이터광장 인증키를 입력해야 사용할 수 있습니다.")
    elif not kakao_key:
        st.warning("사이드바에 카카오맵 REST API 키를 입력해야 주소를 좌표로 변환할 수 있습니다.")
    else:
        def _fmt_seoul_quarter(q):
            return f"{q[:4]}년 {q[4:]}분기"

        st.caption("주소 (서울특별시 + 구 + 동 + 번지)")
        _seoul_sigungu_df = _load_sigungu_list()
        _seoul_gu_df = _seoul_sigungu_df[_seoul_sigungu_df["시도명"] == "서울특별시"].reset_index(drop=True)
        _seoul_gu_options = _seoul_gu_df["시군구명"].tolist()
        col_gu, col_dong, col_bun = st.columns([1, 1, 1])
        seoul_gu = col_gu.selectbox(
            "구", _seoul_gu_options,
            index=_seoul_gu_options.index(addr_gu) if addr_gu in _seoul_gu_options else 0,
            key="seoul_gu",
        )
        seoul_sigungu_code = _seoul_gu_df.loc[_seoul_gu_df["시군구명"] == seoul_gu, "시군구코드"].iloc[0]
        _seoul_dong_options = _load_dong_list(seoul_sigungu_code)
        seoul_dong = col_dong.selectbox(
            "동", _seoul_dong_options,
            index=_seoul_dong_options.index(addr_dong) if addr_dong in _seoul_dong_options else 0,
            key="seoul_dong",
        )
        seoul_bun = col_bun.text_input(
            "번지 (선택)", value="", placeholder="예: 227-1", key="seoul_bun",
            help="비워두면 동 중심 좌표로 상권을 찾습니다.",
        )

        if st.button("상권 찾기", type="primary", key="seoul_submit"):
            addr = f"서울특별시 {seoul_gu} {seoul_dong} {seoul_bun.strip()}".strip()
            with st.spinner("주소 좌표 변환 중..."):
                coord, reason = geocode_address_kakao(kakao_key, addr)
            if not coord:
                st.error(f"좌표를 찾지 못했습니다 ({reason}).")
                st.session_state.seoul_result = None
            else:
                lon, lat = coord
                try:
                    with st.spinner("서울시 상권 위치 데이터를 불러오는 중... (최초 1회, 최대 1분 정도 걸릴 수 있습니다)"):
                        locations_df = _load_seoul_trade_area_locations(seoul_key)

                    if locations_df.empty:
                        st.error("서울시 상권 위치 데이터를 가져오지 못했습니다. 인증키를 확인해주세요.")
                        st.session_state.seoul_result = None
                    else:
                        trdar_row, distance_m = find_nearest_seoul_trade_area(locations_df, lon, lat)

                        with st.spinner("추정매출·점포·생활인구·직장인구 데이터를 불러오는 중... (최초 1회, 최대 1분 정도 걸릴 수 있습니다)"):
                            selng_df, selng_q = _load_seoul_quarter_dataset(seoul_key, SEOUL_TRDAR_SALES_SERVICE)
                            stor_df, stor_q = _load_seoul_quarter_dataset(seoul_key, SEOUL_TRDAR_STORE_SERVICE)
                            flpop_df, flpop_q = _load_seoul_quarter_dataset(seoul_key, SEOUL_TRDAR_FLPOP_SERVICE)
                            wrc_df, wrc_q = _load_seoul_quarter_dataset(seoul_key, SEOUL_TRDAR_WRC_POPLTN_SERVICE)

                        st.session_state.seoul_result = (
                            addr, lon, lat, trdar_row, distance_m,
                            selng_df, selng_q, stor_df, stor_q, flpop_df, flpop_q, wrc_df, wrc_q,
                        )
                except RuntimeError as e:
                    st.error(str(e))
                    st.session_state.seoul_result = None

        result = st.session_state.get("seoul_result")
        if result:
            (addr, lon, lat, trdar_row, distance_m,
             selng_df, selng_q, stor_df, stor_q, flpop_df, flpop_q, wrc_df, wrc_q) = result

            trdar_cd = trdar_row["TRDAR_CD"]
            if distance_m > 1500:
                st.warning(f"가장 가까운 서울시 상권도 {distance_m:,.0f}m 떨어져 있습니다 — 서울시 상권분석 대상 지역이 아닐 수 있습니다.")

            st.success(
                f"'{addr}'에서 가장 가까운 상권: **{trdar_row['TRDAR_CD_NM']}** "
                f"({trdar_row['TRDAR_SE_CD_NM']} · {trdar_row['SIGNGU_CD_NM']} {trdar_row['ADSTRD_CD_NM']} · "
                f"약 {distance_m:,.0f}m)"
            )

            detail = analyze_seoul_trade_area_detail(trdar_row, selng_df, stor_df, flpop_df, wrc_df)
            s = detail["총괄"]

            c1, c2, c3, c4 = st.columns(4)
            c1.metric(f"추정매출 ({_fmt_seoul_quarter(selng_q)})",
                      f"{s['추정매출(원)'] / 1e8:,.1f}억원" if s["추정매출(원)"] else "-")
            c2.metric(f"점포 수 ({_fmt_seoul_quarter(stor_q)})",
                      f"{s['점포수']:,.0f}개" if s["점포수"] else "-")
            c3.metric(f"생활인구 ({_fmt_seoul_quarter(flpop_q)})",
                      f"{s['생활인구']:,.0f}명" if s["생활인구"] is not None else "-")
            c4.metric(f"직장인구 ({_fmt_seoul_quarter(wrc_q)})",
                      f"{s['직장인구']:,.0f}명" if s["직장인구"] is not None else "-")

            if detail["인사이트"]:
                st.markdown("**🔎 자동 인사이트**")
                for line in detail["인사이트"]:
                    st.markdown(f"- {line}")

            if detail["업종별"].empty:
                st.info(f"{_fmt_seoul_quarter(selng_q)} 추정매출 데이터가 없습니다.")
            else:
                st.subheader("📊 업종별 매출·점포 현황")
                display_cols = {
                    "SVC_INDUTY_CD_NM": "업종", "THSMON_SELNG_AMT": "매출액(원)", "THSMON_SELNG_CO": "매출건수",
                    "STOR_CO": "점포수", "OPBIZ_RT": "개업률(%)", "CLSBIZ_RT": "폐업률(%)",
                    "FRC_STOR_CO": "프랜차이즈점포수", "SIMILR_INDUTY_STOR_CO": "유사업종점포수",
                }
                merged = detail["업종별"]
                show_cols = [c for c in display_cols if c in merged.columns]
                st.dataframe(merged[show_cols].rename(columns=display_cols), width='stretch', hide_index=True)

                csv_bytes = merged[show_cols].rename(columns=display_cols).to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "📄 CSV 다운로드", csv_bytes, "seoul_trade_area_industries.csv", "text/csv",
                    key="seoul_csv",
                )

                st.subheader("🧑‍🤝‍🧑 매출 구성 — 성별 · 연령대 · 주중/주말")
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.caption("성별 매출 비중")
                    gm = detail["성별매출"]
                    if gm:
                        st.bar_chart(pd.Series({"남성": gm["남성비율(%)"], "여성": gm["여성비율(%)"]}))
                    else:
                        st.caption("데이터 없음")
                with c2:
                    st.caption("연령대별 매출 비중(%)")
                    age_df = detail["연령대별매출"]
                    if not age_df.empty:
                        st.bar_chart(age_df.set_index("연령대")["비율(%)"])
                    else:
                        st.caption("데이터 없음")
                with c3:
                    st.caption("주중/주말 매출 비중")
                    wm = detail["주중주말매출"]
                    if wm:
                        st.bar_chart(pd.Series({"주중": wm["주중비율(%)"], "주말": wm["주말비율(%)"]}))
                    else:
                        st.caption("데이터 없음")

                c1, c2 = st.columns(2)
                with c1:
                    st.caption("요일별 매출 (억원)")
                    if not detail["요일별매출"].empty:
                        st.bar_chart(detail["요일별매출"].set_index("요일")["매출액(억원)"])
                with c2:
                    st.caption("시간대별 매출 (억원)")
                    if not detail["시간대별매출"].empty:
                        st.bar_chart(detail["시간대별매출"].set_index("시간대")["매출액(억원)"])

            store = detail["점포"]
            if store:
                st.subheader("🏪 점포 개폐업 현황")
                store_cols = st.columns(4)
                labels_units = [
                    ("총점포수", "개"), ("프랜차이즈점포수", "개"), ("개업점포수", "개"), ("폐업점포수", "개"),
                ]
                for col, (label, unit) in zip(store_cols, labels_units):
                    if label in store:
                        col.metric(label, f"{store[label]:,.0f}{unit}")
                rate_cols = st.columns(2)
                if "개업률(%)" in store:
                    rate_cols[0].metric("개업률", f"{store['개업률(%)']}%")
                if "폐업률(%)" in store:
                    rate_cols[1].metric("폐업률", f"{store['폐업률(%)']}%")

            flp_stats, wrc_stats = detail["생활인구"], detail["직장인구"]
            if flp_stats or wrc_stats:
                st.subheader("👥 생활인구 · 직장인구 상세")
                c1, c2 = st.columns(2)
                with c1:
                    st.caption("생활인구")
                    if "남성비율(%)" in flp_stats:
                        st.write(f"성별: 남성 {flp_stats['남성비율(%)']}% · 여성 {flp_stats['여성비율(%)']}%")
                    if "연령대별" in flp_stats:
                        st.bar_chart(flp_stats["연령대별"].set_index("연령대")["인구수"])
                    if not flp_stats:
                        st.caption("데이터 없음")
                with c2:
                    st.caption("직장인구")
                    if "남성비율(%)" in wrc_stats:
                        st.write(f"성별: 남성 {wrc_stats['남성비율(%)']}% · 여성 {wrc_stats['여성비율(%)']}%")
                    if "연령대별" in wrc_stats:
                        st.bar_chart(wrc_stats["연령대별"].set_index("연령대")["인구수"])
                    if not wrc_stats:
                        st.caption("데이터 없음")

            map_df = pd.DataFrame({
                "lat": [trdar_row["lat"], lat],
                "lon": [trdar_row["lon"], lon],
                "표시": [f"🏙️ {trdar_row['TRDAR_CD_NM']}", "📍 입력 주소"],
            })
            render_address_map(map_df, label_col="표시", vworld_key=vworld_key, highlight_lat=lat, highlight_lon=lon)

# ------------------------------------------------------------------
# 탭 11: 자동 pptx 리포트 (주소 하나로 실제 데이터를 채운 부동산 분석 리포트 생성)
# ------------------------------------------------------------------
with tab_autopptx:
    st.write(
        "왼쪽 사이드바에 입력한 주소로 **부동산 분석 리포트(pptx)**를 자동 생성합니다. "
        "건축물대장 · 실거래가 · 공시가격 · 동단위 시장통계 · 상업용부동산 공실률 · 주변 상가업소 · "
        "(서울 소재 시) 서울 상권분석까지, 실제로 조회한 데이터로 슬라이드를 채웁니다."
    )
    st.caption(
        "💡 **종합 리포트(같은 번지) · 동단위 통계 · 노후건축물 · 서울 상권분석** 탭을 이 리포트보다 "
        "먼저 조회해두면, 그 탭이 채워둔 결과/캐시를 그대로 재사용해 해당 부분 조회 시간이 크게 "
        "줄어듭니다 (탭을 안 열어봤다면 리포트가 알아서 새로 조회하니 그냥 눌러도 됩니다)."
    )
    if st.session_state.get("ledger_docs"):
        st.caption(
            f"📎 사이드바에 업로드한 건축물대장 열람본 {len(st.session_state['ledger_docs'])}장에서 OCR로 "
            "소유자현황 · 변동사항을 추출해 리포트 끝에 슬라이드로 포함되고, 위반건축물 표기가 감지되면 "
            "건축물 개요에도 반영됩니다."
        )
    st.caption(
        "※ 상권 성격 · SNS 트렌드 · 개발호재 · SWOT 같은 정성적 항목은 공공데이터로 자동 수집되지 않습니다. "
        "아래에 직접 입력하면 해당 내용 그대로 슬라이드에 반영되고, 비워두면 그 슬라이드는 생략됩니다."
    )

    with st.expander("📈 개발호재 종합 (선택, 직접입력)", expanded=False):
        growth_drivers_text = st.text_area(
            "한 줄에 하나씩 입력하세요",
            value="", height=120, key="autopptx_growth_drivers",
            placeholder="예: 대장홍대선 2025.12 착공, 2031년 준공목표\n마포구 지구단위계획 재정비 (2026년 추진)",
        )

    with st.expander("🔍 SWOT 분석 (선택, 직접입력)", expanded=False):
        swot_col1, swot_col2 = st.columns(2)
        swot_strengths = swot_col1.text_area("S — 강점", value="", height=100, key="autopptx_swot_s")
        swot_weaknesses = swot_col2.text_area("W — 약점", value="", height=100, key="autopptx_swot_w")
        swot_opportunities = swot_col1.text_area("O — 기회", value="", height=100, key="autopptx_swot_o")
        swot_threats = swot_col2.text_area("T — 위협", value="", height=100, key="autopptx_swot_t")

    if not service_key:
        st.warning("사이드바에 공공데이터포털 서비스키를 입력해야 사용할 수 있습니다.")
    elif not kakao_key:
        st.warning("사이드바에 카카오맵 REST API 키를 입력해야 위치 지도를 생성할 수 있습니다.")
    else:
        if st.button("📑 리포트 생성", type="primary", key="autopptx_submit"):
            try:
                with st.spinner("주소를 코드로 변환하는 중..."):
                    sigungu_code, bdong_code, addr_sido, addr_sigungu_name = _resolve_codes()

                progress_box = st.empty()

                def _on_report_progress(msg):
                    progress_box.info(msg)

                # "종합 리포트" 탭이 정확히 같은 번지(sigungu_code/bdong_code/bun/ji)로
                # 이미 조회해둔 결과가 있으면 표제부·주택가격·지역지구구역 재조회를 건너뛴다.
                current_address_key = (
                    sigungu_code, bdong_code, bun or None, ji if ji and ji != "0" else None,
                )
                reusable_report = None
                if st.session_state.get("report_address_key") == current_address_key:
                    reusable_report = st.session_state.get("report")

                with st.spinner("리포트 데이터 수집 중... (최초 조회 시 1~2분 정도 걸릴 수 있습니다)"):
                    report_data = fetch_report_data(
                        service_key=service_key,
                        sido=addr_sido, sigungu_name=addr_sigungu_name, dong_name=addr_dong,
                        sigungu_code=sigungu_code, bdong_code=bdong_code,
                        bun=bun or None, ji=ji if ji and ji != "0" else None,
                        kakao_key=kakao_key, vworld_key=vworld_key, reb_key=reb_key,
                        sangkwon_key=sangkwon_key, seoul_key=seoul_key,
                        growth_drivers_text=growth_drivers_text,
                        swot_text={
                            "strengths": swot_strengths, "weaknesses": swot_weaknesses,
                            "opportunities": swot_opportunities, "threats": swot_threats,
                        },
                        ledger_docs=st.session_state.get("ledger_docs"),
                        progress_callback=_on_report_progress,
                        # 동단위통계·노후건축물·공시가격시계열·서울상권분석 탭이 쓰는 것과
                        # 같은 @st.cache_data 래퍼를 넘긴다 — 그 탭들을 이 리포트보다 먼저
                        # 조회해뒀다면 캐시를 맞아 API 재호출 없이 즉시 재사용된다.
                        district_title_loader=_load_district_titles,
                        district_price_loader=_load_district_prices,
                        seoul_locations_loader=_load_seoul_trade_area_locations,
                        seoul_quarter_loader=_load_seoul_quarter_dataset,
                        single_report=reusable_report,
                    )
                progress_box.empty()

                with st.spinner("pptx 파일 조립 중..."):
                    pptx_bytes = generate_pptx(report_data)

                st.session_state.autopptx_result = (report_data["address"], pptx_bytes)
                st.session_state.autopptx_notes = report_data.get("notes") or []
                st.success(f"'{report_data['address']}' 리포트를 생성했습니다.")
            except Exception as e:
                st.error(f"리포트 생성 실패: {e}")
                st.session_state.autopptx_result = None

    result = st.session_state.get("autopptx_result")
    if result:
        addr_label, pptx_bytes = result
        st.caption(f"'{addr_label}' 리포트 ({len(pptx_bytes) / 1024:.0f} KB)")

        notes = st.session_state.get("autopptx_notes") or []
        if notes:
            with st.expander(f"⚠️ {len(notes)}개 섹션이 비어 있습니다 — 이유 보기", expanded=False):
                for n in notes:
                    st.markdown(f"- {n}")

        # 한글 파일명 + st.download_button -> Content-Disposition이 깨져서 UUID로 저장됨 (수정1).
        # data: URI -> 파일 크기가 브라우저 URL 길이 제한(~2MB)을 넘겨서 링크 자체가 깨짐 (수정2).
        # Blob URL(components.html 안에서 생성) -> 로컬(streamlit run)에서는 완벽히 되는데 실제
        # Streamlit Community Cloud에서는 여전히 UUID로 저장됨 — 아마 컴포넌트 iframe 안에서
        # a.click()으로 트리거하는 다운로드를 Cloud 프록시나 특정 브라우저가 다르게 처리하는 듯
        # (수정3, 원인 미확정). 커스텀 JS를 걷어내고 이 앱의 다른 모든 CSV/엑셀 버튼과 동일하게
        # st.download_button + ASCII 파일명으로 되돌린다 — 가장 단순하고, 이미 이 앱에서 검증된
        # 조합이다 (수정4).
        file_name = f"realestate_analysis_report_{datetime.datetime.now():%Y%m%d_%H%M%S}.pptx"
        st.download_button(
            "📥 pptx 다운로드", pptx_bytes, file_name,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary", width='stretch', key="autopptx_download",
        )
