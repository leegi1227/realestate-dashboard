"""
국토교통부 건축물대장정보 조회 대시보드 (Streamlit)
====================================================

실행 방법
---------
streamlit run dashboard.py

브라우저에서 자동으로 열리며, 안 열리면 터미널에 표시되는
http://localhost:8501 주소를 직접 열면 됩니다.
"""

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
    analyze_price_history,
    analyze_seismic_risk,
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
    resolve_dong_code,
    reverse_match_transactions,
    split_common_and_varying,
)


# 카카오맵 JS SDK를 삽입하는 Custom Components v2 컴포넌트 (기존 pydeck 렌더링을 대체).
# components.v1.declare_component(구 iframe+postMessage 프로토콜)는 이 Streamlit 버전에서
# 최초 마운트 시 컴포넌트가 준비 신호를 보내도 렌더 이벤트를 받지 못하고 조용히 멈추는
# 문제가 있었다(재현 확인 — 공식 트러블슈팅 문서에도 "가장 흔한 v1 실패 증상"으로 명시돼
# 있음). v2(`st.components.v2.component`)는 iframe 없이 같은 페이지에 직접(섀도우 DOM)
# 마운트되고 핸드셰이크가 필요 없어 이 문제가 없다.
_kakao_map_dir = os.path.join(os.path.dirname(__file__), "kakao_map_component")
with open(os.path.join(_kakao_map_dir, "template.html"), encoding="utf-8") as _f:
    _kakao_map_html = _f.read()
with open(os.path.join(_kakao_map_dir, "style.css"), encoding="utf-8") as _f:
    _kakao_map_css = _f.read()
with open(os.path.join(_kakao_map_dir, "script.js"), encoding="utf-8") as _f:
    _kakao_map_js = _f.read()

_kakao_map_component = st.components.v2.component(
    "kakao_map", html=_kakao_map_html, css=_kakao_map_css, js=_kakao_map_js,
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
    kakao_js_key: str = None,
    enable_selection: bool = False,
    selection_key: str = None,
    highlight_lat: float = None,
    highlight_lon: float = None,
):
    """지점을 구글 지도 스타일 핀으로 카카오맵 위에 표시하고, show_labels=True면 핀 옆에

    주소도 표시한다. show_labels=False면 마커만 표시해 지점이 아주 많을 때 더 깔끔하게
    볼 수 있다. kakao_js_key가 없으면 지도를 그리지 않고 안내 문구만 보여준다.

    실제 렌더링은 kakao_map_component/index.html(카카오맵 JS SDK를 삽입한 커스텀
    Streamlit 컴포넌트)이 담당한다 — 이 함수는 파이썬 쪽에서 라벨 중복 제거, 겹침
    방지용 세로 스택 단계 계산까지만 하고 나머지는 컴포넌트에 데이터로 넘긴다.

    enable_selection=True면 마커 클릭을 감지해서, 클릭된 지점의 데이터(주소/lat/lon이
    담긴 dict)를 반환한다(클릭이 없으면 None).

    highlight_lat/highlight_lon을 넘기면(예: 표에서 행을 선택했을 때) 그 좌표에 노란
    테두리 원(halo)을 표시해 "이 마커가 지금 선택된 지점"임을 시각적으로 나타낸다.
    표 → 지도 방향 강조(이 옵션)와 지도 → 표 방향 강조(enable_selection)는 서로
    독립적으로 켤 수 있다.
    """
    if not kakao_js_key:
        st.caption("카카오맵 JavaScript 키를 입력하면 지도를 표시합니다.")
        return None

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

    result = _kakao_map_component(
        key=selection_key,
        data={
            "kakaoJsKey": kakao_js_key,
            "markers": markers,
            "highlight": highlight,
            "showLabels": show_labels,
            "enableSelection": enable_selection,
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
    """동 전체 표제부를 15분 캐시. 동단위 통계/노후건축물/내진 스캔 탭이 공유한다."""
    api = BuildingLedger(service_key)
    return get_building_ledger(
        api,
        ledger_type="표제부",
        sigungu_code=sigungu_code,
        bdong_code=bdong_code,
        max_rows=10000,
        wait_time=0.15,
    )

st.set_page_config(page_title="건축물대장 조회", page_icon="🏢", layout="wide")

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

    kakao_js_key = st.text_input(
        "카카오맵 JavaScript 키 (선택)",
        value="",
        type="password",
        placeholder="지도 화면 표시에 사용 (REST API 키와 다른 값)",
        help="https://developers.kakao.com 앱 설정의 'JavaScript 키'. 위 REST API 키와는 "
             "다른 값이며, 이 키를 쓰려면 카카오 개발자 콘솔의 '플랫폼 > Web'에 이 앱이 "
             "배포된 도메인을 등록해야 지도가 표시됩니다.",
    )
    kakao_js_key = kakao_js_key.strip() if kakao_js_key else None

    reb_key = st.text_input(
        "한국부동산원 인증키 (선택)",
        value="",
        type="password",
        placeholder="상업용부동산 공실률 조회에 사용",
        help="https://www.reb.or.kr/r-one (R-ONE 부동산통계정보시스템)에 로그인 후 "
             "'Open API > 인증키 발급내역'에서 발급. '🏬 상업용부동산 공실률' 탭에서만 쓰입니다.",
    )
    reb_key = reb_key.strip() if reb_key else None

    address = st.text_input("주소 (시/군/구 + 동)", value="성남시 분당구 백현동",
                             help="코드를 몰라도 동 이름으로 자동 검색됩니다.")

    col1, col2 = st.columns(2)
    bun = col1.text_input("번(본번)", value="237")
    ji = col2.text_input("지(부번)", value="0", help="0이면 생략됩니다.")

    st.caption("※ 두 탭(단일 조회 / 종합 리포트)이 위 조건을 공통으로 사용합니다.")


def _resolve_codes():
    sigungu_code, bdong_code, row = resolve_dong_code(address)
    st.success(
        f"인식된 주소: {row['시도명']} {row['시군구명']} {row['동명']} "
        f"(시군구코드={sigungu_code}, 법정동코드={bdong_code})"
    )
    return sigungu_code, bdong_code, row["시도명"], row["시군구명"]


tab_single, tab_report, tab_price, tab_district, tab_old, tab_seismic, tab_priceh, tab_map, tab_geocode, tab_commercial = st.tabs([
    "🔍 단일 조회", "📋 종합 리포트", "💰 실거래가",
    "📊 동단위 통계", "🏚️ 노후건축물", "🧱 내진 취약 스캔", "💹 공시가격 시계열", "🗺️ 지도 업로드",
    "📍 지오코딩", "🏬 상업용부동산 공실률",
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
            except Exception as e:
                st.error(f"리포트 생성 실패: {e}")
                st.session_state.report = None

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

        col_pdf, col_xlsx = st.columns(2)
        try:
            pdf_bytes = generate_pdf_report(report, address_label=addr_label)
            col_pdf.download_button(
                "📄 종합 리포트 PDF 다운로드 (한 문서로 이어짐)",
                pdf_bytes,
                "building_full_report.pdf",
                "application/pdf",
                type="primary",
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
                    "건축물대장 열람 등 별도 인증 경로를 이용해야 합니다."
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

                if tp_with_price and not tp_df.empty:
                    ledger_api = BuildingLedger(service_key)
                    price_progress = st.progress(0.0, text="시가표준액 조회 준비 중...")

                    def _on_price_progress(i, total):
                        price_progress.progress(i / total, text=f"시가표준액 조회 중... ({i}/{total} 지번)")

                    tp_df = add_standard_price_column(
                        tp_df, ledger_api, tp_sigungu, progress_callback=_on_price_progress,
                    )
                    price_progress.empty()

                if tp_with_reverse_match and tp_dong != "(전체)" and not tp_df.empty:
                    has_masked = tp_df["지번"].astype(str).str.contains(r"\*").any() if "지번" in tp_df.columns else False
                    if has_masked:
                        ledger_api = BuildingLedger(service_key)
                        bdong_map = get_bdong_code_map(tp_sigungu)
                        bdong_code = bdong_map.get(tp_dong)
                        if bdong_code:
                            with st.spinner(f"'{tp_dong}' 표제부 전체 수집 중... (역매칭용, 캐시되어 있으면 즉시 완료)"):
                                dong_title_df = _load_district_titles(service_key, tp_sigungu, bdong_code)
                            tp_df = reverse_match_transactions(tp_df, dong_title_df)
                        else:
                            st.warning(f"'{tp_dong}'의 법정동코드를 찾지 못해 역매칭을 건너뜁니다.")
                    else:
                        st.caption("마스킹된 지번이 없어 역매칭을 건너뜁니다.")

                if tp_with_coords and kakao_key and not tp_df.empty and "주소" in tp_df.columns:
                    coord_progress = st.progress(0.0, text="좌표 조회 준비 중...")

                    def _on_coord_progress(i, total):
                        coord_progress.progress(i / total, text=f"주소 좌표 조회 중... ({i}/{total} 고유 주소)")

                    tp_df = add_coordinates_column(
                        tp_df, kakao_key, progress_callback=_on_coord_progress,
                    )
                    coord_progress.empty()

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
                render_address_map(map_df, lat_col="위도", lon_col="경도", label_col="주소", kakao_js_key=kakao_js_key)

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
# 탭 6: 내진 취약 스캔
# ------------------------------------------------------------------
with tab_seismic:
    st.caption(
        "`내진 설계 적용 여부` 필드가 명시돼 있으면 그 값을 쓰고, 없으면 사용승인연도 기준 "
        "내진설계 의무화 연혁(1988 도입 → 2017.12 전면의무화)으로 미적용을 추정합니다. "
        "추정치이며 실제 구조계산서 확인 전 참고용입니다. (동단위 통계와 데이터 공유·캐시)"
    )
    if st.button("내진 취약 스캔", type="primary", key="seismic_submit"):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            try:
                with st.spinner("주소를 코드로 변환하는 중..."):
                    sigungu_code, bdong_code, addr_sido, addr_sigungu_name = _resolve_codes()
                with st.spinner("표제부 전체 수집 중... (캐시되어 있으면 즉시 완료)"):
                    title_df = _load_district_titles(service_key, sigungu_code, bdong_code)
                st.session_state.seismic_result = analyze_seismic_risk(title_df, top_n=50)
            except Exception as e:
                st.error(f"조회 실패: {e}")
                st.session_state.seismic_result = None

    seismic = st.session_state.get("seismic_result")
    if not seismic or seismic["분류별집계"].empty:
        st.info("**내진 취약 스캔** 버튼을 눌러주세요.")
    else:
        st.subheader("분류별 집계")
        st.dataframe(seismic["분류별집계"], width='stretch')
        st.subheader("취약 우선순위 목록 (상위 50)")
        st.dataframe(seismic["취약우선목록"], width='stretch')

# ------------------------------------------------------------------
# 탭 7: 공시가격 시계열
# ------------------------------------------------------------------
with tab_priceh:
    st.caption("사이드바의 번지(bun/ji)를 기준으로, 그 필지의 호(관리건축물대장PK)별 공시가격 추이를 봅니다.")
    if st.button("공시가격 시계열 조회", type="primary", key="priceh_submit"):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            api = BuildingLedger(service_key)
            try:
                with st.spinner("주소를 코드로 변환하는 중..."):
                    sigungu_code, bdong_code, addr_sido, addr_sigungu_name = _resolve_codes()
                with st.spinner("주택가격 조회 중..."):
                    price_df = get_building_ledger(
                        api, ledger_type="주택가격", sigungu_code=sigungu_code, bdong_code=bdong_code,
                        bun=bun or None, ji=ji if ji and ji != "0" else None,
                        max_rows=5000, wait_time=0.3,
                    )
                st.session_state.price_history = analyze_price_history(price_df, top_units=10)
            except Exception as e:
                st.error(f"조회 실패: {e}")
                st.session_state.price_history = None

    ph = st.session_state.get("price_history")
    if not ph:
        st.info("**공시가격 시계열 조회** 버튼을 눌러주세요.")
    elif not ph["단위목록"]:
        st.warning("이 번지는 공시가격(주택가격) 데이터가 없습니다.")
    else:
        st.warning(ph["경고"])
        for unit in ph["단위목록"]:
            pk_short = str(unit["관리건축물대장PK"])[-6:]
            st.markdown(
                f"**호 PK…{pk_short}** · 최신 {unit['최신가격']/1e8:.2f}억({unit['최신연도']}) · "
                f"최초 {unit['최초가격']/1e8:.2f}억({unit['최초연도']}) · "
                f"총증감 {unit['총증감률(%)']}% · 연평균(CAGR) {unit['연평균상승률CAGR(%)']}%"
            )
            st.line_chart(unit["추이"].set_index("연도"))

# ------------------------------------------------------------------
# 탭 8: 지도 업로드 — 주소·위도·경도가 담긴 파일을 올리면 지도에 표시
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
                    show_labels=show_labels, kakao_js_key=kakao_js_key,
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
# 탭 9: 지오코딩 — 주소만 넣으면 경도·위도만 반환
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
                    label_col="주소", kakao_js_key=kakao_js_key,
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
                render_address_map(map_df, lat_col="위도", lon_col="경도", label_col="주소", kakao_js_key=kakao_js_key)

# ------------------------------------------------------------------
# 탭 10: 상업용부동산 공실률 (한국부동산원 R-ONE Open API)
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
