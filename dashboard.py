"""
국토교통부 건축물대장정보 조회 대시보드 (Streamlit)
====================================================

실행 방법
---------
streamlit run dashboard.py

브라우저에서 자동으로 열리며, 안 열리면 터미널에 표시되는
http://localhost:8501 주소를 직접 열면 됩니다.
"""

import base64
import io
import string

import pandas as pd
import pydeck as pdk
import streamlit as st
from PublicDataReader import BuildingLedger, TransactionPrice

from building_example import (
    ALL_LEDGER_TYPES,
    add_address_column,
    add_pyeong_columns,
    add_standard_price_column,
    analyze_district_stats,
    analyze_old_buildings,
    analyze_price_history,
    analyze_seismic_risk,
    build_executive_summary,
    build_master_report,
    combine_zoning_sources,
    generate_master_pdf_report,
    generate_pdf_report,
    get_bdong_code_map,
    get_building_ledger,
    get_dong_list,
    get_full_building_report,
    get_sigungu_list,
    resolve_dong_code,
    reverse_match_transactions,
    split_common_and_varying,
)


_PIN_SVG = """<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">
  <path d="M32 2C19 2 8 13 8 26c0 18 24 36 24 36s24-18 24-36C56 13 45 2 32 2z" fill="#DC2626" stroke="#FFFFFF" stroke-width="2"/>
  <circle cx="32" cy="26" r="9" fill="#FFFFFF"/>
</svg>"""
_PIN_ICON_URL = "data:image/svg+xml;base64," + base64.b64encode(_PIN_SVG.encode("utf-8")).decode("ascii")
# IconLayer는 각 행이 아이콘 정의(url/width/height/anchorY)를 담은 "icon_data" 컬럼을
# 참조하는 방식이 pydeck 공식 예제 패턴이다. anchorY를 아이콘 높이와 같게 두면
# 구글 지도 핀처럼 뾰족한 끝부분이 정확한 좌표를 가리키게 된다.
_PIN_ICON_DATA = {"url": _PIN_ICON_URL, "width": 64, "height": 64, "anchorY": 64}


def render_address_map(df: pd.DataFrame, lat_col: str = "lat", lon_col: str = "lon", label_col: str = None):
    """지점을 구글 지도 스타일 핀으로 표시하고, 핀 옆에 주소를 항상 표시하는 pydeck 지도.

    참고: 이전에는 radius_units="pixels"처럼 리터럴 문자열을 따옴표 없이 넘겨서
    pydeck이 이를 "@@=pixels"라는 (정의되지 않은 변수를 참조하는) JS 표현식으로
    잘못 직렬화하는 버그 때문에 아이콘/텍스트 레이어가 모두 깨졌었다. 리터럴
    문자열 값은 파이썬 문자열 안에 따옴표를 한 번 더 감싸서(예: '"start"')
    넘겨야 pydeck이 accessor가 아닌 고정값으로 취급한다.
    """
    plot_df = df.copy()
    if label_col and label_col in plot_df.columns:
        plot_df["주소"] = plot_df[label_col].astype(str)
    else:
        plot_df["주소"] = "(주소 없음)"
    plot_df["icon_data"] = [_PIN_ICON_DATA] * len(plot_df)

    view_state = pdk.ViewState(
        latitude=float(plot_df[lat_col].mean()),
        longitude=float(plot_df[lon_col].mean()),
        zoom=16 if len(plot_df) <= 1 else 15,
    )
    icon_layer = pdk.Layer(
        "IconLayer",
        data=plot_df,
        get_icon="icon_data",
        get_position=f"[{lon_col}, {lat_col}]",
        get_size=4,
        size_scale=10,
        pickable=True,
    )
    # TextLayer는 기본적으로 아스키 문자만 폰트 아틀라스에 포함시켜서, 한글처럼
    # 기본 문자셋 밖의 글자는 조용히 안 그려진다(주소 뒷자리 번지 숫자만 보이던 원인).
    # 실제 라벨에 쓰이는 문자를 모아 character_set으로 명시해야 한글이 제대로 나온다.
    # 문자열 하나로(리스트가 아니라) 넘겨야 pydeck이 컬럼 접근자로 착각하지 않는다.
    _address_chars = "".join(sorted(set("".join(plot_df["주소"].astype(str))))) + string.printable
    text_layer = pdk.Layer(
        "TextLayer",
        data=plot_df,
        get_position=f"[{lon_col}, {lat_col}]",
        get_text="주소",
        get_size=13,
        get_color=[30, 30, 30, 255],
        get_pixel_offset=[14, -28],
        get_text_anchor='"start"',
        get_alignment_baseline='"center"',
        character_set='"' + _address_chars.replace('"', "") + '"',
        background=True,
        get_background_color=[255, 255, 255, 220],
        pickable=False,
    )
    tooltip = {"html": "<b>{주소}</b>", "style": {"backgroundColor": "white", "color": "black"}}
    st.pydeck_chart(pdk.Deck(
        layers=[icon_layer, text_layer],
        initial_view_state=view_state,
        tooltip=tooltip,
        map_style=None,  # Streamlit 테마 기본 지도 스타일 사용 (Carto/Mapbox 키 불필요)
    ))


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

    vworld_key = st.text_input(
        "브이월드(V-World) 인증키 (선택)",
        value="",
        type="password",
        placeholder="용도지역/지구 상세 정보를 원하면 입력 (없어도 나머지 기능은 정상 동작)",
        help="https://www.vworld.kr 에서 발급. 통합 리포트의 용도지역/지구 상세 조회에만 쓰입니다.",
    )
    vworld_key = vworld_key.strip() if vworld_key else None

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


tab_master, tab_single, tab_report, tab_price, tab_district, tab_old, tab_seismic, tab_priceh, tab_map = st.tabs([
    "🏆 통합 리포트", "🔍 단일 조회", "📋 종합 리포트", "💰 실거래가",
    "📊 동단위 통계", "🏚️ 노후건축물", "🧱 내진 취약 스캔", "💹 공시가격 시계열", "🗺️ 지도 업로드",
])

# ------------------------------------------------------------------
# 탭 0: 통합 리포트 — 단일조회+실거래가+동단위통계+노후건축물+내진+공시가격을 한 번에
# ------------------------------------------------------------------
with tab_master:
    st.write("사이드바에 입력한 주소·번지 하나로, 단일조회·실거래가·노후도·내진·공시가격을 한 번에 모읍니다.")
    months_lookback = st.slider("실거래가 조회 기간 (개월)", 3, 36, 12, key="master_months")
    with_district = st.checkbox(
        "동네 평균과 비교하기 (동 전체를 받아야 해서 최초 조회 시 20~40초 걸릴 수 있음)",
        value=False, key="master_with_district",
    )

    if st.button("통합 리포트 생성", type="primary", key="master_submit"):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            api = BuildingLedger(service_key)
            tp_api = TransactionPrice(service_key)
            try:
                with st.spinner("주소를 코드로 변환하는 중..."):
                    sigungu_code, bdong_code, addr_sido, addr_sigungu_name = _resolve_codes()

                district_title_df = None
                if with_district:
                    with st.spinner("동 전체 표제부 수집 중... (캐시되어 있으면 즉시 완료)"):
                        district_title_df = _load_district_titles(service_key, sigungu_code, bdong_code)

                with st.spinner("단일조회·실거래가·노후도·내진·공시가격 종합 중..."):
                    master = build_master_report(
                        api, tp_api, sigungu_code, bdong_code,
                        bun=bun or None, ji=ji if ji and ji != "0" else None,
                        months_lookback=months_lookback,
                        district_title_df=district_title_df,
                        sido=addr_sido, sigungu_name=addr_sigungu_name,
                        vworld_key=vworld_key,
                    )
                st.session_state.master = master
                st.session_state.master_address_label = address
            except Exception as e:
                st.error(f"리포트 생성 실패: {e}")
                st.session_state.master = None

    master = st.session_state.get("master")
    if not master:
        st.info("**통합 리포트 생성** 버튼을 눌러주세요.")
    else:
        title_df = master.get("표제부")
        core = title_df.iloc[0] if title_df is not None and not title_df.empty else None

        st.subheader("📝 핵심 요약")
        st.info(build_executive_summary(master))

        st.divider()
        if core is not None:
            st.subheader("① 단일 조회 — 핵심 정보")
            addr = core.get("도로명대지위치") or core.get("대지위치", "")
            st.markdown(f"**{addr}** · {core.get('건물명', '') or '(건물명 없음)'}")
            zoning = combine_zoning_sources(master)
            if zoning:
                st.markdown(f"**용도지역/지구**: {', '.join(zoning)}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("주용도", core.get("주용도코드명", "-"))
            c2.metric("구조", core.get("구조코드명", "-"))
            c3.metric("지상/지하층수", f"{core.get('지상층수', '-')} / {core.get('지하층수', '-')}")
            c4.metric("사용승인일", str(core.get("사용승인일", "-")))
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("대지면적(㎡)", core.get("대지면적", "-"))
            c6.metric("연면적(㎡)", core.get("연면적", "-"))
            c7.metric("건폐율(%)", core.get("건폐율", "-"))
            c8.metric("용적률(%)", core.get("용적률", "-"))

            coord = master.get("좌표")
            if coord:
                render_address_map(
                    pd.DataFrame({"lat": [coord[1]], "lon": [coord[0]], "주소": [addr]}),
                    label_col="주소",
                )
            elif vworld_key:
                st.caption("좌표를 확인하지 못해 지도를 표시할 수 없습니다.")
            else:
                st.caption("브이월드 인증키를 입력하면 이 위치를 지도에 표시합니다.")
        else:
            st.warning("표제부 조회 결과가 없습니다 — 주소/번지를 확인해주세요.")

        st.divider()
        st.subheader("② 실거래가")
        tx = master.get("실거래가") or {}
        tx_df = tx.get("df")
        st.caption(f"추정 부동산 유형: {master.get('추정부동산유형') or '판별 불가'}")
        st.write(tx.get("note", ""))
        if tx_df is not None and not tx_df.empty and "거래금액" in tx_df.columns:
            if tx.get("status") != "matched":
                prices = pd.to_numeric(tx_df["거래금액"], errors="coerce").dropna()
                if not prices.empty:
                    p1, p2, p3, p4 = st.columns(4)
                    p1.metric("참고 거래건수", f"{len(prices):,}")
                    p2.metric("평균가(억)", f"{prices.mean()/1e4:.1f}")
                    p3.metric("중앙값(억)", f"{prices.median()/1e4:.1f}")
                    p4.metric("최저~최고(억)", f"{prices.min()/1e4:.1f} ~ {prices.max()/1e4:.1f}")
                st.caption("아래는 이 번지와 정확히 일치하지 않는, 인근 참고 거래 목록입니다.")
                st.dataframe(tx_df.head(20), width='stretch')
            else:
                st.dataframe(tx_df, width='stretch')

        st.divider()
        st.subheader("③ 노후도 · 내진")
        old_df = master.get("노후도")
        seismic = master.get("내진분석") or {}
        seismic_list = seismic.get("취약우선목록")
        c5, c6 = st.columns(2)
        if old_df is not None and not old_df.empty and "경과연수" in old_df.columns:
            c5.metric("경과연수", f"{old_df.iloc[0]['경과연수']:.0f}년")
        else:
            c5.metric("경과연수", "-")
        if seismic_list is not None and not seismic_list.empty:
            c6.metric("내진 분류", seismic_list.iloc[0].get("내진분류", "-"))
        else:
            c6.metric("내진 분류", "-")

        st.divider()
        st.subheader("④ 공시가격(시가표준액) 시계열")
        ph = master.get("공시가격") or {}
        units = ph.get("단위목록") or []
        if not units:
            st.info("이 번지의 공시가격(주택가격) 데이터가 없습니다.")
        else:
            st.warning(ph.get("경고", ""))
            summary_df = pd.DataFrame([
                {
                    "호(PK)": str(u["관리건축물대장PK"])[-6:],
                    "최초연도": u["최초연도"], "최초가격(억)": round(u["최초가격"] / 1e8, 2),
                    "최신연도": u["최신연도"], "최신가격(억)": round(u["최신가격"] / 1e8, 2),
                    "총증감(%)": u["총증감률(%)"], "CAGR(%)": u["연평균상승률CAGR(%)"],
                }
                for u in units
            ])
            st.dataframe(summary_df, width='stretch')
            st.caption(f"최고가 호(…{str(units[0]['관리건축물대장PK'])[-6:]}) 연도별 추이")
            st.line_chart(units[0]["추이"].set_index("연도"))

        district = master.get("동단위통계")
        if district and district.get("총괄"):
            st.divider()
            st.subheader("⑤ 동단위 통계 비교")
            s = district["총괄"]
            d1, d2, d3, d4 = st.columns(4)
            d1.metric("동 전체 총동수", f"{s['총동수']:,}")
            d2.metric("동 총연면적(㎡)", f"{s['총연면적(㎡)']:,.0f}")
            d3.metric("동 평균 층수", s["평균층수"])
            d4.metric("동 평균 경과연수", f"{s['평균경과연수']}년")
            if old_df is not None and not old_df.empty and "경과연수" in old_df.columns and s.get("평균경과연수"):
                diff = old_df.iloc[0]["경과연수"] - s["평균경과연수"]
                st.caption(f"이 건물은 동 평균보다 {abs(diff):.1f}년 {'더 오래됨' if diff > 0 else '더 신축'}")

            col_p, col_a = st.columns(2)
            if district.get("주용도별") is not None and not district["주용도별"].empty:
                col_p.markdown("**주용도별 분포 (상위 5)**")
                col_p.dataframe(district["주용도별"].head(5), width='stretch')
            if district.get("노후도분포") is not None and not district["노후도분포"].empty:
                col_a.markdown("**노후도 분포**")
                col_a.dataframe(district["노후도분포"], width='stretch')

        st.divider()
        try:
            pdf_bytes = generate_master_pdf_report(
                master, address_label=st.session_state.get("master_address_label", "")
            )
            st.download_button(
                "📄 통합 리포트 PDF 다운로드", pdf_bytes, "master_report.pdf", "application/pdf",
                type="primary", key="master_pdf",
            )
        except Exception as e:
            st.error(f"PDF 생성 실패: {e}")

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
                render_address_map(plot_df, label_col="주소" if addr_col else None)

                st.subheader("업로드한 데이터")
                st.dataframe(add_pyeong_columns(map_df), width='stretch')
    else:
        st.info("파일을 업로드해주세요.")
