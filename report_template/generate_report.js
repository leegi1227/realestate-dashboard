// 부동산 분석 리포트 pptx 템플릿 생성기
//
// 이 스크립트는 대시보드(dashboard.py)가 나중에 자동화로 채워 넣을 리포트의
// "생성 로직" 자체다 — 표/차트 행 수가 주소마다 달라지므로 정적 pptx에
// 텍스트만 찾아 바꾸는 방식 대신, 데이터 객체(REPORT_DATA)를 받아 매번 새로
// 슬라이드를 그리는 구조로 만들었다. 지금은 REPORT_DATA에 예시(mock) 데이터를
// 채워 레이아웃을 확인하고, 실제 자동화 단계에서는 이 REPORT_DATA를 만드는
// 부분만 dashboard.py 쪽 함수 호출로 교체하면 된다.
//
// 실행: NODE_PATH="$(npm root -g)" node generate_report.js

const pptxgen = require("pptxgenjs");

// ------------------------------------------------------------------
// 색상 팔레트 — 딥 네이비 베이스 + 비비드(인디고→핑크) 포인트
// (report_generator.py 라이브 리포트와 통일. pptxgenjs는 그라데이션 채우기를
// 지원하지 않아 여기서는 포인트 컬러를 그 그라데이션의 중간톤인 단색 퍼플로 둔다.)
// ------------------------------------------------------------------
const NAVY = "16213E";
const NAVY_DARK = "0B1220";
const NAVY_LIGHT = "2A395C";
const ICE = "E8ECF2";
const TERRACOTTA = "A855F7";
const TERRACOTTA_LIGHT = "E8D9FA";
const WHITE = "FFFFFF";
const TEXT_DARK = "22262E";
const MUTED = "6B7280";
const CARD_BG = "F5F6F8";
const GRID_LINE = "DCE1E8";

// 맑은 고딕을 직접 지정해봤으나(가독성 개선 시도), 이 환경에서는 "굵게"를 준 한글이
// 깨진 장식체로 렌더링되는 버그가 있었다(실제 PowerPoint에서도 재현될 수 있어 위험 —
// 정상체는 멀쩡한데 bold만 깨지는 걸 별도 테스트로 확인). Cambria/Calibri 조합(세리프+산세리프
// 혼용)도 제목·본문 글꼴이 달라 통일감이 떨어졌던 문제가 있어, 굵게/일반 모두 깨끗하게
// 렌더링되는 것을 확인한 Calibri 하나로 통일했다(한글은 OS 기본 대체 글꼴로 표시되며,
// Windows에서는 이것이 사실상 맑은 고딕과 동일하다). 제목/본문 구분은 크기·굵기로만 준다.
const FONT_HEAD = "Calibri";
const FONT_BODY = "Calibri";

// ------------------------------------------------------------------
// 예시 데이터 (실제 자동화에서는 아래 구조를 그대로 두고 값만 dashboard.py의
// 함수 호출 결과로 채우면 된다 — 이 객체가 곧 "이 템플릿이 기대하는 스키마")
// ------------------------------------------------------------------
const REPORT_DATA = {
  address: "서울특별시 마포구 연남동 227-1",
  reportDate: "2026년 7월 30일 기준",
  isSeoul: true,

  summary: {
    text: "대상 건물은 1992년 사용승인된 제2종근린생활시설로, 준공 후 33년이 경과한 노후 건물이며 " +
      "내진설계는 적용되지 않았습니다. 최근 6개월 내 인근 실거래가는 15.2억원 수준이며, 반경 500m 내 " +
      "312개 상가업소가 밀집한 활발한 상권(홍대입구역 3번 골목상권)에 위치합니다.",
    stats: [
      { label: "사용승인일 (경과연수)", value: "1992년", sub: "33년 경과 · 노후" },
      { label: "내진 설계", value: "미적용", sub: "구조: 철근콘크리트조" },
      { label: "최근 실거래가", value: "15.2억", sub: "3층 · 84㎡ · 2026.05" },
      { label: "공실률", value: "4.3%", sub: "중대형 상가 · 2026년 2분기" },
      { label: "반경 500m 상가업소", value: "312개", sub: "소상공인시장진흥공단 기준" },
      { label: "상권 추정매출(월)", value: "84.8억", sub: "홍대입구역 3번 골목상권" },
    ],
  },

  building: {
    core: [
      ["대지면적", "312.4㎡"], ["건축면적", "198.7㎡"], ["연면적", "612.3㎡"],
      ["건폐율", "63.6%"], ["용적률", "196.1%"], ["구조", "철근콘크리트조"],
      ["지붕", "슬래브"], ["지상/지하층수", "4층 / 1층"], ["사용승인일", "1992.11.03"],
      ["허가일", "1991.05.20"], ["착공일", "1991.06.15"], ["내진설계", "미적용"],
    ],
    floors: [
      { label: "1층", value: 152 }, { label: "2층", value: 148 },
      { label: "3층", value: 148 }, { label: "4층", value: 140 }, { label: "지하1층", value: 24 },
    ],
    zoning: ["제2종일반주거지역", "지구단위계획구역", "대공방어협조구역", "가축사육제한구역"],
  },

  location: {
    adongName: "연남동", ldongName: "연남동",
    subway: [
      { name: "홍대입구역 3번 출구", line: "경의중앙선", badge: "경의", color: "77C4A3", dist: "180m" },
      { name: "홍대입구역 3번 출구", line: "공항철도", badge: "공항", color: "0090D2", dist: "180m" },
      { name: "홍대입구역 2번 출구", line: "2호선", badge: "2", color: "00A84D", dist: "340m" },
    ],
  },

  transactions: {
    rows: [
      ["2026.05.12", "3층", "84.3㎡", "15.2억"],
      ["2026.02.03", "2층", "76.1㎡", "13.8억"],
      ["2025.11.20", "4층", "84.3㎡", "14.6억"],
      ["2025.08.07", "1층", "62.5㎡", "18.4억"],
      ["2025.04.15", "3층", "84.3㎡", "13.9억"],
    ],
    trend: [
      { label: "2025.Q2", value: 132 }, { label: "2025.Q3", value: 138 },
      { label: "2025.Q4", value: 141 }, { label: "2026.Q1", value: 149 }, { label: "2026.Q2", value: 156 },
    ],
  },

  district: {
    ageBuckets: [
      { label: "~1980", value: 42 }, { label: "1980s", value: 118 }, { label: "1990s", value: 203 },
      { label: "2000s", value: 156 }, { label: "2010s", value: 97 }, { label: "2020~", value: 34 },
    ],
    callout: { value: "상위 28%", label: "연남동 전체 건물 중 노후 순위" },
  },

  priceHistory: {
    trend: [
      { label: "2020", value: 8.2 }, { label: "2021", value: 8.9 }, { label: "2022", value: 9.6 },
      { label: "2023", value: 9.8 }, { label: "2024", value: 10.1 }, { label: "2025", value: 10.7 }, { label: "2026", value: 11.2 },
    ],
    rows: [
      ["2026", "11.2억", "+4.7%"], ["2025", "10.7억", "+5.9%"], ["2024", "10.1억", "+3.1%"], ["2023", "9.8억", "+2.1%"],
    ],
  },

  commercial: {
    vacancyTrend: [
      { label: "24.Q3", value: 5.1 }, { label: "24.Q4", value: 4.9 }, { label: "25.Q1", value: 4.6 },
      { label: "25.Q2", value: 4.4 }, { label: "25.Q3", value: 4.5 }, { label: "25.Q4", value: 4.2 }, { label: "26.Q1", value: 4.3 },
    ],
    topIndustries: [
      { label: "한식음식점", value: 38 }, { label: "커피-음료", value: 31 }, { label: "일반의류", value: 26 },
      { label: "미용실", value: 19 }, { label: "호프-간이주점", value: 17 },
    ],
  },

  // golmok.seoul.go.kr(서울시 상권분석 서비스) 화면 스타일을 참고한 상권영역 지도 슬라이드용
  // 데이터. TbgisTrdarRelm은 상권 중심점(1점)+면적만 주고 경계 폴리곤 좌표는 안 주므로,
  // locationImage는 정확한 상권 경계가 아니라 대상 위치 중심의 실제 지도 스크린샷이다
  // (map_shot.html로 생성 — Leaflet+OSM, 무료/키 불필요. 실제 자동화 단계에서는 대시보드의
  // geocode_address_kakao() 좌표를 같은 방식으로 넘겨 매 주소마다 새로 캡처하면 된다).
  tradeAreaMap: {
    locationImage: "location_map.png",
    stats: [
      { label: "유동인구 (도로변)", low: "359명", high: "2,932명" },
      { label: "유동인구 (건물 내부)", low: "0명", high: "608명" },
      { label: "인구밀도", low: "0명", high: "0명" },
    ],
    period: "2026년 1분기 기준",
  },

  seoulTradeArea: {
    name: "홍대입구역 3번 골목상권",
    stats: [
      { label: "추정매출 (당월)", value: "84.8억원" },
      { label: "생활인구 (분기)", value: "885,714명" },
      { label: "직장인구 (분기)", value: "10,650명" },
    ],
    topIndustries: [
      ["일반의류", "24.4억"], ["호프-간이주점", "22.0억"], ["한식음식점", "9.1억"],
      ["일반의원", "7.3억"], ["일식음식점", "6.7억"],
    ],
    weekday: [
      { label: "월", value: 9.0 }, { label: "화", value: 9.0 }, { label: "수", value: 8.7 },
      { label: "목", value: 10.1 }, { label: "금", value: 13.5 }, { label: "토", value: 22.0 }, { label: "일", value: 12.5 },
    ],
  },

  // 아래 6개 섹션(toc~swot)은 연남동_상권분석1_보강본.pptx의 3/4/7/8/11/13/26페이지 구성을
  // 참고해 추가한 것. 실거래가·공시가격 등과 달리 이 내용들(상권 성격/SNS 트렌드/개발호재/SWOT)은
  // 공공API로 자동 수집되지 않는 정성적 항목이라, 실제 자동화 단계에서도 수동 리서치 입력이나
  // 별도 LLM 초안 생성 단계가 필요하다 — 지금은 스키마 예시로 연남동 사례를 채워 넣었다.
  coverStats: [
    { value: "21건", label: "확인된 건물 실거래" },
    { value: "+43.8%", label: "실거래 대비 현재 호가 프리미엄" },
    { value: "3노선", label: "홍대입구역 환승" },
  ],

  toc: {
    parts: [
      {
        title: "PART 01 · 입지 및 건축물 분석",
        items: ["표지", "목차", "핵심 요약", "건축물 개요", "위치 및 입지", "입지 및 교통 여건"],
      },
      {
        title: "PART 02 · 시장 데이터 분석",
        items: [
          "실거래가 동향", "동단위 시장 통계", "공시가격 시계열", "상권 개황",
          "상권영역 지도", "상권 특성", "상권 트렌드 지표", "개발호재 종합", "서울 상권 상세",
        ],
      },
      {
        title: "PART 03 · 종합 평가",
        items: ["SWOT 분석", "종합 의견"],
      },
    ],
  },

  locationAssets: [
    {
      title: "홍대입구역",
      desc: "2호선·경의중앙선·공항철도 3개 노선 환승역으로, 일평균 승하차 약 15만~19만명 수준이다. 인근역 대비 압도적으로 높은 유동인구를 기록한다.",
      metric: "3개 노선 환승", sub: "도보 3분 이내",
    },
    {
      title: "경의선숲길(연트럴파크)",
      desc: "홍대입구역 3번 출구 바로 앞에서 시작하는 선형공원으로, 2016년 전 구간 조성이 완료됐다. 상권의 물리적 축을 이룬다.",
      metric: "3번 출구 직결", sub: "전 구간 조성 완료",
    },
    {
      title: "대장홍대선(예정)",
      desc: "2025년 12월 착공, 2031년 준공 목표. 개통 시 홍대입구역은 기존 3개 노선에 더해 4개 노선 환승역으로 격상될 예정이다.",
      metric: "2031년 준공", sub: "4개 노선 환승 예정",
    },
  ],

  marketCharacter: [
    { title: "화교 상권에서 출발", desc: "화교 밀집 거주지의 개성있는 중식당들이 외부에 알려지며 상권의 초기 인지도를 형성했습니다." },
    { title: "홍대발 젠트리피케이션 유입", desc: "홍대 앞 상권의 임대료 상승을 피해 넘어온 예술가·소상공인들이 상권의 개성을 형성했습니다." },
    { title: "청소년기 상권", desc: "로컬맛집·부티크·편집숍·팝업스토어 중심의 청소년기 상권으로, 개인 오너 매장 위주입니다." },
    { title: "20~30대 소비층", desc: "트렌드에 민감하고 SNS 확산력이 높은 20~30대 소비층이 상권 소비의 중심을 이룹니다." },
    { title: "개성있는 골목 매장", desc: "여기서만 만날 수 있는 개성있는 매장들이 인근 골목 일대에 밀집해 있습니다." },
    { title: "오너·디렉터 취향 매장", desc: "표준화된 프랜차이즈보다 오너의 취향이 드러나는 소규모 매장이 소비자를 끌어들입니다." },
  ],

  trend: {
    indicators: [
      { value: "5,000+", label: "#연남동라멘 SNS 게시물", sub: "2026년 7월 기준" },
      { value: "31만+", label: "#라멘맛집 관련 게시물", sub: "2026년 7월 기준" },
      { value: "오전 9시", label: "인기매장 오픈런·예약 시작", sub: "점심시간 전 마감 사례 다수" },
      { value: "20~30대", label: "1인 방문객 중심 소비층", sub: "회전율 높은 업종과 궁합" },
    ],
    insights: [
      { title: "SNS 콘텐츠 최적화 입지", desc: "오픈키친형 인테리어가 SNS 콘텐츠 생산에 최적화돼 있으며, 역세권·공원 인접 입지가 꾸준한 유동인구를 뒷받침합니다." },
      { title: "트렌드 의존 리스크", desc: "SNS 트렌드가 빠르게 순환하는 상권 특성상, 업종·임차인 구성의 변동성에 유의할 필요가 있습니다." },
    ],
  },

  growthDrivers: [
    { title: "대장홍대선", desc: "부천 대장지구~홍대입구역 20.1km, 12개 정차역. 2025.12 착공, 2031년 준공 목표." },
    { title: "철도지하화 통합개발", desc: "경의선 구간 포함, 서울시 2024.10 발표. 국토부 타당성조사 진행 중(미확정)." },
    { title: "지구단위계획 재정비", desc: "역세권 복합개발 인센티브 등 높이기준 재정비가 추진되고 있습니다." },
    { title: "라멘성지 트렌드", desc: "SNS 기반 신규 소비 트렌드로 상권 화제성이 유지되고 있습니다." },
    { title: "플래그십 스토어 진출", desc: "브랜드 오프라인 전략 고도화에 따른 임차 수요 증가가 나타나고 있습니다." },
  ],

  swot: {
    strengths: [
      "홍대입구역(3노선 환승) 중심 압도적 유동인구",
      "경의선숲길(연트럴파크)이라는 대체불가 산책로 자산",
      "최근 실거래가 상승 흐름(전년比 +10.4%)",
    ],
    weaknesses: [
      "트렌드 의존형 상권 특성상 업종 쏠림 변동성 큼",
      "필지 규모가 대체로 소형이라 대형 개발엔 합필 필요",
      "실거래·매물 지번 불일치로 동일 건물 기준 가격비교 데이터 부재",
    ],
    opportunities: [
      "대장홍대선 착공(2031년 4노선 환승역화) 실현 시 가치 확장 가능성",
      "지구단위계획 재정비로 역세권 복합개발 인센티브 도입 추진",
      "인접 필지 동시매물로 합필 개발 여지",
    ],
    threats: [
      "대장홍대선·철도지하화 모두 중장기 계획으로 선반영 리스크",
      "현재 호가가 실거래 대비 높아 고점매수 리스크",
      "공식 통계로 교차검증되지 않은 민간 데이터 의존도 존재",
    ],
  },

  conclusion: [
    "준공 33년 경과·내진 미적용 건물로, 매입 시 리모델링/재건축 비용을 사전 검토할 필요가 있습니다.",
    "홍대입구역 도보 3분 내 위치한 활발한 골목상권으로, 반경 500m 내 업종 다양성과 유동인구가 높은 편입니다.",
    "최근 4개 분기 공실률이 4.2~5.1% 구간에서 안정적으로 유지되고 있어 임대 리스크는 제한적입니다.",
    "실거래가는 최근 1년간 완만한 상승세로, 동일 면적대 매물 대비 소폭 저평가 구간으로 판단됩니다.",
  ],
};

// ------------------------------------------------------------------
// 공통 헬퍼
// ------------------------------------------------------------------
function newSlideBase(pres, { dark = false } = {}) {
  const slide = pres.addSlide();
  slide.background = { color: dark ? NAVY : WHITE };
  return slide;
}

function addSectionTitle(slide, text) {
  slide.addText(text, {
    x: 0.6, y: 0.45, w: 11.3, h: 0.6,
    fontFace: FONT_HEAD, fontSize: 28, bold: true, color: NAVY, align: "left",
  });
}

function addPageFooter(slide, pageLabel) {
  slide.addText(`${REPORT_DATA.address}  ·  ${pageLabel}`, {
    x: 0.6, y: 7.12, w: 10, h: 0.3,
    fontFace: FONT_BODY, fontSize: 9, color: MUTED, align: "left",
  });
}

function card(slide, opts) {
  const { x, y, w, h, radius = 0.12 } = opts;
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: radius,
    fill: { color: opts.fill || CARD_BG },
    line: { color: opts.line || GRID_LINE, width: 0.75 },
    shadow: opts.shadow === false ? undefined : {
      type: "outer", color: "9AA5B1", opacity: 0.25, blur: 6, offset: 2, angle: 90,
    },
  });
}

function statCard(slide, x, y, w, h, stat) {
  card(slide, { x, y, w, h });
  slide.addText(stat.value, {
    x: x + 0.15, y: y + 0.12, w: w - 0.3, h: h * 0.45,
    fontFace: FONT_HEAD, fontSize: 22, bold: true, color: TERRACOTTA, align: "left", valign: "bottom",
  });
  slide.addText(stat.label, {
    x: x + 0.15, y: y + h * 0.55, w: w - 0.3, h: h * 0.22,
    fontFace: FONT_BODY, fontSize: 11, bold: true, color: TEXT_DARK, align: "left",
  });
  if (stat.sub) {
    slide.addText(stat.sub, {
      x: x + 0.15, y: y + h * 0.76, w: w - 0.3, h: h * 0.2,
      fontFace: FONT_BODY, fontSize: 9, color: MUTED, align: "left",
    });
  }
}

function chartAxisDefaults() {
  return {
    catAxisLabelColor: MUTED, catAxisLabelFontSize: 10, catAxisLineColor: GRID_LINE,
    catGridLine: { style: "none" },
    valAxisLabelColor: MUTED, valAxisLabelFontSize: 10, valAxisLineColor: GRID_LINE,
    valGridLine: { color: GRID_LINE, size: 0.75 },
    showLegend: false,
  };
}

// ==================================================================
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
pres.author = "실거래가 대시보드";
pres.title = "부동산 분석 리포트";

// ------------------------------------------------------------------
// 슬라이드 1 — 표지
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres, { dark: true });

  // 배경 모티프: 상권 반경을 연상시키는 동심원 (우하단, 미세한 톤 차이)
  slide.addShape("ellipse", { x: 8.6, y: 3.0, w: 7.5, h: 7.5, fill: { color: NAVY_LIGHT }, line: { type: "none" } });
  slide.addShape("ellipse", { x: 9.6, y: 4.0, w: 5.5, h: 5.5, fill: { color: NAVY_DARK }, line: { type: "none" } });
  slide.addShape("ellipse", { x: 10.4, y: 4.8, w: 3.9, h: 3.9, fill: { color: NAVY }, line: { color: TERRACOTTA, width: 1 } });

  slide.addText("부동산 분석 리포트", {
    x: 0.9, y: 2.35, w: 8, h: 0.4,
    fontFace: FONT_BODY, fontSize: 14, color: TERRACOTTA, bold: true, charSpacing: 3,
  });
  slide.addText(REPORT_DATA.address, {
    x: 0.9, y: 2.8, w: 9.5, h: 1.6,
    fontFace: FONT_HEAD, fontSize: 40, bold: true, color: WHITE, align: "left", valign: "top",
  });
  slide.addText("건축물 · 실거래가 · 상권 통합 분석", {
    x: 0.9, y: 4.05, w: 8, h: 0.5,
    fontFace: FONT_BODY, fontSize: 16, color: ICE,
  });
  // 미니 지표 3종 (실거래/호가/교통) — 표지에서 바로 핵심 임팩트를 전달
  const cs = REPORT_DATA.coverStats;
  const csW = 2.6;
  cs.forEach((s, i) => {
    const x = 0.9 + i * (csW + 0.4);
    slide.addText(s.value, { x, y: 5.5, w: csW, h: 0.55, fontFace: FONT_HEAD, fontSize: 26, bold: true, color: TERRACOTTA });
    slide.addText(s.label, { x, y: 6.05, w: csW, h: 0.5, fontFace: FONT_BODY, fontSize: 10.5, color: ICE });
  });

  slide.addText(REPORT_DATA.reportDate, {
    x: 0.9, y: 6.7, w: 6, h: 0.4,
    fontFace: FONT_BODY, fontSize: 11, color: MUTED,
  });
}

// ------------------------------------------------------------------
// 슬라이드 2 — 목차
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "목차");
  slide.addText(`${REPORT_DATA.address} 분석 리포트 구성`, {
    x: 0.6, y: 1.05, w: 11, h: 0.3, fontFace: FONT_BODY, fontSize: 12, color: MUTED,
  });

  const parts = REPORT_DATA.toc.parts;
  const colW = 3.95, gapX = 0.2, startX = 0.6, startY = 1.55;
  let pageNo = 1;
  parts.forEach((part, pi) => {
    const x = startX + pi * (colW + gapX);
    slide.addText(part.title, {
      x, y: startY, w: colW, h: 0.5, fontFace: FONT_HEAD, fontSize: 13.5, bold: true, color: TERRACOTTA,
    });
    let iy = startY + 0.55;
    part.items.forEach((item) => {
      const badgeW = pageNo >= 10 ? 0.4 : 0.32;
      slide.addShape("roundRect", { x, y: iy, w: badgeW, h: 0.32, rectRadius: 0.06, fill: { color: NAVY }, line: { type: "none" } });
      slide.addText(String(pageNo), {
        x, y: iy, w: badgeW, h: 0.32, align: "center", valign: "middle", fontFace: FONT_BODY,
        fontSize: pageNo >= 10 ? 9 : 10, bold: true, color: WHITE, wrap: false,
      });
      slide.addText(item, {
        x: x + 0.5, y: iy - 0.03, w: colW - 0.53, h: 0.38, fontFace: FONT_BODY, fontSize: 11.5, color: TEXT_DARK, valign: "middle",
      });
      iy += 0.44;
      pageNo++;
    });
  });
  addPageFooter(slide, "목차");
}

// ------------------------------------------------------------------
// 슬라이드 3 — 핵심 요약
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "핵심 요약");

  card(slide, { x: 0.6, y: 1.25, w: 12.1, h: 1.35 });
  slide.addText(REPORT_DATA.summary.text, {
    x: 0.9, y: 1.42, w: 11.5, h: 1.0,
    fontFace: FONT_BODY, fontSize: 13, color: TEXT_DARK, valign: "middle", lineSpacingMultiple: 1.25,
  });

  const stats = REPORT_DATA.summary.stats;
  const cols = 3, gap = 0.25;
  const cardW = (12.1 - gap * (cols - 1)) / cols;
  const cardH = 1.35;
  const startY = 2.85;
  stats.forEach((stat, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = 0.6 + col * (cardW + gap);
    const y = startY + row * (cardH + gap);
    statCard(slide, x, y, cardW, cardH, stat);
  });
  addPageFooter(slide, "핵심 요약");
}

// ------------------------------------------------------------------
// 슬라이드 4 — 건축물 개요
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "건축물 개요");

  // 왼쪽: 핵심정보 카드 (2열 라벨-값)
  const leftX = 0.6, leftY = 1.3, leftW = 6.7, leftH = 5.6;
  card(slide, { x: leftX, y: leftY, w: leftW, h: leftH });
  const core = REPORT_DATA.building.core;
  const rowH = (leftH - 0.4) / Math.ceil(core.length / 2);
  core.forEach(([label, value], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = leftX + 0.3 + col * (leftW / 2 - 0.15);
    const y = leftY + 0.25 + row * rowH;
    slide.addText(label, { x, y, w: leftW / 2 - 0.5, h: rowH * 0.42, fontFace: FONT_BODY, fontSize: 10.5, color: MUTED });
    slide.addText(value, { x, y: y + rowH * 0.38, w: leftW / 2 - 0.5, h: rowH * 0.5, fontFace: FONT_HEAD, fontSize: 15, bold: true, color: NAVY });
  });

  // 오른쪽 위: 층별 면적 바 차트
  const rightX = 7.5, rightW = 5.2;
  slide.addText("층별 면적 (㎡)", { x: rightX, y: 1.3, w: rightW, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addChart("bar", [{
    name: "면적",
    labels: REPORT_DATA.building.floors.map(f => f.label),
    values: REPORT_DATA.building.floors.map(f => f.value),
  }], {
    x: rightX, y: 1.7, w: rightW, h: 2.7,
    barDir: "bar", chartColors: [TERRACOTTA],
    showTitle: false, showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TEXT_DARK, dataLabelFontSize: 9,
    ...chartAxisDefaults(),
  });

  // 오른쪽 아래: 용도지역/지구 pill
  slide.addText("용도지역 · 지구", { x: rightX, y: 4.65, w: rightW, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  let py = 5.05;
  REPORT_DATA.building.zoning.forEach((z) => {
    const pw = Math.min(5.2, 0.35 + z.length * 0.135);
    slide.addShape("roundRect", { x: rightX, y: py, w: pw, h: 0.42, rectRadius: 0.21, fill: { color: ICE }, line: { type: "none" } });
    slide.addText(z, { x: rightX, y: py, w: pw, h: 0.42, align: "center", valign: "middle", fontFace: FONT_BODY, fontSize: 10.5, color: NAVY });
    py += 0.55;
  });
  addPageFooter(slide, "건축물 개요");
}

// ------------------------------------------------------------------
// 슬라이드 5 — 위치 및 입지
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "위치 및 입지");

  // 왼쪽: 지도 placeholder
  const mapX = 0.6, mapY = 1.3, mapW = 7.6, mapH = 5.6;
  slide.addShape("roundRect", {
    x: mapX, y: mapY, w: mapW, h: mapH, rectRadius: 0.12,
    fill: { color: "EDEFF2" }, line: { color: GRID_LINE, width: 1, dashType: "dash" },
  });
  slide.addText("지도 이미지 삽입 영역\n(지오코딩 좌표 + 인근 지하철 마커 스크린샷)", {
    x: mapX, y: mapY + mapH / 2 - 0.5, w: mapW, h: 1.0, align: "center", valign: "middle",
    fontFace: FONT_BODY, fontSize: 13, color: MUTED,
  });

  // 오른쪽: 행정동/법정동 + 지하철
  const rx = 8.5, rw = 4.2;
  card(slide, { x: rx, y: 1.3, w: rw, h: 1.3 });
  slide.addText("행정동", { x: rx + 0.25, y: 1.42, w: rw - 0.5, h: 0.3, fontFace: FONT_BODY, fontSize: 10.5, color: MUTED });
  slide.addText(REPORT_DATA.location.adongName, { x: rx + 0.25, y: 1.68, w: rw - 0.5, h: 0.4, fontFace: FONT_HEAD, fontSize: 18, bold: true, color: NAVY });
  slide.addText("법정동", { x: rx + 0.25, y: 2.1, w: rw - 0.5, h: 0.3, fontFace: FONT_BODY, fontSize: 10.5, color: MUTED });
  slide.addText(REPORT_DATA.location.ldongName, { x: rx + 0.25, y: 2.36, w: rw - 0.5, h: 0.2, fontFace: FONT_BODY, fontSize: 13, color: TEXT_DARK });

  slide.addText("인근 지하철역", { x: rx, y: 2.85, w: rw, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  let sy = 3.25;
  REPORT_DATA.location.subway.forEach((st) => {
    card(slide, { x: rx, y: sy, w: rw, h: 0.85, shadow: false });
    slide.addShape("ellipse", { x: rx + 0.18, y: sy + 0.2, w: 0.44, h: 0.44, fill: { color: st.color }, line: { color: WHITE, width: 1.5 } });
    slide.addText(st.badge, {
      x: rx + 0.18, y: sy + 0.2, w: 0.44, h: 0.44, align: "center", valign: "middle",
      fontFace: FONT_BODY, fontSize: st.badge.length > 1 ? 8.5 : 12, bold: true, color: WHITE, shrinkText: true,
    });
    slide.addText(st.name, { x: rx + 0.78, y: sy + 0.1, w: rw - 1.0, h: 0.35, fontFace: FONT_BODY, fontSize: 11.5, bold: true, color: TEXT_DARK });
    slide.addText(`${st.line} · 도보 약 ${st.dist}`, { x: rx + 0.78, y: sy + 0.45, w: rw - 1.0, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, color: MUTED });
    sy += 1.0;
  });
  addPageFooter(slide, "위치 및 입지");
}

// ------------------------------------------------------------------
// 슬라이드 6 — 입지 및 교통 여건 (핵심 입지 자산)
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "입지 및 교통 여건");
  slide.addText("핵심 입지 자산", { x: 0.6, y: 1.05, w: 10, h: 0.3, fontFace: FONT_BODY, fontSize: 12, color: MUTED });

  const assets = REPORT_DATA.locationAssets;
  const gap = 0.3;
  const cardW = (12.1 - gap * (assets.length - 1)) / assets.length;
  const cardH = 4.7, cardY = 1.55;
  assets.forEach((a, i) => {
    const x = 0.6 + i * (cardW + gap);
    card(slide, { x, y: cardY, w: cardW, h: cardH });
    slide.addShape("ellipse", { x: x + 0.3, y: cardY + 0.3, w: 0.5, h: 0.5, fill: { color: TERRACOTTA }, line: { type: "none" } });
    slide.addText(String(i + 1), {
      x: x + 0.3, y: cardY + 0.3, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 18, bold: true, color: WHITE,
    });
    slide.addText(a.title, { x: x + 0.3, y: cardY + 1.0, w: cardW - 0.6, h: 0.5, fontFace: FONT_HEAD, fontSize: 16, bold: true, color: NAVY });
    slide.addText(a.desc, {
      x: x + 0.3, y: cardY + 1.55, w: cardW - 0.6, h: 2.0, fontFace: FONT_BODY, fontSize: 11.5,
      color: TEXT_DARK, lineSpacingMultiple: 1.3, valign: "top",
    });
    const pillY = cardY + cardH - 1.0;
    slide.addShape("roundRect", { x: x + 0.3, y: pillY, w: cardW - 0.6, h: 0.65, rectRadius: 0.1, fill: { color: ICE }, line: { type: "none" } });
    slide.addText(a.metric, { x: x + 0.45, y: pillY + 0.05, w: cardW - 0.9, h: 0.32, fontFace: FONT_HEAD, fontSize: 13, bold: true, color: NAVY });
    slide.addText(a.sub, { x: x + 0.45, y: pillY + 0.36, w: cardW - 0.9, h: 0.26, fontFace: FONT_BODY, fontSize: 9.5, color: MUTED });
  });
  addPageFooter(slide, "입지 및 교통 여건");
}

// ------------------------------------------------------------------
// 슬라이드 7 — 실거래가 동향
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "실거래가 동향");

  slide.addText("최근 거래 내역", { x: 0.6, y: 1.3, w: 5.6, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addTable(
    [["거래일", "층", "전용면적", "거래가격"].map(t => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11 } }))]
      .concat(REPORT_DATA.transactions.rows.map(r => r.map(t => ({ text: t, options: { fontSize: 11, color: TEXT_DARK } })))),
    {
      x: 0.6, y: 1.7, w: 5.6, h: 2.6, fontFace: FONT_BODY,
      border: { type: "solid", color: GRID_LINE, pt: 0.75 },
      autoPage: false, colW: [1.6, 1.1, 1.4, 1.5],
      rowH: 0.42,
    }
  );

  slide.addText("가격 추이 (평당가 지수, 최근 5개 분기)", { x: 6.6, y: 1.3, w: 6.1, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addChart("line", [{
    name: "평당가지수",
    labels: REPORT_DATA.transactions.trend.map(t => t.label),
    values: REPORT_DATA.transactions.trend.map(t => t.value),
  }], {
    x: 6.6, y: 1.7, w: 6.1, h: 2.7,
    chartColors: [TERRACOTTA], lineSize: 3, lineSmooth: true,
    showTitle: false, showValue: true, dataLabelPosition: "t", dataLabelColor: NAVY, dataLabelFontSize: 9, dataLabelFormatCode: "0.0",
    ...chartAxisDefaults(),
  });

  card(slide, { x: 0.6, y: 4.6, w: 12.1, h: 1.3 });
  slide.addText("최근 6개월 인근 실거래가는 15.2억원 수준으로, 직전 동일 면적대 거래(13.8억) 대비 소폭 상승했습니다.", {
    x: 0.9, y: 4.6, w: 11.5, h: 1.3, valign: "middle", fontFace: FONT_BODY, fontSize: 13, color: TEXT_DARK,
  });
  addPageFooter(slide, "실거래가 동향");
}

// ------------------------------------------------------------------
// 슬라이드 8 — 동단위 시장 통계
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, `동단위 시장 통계 — ${REPORT_DATA.location.adongName}`);

  slide.addText("준공연도대별 건물 수 분포 (동 전체)", { x: 0.6, y: 1.3, w: 8, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addChart("bar", [{
    name: "건물 수",
    labels: REPORT_DATA.district.ageBuckets.map(b => b.label),
    values: REPORT_DATA.district.ageBuckets.map(b => b.value),
  }], {
    x: 0.6, y: 1.7, w: 8.0, h: 4.6,
    chartColors: [NAVY_LIGHT], showTitle: false, showValue: true, dataLabelPosition: "outEnd",
    dataLabelColor: TEXT_DARK, dataLabelFontSize: 10,
    ...chartAxisDefaults(),
  });

  const rx = 9.0, rw = 3.7;
  statCard(slide, rx, 1.7, rw, 1.6, { value: REPORT_DATA.district.callout.value, label: REPORT_DATA.district.callout.label, sub: "1992년 사용승인 기준" });
  card(slide, { x: rx, y: 3.5, w: rw, h: 2.8 });
  slide.addText(
    "동 전체 건물의 절반 이상이 1990~2000년대에 준공되어, 대상 건물과 유사한 노후도 구간에 " +
    "밀집되어 있습니다. 재건축·리모델링 수요가 점진적으로 늘어날 가능성이 있는 지역입니다.",
    { x: rx + 0.25, y: 3.7, w: rw - 0.5, h: 2.4, valign: "top", fontFace: FONT_BODY, fontSize: 12, color: TEXT_DARK, lineSpacingMultiple: 1.3 }
  );
  addPageFooter(slide, "동단위 시장 통계");
}

// ------------------------------------------------------------------
// 슬라이드 9 — 공시가격 시계열
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "공시가격 시계열");

  slide.addText("연도별 공시가격 추이", { x: 0.6, y: 1.3, w: 7.4, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addChart("line", [{
    name: "공시가격(억원)",
    labels: REPORT_DATA.priceHistory.trend.map(t => t.label),
    values: REPORT_DATA.priceHistory.trend.map(t => t.value),
  }], {
    x: 0.6, y: 1.7, w: 7.4, h: 4.6,
    chartColors: [TERRACOTTA], lineSize: 3, lineSmooth: false,
    showTitle: false, showValue: true, dataLabelPosition: "t", dataLabelColor: NAVY, dataLabelFontSize: 10, dataLabelFormatCode: "0.0",
    ...chartAxisDefaults(),
  });

  slide.addText("연도별 변동률", { x: 8.3, y: 1.3, w: 4.4, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addTable(
    [["연도", "공시가격", "전년대비"].map(t => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11 } }))]
      .concat(REPORT_DATA.priceHistory.rows.map(r => r.map((t, i) => ({
        text: t, options: { fontSize: 11, color: i === 2 ? TERRACOTTA : TEXT_DARK, bold: i === 2 },
      })))),
    { x: 8.3, y: 1.7, w: 4.4, h: 2.2, fontFace: FONT_BODY, border: { type: "solid", color: GRID_LINE, pt: 0.75 }, colW: [1.3, 1.6, 1.5], rowH: 0.42 }
  );
  card(slide, { x: 8.3, y: 4.1, w: 4.4, h: 2.2 });
  slide.addText("최근 7년간 연평균 약 5.5% 상승하며 꾸준한 우상향 흐름을 보이고 있습니다.", {
    x: 8.55, y: 4.1, w: 3.9, h: 2.2, valign: "middle", fontFace: FONT_BODY, fontSize: 12, color: TEXT_DARK, lineSpacingMultiple: 1.3,
  });
  addPageFooter(slide, "공시가격 시계열");
}

// ------------------------------------------------------------------
// 슬라이드 10 — 상권 개황 (공실률 + 주변 상가업소)
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "상권 개황 — 공실률 및 주변 업종");

  slide.addText("공실률 추이 (중대형 상가, 분기별)", { x: 0.6, y: 1.3, w: 6.0, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addChart("line", [{
    name: "공실률(%)",
    labels: REPORT_DATA.commercial.vacancyTrend.map(t => t.label),
    values: REPORT_DATA.commercial.vacancyTrend.map(t => t.value),
  }], {
    x: 0.6, y: 1.7, w: 6.0, h: 4.6,
    chartColors: [NAVY_LIGHT], lineSize: 3, lineSmooth: true,
    showTitle: false, showValue: true, dataLabelPosition: "t", dataLabelColor: NAVY, dataLabelFontSize: 9, dataLabelFormatCode: "0.0",
    ...chartAxisDefaults(),
  });

  slide.addText("반경 500m 업종 Top 5 (점포 수)", { x: 6.9, y: 1.3, w: 5.8, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addChart("bar", [{
    name: "점포 수",
    labels: REPORT_DATA.commercial.topIndustries.map(t => t.label),
    values: REPORT_DATA.commercial.topIndustries.map(t => t.value),
  }], {
    x: 6.9, y: 1.7, w: 5.8, h: 4.6,
    barDir: "bar", chartColors: [TERRACOTTA],
    showTitle: false, showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TEXT_DARK, dataLabelFontSize: 10,
    ...chartAxisDefaults(),
  });
  addPageFooter(slide, "상권 개황");
}

// ------------------------------------------------------------------
// 슬라이드 11 — 상권영역 지도 (서울시 상권분석 서비스 화면 스타일 참고, 서울 소재 주소만)
// ------------------------------------------------------------------
if (REPORT_DATA.isSeoul) {
  const slide = newSlideBase(pres);
  addSectionTitle(slide, `상권영역 지도 — ${REPORT_DATA.seoulTradeArea.name}`);

  const mapX = 0.6, mapY = 1.3, mapW = 6.5, mapH = 5.2;
  const tm = REPORT_DATA.tradeAreaMap;

  // 실제 위치 지도 스크린샷 (OpenStreetMap/Leaflet, map_shot.html로 생성 — 300m 반경 표시).
  // 이미지 비율(2600x2080 = 1.25)을 mapW/mapH 비율과 맞춰뒀기 때문에 찌그러짐 없이 꽉 채워진다.
  slide.addImage({ path: tm.locationImage, x: mapX, y: mapY, w: mapW, h: mapH, sizing: { type: "cover", w: mapW, h: mapH } });
  slide.addShape("roundRect", {
    x: mapX, y: mapY, w: mapW, h: mapH, rectRadius: 0.1,
    fill: { type: "none" }, line: { color: GRID_LINE, width: 1 },
  });

  slide.addText(
    "※ OpenStreetMap 기반 위치 참고 지도(점선: 대상 건물 기준 반경 300m). 상권영역 경계 자체는 서울시 API가 " +
    "중심좌표+면적만 제공해 정확한 폴리곤으로 표시할 수 없습니다.",
    { x: mapX, y: mapY + mapH + 0.08, w: mapW, h: 0.3, fontFace: FONT_BODY, fontSize: 8.5, italic: true, color: MUTED }
  );

  // 오른쪽 통계 카드 (골목상권 화면의 유동인구/인구밀도 패널 참고)
  const rx = 7.4, rw = 5.3;
  let ry = 1.3;
  tm.stats.forEach((stat) => {
    card(slide, { x: rx, y: ry, w: rw, h: 1.35 });
    slide.addText(stat.label, { x: rx + 0.25, y: ry + 0.18, w: rw - 0.5, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
    slide.addText([
      { text: "적음  ", options: { color: MUTED, fontSize: 12 } },
      { text: stat.low, options: { color: NAVY_LIGHT, bold: true, fontSize: 14 } },
      { text: "     많음  ", options: { color: MUTED, fontSize: 12 } },
      { text: stat.high, options: { color: TERRACOTTA, bold: true, fontSize: 14 } },
    ], { x: rx + 0.25, y: ry + 0.6, w: rw - 0.5, h: 0.35, fontFace: FONT_BODY });
    slide.addText(tm.period, { x: rx + 0.25, y: ry + 0.98, w: rw - 0.5, h: 0.3, fontFace: FONT_BODY, fontSize: 9.5, color: MUTED });
    ry += 1.55;
  });
  addPageFooter(slide, "상권영역 지도");
}

// ------------------------------------------------------------------
// 슬라이드 12 — 상권 특성 (시장 성격, 정성적 분석)
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "상권 특성");
  slide.addText("상권의 성격과 특징 (공식 통계가 아닌 시장 관찰 기반)", {
    x: 0.6, y: 1.05, w: 11, h: 0.3, fontFace: FONT_BODY, fontSize: 12, color: MUTED,
  });

  const items = REPORT_DATA.marketCharacter;
  const cols = 3, gap = 0.25;
  const cardW = (12.1 - gap * (cols - 1)) / cols;
  const cardH = 2.3, startY = 1.5;
  items.forEach((it, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = 0.6 + col * (cardW + gap), y = startY + row * (cardH + gap);
    card(slide, { x, y, w: cardW, h: cardH });
    slide.addShape("ellipse", { x: x + 0.25, y: y + 0.22, w: 0.36, h: 0.36, fill: { color: TERRACOTTA_LIGHT }, line: { type: "none" } });
    slide.addShape("ellipse", { x: x + 0.36, y: y + 0.33, w: 0.14, h: 0.14, fill: { color: TERRACOTTA }, line: { type: "none" } });
    slide.addText(it.title, {
      x: x + 0.75, y: y + 0.15, w: cardW - 1.0, h: 0.55, fontFace: FONT_HEAD, fontSize: 12.5, bold: true, color: NAVY, valign: "middle",
    });
    slide.addText(it.desc, {
      x: x + 0.25, y: y + 0.8, w: cardW - 0.5, h: cardH - 1.0, fontFace: FONT_BODY, fontSize: 10.5,
      color: TEXT_DARK, lineSpacingMultiple: 1.25, valign: "top",
    });
  });
  addPageFooter(slide, "상권 특성");
}

// ------------------------------------------------------------------
// 슬라이드 13 — 상권 트렌드 지표 (SNS 데이터 기반)
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "상권 트렌드 지표");
  slide.addText("SNS 데이터로 보는 상권 트렌드", { x: 0.6, y: 1.05, w: 11, h: 0.3, fontFace: FONT_BODY, fontSize: 12, color: MUTED });

  const inds = REPORT_DATA.trend.indicators;
  const cols = inds.length, gap = 0.25;
  const cardW = (12.1 - gap * (cols - 1)) / cols;
  inds.forEach((s, i) => {
    const x = 0.6 + i * (cardW + gap);
    statCard(slide, x, 1.5, cardW, 1.55, { value: s.value, label: s.label, sub: s.sub });
  });

  const insights = REPORT_DATA.trend.insights;
  const iw = (12.1 - 0.3) / 2;
  insights.forEach((ins, i) => {
    const x = 0.6 + i * (iw + 0.3);
    card(slide, { x, y: 3.4, w: iw, h: 3.5 });
    slide.addText(ins.title, { x: x + 0.3, y: 3.65, w: iw - 0.6, h: 0.45, fontFace: FONT_HEAD, fontSize: 15, bold: true, color: NAVY });
    slide.addText(ins.desc, {
      x: x + 0.3, y: 4.2, w: iw - 0.6, h: 2.5, fontFace: FONT_BODY, fontSize: 12.5, color: TEXT_DARK, lineSpacingMultiple: 1.35, valign: "top",
    });
  });
  addPageFooter(slide, "상권 트렌드 지표");
}

// ------------------------------------------------------------------
// 슬라이드 14 — 개발호재 종합
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "개발호재 종합");
  slide.addText("상권을 둘러싸고 동시에 진행 중인 교통·제도·트렌드 호재", {
    x: 0.6, y: 1.05, w: 11, h: 0.3, fontFace: FONT_BODY, fontSize: 12, color: MUTED,
  });

  const drivers = REPORT_DATA.growthDrivers;
  const cols = 3, gap = 0.25;
  const cardW = (12.1 - gap * (cols - 1)) / cols;
  const cardH = 2.5, startY = 1.55;
  drivers.forEach((d, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = 0.6 + col * (cardW + gap), y = startY + row * (cardH + gap);
    card(slide, { x, y, w: cardW, h: cardH });
    slide.addShape("ellipse", { x: x + 0.25, y: y + 0.25, w: 0.5, h: 0.5, fill: { color: NAVY }, line: { type: "none" } });
    slide.addText(String(i + 1), {
      x: x + 0.25, y: y + 0.25, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 18, bold: true, color: WHITE,
    });
    slide.addText(d.title, { x: x + 0.25, y: y + 0.9, w: cardW - 0.5, h: 0.4, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: NAVY });
    slide.addText(d.desc, {
      x: x + 0.25, y: y + 1.35, w: cardW - 0.5, h: cardH - 1.5, fontFace: FONT_BODY, fontSize: 10.5,
      color: TEXT_DARK, lineSpacingMultiple: 1.25, valign: "top",
    });
  });
  addPageFooter(slide, "개발호재 종합");
}

// ------------------------------------------------------------------
// 슬라이드 15 — 서울 상권 상세 (서울 소재 주소만)
// ------------------------------------------------------------------
if (REPORT_DATA.isSeoul) {
  const slide = newSlideBase(pres);
  addSectionTitle(slide, `서울 상권 상세 — ${REPORT_DATA.seoulTradeArea.name}`);

  const stats = REPORT_DATA.seoulTradeArea.stats;
  const cols = 3, gap = 0.25;
  const cardW = (12.1 - gap * (cols - 1)) / cols;
  stats.forEach((stat, i) => {
    const x = 0.6 + i * (cardW + gap);
    statCard(slide, x, 1.3, cardW, 1.2, { value: stat.value, label: stat.label });
  });

  slide.addText("업종별 매출 Top 5", { x: 0.6, y: 2.85, w: 5.6, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addTable(
    [["업종", "당월 매출"].map(t => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11 } }))]
      .concat(REPORT_DATA.seoulTradeArea.topIndustries.map(r => r.map(t => ({ text: t, options: { fontSize: 11.5, color: TEXT_DARK } })))),
    { x: 0.6, y: 3.25, w: 5.6, h: 3.0, fontFace: FONT_BODY, border: { type: "solid", color: GRID_LINE, pt: 0.75 }, colW: [3.6, 2.0], rowH: 0.5 }
  );

  slide.addText("요일별 매출 (억원)", { x: 6.6, y: 2.85, w: 6.1, h: 0.35, fontFace: FONT_BODY, fontSize: 13, bold: true, color: NAVY });
  slide.addChart("bar", [{
    name: "매출(억원)",
    labels: REPORT_DATA.seoulTradeArea.weekday.map(d => d.label),
    values: REPORT_DATA.seoulTradeArea.weekday.map(d => d.value),
  }], {
    x: 6.6, y: 3.25, w: 6.1, h: 3.0,
    chartColors: [TERRACOTTA], showTitle: false, showValue: true, dataLabelPosition: "outEnd",
    dataLabelColor: TEXT_DARK, dataLabelFontSize: 9, dataLabelFormatCode: "0.0",
    ...chartAxisDefaults(),
  });
  addPageFooter(slide, "서울 상권 상세");
}

// ------------------------------------------------------------------
// 슬라이드 16 — SWOT 분석
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres);
  addSectionTitle(slide, "SWOT 분석");
  slide.addText(`상권 및 건물가치 종합 진단 · ${REPORT_DATA.reportDate}`, {
    x: 0.6, y: 1.05, w: 11, h: 0.3, fontFace: FONT_BODY, fontSize: 12, color: MUTED,
  });

  const quadW = (12.1 - 0.25) / 2, quadH = (5.5 - 0.25) / 2;
  const quads = [
    { key: "S", title: "Strengths 강점", bg: "E3EEE6", accent: "3F7A5C", items: REPORT_DATA.swot.strengths, x: 0.6, y: 1.5 },
    { key: "W", title: "Weaknesses 약점", bg: "F5E6DE", accent: TERRACOTTA, items: REPORT_DATA.swot.weaknesses, x: 0.6 + quadW + 0.25, y: 1.5 },
    { key: "O", title: "Opportunities 기회", bg: "E3E9F5", accent: NAVY, items: REPORT_DATA.swot.opportunities, x: 0.6, y: 1.5 + quadH + 0.25 },
    { key: "T", title: "Threats 위협", bg: "FBEFD9", accent: "A66A1E", items: REPORT_DATA.swot.threats, x: 0.6 + quadW + 0.25, y: 1.5 + quadH + 0.25 },
  ];
  quads.forEach((q) => {
    slide.addShape("roundRect", { x: q.x, y: q.y, w: quadW, h: quadH, rectRadius: 0.1, fill: { color: q.bg }, line: { type: "none" } });
    slide.addShape("ellipse", { x: q.x + 0.25, y: q.y + 0.22, w: 0.44, h: 0.44, fill: { color: q.accent }, line: { type: "none" } });
    slide.addText(q.key, {
      x: q.x + 0.25, y: q.y + 0.22, w: 0.44, h: 0.44, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 18, bold: true, color: WHITE,
    });
    slide.addText(q.title, {
      x: q.x + 0.8, y: q.y + 0.22, w: quadW - 1.0, h: 0.44, fontFace: FONT_HEAD, fontSize: 14, bold: true, color: q.accent, valign: "middle",
    });
    let iy = q.y + 0.85;
    q.items.forEach((item) => {
      slide.addShape("ellipse", { x: q.x + 0.3, y: iy + 0.08, w: 0.08, h: 0.08, fill: { color: q.accent }, line: { type: "none" } });
      slide.addText(item, {
        x: q.x + 0.5, y: iy, w: quadW - 0.8, h: 0.55, fontFace: FONT_BODY, fontSize: 10.5, color: TEXT_DARK, valign: "top", lineSpacingMultiple: 1.15,
      });
      iy += 0.58;
    });
  });
  addPageFooter(slide, "SWOT 분석");
}

// ------------------------------------------------------------------
// 슬라이드 17 — 종합 의견
// ------------------------------------------------------------------
{
  const slide = newSlideBase(pres, { dark: true });
  slide.addShape("ellipse", { x: -2.5, y: 4.2, w: 6.5, h: 6.5, fill: { color: NAVY_LIGHT }, line: { type: "none" } });

  slide.addText("종합 의견", {
    x: 0.9, y: 0.7, w: 8, h: 0.7, fontFace: FONT_HEAD, fontSize: 32, bold: true, color: WHITE,
  });

  let cy = 1.9;
  REPORT_DATA.conclusion.forEach((point, i) => {
    slide.addShape("ellipse", { x: 0.9, y: cy + 0.05, w: 0.36, h: 0.36, fill: { color: TERRACOTTA }, line: { type: "none" } });
    slide.addText(String(i + 1), { x: 0.9, y: cy + 0.05, w: 0.36, h: 0.36, align: "center", valign: "middle", fontFace: FONT_HEAD, fontSize: 14, bold: true, color: WHITE });
    slide.addText(point, {
      x: 1.5, y: cy, w: 10.6, h: 0.9, valign: "top",
      fontFace: FONT_BODY, fontSize: 15, color: ICE, lineSpacingMultiple: 1.3,
    });
    cy += 1.15;
  });

  slide.addText(`${REPORT_DATA.address}  ·  ${REPORT_DATA.reportDate}`, {
    x: 0.9, y: 6.9, w: 10, h: 0.4, fontFace: FONT_BODY, fontSize: 10, color: MUTED,
  });
}

pres.writeFile({ fileName: "report_template.pptx" }).then(() => {
  console.log("wrote report_template.pptx");
});
