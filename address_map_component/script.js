// 기존 pydeck 버전과 동일한 빨간 핀 SVG.
const PIN_SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">'
  + '<path d="M32 2C19 2 8 13 8 26c0 18 24 36 24 36s24-18 24-36C56 13 45 2 32 2z" fill="#DC2626" stroke="#FFFFFF" stroke-width="2"/>'
  + '<circle cx="32" cy="26" r="9" fill="#FFFFFF"/></svg>';
const PIN_IMAGE_URL = "data:image/svg+xml;base64," + btoa(PIN_SVG);

const LEAFLET_CSS_URL = "https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css";
const LEAFLET_JS_URL = "https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js";

// 서울 지하철역 목록: [역명, 위도, 경도, [[노선번호/약칭, 노선색], ...]]. OpenStreetMap의
// route=subway/light_rail/train(수도권) relation에서 정거장(stop) 멤버를 모아 역명으로
// 묶고, 각 relation의 ref/colour(공식 노선 번호·색)를 그대로 붙인 것 — dashboard.py가
// import 시점에 subway_stations.json 내용을 아래 자리에 그대로 채워 넣는다(매 rerun마다
// 35KB를 다시 보내지 않도록, data=가 아니라 컴포넌트 js= 문자열에 정적으로 박아 넣는
// 방식). 카카오맵/네이버처럼 역마다 노선색 원형 배지로 표시하는 데 쓴다.
const SUBWAY_STATIONS = "__SUBWAY_STATIONS_JSON__";

// 노선 실제 선형: [노선번호/약칭, 노선색, [[[위도,경도],...] 구간, ...]]. route relation의
// way(선로) 멤버 geometry를 그대로 이어붙인 것 — 여러 구간(segments)으로 나뉘어 있는 건
// way가 원래 조각조각이기 때문이며 Leaflet은 끊어진 구간을 그대로 여러 선으로 그려도
// 자연스럽게 이어져 보인다. 좌표점이 과도하게 많아(원본 3만 개) Douglas-Peucker로
// 압축했다(허용 오차 약 11m — 지도에서 보이는 굵은 선 두께보다 작아 시각적으로 거의
// 차이가 없다).
const SUBWAY_LINES = "__SUBWAY_LINES_JSON__";

// 출구: [[위도, 경도, 출구번호], ...]. railway=subway_entrance 노드의 ref 태그 —
// 카카오맵처럼 역명과 무관하게 항상 같은 중립색 원으로 표시한다(출구 자체는 특정
// 노선 색이 없으므로).
const SUBWAY_EXITS = "__SUBWAY_EXITS_JSON__";

// Leaflet의 JS 라이브러리(window.L)는 페이지에 하나만 있으면 되는 전역이라
// 모듈 스코프에 캐시해서 앱 전체(탭이 여러 개라 컴포넌트 인스턴스도 여러 개)가
// 공유한다 — CSS와 달리 섀도 DOM 경계와 무관하게 그냥 전역 객체라 공유해도 된다.
let leafletJsPromise = null;

function loadLeafletJs() {
  if (leafletJsPromise) return leafletJsPromise;
  leafletJsPromise = window.L
    ? Promise.resolve(window.L)
    : new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = LEAFLET_JS_URL;
        script.onload = () => resolve(window.L);
        script.onerror = () => reject(new Error("Leaflet 스크립트를 불러오지 못했습니다 (네트워크 오류)"));
        document.head.appendChild(script);
      });
  return leafletJsPromise;
}

// CSS는 JS와 달리 공유할 수 없다: CCv2 컴포넌트는 기본적으로(isolate_styles=True)
// 각자 자기만의 섀도 루트에 마운트되고, document.head에 넣은 <link>/<style>은
// 섀도 DOM 경계를 넘지 못해 그 안의 엘리먼트(.leaflet-container 등)에는 전혀
// 적용되지 않는다 — 이게 바로 지도가 찌그러져 보이던 진짜 원인이었다(CSS 로드
// "타이밍"이 아니라 애초에 적용될 수 없는 위치에 넣고 있었음: 브라우저는 계속
// 정상적으로 CSS를 받아왔지만 그 CSS가 유효한 범위가 이 컴포넌트의 섀도 루트가
// 아니었을 뿐). 그래서 CSS는 인스턴스별 섀도 루트(root)에 각각 주입해야 한다.
const cssLoadedRoots = new WeakSet();

function ensureLeafletCss(root) {
  if (cssLoadedRoots.has(root)) return Promise.resolve();
  return new Promise((resolve) => {
    const existing = root.querySelector(`link[href="${LEAFLET_CSS_URL}"]`);
    if (existing) {
      cssLoadedRoots.add(root);
      if (existing.sheet) {
        resolve();
      } else {
        existing.addEventListener("load", () => resolve(), { once: true });
        existing.addEventListener("error", () => resolve(), { once: true });
      }
      return;
    }
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = LEAFLET_CSS_URL;
    link.onload = () => {
      cssLoadedRoots.add(root);
      resolve();
    };
    // CSS 로드 실패는 지도가 못생기게 나오는 정도지 기능이 죽는 건 아니므로,
    // JS 로드 실패(showError로 이어짐)와 달리 그냥 resolve해서 진행시킨다.
    link.onerror = () => {
      cssLoadedRoots.add(root);
      resolve();
    };
    root.appendChild(link);
  });
}

// 컴포넌트 인스턴스(mapEl)별로 마지막에 만든 Leaflet map과, 그걸 만들 때 쓴
// data를 기억해둔다. Streamlit은 이 앱 어디를 조작하든(사이드바 입력, 다른 탭의
// 위젯 등) 전체 스크립트를 재실행하고, 그때마다 이 컴포넌트도 다시 호출된다 —
// 지도에 실제로 영향을 주는 data가 그대로라면 다시 그릴 이유가 없다(오히려
// 매번 지도를 부수고 새로 만들면 OSM 타일을 매번 새로 받아오고, 사용자가
// 손으로 옮겨둔 팬/줌 위치도 계속 초기화돼서 매우 느리고 거슬리게 느껴진다).
const mapInstances = new WeakMap();
const lastRenderedData = new WeakMap();

function getSubwayStations() {
  // dashboard.py가 import 시점에 "__SUBWAY_STATIONS_JSON__" 자리(따옴표째로)를 실제
  // 배열 리터럴 텍스트로 치환해 넣으므로, 정상적인 경우 SUBWAY_STATIONS는 이미 배열이다
  // (JSON.parse 불필요). 치환이 어떤 이유로든 안 되면 문자열 그대로 남아 .forEach가
  // 없어 에러가 나는데, 그래도 지도 전체가 죽지 않도록 역 표시만 조용히 건너뛴다.
  return Array.isArray(SUBWAY_STATIONS) ? SUBWAY_STATIONS : [];
}

function addSubwayStations(L, map) {
  getSubwayStations().forEach(([name, lat, lon, lines]) => {
    const badges = lines
      .map(([ref, colour]) => {
        const label = [...ref].length > 2 ? [...ref].slice(0, 2).join("") : ref;
        // 주의: 이 문자열은 HTML style="..." 속성값(큰따옴표)이라, 안의 CSS에서
        // font-family를 큰따옴표로 감싸면("Malgun Gothic") 브라우저 HTML 파서가
        // 거기서 style 속성이 끝난 걸로 오해해 그 뒤(border/box-shadow/nowrap/
        // word-break 전부)가 조용히 엉뚱한 속성으로 깨진다(겉보기엔 에러 없이
        // 그냥 스타일만 안 먹힘 — 두 글자 라벨이 원 안에서 위아래로 쪼개져 보이던
        // 진짜 원인이 이거였다). CSS 문자열 쪽은 항상 작은따옴표로 감싼다.
        return (
          '<span style="display:inline-flex;align-items:center;justify-content:center;' +
          "width:22px;height:22px;border-radius:50%;background:" + colour + ";color:#fff;" +
          "font:bold 8.5px/1 -apple-system,'Malgun Gothic',sans-serif;border:1.5px solid #fff;" +
          "box-shadow:0 0 2px rgba(0,0,0,.6);margin-left:-6px;" +
          'white-space:nowrap;word-break:keep-all;overflow:visible;">' + label + "</span>"
        );
      })
      .join("");
    const icon = L.divIcon({
      className: "kmc-subway-icon",
      html: '<div style="display:flex;padding-left:5px;">' + badges + "</div>",
      iconSize: null,
    });
    L.marker([lat, lon], { icon, keyboard: false })
      .bindTooltip(name, { direction: "top", offset: [0, -10] })
      .addTo(map);
  });
}

function getSubwayLines() {
  return Array.isArray(SUBWAY_LINES) ? SUBWAY_LINES : [];
}

function getSubwayExits() {
  return Array.isArray(SUBWAY_EXITS) ? SUBWAY_EXITS : [];
}

function addSubwayLines(L, map) {
  // 역 배지(addSubwayStations)보다 먼저 그려서 배지가 선 위에 얹히게 한다.
  getSubwayLines().forEach(([ref, colour, segments]) => {
    segments.forEach((seg) => {
      L.polyline(seg, {
        color: colour,
        weight: 5,
        opacity: 0.85,
        lineCap: "round",
        lineJoin: "round",
      }).addTo(map);
    });
  });
}

function addSubwayExits(L, map) {
  const EXIT_COLOR = "#FFC400";
  getSubwayExits().forEach(([lat, lon, ref]) => {
    const icon = L.divIcon({
      className: "kmc-subway-icon",
      html:
        '<span style="display:flex;align-items:center;justify-content:center;' +
        "width:16px;height:16px;border-radius:50%;background:" + EXIT_COLOR + ";color:#222;" +
        "font:bold 9px/1 -apple-system,'Malgun Gothic',sans-serif;border:1.5px solid #fff;" +
        'box-shadow:0 0 2px rgba(0,0,0,.6);">' + ref + "</span>",
      iconSize: null,
    });
    L.marker([lat, lon], { icon, keyboard: false }).addTo(map);
  });
}

// data와 setTriggerValue를 그대로 인자로 받아 클로저에 가둔다 — 컴포넌트
// 인스턴스가 여러 개(대시보드 탭마다 하나씩)라도 서로의 콜백을 침범하지 않는다.
function renderMap(L, mapEl, data, setTriggerValue) {
  const dataJson = JSON.stringify(data);
  if (lastRenderedData.get(mapEl) === dataJson && mapInstances.has(mapEl)) {
    return mapInstances.get(mapEl);
  }
  lastRenderedData.set(mapEl, dataJson);

  const prevMap = mapInstances.get(mapEl);
  if (prevMap) {
    prevMap.remove();
  }
  mapEl.innerHTML = "";
  const markers = data.markers || [];
  const first = markers[0];

  const map = L.map(mapEl, { attributionControl: true }).setView(
    [first ? first.lat : 37.5665, first ? first.lon : 126.9780],
    data.singleZoom || 17,
  );
  function addOsmLayer(attributionSuffix) {
    // 표준 OSM 타일(tile.openstreetmap.org)은 상점·업종 아이콘/라벨(예: "F&B" 같은
    // 프랜차이즈·업종 텍스트)이 촘촘히 박혀 있어, 이미 지하철 노선·역·출구까지 얹은
    // 지도에서는 너무 산만해 보인다는 피드백을 받았다. 라벨을 최소화한 CartoDB
    // Positron("light_all") 스타일로 교체 — 데이터 출처는 동일하게 OSM이고 키도
    // 필요 없다(연남동 상권분석 PPTX 지도 작업에서도 같은 스타일을 썼다).
    L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}.png", {
      maxZoom: 19,
      subdomains: "abcd",
      attribution: "&copy; OpenStreetMap contributors &copy; CARTO" + (attributionSuffix || ""),
    }).addTo(map);
  }

  if (data.vworldKey) {
    // 브이월드 배경지도(WMTS). 경로가 {z}/{y}/{x} 순서인 건 브이월드 쪽 규격 —
    // Leaflet은 URL 안의 플레이스홀더를 그대로 치환할 뿐이라 순서를 바꿔 써도 된다.
    // "Base"(기본 참조지도)는 건물을 옅은 외곽선으로만 그려 건물이 잘 안 보인다는
    // 피드백을 받아 "Hybrid"(항공사진 + 라벨)로 바꿨다 — 실제 건물 형태가 항공
    // 사진으로 그대로 보인다. 다른 선택지: Satellite(라벨 없는 순수 항공사진),
    // White(최소한의 백지도), Midnight(Base의 야간 버전).
    const vworldLayer = L.tileLayer(
      `https://api.vworld.kr/req/wmts/1.0.0/${data.vworldKey}/Hybrid/{z}/{y}/{x}.png`,
      { maxZoom: 19, minZoom: 5, tileSize: 256, attribution: "&copy; VWorld" },
    );
    // 브이월드는 해외 클라우드(AWS/GCP 등) IP를 정책적으로 차단한다(공간정보관리법
    // 제16조 국외반출 제한 근거 — 지오코딩 API에서 먼저 확인된 것과 동일한 제한이
    // 지도 타일 API에도 적용되는 것으로 보인다). 이 경우 Streamlit Cloud 배포본에서는
    // 타일이 전부 회색 빈 화면으로 나온다. 키 오류든 이 차단이든, 타일이 연달아
    // 여러 장 실패하면(일시적 네트워크 hiccup과 구분하기 위해 즉시 전환하지 않고
    // 몇 장은 봐준다) 사용자가 빈 지도만 보는 것보다는 OSM으로 자동 전환해 최소한
    // 지도 자체는 항상 뜨게 한다.
    let vworldFailCount = 0;
    let fellBack = false;
    vworldLayer.on("tileerror", () => {
      vworldFailCount += 1;
      if (!fellBack && vworldFailCount >= 3) {
        fellBack = true;
        map.removeLayer(vworldLayer);
        addOsmLayer(" (브이월드 타일 로드 실패로 자동 전환됨 — 배포 환경에서는 해외 IP 차단 정책 때문일 수 있음)");
      }
    });
    vworldLayer.addTo(map);
  } else {
    addOsmLayer();
  }

  // 지하철 노선·출구·역 — 카카오맵/네이버맵처럼 굵은 노선색 선 + 출구 번호 원 +
  // 역마다 노선색 배지로 표시한다(SUBWAY_LINES/SUBWAY_EXITS/SUBWAY_STATIONS, 파일
  // 위쪽에서 정의). 이전엔 OpenRailwayMap 타일 오버레이를 썼는데, 실제 벡터 선을
  // 직접 그리게 되면서 타일까지 같이 켜두면 얇은 타일 선과 굵은 벡터 선이 겹쳐
  // 지저분해 보여 제거했다. 그리는 순서가 중요 — 선을 맨 밑에 깔고, 그 위에
  // 출구·역 마커를 얹어야 마커가 선에 가려지지 않는다.
  addSubwayLines(L, map);
  addSubwayExits(L, map);
  addSubwayStations(L, map);

  if (markers.length > 1) {
    const bounds = L.latLngBounds(markers.map((m) => [m.lat, m.lon]));
    map.fitBounds(bounds, { padding: [50, 50] });
  }

  const pinIcon = L.icon({
    iconUrl: PIN_IMAGE_URL,
    iconSize: [36, 36],
    iconAnchor: [18, 36],
  });

  markers.forEach((m) => {
    const marker = L.marker([m.lat, m.lon], { icon: pinIcon }).addTo(map);

    if (data.showLabels && m.label) {
      const dy = 24 + (m.offsetLevel || 0) * 18;
      marker.bindTooltip(m.label, {
        permanent: true,
        direction: "right",
        offset: [12, -dy],
        className: "kmc-label",
      }).openTooltip();
    }

    if (data.enableSelection) {
      marker.on("click", () => {
        setTriggerValue("selected", { "주소": m.label, lat: m.lat, lon: m.lon });
      });
    }
  });

  if (data.highlight) {
    L.circleMarker([data.highlight.lat, data.highlight.lon], {
      radius: 28,
      color: "rgba(255, 170, 0, 0.9)",
      weight: 3,
      fillColor: "rgba(255, 214, 0, 0.43)",
      fillOpacity: 1,
    }).addTo(map);
  }

  mapInstances.set(mapEl, map);

  // Leaflet은 L.map() 실행 시점의 컨테이너 크기를 그대로 굳혀서 타일/줌 컨트롤을
  // 배치한다 — 그 시점에 섀도우 DOM 레이아웃이나 CSS 적용이 아직 완전히
  // 끝나지 않았으면(특히 컴포넌트가 막 마운트된 첫 렌더) 실제 최종 크기와 어긋난
  // 채로 굳어버려 지도가 찌그러져 보일 수 있다. 다음 프레임에 실제 크기를 다시
  // 재보게 강제해 이 어긋남을 스스로 바로잡는다 — 크기가 이미 맞았으면 아무
  // 효과가 없는 안전한 호출이라 항상 걸어둔다.
  requestAnimationFrame(() => map.invalidateSize());

  return map;
}

// 전체화면 버튼에 클릭 리스너를 단 적이 있는 버튼 엘리먼트를 기억해둔다 —
// export default function은 rerun마다 다시 호출되는데, 버튼 자체는 #map-wrap의
// 형제 엘리먼트라 mapEl.innerHTML 초기화의 영향을 안 받고 계속 살아있으므로,
// 매번 리스너를 새로 달면 중복 등록돼 한 번 클릭에 여러 번 토글되는 버그가 난다.
const expandWired = new WeakSet();

function wireExpandButton(wrapEl, btnEl, getMap) {
  if (expandWired.has(btnEl)) return;
  expandWired.add(btnEl);
  btnEl.addEventListener("click", () => {
    const isFull = wrapEl.classList.toggle("kmc-fullscreen");
    btnEl.textContent = isFull ? "✕" : "⛶";
    btnEl.title = isFull ? "지도 원래 크기로" : "지도 크게 보기";
    const map = getMap();
    if (map) {
      // 크기가 바뀐 다음 프레임에 재보게 해야 Leaflet이 새 크기로 타일을 다시 배치한다.
      requestAnimationFrame(() => map.invalidateSize());
    }
  });
}

export default function (component) {
  const { data, parentElement, setTriggerValue } = component;
  const wrapEl = parentElement.querySelector("#map-wrap");
  const mapEl = parentElement.querySelector("#map");
  const expandBtn = parentElement.querySelector("#kmc-expand");
  const errorEl = parentElement.querySelector("#error");
  if (!mapEl || !errorEl) return;

  if (wrapEl && expandBtn) {
    wireExpandButton(wrapEl, expandBtn, () => mapInstances.get(mapEl));
  }

  function showError(msg) {
    errorEl.textContent = msg;
    errorEl.style.display = "block";
    mapEl.style.display = "none";
  }
  function clearError() {
    errorEl.style.display = "none";
    mapEl.style.display = "block";
  }

  Promise.all([loadLeafletJs(), ensureLeafletCss(parentElement)])
    .then(([L]) => {
      clearError();
      renderMap(L, mapEl, data || {}, setTriggerValue);
    })
    .catch((err) => {
      showError("지도를 불러오지 못했습니다: " + (err && err.message ? err.message : err));
    });
}
