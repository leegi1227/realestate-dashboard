// 기존 pydeck 버전과 동일한 빨간 핀 SVG.
const PIN_SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">'
  + '<path d="M32 2C19 2 8 13 8 26c0 18 24 36 24 36s24-18 24-36C56 13 45 2 32 2z" fill="#DC2626" stroke="#FFFFFF" stroke-width="2"/>'
  + '<circle cx="32" cy="26" r="9" fill="#FFFFFF"/></svg>';
const PIN_IMAGE_URL = "data:image/svg+xml;base64," + btoa(PIN_SVG);

const LEAFLET_CSS_URL = "https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css";
const LEAFLET_JS_URL = "https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js";

// 이 모듈은 address_map 컴포넌트의 모든 인스턴스(대시보드의 여러 탭)가 공유한다 —
// Leaflet 로드는 앱 전체에서 한 번만 하면 되므로 모듈 스코프(함수 바깥)에 캐시한다.
// 카카오맵과 달리 API 키/도메인 등록이 필요 없어(OpenStreetMap 무료 타일) 로드
// 실패 가능성이 훨씬 낮다.
let leafletLoadPromise = null;

function loadLeaflet() {
  if (leafletLoadPromise) return leafletLoadPromise;
  leafletLoadPromise = new Promise((resolve, reject) => {
    if (window.L) {
      resolve(window.L);
      return;
    }
    if (!document.querySelector(`link[href="${LEAFLET_CSS_URL}"]`)) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = LEAFLET_CSS_URL;
      document.head.appendChild(link);
    }
    const script = document.createElement("script");
    script.src = LEAFLET_JS_URL;
    script.onload = () => resolve(window.L);
    script.onerror = () => reject(new Error("Leaflet 스크립트를 불러오지 못했습니다 (네트워크 오류)"));
    document.head.appendChild(script);
  });
  return leafletLoadPromise;
}

// 컴포넌트 인스턴스(mapEl)별로 마지막에 만든 Leaflet map과, 그걸 만들 때 쓴
// data를 기억해둔다. Streamlit은 이 앱 어디를 조작하든(사이드바 입력, 다른 탭의
// 위젯 등) 전체 스크립트를 재실행하고, 그때마다 이 컴포넌트도 다시 호출된다 —
// 지도에 실제로 영향을 주는 data가 그대로라면 다시 그릴 이유가 없다(오히려
// 매번 지도를 부수고 새로 만들면 OSM 타일을 매번 새로 받아오고, 사용자가
// 손으로 옮겨둔 팬/줌 위치도 계속 초기화돼서 매우 느리고 거슬리게 느껴진다).
const mapInstances = new WeakMap();
const lastRenderedData = new WeakMap();

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
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

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
  return map;
}

export default function (component) {
  const { data, parentElement, setTriggerValue } = component;
  const mapEl = parentElement.querySelector("#map");
  const errorEl = parentElement.querySelector("#error");
  if (!mapEl || !errorEl) return;

  function showError(msg) {
    errorEl.textContent = msg;
    errorEl.style.display = "block";
    mapEl.style.display = "none";
  }
  function clearError() {
    errorEl.style.display = "none";
    mapEl.style.display = "block";
  }

  loadLeaflet()
    .then((L) => {
      clearError();
      renderMap(L, mapEl, data || {}, setTriggerValue);
    })
    .catch((err) => {
      showError("지도를 불러오지 못했습니다: " + (err && err.message ? err.message : err));
    });
}
