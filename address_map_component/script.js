// 기존 pydeck 버전과 동일한 빨간 핀 SVG.
const PIN_SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">'
  + '<path d="M32 2C19 2 8 13 8 26c0 18 24 36 24 36s24-18 24-36C56 13 45 2 32 2z" fill="#DC2626" stroke="#FFFFFF" stroke-width="2"/>'
  + '<circle cx="32" cy="26" r="9" fill="#FFFFFF"/></svg>';
const PIN_IMAGE_URL = "data:image/svg+xml;base64," + btoa(PIN_SVG);

const LEAFLET_CSS_URL = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
const LEAFLET_JS_URL = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";

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

// 컴포넌트 인스턴스(mapEl)별로 마지막에 만든 Leaflet map을 기억해둔다 — 마커
// 클릭(setTriggerValue) 자체가 Streamlit 재실행을 유발해서 같은 mapEl에 대해
// 이 함수가 다시 호출되는데, 이전 map을 remove() 없이 L.map()을 다시 호출하면
// Leaflet이 "Map container is already initialized" 에러를 던진다.
const mapInstances = new WeakMap();

// data와 setTriggerValue를 그대로 인자로 받아 클로저에 가둔다 — 컴포넌트
// 인스턴스가 여러 개(대시보드 탭마다 하나씩)라도 서로의 콜백을 침범하지 않는다.
function renderMap(L, mapEl, data, setTriggerValue) {
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
