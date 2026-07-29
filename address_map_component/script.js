// 기존 pydeck 버전과 동일한 빨간 핀 SVG.
const PIN_SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">'
  + '<path d="M32 2C19 2 8 13 8 26c0 18 24 36 24 36s24-18 24-36C56 13 45 2 32 2z" fill="#DC2626" stroke="#FFFFFF" stroke-width="2"/>'
  + '<circle cx="32" cy="26" r="9" fill="#FFFFFF"/></svg>';
const PIN_IMAGE_URL = "data:image/svg+xml;base64," + btoa(PIN_SVG);

const LEAFLET_CSS_URL = "https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css";
const LEAFLET_JS_URL = "https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js";

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
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: "&copy; OpenStreetMap contributors" + (attributionSuffix || ""),
    }).addTo(map);
  }

  if (data.vworldKey) {
    // 브이월드 배경지도(WMTS). 경로가 {z}/{y}/{x} 순서인 건 브이월드 쪽 규격 —
    // Leaflet은 URL 안의 플레이스홀더를 그대로 치환할 뿐이라 순서를 바꿔 써도 된다.
    const vworldLayer = L.tileLayer(
      `https://api.vworld.kr/req/wmts/1.0.0/${data.vworldKey}/Base/{z}/{y}/{x}.png`,
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
