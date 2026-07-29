// 기존 pydeck 버전과 동일한 빨간 핀 SVG.
const PIN_SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">'
  + '<path d="M32 2C19 2 8 13 8 26c0 18 24 36 24 36s24-18 24-36C56 13 45 2 32 2z" fill="#DC2626" stroke="#FFFFFF" stroke-width="2"/>'
  + '<circle cx="32" cy="26" r="9" fill="#FFFFFF"/></svg>';
const PIN_IMAGE_URL = "data:image/svg+xml;base64," + btoa(PIN_SVG);

// 이 모듈은 kakao_map 컴포넌트의 모든 인스턴스(대시보드의 여러 탭)가 공유한다 —
// SDK 로드는 앱 전체에서 한 번만 하면 되므로 모듈 스코프(함수 바깥)에 캐시한다.
let kakaoSdkPromise = null;
let kakaoSdkKey = null;

function loadKakaoSdk(appKey) {
  if (kakaoSdkPromise && kakaoSdkKey === appKey) return kakaoSdkPromise;
  kakaoSdkKey = appKey;
  kakaoSdkPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://dapi.kakao.com/v2/maps/sdk.js?appkey=" + encodeURIComponent(appKey) + "&autoload=false";
    script.onload = () => {
      try {
        window.kakao.maps.load(() => resolve());
      } catch (e) {
        reject(e);
      }
    };
    script.onerror = () => reject(new Error("SDK 스크립트를 불러오지 못했습니다 (네트워크 오류 또는 키 오류)"));
    document.head.appendChild(script);
  });
  return kakaoSdkPromise;
}

function renderMap(mapEl, data) {
  mapEl.innerHTML = "";
  const markers = data.markers || [];
  const first = markers[0];

  const map = new kakao.maps.Map(mapEl, {
    center: new kakao.maps.LatLng(first ? first.lat : 37.5665, first ? first.lon : 126.9780),
    level: 3,
  });

  if (markers.length > 1) {
    const bounds = new kakao.maps.LatLngBounds();
    markers.forEach((m) => bounds.extend(new kakao.maps.LatLng(m.lat, m.lon)));
    map.setBounds(bounds, 60, 60, 60, 60);
  }

  const markerImage = new kakao.maps.MarkerImage(
    PIN_IMAGE_URL,
    new kakao.maps.Size(36, 36),
    { offset: new kakao.maps.Point(18, 36) },
  );

  return { map, markerImage };
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

  if (!data || !data.kakaoJsKey) {
    showError("카카오맵 JavaScript 키가 없습니다.");
    return;
  }

  loadKakaoSdk(data.kakaoJsKey)
    .then(() => {
      clearError();
      const { map, markerImage } = renderMap(mapEl, data);
      const markers = data.markers || [];

      markers.forEach((m) => {
        const pos = new kakao.maps.LatLng(m.lat, m.lon);
        const marker = new kakao.maps.Marker({ position: pos, image: markerImage, map });

        if (data.showLabels && m.label) {
          const labelEl = document.createElement("div");
          labelEl.className = "kmc-label";
          labelEl.textContent = m.label;
          const dy = 24 + (m.offsetLevel || 0) * 18;
          labelEl.style.transform = `translate(12px, -${dy}px)`;
          new kakao.maps.CustomOverlay({
            position: pos, content: labelEl, xAnchor: 0, yAnchor: 0.5, zIndex: 2,
          }).setMap(map);
        }

        if (data.enableSelection) {
          kakao.maps.event.addListener(marker, "click", () => {
            setTriggerValue("selected", { "주소": m.label, lat: m.lat, lon: m.lon });
          });
        }
      });

      if (data.highlight) {
        const haloEl = document.createElement("div");
        haloEl.className = "kmc-halo";
        new kakao.maps.CustomOverlay({
          position: new kakao.maps.LatLng(data.highlight.lat, data.highlight.lon),
          content: haloEl, xAnchor: 0.5, yAnchor: 0.5, zIndex: 1,
        }).setMap(map);
      }
    })
    .catch((err) => {
      showError("카카오맵을 불러오지 못했습니다: " + (err && err.message ? err.message : err));
    });
}
