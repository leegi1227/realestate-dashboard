# 지도 렌더링 전환 설계 (카카오맵 → Leaflet+OpenStreetMap → +브이월드)

> **최종 결과 요약**: 처음 요청은 "카카오맵으로 변경"이었으나, 실제 배포
> 환경(Streamlit Community Cloud)에서 유효한 키·정확히 등록된 도메인 조합으로도
> 카카오맵 JS SDK 요청이 브라우저의 ORB(Opaque Response Blocking)에 막혀 계속
> 실패했다. 사용자가 "다른 지도라도 불러와"로 요청을 바꿔, API 키/도메인 등록이
> 아예 필요 없는 **Leaflet + OpenStreetMap**으로 구현했다. 그 다음 사용자가
> "브이월드 지도로 만들어줘"로 다시 요청해서, Leaflet은 그대로 두고 타일
> 소스만 브이월드로 바꿨다(둘 다 Leaflet의 `L.tileLayer`가 그냥 래스터
> 이미지 URL을 받는 것뿐이라 지도 라이브러리를 바꿀 필요는 없었음). 브이월드
> 키가 없으면 OpenStreetMap으로 자동 대체된다. 아래 문서는 그 과정에서 내린
> 결정들을 시간순으로 남긴다 — 카카오맵 관련 내용은 "왜 안 됐는지"의 기록으로
> 유효하다.

## 배경

기존 `dashboard.py`의 `render_address_map()`은 pydeck(deck.gl)으로 지도를 그리며,
배경지도는 기본 Carto 또는 사이드바에 입력한 브이월드(V-World) 키를 쓰는 래스터
타일이었다.

## 결정 사항 (시간순)

1. 카카오맵 JS SDK로 전환 (사용자 최초 요청). 브이월드 키 입력란/기능 제거
   (카카오맵 자체 배경지도로 대체되어 불필요하다고 판단 — 이후 4번에서 다시
   뒤집힘).
2. 지도 업로드 탭의 "지도 마커 클릭 → 표 행 강조" 양방향 상호작용은 계속 유지.
3. **카카오맵 → Leaflet+OpenStreetMap으로 재전환** (아래 "카카오맵이 실패한 이유"
   참고). 카카오맵 REST API 키(지오코딩용, `kakao_key`)는 지도 표시와 무관한
   별개 기능이라 그대로 유지. "카카오맵 JavaScript 키" 사이드바 입력란은 제거.
4. **브이월드 배경지도 옵션을 되살림** (사용자 요청). 1번에서 지웠던 "브이월드
   인증키" 사이드바 입력란과 `render_address_map()`의 `vworld_key` 파라미터를
   원래 있던 문구 그대로 복원했다. Leaflet 자체는 그대로 두고, `L.tileLayer`가
   받는 타일 URL만 `vworld_key` 유무에 따라 브이월드 WMTS ↔ OpenStreetMap으로
   바꾼다 — 지도 라이브러리를 다시 바꿀 필요는 없었다.

## 아키텍처

- pydeck 기반 렌더링을 제거하고, 지도를 삽입하는 **Custom Components v2**
  (`st.components.v2.component(name, html=, css=, js=)`) 컴포넌트로 교체했다.

  > **구현 중 변경 1 (컴포넌트 프로토콜: v1 → v2)**: 최초 설계는
  > `st.components.v1.declare_component(path=...)` + 직접 구현한 iframe
  > `postMessage` 핸드셰이크(`streamlit:componentReady` 등)였다. 로컬 검증 중
  > 이 조합이 이 Streamlit 버전(1.60)에서 **최초 마운트 시 컴포넌트가 준비
  > 신호를 반복 전송해도 렌더 이벤트를 영영 받지 못하고 조용히 빈 화면으로
  > 남는** 문제를 재현했다(Streamlit 자체에 번들된
  > `.agents/skills/developing-with-streamlit/references/ccv2-troubleshooting.md`
  > 문서에도 "v1 오염 — 가장 흔한 실패 증상: 컴포넌트가 빈 iframe으로 렌더되고
  > 파이썬과 통신이 안 됨"으로 정확히 명시돼 있었다). 같은 문서가 v1은 legacy이며
  > 신규 컴포넌트는 반드시 v2를 쓰라고 명시하고 있어 v2로 전환했다. v2는 iframe이
  > 아니라 같은 페이지의 섀도우 DOM에 직접 마운트되고 준비 핸드셰이크 자체가
  > 없어 이 문제가 구조적으로 발생하지 않는다 — 전환 직후 첫 시도부터 정상
  > 동작을 확인했다.

  > **구현 중 변경 2 (지도 제공자: 카카오맵 → Leaflet+OSM)**: v2로 전환한 뒤
  > 로컬 검증은 전부 통과했지만, 실제 Streamlit Community Cloud 배포본에서
  > 사용자가 유효한 카카오 JS 키 + 정확히 등록된 도메인으로도 계속 "SDK 스크립트를
  > 불러오지 못했습니다" 에러를 만났다. Playwright로 배포본을 직접 재현해보니
  > 원인은 브라우저 네트워크 탭의 `net::ERR_BLOCKED_BY_ORB`였다 — 카카오
  > SDK 엔드포인트가 (예: 형식이 안 맞는 키에 대해) `Content-Type:
  > application/json`인 401 에러 응답을 돌려주면, `<script>` 태그로 요청한
  > 리소스치고는 응답 형식이 이상하다고 판단한 크롬이 스펙터 완화 목적의
  > ORB로 그 응답 자체를 페이지에서 아예 못 읽게 차단한다. 이 때문에 JS
  > 쪽에서는 카카오가 정확히 어떤 이유로 거절했는지(키 오류/도메인 미등록/
  > 상품 미활성화 등) 구분할 방법이 없었고, 사용자와 함께 도메인 등록·키
  > 종류를 하나씩 확인해도 계속 같은 뭉뚱그려진 에러만 재현됐다. 사용자가
  > "다른 지도라도 불러와"로 요청을 바꿔, API 키·도메인 등록이 전혀 필요 없는
  > Leaflet(지도 라이브러리, CDN에서 무료로 로드) + OpenStreetMap(무료 타일)
  > 조합으로 교체했다 — 이 조합은 애초에 인증이 필요 없어 같은 종류의 실패가
  > 구조적으로 불가능하다.

- 컴포넌트 정의는 `address_map_component/`(카카오 전용이 아니게 된 후 이름을
  `kakao_map_component`에서 변경) 아래 `template.html`(내부 마크업),
  `style.css`, `script.js`(ES 모듈, `export default function(component)`) 세
  파일로 두고, `dashboard.py`가 모듈 임포트 시점에 파일 내용을 읽어 문자열로
  `st.components.v2.component(...)`에 전달한다(경로가 아니라 내용 자체를 넘겨야
  "인라인 콘텐츠"로 확실히 인식됨 — CCv2는 여러 줄 문자열을 항상 인라인으로
  취급).
- JS → 파이썬 값 전달은 v2의 `setTriggerValue("selected", {...})` + 마운트 시
  `on_selected_change=lambda: None` 콜백 등록 + 반환된 `result.selected`
  조합을 쓴다.
- `render_address_map()`은 지도 제공자 관련 파라미터(`vworld_key`, 이후
  `kakao_js_key`)를 전부 제거했다 — Leaflet+OSM은 아무 키도 필요 없다. 4곳의
  호출부 모두 지도 관련 인자 없이 호출한다.

## 데이터 흐름

1. 파이썬에서 기존과 동일하게 좌표 중복 제거("외 N건" 라벨), 표시할 주소 텍스트를
   계산한다.
2. 라벨 겹침 방지를 위해 기존 `_mercator_pixel` / `_fit_zoom` /
   `_declutter_label_levels`를 그대로 재사용해 마커별 세로 스택 단계(level)를
   구한다. `_fit_zoom`이 반환하는 표준 웹 메르카토르 줌은 Leaflet의 줌 레벨과
   같은 체계라 단일 지점 표시 시의 줌(`singleZoom`)으로 그대로 재사용한다
   (카카오맵 때는 반대로 매기는 "레벨" 체계라 못 썼음).
3. 마커 레코드 목록(`lat`, `lon`, `label`, `offsetLevel`), 강조 좌표
   (`highlight`), `enableSelection` 여부, `singleZoom`을 `data=` 딕셔너리로
   컴포넌트에 전달한다.
4. `script.js`의 `export default function(component)`:
   - Leaflet JS/CSS를 CDN(unpkg)에서 동적 로드(모듈 스코프에 로드 Promise를
     캐시해 컴포넌트 인스턴스가 여러 개여도 한 번만 로드).
   - 지점이 2개 이상이면 `L.latLngBounds`+`fitBounds`로 자동 범위 맞춤, 1개면
     `singleZoom` 레벨로 표시.
   - 기존과 동일한 빨간 핀 SVG를 `L.icon`/`L.marker`로, 주소 라벨은
     `marker.bindTooltip(..., {permanent: true, offset: [...]})`으로 표시하며
     `offsetLevel`만큼 세로로 밀어 배치.
   - 강조 좌표가 있으면 `L.circleMarker`(픽셀 단위 반지름)로 노란 반투명 원을
     표시.
   - `enableSelection=true`면 각 마커 클릭 시
     `setTriggerValue("selected", {"주소": ..., "lat": ..., "lon": ...})`
     형태(기존 pydeck 선택 결과와 동일한 키)로 값을 돌려보낸다.
   - **버그 수정**: 마커 클릭(`setTriggerValue`) 자체가 Streamlit 재실행을
     유발해 같은 DOM 컨테이너에 대해 이 함수가 다시 호출되는데, 이전 `L.map()`
     인스턴스를 `remove()`하지 않고 새로 만들면 Leaflet이 "Map container is
     already initialized" 에러를 던진다. 컨테이너별로 마지막 map 인스턴스를
     `WeakMap`에 저장해두고, 재렌더링 시 이전 인스턴스를 `remove()`한 뒤 새로
     만들도록 수정했다(Playwright로 마커를 연속 클릭해 재현 및 수정 확인).
5. 파이썬 래퍼 함수의 반환값은 기존과 동일한 dict/None이라(`result.selected`),
   지도 업로드 탭에서 `selected.get("lat")` 등을 쓰는 다운스트림 코드는 수정
   없이 그대로 동작한다.

## 에러 처리

- Leaflet/OSM은 키가 필요 없어 "키 없음" 안내 캡션 자체가 사라졌다 — 항상
  지도를 그린다.
- Leaflet 스크립트 로드가 실패하면(순수 네트워크 문제) 빈 화면 대신, 컴포넌트
  내부 에러 배너에 원인을 표시한다.

## 테스트 계획 (전부 실행 완료)

- 이 프로젝트에는 자동화 테스트가 없다 (확인됨: `test_*` 패턴/`pytest`/`unittest`
  사용 없음).
- 로컬에서 `streamlit run dashboard.py` + Playwright로 실제 렌더링을 확인했다:
  - 실제 OpenStreetMap 타일, 빨간 핀 마커, 흰 배경 라벨이 화면에 정상 표시됨
    (스크린샷으로 육안 확인).
  - 마커를 실제로 클릭(JS `dispatchEvent`)해서 지도 업로드 탭의
    "🔎 지도에서 OO을(를) 클릭했습니다 — 아래 표에서 강조 표시된 행입니다"
    캡션이 뜨는 것까지 왕복 확인.
  - 마커를 연속으로 두 번 클릭해도(재실행 반복) 에러 없이 정상 동작하는 것을
    확인(위 "Map container is already initialized" 수정 검증).
  - 브라우저 콘솔 에러/경고 0건.
- 표 → 지도 방향 강조(halo)는 구조적으로 동일한 `data.highlight` 전달 경로를
  타므로(halo 좌표 계산 로직 자체는 이번에 변경하지 않음) 별도 상호작용
  테스트는 생략했다 — 사용자가 실제 사용 중 이상이 있으면 알려달라고 안내.
- **성능 버그 발견 및 수정**: Streamlit은 이 앱 어디를 조작하든 전체 스크립트를
  재실행하고, CCv2는 그때마다 마운트된 컴포넌트의 JS를 다시 부른다 — 지도
  컴포넌트가 데이터 변화와 무관하게 매번 지도를 통째로 부수고 새로 만들고
  있어서(OSM 타일 재요청 포함) "느리다"는 사용자 보고로 발견했다. 컨테이너별로
  마지막에 그린 data를 캐시해두고 실제로 안 바뀌었으면 다시 그리지 않도록
  고쳤다. Playwright로 지도와 무관한 사이드바 입력을 바꿔도 Leaflet map
  인스턴스(`_leaflet_id`)와 타일 DOM이 그대로 유지되는 것으로 검증.
- 브이월드 타일 전환 검증: 더미 키로 실제 브이월드 WMTS 엔드포인트에 요청이
  정확한 URL로 나가는 것을 확인했고, 브이월드가 잘못된 키에 어떻게 응답하는지도
  직접 확인했다(`Content-Type: application/xml`인 200 OK, OGC
  `ExceptionReport`, "등록되지 않은 인증키입니다") — `<img>` 태그로 받는
  래스터 타일이라 카카오 SDK 때와 달리 ORB에 막히지 않고, 그냥 빈 타일로
  깨질 뿐 지도 전체가 죽지는 않는다(스크린샷으로 확인). 실제 유효한 키로
  브이월드 지도가 정상 표시되는지는 사용자가 직접 확인해야 한다.

## 범위 밖

- 마커 클러스터링, 커스텀 지도 스킨/테마 등 기존에 없던 신규 기능은 추가하지 않는다.
- REST API 지오코딩(`geocode_address_kakao`) 로직은 변경하지 않는다.
