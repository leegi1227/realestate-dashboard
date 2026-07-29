# 카카오맵 전환 설계

## 배경

현재 `dashboard.py`의 `render_address_map()`은 pydeck(deck.gl)으로 지도를 그리며,
배경지도는 기본 Carto 또는 사이드바에 입력한 브이월드(V-World) 키를 쓰는 래스터
타일이다. 이를 카카오맵 JavaScript SDK 기반 지도로 교체한다.

## 결정 사항 (사용자 확인 완료)

- 카카오맵 JavaScript 키는 이미 발급 및 배포 도메인 등록 완료됨.
- 브이월드 키 입력란/기능은 제거한다 (카카오맵 자체 배경지도로 대체되어 불필요).
- 지도 업로드 탭의 "지도 마커 클릭 → 표 행 강조" 양방향 상호작용은 유지한다.

## 아키텍처

- 사이드바에 "카카오맵 JavaScript 키" 입력란을 신설한다 (`kakao_js_key`). 기존
  "카카오맵 REST API 키" (`kakao_key`, 지오코딩용)는 그대로 둔다. 두 키는 카카오
  개발자 콘솔에서 같은 앱 안에 별도로 발급되는 서로 다른 값이다.
- "브이월드 키" 입력란, `_vworld_basemap_style()` 함수, `render_address_map()`의
  `vworld_key` 파라미터, 그리고 모든 호출부의 `vworld_key=vworld_key` 인자를 제거한다.
- pydeck 기반 렌더링을 제거하고, 카카오맵 JS SDK를 삽입하는 **커스텀 Streamlit
  컴포넌트**로 교체한다. `st.components.v1.html()`(단방향)이 아니라
  `st.components.v1.declare_component(path=...)` 로 등록하는 정적 `index.html` 한
  장을 사용한다 — 빌드 과정 없이, Streamlit이 공개적으로 사용하는 iframe
  `postMessage` 프로토콜(`streamlit:componentReady` /
  `streamlit:render` / `streamlit:setFrameHeight` /
  `streamlit:setComponentValue`)을 직접 구현해 마커 클릭 값을 파이썬으로 돌려받는다.
- 새 디렉터리 `kakao_map_component/index.html` 1개 파일만 추가한다. 좌표/라벨
  데이터는 매 호출마다 컴포넌트 인자(args)로 전달되므로 이 파일은 재생성할 필요가
  없다.
- `render_address_map()`의 함수 시그니처는 최대한 유지한다
  (`vworld_key` 제거, `kakao_js_key` 추가). 5곳의 호출부는 인자 이름만 바뀐다.

## 데이터 흐름

1. 파이썬에서 기존과 동일하게 좌표 중복 제거("외 N건" 라벨), 표시할 주소 텍스트를
   계산한다.
2. 라벨 겹침 방지를 위해 기존 `_mercator_pixel` / `_fit_zoom` /
   `_declutter_label_levels`를 그대로 재사용해 마커별 세로 스택 단계(level)를
   구한다. 이 zoom 값은 실제 카카오맵이 표시할 줌과 정확히 같을 필요가 없는,
   라벨 간 픽셀 거리 추정용 내부 근사치일 뿐이다.
3. 마커 레코드 목록(`lat`, `lon`, `label`, `offset_level`), 강조 좌표
   (`highlight_lat`/`highlight_lon`), `enable_selection` 여부, 카카오 JS 키를
   컴포넌트 인자로 전달한다.
4. `index.html`:
   - `https://dapi.kakao.com/v2/maps/sdk.js?appkey=...&autoload=false` 로 SDK를
     동적 로드.
   - 지점이 2개 이상이면 `kakao.maps.LatLngBounds`로 자동 범위 맞춤, 1개면 건물
     단위 확대 레벨(레벨 3 고정)로 표시.
   - 기존과 동일한 빨간 핀 SVG를 `kakao.maps.MarkerImage`/`Marker`로, 주소
     라벨은 흰 배경 `kakao.maps.CustomOverlay`로 그리며 `offset_level`만큼
     세로로 밀어 배치(픽셀 단위 CSS transform).
   - 강조 좌표가 있으면 고정 픽셀 크기의 노란 반투명 원을 `CustomOverlay`로 표시.
   - `enable_selection=true`면 각 마커 클릭 시
     `Streamlit.setComponentValue({"주소": ..., "lat": ..., "lon": ...})`
     형태(기존 pydeck 선택 결과와 동일한 키)로 값을 돌려보낸다.
5. 파이썬 래퍼 함수의 반환값은 기존과 동일한 dict/None이라, 지도 업로드 탭에서
   `selected.get("lat")` 등을 쓰는 다운스트림 코드는 수정 없이 그대로 동작한다.

## 에러 처리

- `kakao_js_key`가 비어 있으면 지도를 그리지 않고, 기존 REST 키 안내 문구와
  같은 톤의 안내 캡션("카카오맵 JavaScript 키를 입력하면 지도를 표시합니다")을
  보여준다.
- 카카오 SDK 로드가 실패하면(키 오류/도메인 미등록 등) 빈 화면 대신, 컴포넌트
  내부에 카카오 SDK가 반환한 에러 메시지를 표시해 원인을 바로 알 수 있게 한다.

## 테스트 계획

- 이 프로젝트에는 자동화 테스트가 없다 (확인됨: `test_*` 패턴/`pytest`/`unittest`
  사용 없음).
- 로컬에서 `streamlit run dashboard.py`로 실행해 파이썬 예외 없이 5개 호출부
  모두 렌더링되는지 확인한다.
- 가능하면 Playwright로 브라우저 콘솔 에러 유무와 컴포넌트 마운트(iframe 로드)
  여부를 점검한다. 다만 실제 카카오 타일이 정상 표시되는지는 사용자의 키/도메인
  설정이 유효해야 최종 확인 가능하므로, 배포 후 최종 확인은 사용자가 직접 한다.

## 범위 밖

- 마커 클러스터링, 커스텀 지도 스킨/테마 등 기존에 없던 신규 기능은 추가하지 않는다.
- REST API 지오코딩(`geocode_address_kakao`) 로직은 변경하지 않는다.
