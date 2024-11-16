### INTRO

- 구글 스프레드 시트 API 가 이미지를 반환하지 않는 이슈로 이미지를 파싱하고 json 파일을 생성하는 모듈 개발
- json 파일에는 동아리정보, 활동들이 있으며 이미지 관련 값은 이미지 경로만 가지고 있음.

### 사용법

1. 계보 프로젝트 데이터 정리 시트에서 `파일 > 다운로드 > .xlsx` 선택해 엑셀 파일 다운로드

2. 파일 이름을 `original.xlsx` 로 변경

3. 프로젝트 최상단 경로(/) 에 `original.xlsx` 파일 추가

4. `parse.js` 파일 실행 (VSCode 에서 `Run Code` 실행)

5. 서버가 필요한 경우 `server.js` 파일 실행 (`http://localhost:3000/`)

6. 모듈 실행 시 생성되는 `riu/impages` 와 `riu/output/results.json` 을 계보 프로젝트 웹 레포에 복사
