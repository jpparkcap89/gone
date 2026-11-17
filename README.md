# G-ONE 쉽게 배우기 · 자료실

G-ONE 일정·브리티 메신저 안내 영상과 PDF 안내서를 모아 공유하는 사이트입니다.

- 프론트엔드: GitHub Pages (`index.html` 한 파일) → `https://gnoe.cap89.com`
- 백엔드·DB: Google Apps Script 웹앱 + Google 스프레드시트 (`apps-script/Code.gs`)
- 디자인: cap89 스타일 + Lucian UX/UI 요소 (#039 알약 탭, #044 FAQ 아코디언, #070 토스트, MIT)

## 1. 구글 시트 + Apps Script

1. 새 구글 스프레드시트를 만듭니다 (예: `G-ONE 자료실 DB`).
2. 확장 프로그램 → Apps Script → `Code.gs` 내용을 붙여 넣고 저장합니다.
3. 편집기에서 `setup` 함수를 실행합니다 → `resources` 시트와 예시 자료 11건이 만들어집니다.
4. 프로젝트 설정(톱니) → 스크립트 속성 → `ADMIN_KEY` 추가 (관리자 등록용 비밀번호).
5. 배포 → 새 배포 → 유형 **웹 앱**
   - 실행: **나**
   - 액세스 권한: **모든 사용자**
6. 받은 웹 앱 주소(`https://script.google.com/macros/s/…/exec`)를 복사합니다.

> 코드를 고친 뒤에는 배포 관리 → 기존 배포 편집 → 새 버전으로 다시 배포해야 반영됩니다.

### resources 시트 열

| 열 | 설명 |
|---|---|
| id | 고유값 (예: m1) |
| order | 표시 순서 (작을수록 앞) |
| category | `messenger` / `calendar` / `guide` / `etc` |
| type | `video` / `pdf` / `link` |
| title, description | 제목, 한두 문장 설명 |
| url | 구글 드라이브 공유 링크 또는 유튜브 주소 |
| thumbnail | (선택) 썸네일 이미지 주소. 비우면 드라이브·유튜브 썸네일 자동 |
| duration | `1:30`, `7쪽` 등 |
| tags | 쉼표 구분 검색어 |
| featured | TRUE면 첫 화면 ‘TODAY'S PICK’ |
| isNew | TRUE면 NEW 표시 |
| visible | FALSE면 숨김 |
| views | 조회수 (자동 증가) |

**드라이브 파일 공유**: 영상·PDF를 드라이브에 올리고 공유를 ‘링크가 있는 모든 사용자 – 뷰어’로 해야 사이트에서 재생됩니다. 학교 계정 정책으로 외부 공개가 막혀 있으면 같은 도메인 로그인 사용자에게만 보입니다.

## 2. 프론트엔드 연결

`index.html`의 아래 줄에 웹 앱 주소를 넣습니다.

```js
var API_URL = 'https://script.google.com/macros/s/…/exec';
```

비워 두면 예시 데이터로 미리보기합니다. 관리자 등록 화면은 `https://gnoe.cap89.com/?admin` 으로 엽니다.

## 3. GitHub Pages + 도메인

1. GitHub에 새 저장소를 만들고 `index.html`, `CNAME`, `.nojekyll`을 올립니다 (`apps-script/`는 참고용, 올려도 무방).
2. Settings → Pages → Source: `main` 브랜치 / `(root)`.
3. Custom domain에 `gnoe.cap89.com` 입력 (CNAME 파일이 이미 들어 있음).
4. cap89.com DNS 관리 화면에서 레코드 추가:
   - 유형 `CNAME`, 호스트 `gnoe`, 값 `<GitHub 아이디>.github.io`
5. DNS가 반영되면(수 분~수 시간) Pages 설정에서 **Enforce HTTPS**를 켭니다.

## 라이선스 표기

Lucian UX/UI Prompt 요소 #039, #044, #070 · MIT License (c) 2026 Seokwoooo · https://lucianlabs.dev/uxui-prompt/
