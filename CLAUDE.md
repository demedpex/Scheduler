# 내부망 업무 스케줄러

망분리 내부망에서 도는 업무 스케줄러. 서버도 DB도 공유폴더도 쓰지 않는다.
파일 하나가 프로그램이자 데이터베이스다.

- 웹 버전: `dist/배포/sched.html` — HTML 한 장. 열면 그대로 돈다
- 엑셀 버전: `src/` 의 VBA 를 `build/build.ps1` 이 `Scheduler.xlsm` 으로 묶는다
- 두 버전은 **메신저에 붙여넣는 텍스트**로 데이터를 주고받는다 (공유폴더·첨부를 못 쓴다)

## 절대 규칙

### R1. 화면은 `dist/배포/sched.html` 하나만 고친다

이 파일이 **유일한 기준이자 최신본이다.** 다른 곳에 사본을 만들지 않는다.

전에 `web/스케줄러.html` 에 같은 내용이 한 벌 더 있어서 한쪽만 고치면 조용히 어긋났다.
지금은 지웠고, 검증 스크립트도 전부 `dist/배포/sched.html` 을 읽는다.

기준이 **아닌** 파일 — 참고용으로만 둔다. 고치지 않는다.

| 파일 | 무엇인가 |
|---|---|
| `dist/AX금융혁신팀_업무스케줄러_20260827_v3/업무스케줄러.html` | 옛 판. 기준본이 이것의 확장이다 |
| `web/스케줄러_개편전.html` | 개편 전 화면 |
| `build/.fixture/design_v2.html` · `design_v3.html` | 디자인 시안 |

### R2. 외부에서 아무것도 받아오지 않는다

망분리라 인터넷이 없다. 지금 `sched.html` 은 CDN·외부 js/css·`fetch` 를 **하나도 쓰지 않는다.**

- 라이브러리를 링크로 걸지 않는다. 필요하면 코드를 파일 안에 넣는다
- 이미지는 인라인 SVG 나 data URI 로 넣는다
- 폰트를 외부에서 불러오지 않는다

### R3. 파일 밖에 데이터를 두지 않는다

서버·DB·공유폴더가 없다. 저장은 이 파일(엑셀판은 통합문서) 안에서 끝난다.
업무공유만 **텍스트로 내보내기 → 메신저 붙여넣기 → 상대가 가져오기** 로 오간다.

엑셀판과 웹판이 같은 텍스트를 주고받아야 하므로, **동기화 형식을 바꾸면 양쪽을 같이 고친다.**
`build/interop.js` 가 두 구현이 바이트 단위로 같은 텍스트를 만드는지 검사한다.

## 구조

```
dist/배포/sched.html   ← 웹 버전. 여기만 고친다
dist/배포/사용안내.txt   사용자용 안내
src/                   VBA 모듈(.bas) · 시트 문서 · 폼 코드
build/                 빌드·검증 스크립트
web/                   개편 전 화면 (참고용)
sample/                직원명부 양식 (더미 데이터)
README.md              설치·사용 안내
SPEC.md                기능 명세
docs_아키텍처.md        구조 문서
WORKING.md             어느 파일이 기준인지
```

## 검사

고친 뒤에는 전부 돌린다. Node 만 있으면 된다.

> 줄바꿈은 **LF 로 둔다.** `review-test.js` 같은 검사가 `
` 로 위치를 잡아서,
> CRLF 로 바뀌면 앵커를 못 찾고 통째로 실패한다.

```bash
node build/undef-test.js      # 선언 없이 쓰는 이름
node build/review-test.js     # 코드 리뷰 규칙 (139건)
node build/security-test.js   # 위험 패턴
node build/dday-test.js       # 날짜 계산 (20건)
node build/xlsx-test.js       # 명부 가져오기 (18건)
node build/batch-test.js      # 묶음 처리 (5건)
```

엑셀판까지 볼 때 (Windows + Excel 필요):

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1     # Scheduler.xlsm 생성
powershell -ExecutionPolicy Bypass -File build\simulate.ps1  # 동기화 기준 파일
node build\interop.js                                        # 웹판 ↔ 엑셀판 호환
```

## 문구

- 화면 문구는 **해요체**. 버튼·안내·오류 전부
- 사용자는 개발자가 아니다. "파싱 실패" 대신 "가져올 수 없는 형식이에요" 처럼 쓴다
- 오류에서 끝내지 말고 다음에 할 일을 알려준다
