# Open Excel MVP

Open Excel은 **Windows에서 이미 실행 중인 Microsoft Excel에 연결해서** 자연어로 셀, 행, 시트 작업을 수행하는 데스크톱 앱입니다.

## 현재 기능

- OpenAI OAuth 로그인
- Excel Live Mode 연결 (`Connect Excel`)
- 셀 쓰기 / 셀 삭제
- 행 추가 / 행 삭제
- 시트 생성
- 표 붙여넣기 인식
- 웹 검색 결과를 표로 정리
- Windows 설치 파일 배포

## 개발

```bash
npm install
npm run dev
```

추가 검증:

```bash
npm run typecheck
npm run build
```

## 사용 방법

1. Windows에서 Excel 실행
2. workbook 열기
3. Open Excel 실행
4. `Login with OpenAI`
5. `Connect Excel`
6. 자연어로 작업 요청

예:

- `A1 셀에 제목 넣어줘`
- `3행 삭제해줘`
- `B4 셀 내용 지워줘`
- `새 시트 Summary 만들어줘`

## 테스트

기본 확인:

```bash
npm run typecheck
npm run build
```

수동 확인:

- Excel 실행 후 workbook 열기
- `Connect Excel` 연결 확인
- 셀/행/시트 작업이 실제 Excel에 반영되는지 확인

## 배포

로컬 Windows 설치 파일 생성:

```bash
npm run dist:win
```

GitHub Release 배포:

- `.github/workflows/release.yml` 사용
- `v*` 태그를 push하면 Windows installer가 release에 업로드됨

## 참고

- Live Mode는 **Windows 전용**입니다.
- macOS에서는 Excel attach가 동작하지 않습니다.
