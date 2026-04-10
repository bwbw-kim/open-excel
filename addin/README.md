# Excel Copilot Add-in

Excel에서 자연어로 스프레드시트 작업을 수행하는 AI 기반 Office Add-in입니다.

## 주요 기능

- 🗣️ **자연어 명령** - "A1 셀에 제목 넣어줘", "3행 삭제해줘" 등 한국어로 Excel 조작
- 📊 **셀/범위 읽기/쓰기** - 단일 셀부터 넓은 범위까지 데이터 조작
- 📋 **표 붙여넣기** - 클립보드에서 표를 인식하여 Excel에 삽입
- 📄 **시트 관리** - 시트 생성, 이름 변경 등
- 🌐 **크로스 플랫폼** - Windows, Mac, Excel Online 모두 지원

## 기술 스택

- **UI**: React 18 + Fluent UI React v9
- **Excel API**: Office.js
- **빌드**: Webpack 5 + TypeScript
- **AI**: OpenAI Codex API

---

## 🧪 로컬 테스트 (Step-by-Step)

처음 테스트하는 경우 아래 단계를 순서대로 따라하세요.

### Step 1: 프로젝트 클론 및 의존성 설치

```bash
# 프로젝트 루트로 이동
cd open-excel/addin

# 의존성 설치 (최초 1회)
npm install
```

### Step 2: 개발 서버 시작 + Excel 실행

```bash
npm start
```

**실행 시 발생하는 일:**
1. HTTPS 개발 서버가 `https://localhost:3000`에서 시작됨
2. 자체 서명 SSL 인증서가 자동 생성됨 (첫 실행 시)
3. Excel이 자동으로 실행되고 Add-in이 사이드로드됨

**첫 실행 시 프롬프트:**
```
? Allow localhost loopback for Microsoft Edge WebView? (Y/n)
```
→ `Y` 입력 후 Enter

**인증서 신뢰 팝업이 나타나면:**
→ "예" 또는 "Trust" 클릭

### Step 3: Excel에서 Add-in 열기

1. Excel이 자동으로 열림
2. **Home** 탭으로 이동
3. **Open Copilot** 버튼 클릭 (리본 메뉴에 추가됨)
4. 우측에 Task Pane이 열림

### Step 4: 기능 테스트

Task Pane 입력창에 다음을 입력해보세요:

```
A1 셀에 '안녕하세요' 넣어줘
```

Enter를 누르면 A1 셀에 값이 입력됩니다.

**더 많은 테스트:**
```
B2:D4 범위 읽어줘
3행 삭제해줘
'테스트시트' 시트 만들어줘
```

### Step 5: 코드 수정 후 테스트

1. `src/` 폴더의 코드 수정
2. 저장하면 자동으로 Hot Reload됨
3. Excel Task Pane에서 바로 변경사항 확인

### Step 6: 개발 서버 종료

```bash
npm stop
```

이 명령어는:
- 개발 서버 종료
- Excel에서 Add-in 언로드
- 관련 프로세스 정리

---

## 🚀 빠른 시작 (요약)

### 필수 조건

1. **Node.js** (v18 LTS 이상)
   ```bash
   node -v  # v18.0.0 이상 확인
   ```

2. **Microsoft 365 구독** (Excel 필요)
   - 개발용 무료 구독: https://aka.ms/m365devprogram

### 전체 명령어 (복사용)

```bash
# 1. 프로젝트로 이동
cd open-excel/addin

# 2. 의존성 설치 (최초 1회)
npm install

# 3. 개발 서버 + Excel 실행
npm start

# 4. 종료 시
npm stop
```

---

## 📦 사이드로드 방법

### Windows/Mac 데스크톱 Excel

```bash
# 자동 사이드로드 (권장)
npm start

# Excel에서 Home 탭 → "Open Copilot" 버튼 클릭
```

### Mac에서 수동 사이드로드

```bash
# 터미널 1: 개발 서버 실행
npm run dev

# 터미널 2: Excel에 사이드로드
npm run start:desktop
```

### Excel Online (웹)

**방법 1: 자동 사이드로드**
```bash
# SharePoint/OneDrive의 Excel 파일 URL 필요
npm run start:web -- --document "https://contoso.sharepoint.com/:x:/r/sites/..."
```

**방법 2: 수동 업로드**
1. Excel Online에서 파일 열기
2. **Home** → **Add-ins** → **More Add-ins**
3. **Upload My Add-in** 클릭
4. `manifest.json` 파일 선택 후 Upload
5. Home 탭에서 "Open Copilot" 클릭

---

## 🛠️ 개발 명령어

| 명령어 | 설명 |
|--------|------|
| `npm start` | 개발 서버 + Excel 자동 사이드로드 |
| `npm stop` | 개발 서버 종료 및 Add-in 언로드 |
| `npm run dev` | 개발 서버만 실행 (수동 사이드로드 시) |
| `npm run build` | 프로덕션 빌드 |
| `npm run typecheck` | TypeScript 타입 검사 |
| `npm run validate` | Manifest 유효성 검사 |

---

## 📁 프로젝트 구조

```
addin/
├── manifest.json              # Office Add-in 매니페스트
├── package.json
├── tsconfig.json
├── webpack.config.js
├── assets/                    # 아이콘 (16x16, 32x32, 80x80)
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── src/
    ├── taskpane/              # UI 레이어
    │   ├── index.tsx          # React 진입점
    │   ├── App.tsx            # 메인 컴포넌트
    │   ├── styles.css         # 좁은 Task Pane 최적화 스타일
    │   ├── taskpane.html
    │   └── components/
    │       ├── Header.tsx     # 상단 헤더 (연결 상태)
    │       ├── ChatMessages.tsx  # 채팅 메시지 목록
    │       └── Composer.tsx   # 입력창
    ├── services/              # 비즈니스 로직
    │   ├── excel-service.ts   # Office.js Excel API 래퍼
    │   ├── llm-service.ts     # OpenAI API 호출
    │   ├── auth-service.ts    # OAuth 인증
    │   └── agent-types.ts     # Agent 타입 정의
    └── shared/
        └── types.ts           # 공유 타입
```

---

## 💬 사용 예시

### 셀 작업
```
"A1 셀에 '제목' 넣어줘"
"B2 셀 내용 지워줘"
```

### 범위 읽기
```
"B2:D10 범위 읽어줘"
"현재 시트 데이터 보여줘"
```

### 행 관리
```
"3행 삭제해줘"
"5행 아래에 새 행 추가해줘"
```

### 시트 관리
```
"'Summary' 시트 만들어줘"
```

### 표 붙여넣기
1. 다른 앱에서 표 복사 (Ctrl+C)
2. 입력창에 붙여넣기 (Ctrl+V)
3. 자동으로 표로 인식됨
4. "A1부터 붙여넣어줘" 명령

---

## 🚀 배포

### 1. 프로덕션 빌드

```bash
npm run build
```

`dist/` 폴더에 빌드 결과물이 생성됩니다.

### 2. 호스팅

빌드 결과물을 HTTPS 지원 서버에 배포:
- Azure Static Web Apps
- AWS S3 + CloudFront
- Vercel / Netlify

### 3. Manifest URL 수정

`manifest.json`의 모든 `localhost:3000`을 프로덕션 URL로 변경:

```json
"page": "https://your-domain.com/taskpane.html"
```

### 4. 배포 옵션

| 방법 | 대상 | 비용 |
|------|------|------|
| **Microsoft 365 Admin Center** | 조직 내 사용자 | 무료 |
| **AppSource (Microsoft Marketplace)** | 전 세계 사용자 | ~$99/년 |

#### 조직 내 배포

1. Microsoft 365 Admin Center 접속
2. **Settings** → **Integrated apps** → **Upload custom app**
3. `manifest.json` 업로드
4. 배포 대상 선택 (전체 조직 또는 특정 그룹)

---

## 🔧 문제 해결

### 캐시 문제

오래된 Add-in이 로드되는 경우:

```bash
# Add-in 언로드 후 재시작
npm stop
npm start
```

또는 Office 캐시 수동 삭제:
- **Windows**: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
- **Mac**: `~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Office/16.0/Wef`

### 포트 충돌

기본 포트(3000)가 사용 중인 경우 `webpack.config.js`에서 변경:

```javascript
devServer: {
  port: 3001,  // 다른 포트로 변경
}
```

### HTTPS 인증서 오류

```bash
# 인증서 재생성
npx office-addin-dev-certs install --days 365
```

### Excel에서 Add-in이 보이지 않음

1. Excel 완전 종료
2. `npm stop` 실행
3. `npm start` 재실행

---

## 📚 참고 자료

- [Office Add-in 공식 문서](https://learn.microsoft.com/office/dev/add-ins/)
- [Office.js API 레퍼런스](https://learn.microsoft.com/javascript/api/excel)
- [Fluent UI React](https://react.fluentui.dev/)
- [Task Pane 디자인 가이드](https://learn.microsoft.com/office/dev/add-ins/design/task-pane-add-ins)

---

## 지원 플랫폼

| 플랫폼 | 지원 |
|--------|------|
| Windows Excel 2016+ | ✅ |
| Mac Excel 2016+ | ✅ |
| Excel Online | ✅ |
| iPad Excel | ✅ |

---

## 라이선스

MIT
