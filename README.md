# 지멘스 공문 자동 등록·요약 (`siemens_letters.py`)

모바일에서 실행하면 **모바일 Outlook 앱과 같은 사서함**을 조회해 최근 3일간 접수된 지멘스(Siemens) 공문을 찾고,
**첨부 공문을 먼저 추출**하여 **'지멘스 대응' 프로젝트에 등록**한 뒤 Claude로 **요약**합니다.

> 모바일 OS는 앱끼리 데이터를 공유하지 않으므로(샌드박스), 스크립트가 Outlook 앱의 내부 저장소를 직접 읽을 수는 없습니다.
> 대신 Outlook 앱이 동기화하는 것과 동일한 사서함(Microsoft 365 / Outlook.com)을 **Microsoft Graph API**로 읽습니다.
> 결과는 Outlook 앱에서 보는 메일과 같습니다.

## 처리 흐름

1. **공문 조회** – 최근 N일(기본 3일) 수신 메일 중
   - 발신 도메인에 `siemens`가 포함된 메일(siemens.com, siemens-energy.com 등), 또는
   - 제목/발신자명에 `지멘스`·`Siemens`가 있고 공문 키워드(공문, Letter, Notice, Claim, 통지 등)가 있는 메일(전달 메일 포함)
   - **첨부가 있는 공문부터** 처리
2. **첨부 우선 추출** – 파일 첨부를 `output/지멘스 대응/<수신일>_<제목>/`에 저장 (서명 등 본문 삽입 이미지는 제외)
3. **요약 (Claude)** – 첨부 공문을 1차 근거로, 메일 본문을 보조로 분석
   - 지원: PDF(스캔본 포함), 이미지, DOCX, XLSX, PPTX, HWPX, TXT/CSV
   - HWP(구형 한글)는 텍스트 추출 불가 → 요약에 "PDF 변환 권장"으로 표시
   - 결과: 공문번호, 발신, 일자, 분류, 요약, 주요 내용, 요청 조치사항·기한, 중요도, 리스크
4. **'지멘스 대응' 프로젝트 등록**
   - OneDrive `지멘스 대응/<수신일>_<제목>/` 폴더에 첨부 + `요약.md` 업로드
   - Microsoft To Do 목록 **'지멘스 대응'**(없으면 자동 생성)에 작업 등록
     - 제목 `[지멘스 공문] [문서번호] 제목`, 메모에 요약, 기한·중요도·분류 설정
     - 요청 조치사항 → 체크리스트 단계, 3MB 이하 첨부 → 작업에 직접 첨부
     - 공문 메일 원문 / OneDrive 폴더 링크 연결
   - 또는 Planner 플랜(`project_backend: "planner"`)의 '공문 접수' 버킷에 등록
   - 이미 등록한 공문은 `registry.json`으로 중복 등록 방지
5. **요약 보고서** – 화면 출력 + `output/지멘스 대응/요약보고_<날짜>.md`

## 사전 준비 (최초 1회)

### 1) Azure 앱 등록 (Graph 접근용, 무료)
1. <https://portal.azure.com> → **Microsoft Entra ID → 앱 등록 → 새 등록**
2. 지원되는 계정 유형: *모든 조직 디렉터리의 계정 및 개인 Microsoft 계정*
3. 등록 후 **인증 → 고급 설정 → 공용 클라이언트 흐름 허용: 예**
4. **API 사용 권한 → Microsoft Graph → 위임된 권한**: `User.Read`, `Mail.Read`, `Tasks.ReadWrite`, `Files.ReadWrite`
5. **애플리케이션(클라이언트) ID**를 복사
   (회사 계정은 테넌트 정책상 관리자 동의가 필요할 수 있습니다.)

### 2) Claude API 키
<https://console.anthropic.com>에서 발급 → `config.json`의 `anthropic_api_key` 또는 환경변수 `ANTHROPIC_API_KEY`

### 3) 설정 파일
```bash
cp config.example.json config.json   # azure_client_id, anthropic_api_key 입력
```

## 모바일에서 실행

### Android (Termux)
```bash
pkg install python git rust    # rust: anthropic 의존성(pydantic-core) 빌드용
git clone <이 저장소> && cd Outlook
pip install -r requirements.txt
cp config.example.json config.json && nano config.json
python siemens_letters.py
```
최초 실행 시 표시되는 주소·코드로 **모바일 Outlook 앱과 같은 계정**으로 로그인하면, 이후에는 토큰 캐시로 자동 로그인됩니다.
Termux:Widget을 쓰면 홈 화면 버튼 한 번으로 실행할 수 있습니다.

### iPhone / iPad
Pyto 등 Python 앱에서 `requirements.txt` 설치가 가능하면 그대로 실행하면 됩니다.
컴파일 의존성 설치가 막히는 경우에는 PC·서버에서 이 스크립트를 실행하고 iOS 단축어(SSH 실행)로 호출하는 방식을 권장합니다.

## 옵션
```bash
python siemens_letters.py --days 7    # 최근 7일
python siemens_letters.py --dry-run   # 프로젝트 등록 없이 첨부 추출·요약만 (로컬 저장)
python siemens_letters.py --force     # 이미 등록된 공문도 다시 처리
```

| 설정 | 설명 |
|---|---|
| `project_backend` | `todo`(기본, 개인/회사 계정 모두) · `planner`(회사 계정, `planner_plan_id` 필요) · `local`(로컬 저장만) |
| `upload_to_onedrive` | 첨부·요약을 OneDrive `지멘스 대응` 폴더에 업로드 |
| `sender_domain_keywords` / `text_keywords` / `letter_keywords` | 지멘스 공문 판별 키워드 |
| `only_with_attachments` | `true`면 첨부가 있는 공문만 처리 |
| `claude_model` | 요약 모델 (기본 `claude-opus-5`) |

## 보안 참고
- `config.json`(API 키)과 `.token_cache.json`(로그인 토큰)은 `.gitignore`에 포함되어 커밋되지 않습니다.
- 공문 내용은 요약을 위해 Anthropic API로 전송됩니다. 회사 보안 정책을 확인한 뒤 사용하세요.
