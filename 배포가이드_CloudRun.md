# 📦 소포수령증 자동화 앱 — Google Cloud Run 배포 가이드

## 전체 순서

```
1. google_credentials.json 수정
2. 직원 이메일 목록 추가 (app.py)
3. Google Cloud SDK 설치
4. gcloud 로그인 및 프로젝트 설정
5. 배포 실행
6. OAuth 리디렉션 URI 업데이트
```

---

## STEP 1: google_credentials.json 수정

`google_credentials.json` 파일을 열어서 아래 3곳을 수정하세요:

```json
{
  "web": {
    "client_id": "123456789-abc.apps.googleusercontent.com",   ← 실제 클라이언트 ID
    "project_id": "my-project-id",                             ← 프로젝트 ID
    "client_secret": "GOCSPX-xxxxxxxxxxxxxxxxx",               ← 클라이언트 보안 비밀번호
    "redirect_uris": ["https://배포후주소.run.app"],            ← 일단 임시 입력, 나중에 수정
    ...
  }
}
```

> ⚠️ 이 파일은 외부에 절대 공유하지 마세요 (비밀번호와 같음)

---

## STEP 2: 허용 직원 이메일 추가 (app.py)

`app.py` 상단의 `ALLOWED_EMAILS` 목록을 수정하세요:

```python
ALLOWED_EMAILS = [
    "직원1@회사.com",
    "직원2@회사.com",
    "hj@taxexpert.kr",
    # 필요한 만큼 추가
]
```

> 이 목록이 **비어있으면** Google 계정을 가진 누구나 로그인 가능합니다.

---

## STEP 3: Google Cloud SDK 설치

### Windows
1. https://cloud.google.com/sdk/docs/install 에서 설치 파일 다운로드
2. 설치 후 **Google Cloud SDK Shell** 실행

### Mac
```bash
brew install --cask google-cloud-sdk
```

---

## STEP 4: 로그인 및 프로젝트 설정

```bash
# Google 계정으로 로그인
gcloud auth login

# 프로젝트 설정 (Google Cloud Console에서 프로젝트 ID 확인)
gcloud config set project 여기에_프로젝트_ID

# Cloud Run API 활성화
gcloud services enable run.googleapis.com
gcloud services enable cloudbuild.googleapis.com
```

---

## STEP 5: 배포 실행

`소포수령증_자동화` 폴더에서 아래 명령 실행:

```bash
cd 소포수령증_자동화

gcloud run deploy soposuryjungjeung \
  --source . \
  --region asia-northeast3 \
  --platform managed \
  --allow-unauthenticated \
  --set-env-vars "COOKIE_SECRET=랜덤한문자열아무거나" \
  --memory 1Gi
```

> 배포 완료 후 앱 주소가 출력됩니다:
> `https://soposuryjungjeung-xxxxxxxxxx-de.a.run.app`

---

## STEP 6: OAuth 리디렉션 URI 업데이트

배포 후 앱 주소를 받으면:

### 6-1. google_credentials.json 수정
```json
"redirect_uris": ["https://soposuryjungjeung-xxxxxxxxxx-de.a.run.app"]
```

### 6-2. Google Cloud Console에서도 추가
1. Cloud Console → **API 및 서비스** → **사용자 인증 정보**
2. OAuth 클라이언트 ID 클릭
3. **승인된 리디렉션 URI** 에 배포 주소 추가
4. 저장

### 6-3. 환경변수에도 추가하여 재배포
```bash
gcloud run deploy soposuryjungjeung \
  --source . \
  --region asia-northeast3 \
  --platform managed \
  --allow-unauthenticated \
  --set-env-vars "COOKIE_SECRET=아까_입력한_같은_문자열,REDIRECT_URI=https://soposuryjungjeung-xxxxxxxxxx-de.a.run.app" \
  --memory 1Gi
```

---

## 배포 완료 후 확인

- 앱 주소로 접속하면 Google 로그인 화면이 나타납니다
- 허용된 이메일로 로그인하면 앱 사용 가능
- 허용되지 않은 이메일은 "접근 권한 없음" 메시지 표시

---

## 비용 안내

Google Cloud Run은 **사용량 기반 과금**입니다:
- 월 200만 요청까지 **무료**
- 소규모 사내 사용은 실질적으로 거의 무료
- https://cloud.google.com/run/pricing 참고

---

## 로컬 실행 (기존 방식 유지)

`google_credentials.json`의 값이 "여기에..."로 되어 있으면 인증 없이 로컬 실행됩니다:
```bash
streamlit run app.py
```
