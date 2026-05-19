# 1차 무료 운영 배포 가이드

목표는 결제 없이 가능한 범위에서 운영하는 것입니다.

```text
Streamlit Community Cloud
  -> Google Sheets: 계측기/검교정/담당자 데이터 저장
  -> Google Drive: 성적서/폐기보고서 파일 저장
  -> OCR: 1차 제외, QC가 번호/보정값 확인 입력
```

## 준비물

1. GitHub 계정
2. Streamlit Community Cloud 계정
3. Google Cloud 프로젝트
   - 결제 연결 없이 Google Drive API / Google Sheets API만 사용
4. Google 서비스 계정 JSON
5. Google Drive 폴더 ID
6. Google Sheet ID

## Google Cloud에서 켤 API

결제 없이 1차 무료 운영만 할 경우 아래 2개만 우선 사용합니다.

- Google Drive API
- Google Sheets API

Cloud Run, Cloud Build, Artifact Registry, Cloud Vision API는 1차 무료 운영에서는 사용하지 않습니다.

## 서비스 계정

1. Google Cloud Console > IAM 및 관리자 > 서비스 계정
2. 서비스 계정 생성
3. 키 > 새 키 만들기 > JSON 다운로드
4. JSON 안의 `client_email` 값을 복사
5. Google Drive 폴더와 Google Sheet에 해당 이메일을 `편집자`로 공유

## Drive 폴더 ID

Drive 폴더 URL 예:

```text
https://drive.google.com/drive/folders/1AbCDefGhijkLMNopQRsTuv
```

폴더 ID:

```text
1AbCDefGhijkLMNopQRsTuv
```

## Sheet ID

Google Sheet URL 예:

```text
https://docs.google.com/spreadsheets/d/1XyzABC1234567890/edit
```

Sheet ID:

```text
1XyzABC1234567890
```

## Streamlit Cloud Secrets

Streamlit Cloud 앱 설정의 `Secrets`에 아래 형식으로 넣습니다.

```toml
QC_APP_PASSWORD = "QC2026"
GOOGLE_DRIVE_FOLDER_ID = "..."
GOOGLE_SHEET_ID = "..."
ENABLE_VISION_OCR = "false"
GOOGLE_SERVICE_ACCOUNT_JSON = '''
{
  "type": "service_account",
  "project_id": "...",
  "private_key_id": "...",
  "private_key": "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n",
  "client_email": "...",
  "client_id": "...",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "..."
}
'''
```

## 앱 동작

- 위 설정이 있으면 앱 상단 저장소가 `Google Sheets`로 표시됩니다.
- 설정이 없으면 `Local SQLite`로 표시됩니다.
- `데이터 가져오기`에서 기존 엑셀 파일을 업로드하면 Google Sheet에 데이터가 생성됩니다.
- 성적서/폐기보고서 파일은 Google Drive 폴더로 업로드되고, 앱에는 Drive 링크가 저장됩니다.

## OCR

1차 무료 운영에서는 OCR을 제외합니다. OCR이 필요해지면 `ENABLE_VISION_OCR=true`와 Cloud Vision API가 필요하며, 이 단계에서는 결제 계정 요구가 생길 수 있습니다.
