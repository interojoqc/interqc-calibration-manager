# 계측기 검교정 관리 시스템

엑셀 계측기 등록 대장을 가져와서 검교정 이력, 3개월 도래 알림, 보정값, 폐기 계측기, 내부교정 성적서 생성을 관리하는 Streamlit 앱입니다.

## 실행

```powershell
cd "C:\Users\김상훈\Documents\New project\instrument_calibration_manager"
pip install -r requirements.txt
streamlit run app.py --server.port 8503
```

## 기본 사용 흐름

1. 사이드바에서 `QC` 역할과 암호를 입력합니다. 기본 암호는 `QC2026`이며 환경변수 `QC_APP_PASSWORD`로 바꿀 수 있습니다.
2. `데이터 가져오기`에서 기존 엑셀 파일을 업로드하거나 기본 경로 `C:\Users\김상훈\Desktop\계측기 등록 대장_2026.05.15.xlsx`를 불러옵니다.
3. `대시보드`에서 향후 3개월 검교정 도래 계측기를 확인합니다.
4. `알림 문구`에서 사용부서 담당자에게 보낼 카카오톡 메시지를 생성합니다.
5. `검교정 입력`에서 검교정 완료 내용을 추가합니다.
6. `내부교정/보정값`에서 측정값과 보정값을 입력하면 보정 환산값이 계산되고 내부교정 성적서 엑셀 파일이 생성됩니다.

## 권한

- QC: 등록, 수정, 검교정 입력, 폐기 처리, 담당자 관리 가능
- 사용부서: 열람과 알림 문구 확인만 가능

현재 버전은 현장 적용 전 프로토타입이므로 사내 계정 연동은 포함하지 않았습니다.

## 클라우드 운영

1차 무료 운영은 Streamlit Community Cloud, Google Sheets, Google Drive 기반입니다.
자세한 내용은 [STREAMLIT_CLOUD_DEPLOY.md](STREAMLIT_CLOUD_DEPLOY.md)를 확인하세요.

Cloud Run과 OCR까지 포함한 유료 가능 운영안은 [CLOUD_DEPLOY.md](CLOUD_DEPLOY.md)에 정리했습니다.
