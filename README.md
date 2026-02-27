# Univ3 Automation Python

## 설치

```powershell
cd C:\univ3-automation-python
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 설정

1. `config/local.json`에서 `paths.base_dir`를 작업할 학기 루트로 설정
2. 학기 전환은 `config/local.json`의 `paths.base_dir` 값을 변경
3. Google 서비스 계정 키는 아래 중 한 곳에 배치

- `config/secrets/google-service-account.json` (권장)
- `<base_dir>\코드\credentials.json`

4. Gmail 앱 비밀번호는 `.env` 파일에 설정

```env
GMAIL_APP_PASSWORD=...
```

## 실행

```powershell
# package entrypoint
$env:PYTHONPATH = "C:\univ3-automation-python\src"
python -m univ3_automation.cli

# 또는 직접 CLI 스크립트
python C:\univ3-automation-python\src\univ3_automation\commands\admin_cli.py
```

## 다중 폴더 적용 방식

코드는 동일하게 유지하고 `paths.base_dir`만 바꿔 적용합니다.

- 예: `G:/내 드라이브/대학3부/하드 행정팀/25-2`
- 예: `G:/내 드라이브/대학3부/하드 행정팀/26-1`

필요하면 실행 시점에 환경변수로 임시 오버라이드할 수 있습니다.

```powershell
$env:UNIV3_BASE_DIR = "G:/내 드라이브/대학3부/하드 행정팀/26-1"
python C:\univ3-automation-python\src\univ3_automation\commands\admin_cli.py
```
