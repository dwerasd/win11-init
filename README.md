# win11-init

## 한 줄 소개

Windows 11 신규 설치·재설치 환경을 개인 설정값으로 자동 세팅하는 스크립트 모음. 레지스트리 값 일괄 적용, 불필요한 기본 앱/서비스 제거, 사용자 폴더 백업·복구를 각각 JSON 설정 파일 기반으로 수행한다.

## 주요 기능

- **setup.py** — 세 단계로 시스템을 설정한다.
  1. `registry_config.json`에 정의된 43개 레지스트리 값(원격 분석 비활성화, 애니메이션/딜레이 끄기, 마우스 가속 끄기, HTTP 연결 수 조정 등)을 `winreg`로 일괄 기록.
  2. 현재 로컬 IP 대역을 감지해 인터넷 옵션의 로컬 인트라넷 영역에 자동 등록.
  3. `commands_config.json`에 정의된 59개 cmd/PowerShell 명령(Xbox Game Bar·Cortana·Your Phone 등 기본 앱 제거, 텔레메트리·SysMain 등 서비스 중지/비활성화, 우클릭 메뉴 복원 등)을 순차 실행.
  `--dry-run`으로 실제 변경 없이 시뮬레이션, `--list`로 전체 적용 항목 목록 확인 가능.
- **folder.py** — `folder_config.json`에 등록된 경로(NPKI, qBittorrent 설정, 카카오톡 받은 파일, 스크린샷 등)를 지정 대상으로 백업/복구한다. 백업 대상에 연결된 Windows 서비스가 있으면 백업/복구 전후로 자동 중지·재시작하며, `--add`/`--remove`/`--list`로 백업 경로 목록을 관리한다.

## 스택

- Python 3 (`argparse`, `json`, `subprocess`, `winreg`, `socket`, `shutil`, `pathlib`)
- 외부 의존성 없음 — 표준 라이브러리만 사용(단, `winreg`는 Windows 전용 모듈이라 Windows에서만 실행 가능)
- Windows `sc`(서비스 제어), `cmd`, `powershell` CLI 호출

## 폴더 구성

```
win11-init/
├── setup.py               # 레지스트리·인트라넷·명령어 자동 적용
├── folder.py              # 사용자 폴더 백업/복구
├── registry_config.json   # 적용할 레지스트리 항목 정의
├── commands_config.json   # 실행할 cmd/PowerShell 명령 정의
├── folder_config.json     # 백업 대상 경로 및 마지막 백업 위치
└── LICENSE                # MIT
```

## 빌드·실행

Windows 11 전용(`winreg`, 레지스트리 경로, `AppData` 경로 등 Windows 종속 코드). 관리자 권한이 필요한 레지스트리 항목은 비관리자 실행 시 권한 오류만 출력하고 계속 진행한다.

```
python setup.py               # registry_config.json + commands_config.json 전체 적용
python setup.py --dry-run     # 실제 변경 없이 적용 내역만 출력
python setup.py --list        # 적용 대상 항목 목록 확인

python folder.py --list                       # 등록된 백업 경로 목록
python folder.py --backup "E:\backup"         # 백업 실행
python folder.py --restore "E:\backup\NPKI"   # 선택 복구
python folder.py --add "C:\경로\폴더"         # 백업 경로 추가
```

`folder_config.json`의 `backup_paths`는 특정 사용자 계정(`C:\Users\Administrator\...`) 경로가 그대로 하드코딩되어 있어, 다른 환경에서 쓰려면 직접 값을 수정해야 한다.

## 상태

개인 PC 초기 세팅용 스크립트로, 별도의 테스트 코드나 CI는 없다. 레지스트리 변경과 서비스 비활성화는 시스템 동작에 직접 영향을 주므로 `--dry-run`으로 먼저 확인 후 적용하는 것을 권장한다.
