# OLE Object Load & Activation Tester

본 도구는 특정 CLSID를 포함한 RTF 파일을 생성하여, 해당 객체가 Microsoft Word 환경에서 정상적으로 로드 및 활성화(Instantiate)되는지 검증하는 파이썬 기반 CLI 프로그램입니다.

## ✨ 주요 기능

- **RTF 페이로드 자동 생성**: 지정한 CLSID(단일 또는 로컬 레지스트리 전체)를 주입한 테스트용 RTF 문서를 생성합니다.
- **자동 모니터링**: MS Word 프로세스를 실행하고 지정된 시간(`timeout`) 동안 DLL 로드 여부를 감시합니다.
- **일괄 스캔 모드 (`--all-clsids`)**: 시스템에 등록된 모든 CLSID를 순차적으로 테스트할 수 있습니다.
- **다양한 출력 포맷 제공**: 눈으로 확인하기 쉬운 `console` 출력 외에도 대량 데이터 관리를 위한 `csv`, `json` 형태의 리포트를 지원합니다.

---

## 🚀 사용 환경 및 요구사항

- **OS**: Windows (Microsoft Word 설치 필수)
- **Python**: Python 3.6 이상 권장
- **필수 패키지**: `psutil`, `pywin32` (`pip install -r requirements.txt` 또는 `pip install psutil pywin32` 등으로 설치 가능)

---

## 📖 사용 방법

도구는 `cli.py`를 통해 실행하며, 터미널(또는 명령 프롬프트, 파워쉘)에서 아래와 같은 옵션을 사용하여 실행합니다.

### 1. 단일 CLSID 테스트
특정 CLSID 하나만 테스트하고 싶을 때 사용합니다.
```bash
python cli.py --clsid "0002CE02-0000-0000-C000-000000000046"
```

### 2. 전체 CLSID 테스트 (일괄 스캔)
시스템에 등록된 전체 CLSID를 스캔해서 각각 테스트할 때 사용합니다.
```bash
python cli.py --all-clsids
```

#### 옵션: 테스트 개수 제한 (`--limit`)
전체 테스트 시 수천 개의 CLSID를 검사하므로 시간이 오래 걸릴 수 있습니다. 테스트 목적으로 앞의 N개만 시도하고 싶을 때 사용합니다.
```bash
python cli.py --all-clsids --limit 10
```

### 3. 결과 리포트 저장 포맷 설정 (`--format`, `--output`)
출력 포맷(`console`, `csv`, `json`)과 저장될 파일명(확장자를 뺀 이름)을 지정할 수 있습니다. 
CSV 등 파일로 저장할 때도 기본적으로 콘솔에는 진행 상황이 출력됩니다.
```bash
# CSV 형식으로 결과를 scan_results.csv 파일에 저장
python cli.py --all-clsids --limit 50 --format csv --output scan_results

# JSON 형식으로 결과를 my_test.json 파일에 저장
python cli.py --clsid "0002CE02-0000-0000-C000-000000000046" --format json --output my_test
```

### 4. 기타 설정 (Word 경로 수동 지정, 타임아웃 변경)
MS Word가 기본 경로에 설치되어 있지 않거나, DLL이 로딩되기까지 더 긴/짧은 시간을 기다려야 할 때 사용합니다.
```bash
# 타임아웃을 15초로 늘리고, Word 실행 경로를 수동으로 지정
python cli.py --clsid "0002CE02-0000-0000-C000-000000000046" --timeout 15 --word-path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
```

---

## 🛠 도움말 확인
모든 옵션에 대한 상세한 설명은 아래 명령어로 확인할 수 있습니다.
```bash
python cli.py --help
```
