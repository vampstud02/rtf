MS 오피스 내에서 특정 CLSID를 가진 OLE 객체의 로드 및 활성화 여부를 검증하는 테스트 도구를 위한 PRD(제품 요구사항 정의서) 초안을 작성해 드립니다.

이 도구의 핵심은 **"임의의 CLSID를 주입한 RTF 생성"**과 **"실제 로드 여부 판정"**입니다.

[PRD] OLE Object Load & Activation Tester
1. 개요 (Introduction)
목적: 특정 CLSID를 포함한 RTF 파일을 생성하고, MS 워드(Word) 환경에서 해당 객체가 정상적으로 로드 및 활성화(Instantiate)되는지 자동화된 방식으로 검증함.

배경: 보안 취약점(CVE) 재현, 특정 ActiveX 컨트롤의 오피스 호환성 테스트, 또는 Kill-bit 설정 적용 여부 확인.

2. 목표 (Goals)
사용자가 입력한 CLSID를 기반으로 유효한 RTF 문서를 생성한다.

생성된 문서를 MS Word로 실행하고, 프로세스 수준에서 객체 로드를 감지한다.

테스트 결과를 성공/실패(로드됨/차단됨)로 리포팅한다.

3. 사용자 요구사항 (User Requirements)
CLSID 입력: 사용자는 테스트하고자 하는 객체의 CLSID를 문자열 형태로 입력할 수 있어야 함.

RTF 자동 생성: OLE 스트림 구조를 유지하면서 CLSID만 교체한 테스트용 RTF를 생성해야 함.

로드 모니터링: Word가 실행될 때 해당 CLSID와 연결된 DLL이 로드되거나 레지스트리 접근이 발생하는지 확인해야 함.

환경 설정: Word의 경로 및 테스트 대기 시간(Timeout)을 설정할 수 있어야 함.

인터페이스: 자동화 스크립트 특성상 CLI(Command Line Interface) 형태로 제공되어야 함.

4. 상세 기능 요구사항 (Functional Requirements)
4.1 RTF 페이로드 생성 모듈
기능: 표준 OLE 컨테이너 바이너리 구조를 생성하고 \objdata 섹션에 삽입.

세부사항:

OLE 1.0/2.0 헤더 및 구조체 생성 기능.

입력받은 CLSID를 Little-endian 바이너리 형태로 변환하여 주입.

\object\objemb{\objclass [ProgID]}\objdata [HexData] 형식의 RTF 래핑.

4.2 실행 및 모니터링 모듈
기능: 생성된 RTF를 MS Word로 열고 시스템 상태를 감시.

검증 방식 (택 1 또는 조합):

모듈 로드 확인: GetModuleHandle 등을 이용해 Word 프로세스 내에 대상 DLL이 로드되었는지 체크.

레지스트리 감시: RegQueryValue 호출 발생 여부 모니터링.

VBA 이벤트: 문서 오픈 시 특정 매크로를 실행하여 객체 상태 확인 (신뢰할 수 있는 위치 설정 필요).

4.3 결과 리포팅
출력: 테스트 성공 여부, 로드된 시간, 탐지된 모듈 경로(DLL) 표시.

저장 포맷: 대량 테스트 및 이력 관리를 위해 결과를 .csv 또는 .json 파일 형태로 저장 지원.

5. 기술 스택 (Technical Stack)
Language: Python (가장 권장 - rtf 관련 라이브러리 및 psutil, pywin32 활용 용이)

Target App: Microsoft Word 2016 / 2019 / Office 365 (32/64 bit)

Library:

oletools (OLE 구조 분석용)

pywin32 (Word COM 제어용)

frida 또는 Process Monitor (정밀 모니터링 필요 시)

6. 테스트 시나리오 (Test Scenarios)
정상 시나리오: 등록된 CLSID를 넣고 실행 -> Word 실행 후 해당 객체 핸들러(DLL) 로드 확인 -> "Success".

Kill-bit 시나리오: 레지스트리에서 차단된 CLSID 사용 -> Word 실행 후 로드 실패 확인 -> "Blocked/Fail".

미등록 CLSID: 존재하지 않는 CLSID 사용 -> 오류 메시지 또는 무반응 확인 -> "Not Found".

7. 향후 확장성 (Future Scope)
멀티 파일 생성: 여러 개의 CLSID를 리스트로 넣어 대량의 테스트 세트를 한 번에 생성.

Sandbox 연동: 가상 환경에서 Word를 실행하고 행위를 분석하는 기능.

8. 제한 사항 및 전제 조건 (Limitations & Prerequisites)
대상 장비: 테스트 대상인 MS Word(2016/2019/365)가 설치되어 있어야 함.

실행 권한: 프로세스(Word) 내부 모니터링 및 레지스트리 접근 감시를 위해 스크립트를 관리자 권한(Administrator)으로 실행해야 할 수 있음.

보안 설정: Office의 제한된 보기(Protected View), 매크로 차단 등 기본 보안 정책으로 인해 테스트가 멈추지 않도록 적절한 환경 구성(신뢰할 수 있는 위치 등)이 필요함.

9. 보안 및 안전 지침 (Security Guidelines)
격리 환경 권장: 취약점(CVE) 테스트 또는 악성 OLE 객체의 가능성이 있는 CLSID를 검증하는 도구이므로, 호스트 OS가 아닌 스냅샷 복원이 용이한 격리된 가상 머신(VM)에서 실행할 것을 강력히 권장함.