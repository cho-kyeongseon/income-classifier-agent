# HANDOFF

## 지금까지 한 작업
- 소득구분팀(A~D) 서브에이전트 작성: `.claude/agents/`
  - A `a-document-writer`: 소득대장 → 인원별 필요서류 (`templates/필요서류_양식.md`)
  - B `b-criteria-reviewer`: 법령 확인(법령 MCP 먼저, 안 되면 국세법령정보시스템) → 소득구분기준.md + rules.json
  - C `c-income-calculator`: 구분 확정 + `scripts/calc_withholding.py`로 원천징수세액 계산
  - D `d-kakao-writer`: 대표님 카톡 메시지(파일로만 저장, 보내지 않음)
- 실행 명령 `/소득구분 업체명`으로 변경(`.claude/commands/소득구분.md`). 보고는 한글.
- 결과 폴더를 `output/<업체명>_<YYYYMMDD_HHMM>/`로, 맨 앞 `00_원본자료/`에 쓴 원본 소득대장을 복사해 보관.
- 2026-11-14 해인커피 실습자료로 전체 흐름 실행 → `output/해인커피_20261114_1047/` (USB에만 있음)

## 현재 상태 (2026-11-14)
- 명령·에이전트·계산 스크립트는 동작함. 계산 스크립트는 가상 샘플과 해인커피 일용 13건 손계산과 일치.
- 해인커피 결과: 10명, 지급 93,115,000원. 확정은 일용 2명(송예린·윤태호, 원천징수 4,440원)뿐, 8명 보류.
  - 이서준·박지민: 근로_상용 확정, 세액은 간이세액표 필요(스크립트 미지원).
  - 최유나(대표 배우자)·한도윤(3개월 이상 계속 근무 → 일용 아님 가능성)·강민혁·이하람·오세라·배상훈: 서류 받고 확정.
- 법령 미확인(verified=false): 시행령 제20조(일용 범위)·제87조(필요경비 60%)·제184조, 제86조 단서, 지방소득세 끝수 근거,
  간이세액표, 식대·자가운전보조금 비과세(시행령 제17조의2·제12조). 그날 법령 MCP 오류 + law.go.kr 점검 때문.
- `/소득구분`은 이 저장소를 연 **새 세션**에서만 목록에 뜬다. 새 세션에서 바로 동작하는지는 아직 확인 안 함.

## 남은 일
- 새 세션에서 `/소득구분 해인커피`로 명령·서브에이전트가 바로 동작하는지 확인.
- 법령 접속이 되면 위 미확인 항목을 확인해 `reference/rules.template.json`에 출처 기입.
- 필요하면 상용 근로소득 간이세액표 계산 추가.
- 해인커피: 대표님께 서류 받으면 보류 8명 확정.

## 주의할 점
- `reference/rules.template.json`의 세율·산식은 초안이다. B가 실행 때마다 확인해 `verified: true`로 채워야 한다.
- 이 저장소 전용 클라우드 환경 `income-classifier-agent`에 taxlaw.nts.go.kr, www.law.go.kr 접속을 허용해 두었다.
- `output/`(결과, 개인정보)과 `data/raw/`(원본)은 GitHub에 올리지 않는다(.gitignore).
- USB(exFAT 등)라 git이 "dubious ownership" 오류를 내면:
  `git config --global --add safe.directory "<이 폴더 경로>"` (다른 컴퓨터·다른 드라이브 문자마다 한 번)
- 이 PC에는 git 사용자 이름·이메일 설정이 없어 커밋 작성자를 명령마다 지정했다.

## 실행 방법
- 설치: Git, GitHub CLI(`gh`, 푸시하려면 `gh auth login`), Python 3(openpyxl, pandas). Windows는 `python3` 대신 `python`.
- 원본 소득대장을 `data/raw/`에 넣는다. 파일 이름에 업체명이 들어가야 한다(예: `해인커피_실습자료_수강생용.xlsx`).
- Claude 데스크톱 앱에서 이 폴더로 새 세션을 열고: `/소득구분 해인커피`
- 계산 스크립트만 직접: `python scripts/calc_withholding.py <구분확정.csv> <rules.json> <출력폴더>/`

## 원본 자료 위치 (GitHub에 없음, USB에만)
- 원본 소득대장: `data/raw/` (지금: `data/raw/해인커피_실습자료_수강생용.xlsx`)
  - 처음 받은 곳: `D:\강의안 및 실습자료\해인커피_실습자료_수강생용.xlsx`
- 실행 결과: `output/<업체명>_<실행ID>/` (각 결과의 `00_원본자료/`에 원본 사본)
- 프로그램은 모두 작업 폴더 기준 상대 경로(`data/raw/`, `output/`)로 읽는다. 드라이브 문자가 바뀌어도 된다.
