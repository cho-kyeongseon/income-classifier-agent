# HANDOFF

## 현재 상태 (2026-11-14)
- 소득구분팀(A~D) 서브에이전트 작성: `.claude/agents/`
  - A `a-document-writer`: 소득대장 → 인원별 필요서류 (`templates/필요서류_양식.md`)
  - B `b-criteria-reviewer`: 법령 확인(mcp_law 먼저, 안 되면 국세법령정보시스템) → 소득구분기준.md + rules.json
  - C `c-income-calculator`: 구분 확정 + `scripts/calc_withholding.py`로 원천징수세액 계산
  - D `d-kakao-writer`: 대표님 카톡 메시지
- 한 번에 실행: `/소득구분 업체명` (예: `/소득구분 해인커피`, `.claude/commands/소득구분.md`)
  - `data/raw/`에서 파일 이름에 업체명이 들어간 소득대장을 찾는다(여러 개면 최신 파일).
  - 결과는 `output/<업체명>_<YYYYMMDD_HHMM>/`에 저장(USB에만, GitHub 제외).
  - 결과 폴더 안 `00_원본자료/`에 그 실행에 쓴 원본 소득대장을 복사해 보관한다.
- 계산 스크립트는 가상 샘플(`data/sample/구분확정_샘플.csv`)로 손계산과 일치 확인함.
- 2026-11-14 해인커피 실습자료로 전체 흐름 실행 → `output/해인커피_20261114_1047/`
  - 확정은 일용 2명(원천징수 4,440원)뿐, 8명 보류(서류 필요·간이세액표).
  - 그날 법령 MCP 오류, law.go.kr 점검으로 시행령 제20조·제87조·제184조 등 미확인(verified=false).

## 주의
- `reference/rules.template.json`의 세율·산식은 **법령 미확인 초안**이다. 작성 시 mcp_law는 오류,
  국세법령정보시스템은 연결 끊김으로 확인하지 못했다. B가 실행 때마다 확인해 `verified: true`로 채워야 한다.
- 이 저장소 전용 클라우드 환경 `income-classifier-agent`에 taxlaw.nts.go.kr, www.law.go.kr 접속을 허용해 두었다.
- 상용 근로소득(간이세액표)은 스크립트가 계산하지 않고 보류로 둔다.
- `output/`은 개인정보가 있어 GitHub에 올리지 않는다.

## 다음 할 일
- 새 세션에서 `/소득구분 해인커피`로 명령·서브에이전트가 바로 동작하는지 확인.
- 법령 접속이 되면 rules.template.json 값 확인·출처 기입.
- 필요하면 상용 근로소득 간이세액표 계산 추가.
