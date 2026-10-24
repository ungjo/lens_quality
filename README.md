# 렌즈 검사 프로젝트

Python 3.10 환경에서 Excel의 측정값과 테스트용 기준을 읽어 SQLite에 저장하고 판정합니다.
이 기준은 프로그램 개발 실습용이며 실제 품질 인증·출하 기준이 아닙니다.

## 상세 문서

- [데이터 흐름](docs/data-flow.md): Excel 입력부터 검사, DB·Excel 저장, HTML 조회까지의 과정
- [DB 명세서](docs/database-spec.md): 실제 테이블·컬럼 구조, 검사 기준, 제약조건과 예외 처리
- [테이블 생성 SQL](docs/schema.sql): 현재 DB 구조를 재현하는 CREATE TABLE 문장 (데이터 제외)

## 구성
- Lens.py: 검사 프로그램
- data/lens_qualityCheck.xlsx: 기준표와 30개 측정데이터 (경도지수 제외)
- .venv/: 이 프로젝트의 Python 가상환경
- ../../resources/lens_quality/outputs/results/: 실행 시 생성되는 SQLite DB와 결과 Excel
- requirements.txt: 직접 사용하는 라이브러리 버전

## VSCode 실행
VSCode에서 이 lens_quality 폴더 자체를 여세요.
Python 인터프리터로 .venv/Scripts/python.exe를 선택하세요.

PowerShell 터미널에서 실행:

```powershell
.\.venv\Scripts\python.exe .\Lens.py
```

예상 결과: 합격 14개, 불합격 16개.
결과 파일은 ../../resources/lens_quality/outputs/results/lens_quality.db와 ../../resources/lens_quality/outputs/results/lens_inspection_result.xlsx입니다.
실행할 때마다 최신 데이터로 DB 테이블과 결과 Excel을 덮어씁니다.
결과 Excel을 열어둔 경우 닫고 실행하세요.

## 환경 재생성
가상환경 폴더는 복사하거나 이동하지 않고 새 위치에서 다시 만듭니다.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

VSCode에서 열 폴더: C:\Hitech_subsidy\AI_automize\workspace\lens_quality
기존 Excel과 미리보기: C:\Hitech_subsidy\AI_automize\resources\lens_quality\outputs\lens_test_20260930

