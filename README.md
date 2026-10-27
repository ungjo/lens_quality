# 렌즈 검사 프로젝트

Python 3.10 환경에서 Excel의 측정값과 테스트용 기준을 읽어 SQLite에 저장하고 판정합니다.
이 기준은 프로그램 개발 실습용이며 실제 품질 인증·출하 기준이 아닙니다.

## 상세 문서

- [데이터 흐름](docs/data-flow.md): Excel 입력부터 검사, DB·Excel 저장, HTML 조회까지의 과정
- [DB 명세서](docs/database-spec.md): 실제 테이블·컬럼 구조, 검사 기준, 제약조건과 예외 처리
- [테이블 생성 SQL](docs/schema.sql): 현재 DB 구조를 재현하는 CREATE TABLE 문장 (데이터 제외)

## 구성
- gui.py: GUI 실행, 파일 선택, 검사·조회·저장 버튼 연결
- Lens.py: 검사 및 DB·결과 Excel 저장 (GUI와 터미널에서 함께 사용)
- table_model.py: GUI 표의 값·색상·검색·필터·정렬 처리 (단독 실행 불필요)
- view_db.py: 프로젝트 전체 도움말, 기본 DB 조회, HTML 생성과 SQL 문장 확인
- export_sql.py: 기본 DB 전체를 SQL 파일로 내보내기 (GUI에서도 함수 호출)
- data/lens_qualityCheck.xlsx: 기준표와 30개 측정데이터 (경도지수 제외)
- .venv/: 이 프로젝트의 Python 가상환경
- ../../resources/lens_quality/outputs/results/: 실행 시 생성되는 SQLite DB와 결과 Excel
- requirements.txt: 직접 사용하는 pandas, openpyxl, PyQt6 버전
- requirements-lock.txt: 전이 의존성까지 포함한 전체 라이브러리 버전

## VSCode 실행
VSCode에서 이 lens_quality 폴더 자체를 여세요.
Python 인터프리터로 .venv/Scripts/python.exe를 선택하세요.

프로젝트 폴더의 PowerShell에서 명령을 한 줄씩 실행하세요.
가상환경 Python을 직접 지정하면 활성화 여부와 관계없이 같은 환경을 사용합니다.

GUI 실행:

```powershell
.\.venv\Scripts\python.exe .\gui.py
```

- 창을 열면 입력을 기다립니다. 원본 Excel과 결과 저장 폴더를 선택하고 **검사 실행**을 누르세요.
- **저장된 DB 열기**는 기존 판정 결과를 읽기 전용으로 조회합니다. 새 판정은 하지 않습니다.
- DB 조회 후 **검사 실행**을 누르면 상단에 지정된 원본 Excel을 다시 읽어 판정합니다.
- 검색·정렬·판정 및 소재 필터, 제품별 상세·사유, SQL 미리보기를 사용할 수 있습니다.
- GUI 표는 HTML 파일을 생성하지 않습니다. **SQL 내보내기**를 누르면 현재 DB 전체를 `.sql`로 저장합니다.
- 결과 Excel 열기·다른 이름으로 저장은 이 창에서 검사한 결과에 사용할 수 있습니다. Excel 저장에 실패하면 저장만 다시 시도할 수 있습니다.

초기 경로를 직접 지정하려면:

```powershell
.\.venv\Scripts\python.exe .\gui.py --input ".\data\lens_qualityCheck.xlsx" --output ".\results"
```

`--output`은 폴더 경로입니다. 위 예시는 프로젝트 안의 `results` 폴더를 지정하며, 검사 버튼을 눌러야 파일이 생성됩니다.
GUI에서 변경한 경로는 터미널 명령의 기본 경로를 변경하지 않습니다.

터미널에서 기본 경로로 바로 검사하려면:

```powershell
.\.venv\Scripts\python.exe .\Lens.py
```

제공된 샘플 데이터의 예상 결과는 합격 14개, 불합격 16개입니다.
기본 결과 파일은 다음과 같습니다.

- `../../resources/lens_quality/outputs/results/lens_quality.db`
- `../../resources/lens_quality/outputs/results/lens_inspection_result.xlsx`

검사할 때 같은 경로의 기존 DB 테이블과 결과 Excel을 갱신합니다.
DB 저장 후 Excel 저장에 실패할 수 있으므로 오류 메시지를 확인하세요.
검사만으로 `.sql` 파일이 생성되지는 않습니다.

기본 DB 조회와 SQL 내보내기:

```powershell
.\.venv\Scripts\python.exe .\view_db.py 결과
.\.venv\Scripts\python.exe .\view_db.py create 결과
.\.venv\Scripts\python.exe .\export_sql.py
```

`view_db.py`의 일반 조회는 HTML 파일을 저장하고 연결된 브라우저 등으로 엽니다.
`--터미널`은 터미널 출력만, `--열지않기`는 HTML 저장만 수행합니다.
SQL 확인 명령은 SQL 문장을 출력하며 실행하지 않습니다.
`export_sql.py`는 입력 데이터와 검사 결과를 포함한 기본 DB 전체를 같은 폴더의 `lens_quality.sql`로 저장합니다.

GUI에서 다른 폴더에 저장한 DB는 GUI의 **저장된 DB 열기**로 선택하여 확인·내보내기하세요.
`view_db.py`와 `export_sql.py`는 `Lens.py`에 설정된 기본 DB를 사용합니다.
DB를 갱신한 뒤에는 HTML 조회와 SQL 내보내기도 다시 실행해야 최신 내용이 반영됩니다.

## 명령어 도움말

전체 명령과 파일별 역할, GUI 사용법, 기본 입출력 경로:

```powershell
.\.venv\Scripts\python.exe .\view_db.py --help
```

`view_db.py`는 아래 다섯 방식으로 동일한 도움말을 표시합니다.

```powershell
.\.venv\Scripts\python.exe .\view_db.py
.\.venv\Scripts\python.exe .\view_db.py help
.\.venv\Scripts\python.exe .\view_db.py 도움말
.\.venv\Scripts\python.exe .\view_db.py --help
.\.venv\Scripts\python.exe .\view_db.py -h
```

각 실행 파일의 개별 도움말:

```powershell
.\.venv\Scripts\python.exe .\gui.py --help
.\.venv\Scripts\python.exe .\Lens.py --help
.\.venv\Scripts\python.exe .\export_sql.py --help
```

네 실행 파일 모두 `--help` 대신 `-h`를 사용할 수 있습니다.
도움말은 DB가 없어도 볼 수 있으며 검사, DB 조회, 결과 파일 생성, GUI·브라우저 열기를 실행하지 않습니다.
필요한 라이브러리는 설치되어 있어야 합니다.

**인수 생략, `help`, `도움말` 방식은 `view_db.py`에서만 도움말로 처리합니다.**
인수 없이 `gui.py`를 실행하면 창이 열리고, `Lens.py`는 검사, `export_sql.py`는 내보내기를 실행합니다.
모듈을 `import`하는 것만으로 이러한 작업을 시작하지 않습니다.

## 환경 재생성
가상환경 폴더는 복사하거나 이동하지 않고 새 위치에서 다시 만듭니다.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

전체 의존성을 기록된 버전으로 설치하려면 위 설치 명령에서 `requirements.txt` 대신 `requirements-lock.txt`를 사용하세요. 두 파일을 모두 설치할 필요는 없습니다.

VSCode에서 열 폴더: C:\Hitech_subsidy\AI_automize\workspace\lens_quality
기존 Excel과 미리보기: C:\Hitech_subsidy\AI_automize\resources\lens_quality\outputs\lens_test_20260930

