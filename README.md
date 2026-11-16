# 렌즈 검사 프로젝트

**Excel에 기록된 렌즈 측정값을 검사하고, 합격·불합격 결과와 그 이유를 보여주는 프로그램입니다.**

측정값을 검사 기준과 하나씩 비교하는 작업을 Python으로 처리하고, 결과를 Excel과 SQLite 데이터베이스에 저장합니다. 버튼과 표로 구성된 GUI 화면에서 검사 실행, 결과 확인, 데이터 조회와 내보내기를 할 수 있습니다.

QR 코드로 방문했다면 아래의 **실행 화면 → 판정 예시 → 단계별 제작 과정** 순서로 살펴보세요. 직접 사용하려면 [내 컴퓨터에서 실행하기](#내-컴퓨터에서-실행하기)를 따라 하면 됩니다. 이 프로그램은 PC에서 실행하는 Python 데스크톱 앱입니다.

> 제공된 제품 데이터와 검사 기준은 프로그램 개발 실습용입니다. 실제 품질 인증·출하 기준으로 검증된 값은 아니며, 경도지수는 검사 대상에서 제외했습니다.

## QR 코드로 프로젝트 공유하기

<a href="https://github.com/ungjo/lens_quality">
  <img src="docs/images/lens_quality_github_qr.png" alt="GitHub 렌즈 검사 프로젝트로 연결되는 QR 코드" width="246" height="246">
</a>

휴대폰 카메라로 QR 코드를 스캔하면 이 GitHub 프로젝트 페이지가 열립니다. 발표자료나 포스터에는 아래 원본 이미지 파일을 내려받아 사용하세요.

[프로젝트 페이지 열기](https://github.com/ungjo/lens_quality) · [QR 코드 PNG 원본](docs/images/lens_quality_github_qr.png) · [QR 코드 SVG 원본](docs/images/lens_quality_github_qr.svg)

## 프로그램으로 할 수 있는 일

- Excel의 측정값과 기준표를 읽어 제품을 검사합니다.
- 제품별 **합격·불합격·판정불가**와 항목별 사유를 확인합니다.
- 제품명 검색, 소재·판정 필터, 정렬로 원하는 데이터를 찾습니다.
- 검사 결과를 색상이 적용된 Excel 파일로 저장합니다.
- 저장된 DB를 다시 열거나 구조와 데이터를 SQL 파일로 내보냅니다.

### 실제 실행 화면

![렌즈 검사 GUI — 샘플 결과와 L002의 불합격 사유](docs/images/gui-overview.png)

현재 프로그램에 저장된 샘플 DB를 불러온 화면입니다. 선택한 제품의 측정값과 불합격 사유를 함께 확인할 수 있습니다. [화면을 크게 보기](docs/images/gui-overview.png)

## 샘플로 이해하는 검사 결과

샘플은 **제품 30개, 제품당 검사 항목 13개**로 구성되어 있습니다. 소재는 CR-39, 하이인덱스, 폴리카보네이트 3종입니다.

| 전체 제품 | 합격 | 불합격 | 판정불가 |
|---:|---:|---:|---:|
| 30개 | 14개 | 16개 | 0개 |

위 수치는 제공된 원본과 기준을 사용한 예상 결과입니다. 측정값이나 기준을 바꾸면 결과도 달라집니다.

| 제품 | 판정 | 이유 |
|---|---|---|
| L001 | 합격 | 13개 항목이 모두 허용 범위 이내 |
| L002 | 불합격 | 렌즈직경이 69.61 mm로 하한 70 mm보다 작음 |

예를 들어 렌즈직경 기준은 **70~74 mm**입니다. 70 또는 74처럼 경계값과 같아도 해당 항목은 합격입니다. 제품에 불합격 항목이 하나라도 있으면 최종 불합격이며, 불합격 없이 측정값·기준 문제로 판단할 수 없는 항목이 있으면 최종 판정불가입니다.

## 데이터는 이렇게 이동합니다

```text
원본 Excel: 측정데이터 + 기준표
                ↓
       입력 확인 · 숫자 변환
                ↓
     SQLite DB에 입력 데이터 저장
                ↓
       SQL로 저장한 데이터 조회
                ↓
   Python 조건문으로 제품·항목별 판정
                ↓
         DB에 검사 결과 저장
                ↓
      결과 Excel 저장 · GUI에 표시
```

**Python**은 데이터를 읽고 판정하는 역할을, **SQL**은 DB에 저장된 데이터를 조회하는 역할을 합니다. **SQLite**는 표들을 하나의 `.db` 파일에 보관하는 데이터베이스입니다.

GUI의 **검사 실행**은 위 과정을 수행합니다. **저장된 DB 열기**는 이미 저장된 결과를 읽어 보여주는 기능입니다.

## 단계별 제작 과정

아래는 프로그램을 구성한 기능을 6단계로 나누어 정리한 제작 과정입니다. 각 단계의 관련 파일을 누르면 실제 코드를 볼 수 있습니다.

### 1단계 — Excel에 측정 데이터와 기준 준비

먼저 제품 정보, 측정값, 검사 기준을 일정한 표 형태로 정리했습니다. 기존 Excel 판정을 그대로 가져오지 않고, 측정값과 기준을 사용해 Python에서 다시 판정하도록 입력 범위를 정했습니다.

| 입력 시트 | 사용하는 범위 | 내용 |
|---|---|---|
| 측정데이터 | A:P, 1행은 컬럼명 | 제품 정보 3개 + 측정 항목 13개 |
| 기준표 | A:F, 컬럼명 아래 최대 15행 | 검사항목·적용소재·단위·하한·목표치·상한 |

기준은 **공통 12행 + 소재별 굴절률 3행**입니다. 목표치는 참고용으로 저장하며, 합격 여부는 하한과 상한으로 결정합니다. Q열 이후에 있는 원본의 기존 판정·사유·보조 계산은 검사 입력에 포함하지 않습니다.

관련 파일: [원본 Excel](data/lens_qualityCheck.xlsx) · [입력 명세](docs/database-spec.md)

### 2단계 — Excel 데이터를 Python과 DB로 연결

pandas의 `read_excel()`로 Excel을 읽어 **DataFrame이라는 메모리 안의 표**로 만들었습니다. 필수 컬럼, 중복 제품번호, 누락된 검사 항목, 숫자로 변환할 수 없는 값 등을 검사한 후 SQLite에 저장했습니다.

저장할 때는 `to_sql()`이 테이블 생성과 데이터 삽입을 처리합니다. 저장한 표를 SQL로 다시 조회하여 다음 판정 단계에서 사용합니다.

```python
measurements.to_sql("measurements", conn, if_exists="replace", index=False)
criteria.to_sql("criteria", conn, if_exists="replace", index=False)
```

위 코드는 검사 함수 안에서 사용하는 일부입니다. `measurements`에는 측정 데이터를, `criteria`에는 검사 기준을 저장합니다.

관련 파일: [Lens.py — 입력·DB 처리](Lens.py) · [테이블 구조](docs/schema.sql)

### 3단계 — 조건문으로 합격·불합격과 사유 계산

제품 하나를 선택하고, 그 제품의 13개 검사 항목을 반복해서 확인하도록 만들었습니다. 검사항목과 소재에 맞는 기준을 찾은 뒤 `check_value()`에서 측정값을 비교합니다.

```text
하한 ≤ 측정값 ≤ 상한  → 합격
측정값 < 하한         → 불합격
측정값 > 상한         → 불합격
측정값·적용 기준 문제 → 판정불가
```

숫자 비교에 앞서 입력 오류와 결측값을 확인합니다. 예를 들어 **필수 검사 항목 자체가 기준표에서 빠지면 검사를 중단**하고, 입력 검증을 통과한 제품의 **개별 측정값이 비어 있으면 해당 항목을 판정불가**로 기록합니다.

제품별 최종 판정과 함께 항목별 측정값·하한·상한·판정·사유를 남겨, 어떤 값 때문에 불합격했는지 확인할 수 있도록 했습니다. 30개 제품을 13개 항목으로 검사하므로 상세 결과는 **390행**입니다.

관련 파일: [Lens.py — 판정 함수](Lens.py) · [판정 규칙과 예외 처리](docs/database-spec.md)

### 4단계 — 결과 Excel과 색상 표시 만들기

검사 결과를 DB에 저장한 뒤, openpyxl을 이용해 결과 Excel에 색상을 적용했습니다.

| 표시 대상 | 색상 |
|---|---|
| 합격 판정 셀 | 연두색 |
| 불합격 판정 셀 | 분홍색 |
| 불합격을 일으킨 측정값 셀 | 회색 |

결과 Excel은 **검사결과**, **항목별상세**, **적용기준**의 3개 시트로 구성했습니다. 제품 전체 결과와 개별 항목의 근거를 함께 확인할 수 있습니다.

저장은 임시 Excel을 완성한 뒤 최종 파일로 교체하는 방식입니다. Excel 파일이 열려 있어 저장에 실패한 경우, 열려 있는 결과 Excel을 닫은 뒤 GUI에서 저장을 재시도하거나 다른 이름으로 저장할 수 있도록 했습니다.

관련 파일: [Lens.py — Excel 저장·서식 처리](Lens.py)

### 5단계 — DB 조회와 SQL 내보내기 추가

저장된 데이터를 검사 후에도 살펴볼 수 있도록 `view_db.py`를 만들었습니다. 측정값, 기준, 결과, 상세 데이터를 각각 조회하고, 브라우저에서 표로 확인하거나 터미널에 출력하도록 했습니다.

또한 `export_sql.py`에서 DB의 테이블 구조와 현재 데이터를 SQL 파일로 내보내도록 했습니다. SQL에는 입력 데이터뿐 아니라 검사 결과도 포함됩니다. 이 작업들은 새 검사와 별도로 실행합니다.

관련 파일: [view_db.py — 조회·도움말](view_db.py) · [export_sql.py — SQL 내보내기](export_sql.py)

### 6단계 — 버튼으로 사용하는 GUI 제작

PyQt6로 파일 선택, 검사 실행, 검색·필터·정렬, 제품별 상세 확인, Excel·SQL 내보내기를 연결했습니다.

검사 기능을 `run_inspection()` 같은 함수로 분리하여 터미널과 GUI가 같은 기능을 사용하도록 했습니다. `table_model.py`는 DataFrame의 값을 GUI 표에 표시하고 판정 색상을 제공하며, `gui.py`는 버튼과 기능을 연결합니다.

검사·DB 조회·내보내기는 작업 스레드에서 처리하고 진행 상태를 화면에 전달하도록 구성했습니다.

관련 파일: [gui.py — 화면과 버튼](gui.py) · [table_model.py — 표 표시](table_model.py)

## 내 컴퓨터에서 실행하기

아래 안내는 **Windows PC에서 ZIP 파일을 내려받아 VSCode와 PowerShell로 실행하는 방법**입니다. 휴대폰으로 QR 코드를 찍었다면, 연결된 [프로젝트 페이지](https://github.com/ungjo/lens_quality)를 실행할 PC에서도 열어 주세요.

**처음 한 번:** 다운로드 → 압축 풀기 → Python·VSCode 준비 → 프로젝트 폴더 열기 → 가상환경 만들기 → 인터프리터 선택 → 라이브러리 설치 → GUI 실행

**두 번째부터:** 같은 프로젝트 폴더 열기 → [GUI 실행 명령](#8-gui-실행하고-검사하기) 입력

### 1. GitHub에서 프로젝트 다운로드하기

1. PC 브라우저에서 [프로젝트 페이지](https://github.com/ungjo/lens_quality)를 엽니다.
2. 파일 목록 위의 초록색 **Code** 버튼을 누릅니다.
3. **Download ZIP**을 누르면 프로젝트 파일이 ZIP으로 다운로드됩니다. Git이나 GitHub 계정을 따로 준비할 필요는 없습니다.
4. 파일 탐색기에서 다운로드한 ZIP을 마우스 오른쪽 버튼으로 누르고 **모두 압축 풀기**를 선택합니다.

다운로드·문서 폴더 안에 압축을 풀어 주세요. ZIP을 더블클릭해서 내부 파일을 보고 있는 상태에서는 다음 단계를 진행하지 않습니다. 기본 브랜치를 다운로드하면 폴더 이름은 보통 `lens_quality-main`입니다.

예시 위치는 다음과 같습니다. `사용자이름`은 자신의 Windows 사용자 폴더 이름입니다.

```text
C:\Users\사용자이름\Downloads\lens_quality-main
```

기존 개발자의 `C:\Hitech_subsidy\AI_automize` 경로를 똑같이 만들 필요는 없습니다. 다만 현재 코드의 기본 경로 계산 때문에 `C:\lens_quality`처럼 드라이브 바로 아래에 두지 말고, 위 예시처럼 **상위 폴더가 있는 위치**에 보관해 주세요.

다운로드 방법 참고: [GitHub 공식 ZIP 다운로드 안내](https://docs.github.com/en/repositories/working-with-files/using-files/downloading-source-code-archives)

### 2. Python과 VSCode 준비하기

| 준비할 프로그램 | 용도 | 준비 방법 |
|---|---|---|
| Python | 검사 프로그램을 실제로 실행 | [Python 공식 Windows 다운로드](https://www.python.org/downloads/windows/) |
| Visual Studio Code | 프로젝트 폴더·코드·터미널을 한곳에서 사용 | [VSCode 공식 다운로드](https://code.visualstudio.com/download) |
| VSCode의 Python 확장 | 프로젝트의 Python 인터프리터 선택 | VSCode에서 `Ctrl + Shift + X` → `Python` 검색 → 게시자가 **Microsoft**인 확장 설치 |

새 PC의 설치 예시는 **Python 3.13, 일반 Windows x64 PC용 64비트 버전**을 기준으로 설명합니다. Python 다운로드 페이지에서 3.13 계열의 **Windows installer (64-bit)**를 선택하세요. 설치 화면에서 **Add python.exe to PATH**를 선택하고, `pip`와 Python Launcher(`py`)를 함께 설치합니다. 설치가 끝나면 VSCode를 다시 열어 주세요.

> **버전 확인 범위:** 이 프로젝트를 개발하고 실행을 확인한 환경은 Python 3.10.11입니다. Python 3.10은 지원이 종료되어 새 설치 예시는 3.13으로 구분했습니다. 현재 라이브러리의 배포 정보를 바탕으로 선택한 버전이며, Python 3.13에서 프로그램 전체의 실행 검증은 아직 완료하지 않았습니다. 기존에 정상 동작하는 환경을 사용한다면 아래의 **기존 Python 사용 방법**을 참고하세요.
>
> 근거: [Python 3.10 지원 상태](https://www.python.org/downloads/release/python-31011/) · [pandas 배포 파일](https://pypi.org/project/pandas/2.3.3/#files) · [PyQt6 설치 조건](https://pypi.org/project/PyQt6/6.11.0/) · [Python Windows 설치 안내](https://docs.python.org/3.13/using/windows.html)

프로그램의 검사, GUI 표 조회, Excel 파일 생성에는 Microsoft Excel이나 별도 DB 서버가 필요하지 않습니다. 다만 **결과 Excel 열기** 버튼으로 `.xlsx` 파일을 직접 열려면 Excel처럼 해당 파일을 열 수 있는 연결 프로그램이 필요합니다.

### 3. VSCode에서 압축을 푼 폴더 열기

1. VSCode에서 **파일 → 폴더 열기**를 선택합니다.
2. 압축을 푼 폴더 안에서 **`gui.py`와 `requirements.txt`가 함께 있는 폴더**를 선택합니다.
3. 왼쪽 탐색기에 아래 파일들이 보이는지 확인합니다.

```text
lens_quality-main/
├─ data/
│  └─ lens_qualityCheck.xlsx
├─ docs/
├─ Lens.py
├─ gui.py
├─ table_model.py
├─ view_db.py
├─ export_sql.py
├─ requirements.txt
├─ requirements-lock.txt
└─ README.md
```

원본 데이터는 함께 내려받은 `data/lens_qualityCheck.xlsx`입니다. 데이터 파일을 따로 내려받거나 예전 `resources` 폴더를 복사할 필요는 없습니다.

처음 다운로드했을 때 `.venv`와 `results` 폴더가 없는 것은 정상입니다. `.venv`는 아래 단계에서 만들고, `results`는 검사를 실행할 때 생성합니다.

### 4. PowerShell 터미널과 Python 확인하기

VSCode 상단 메뉴에서 **터미널 → 새 터미널**을 누릅니다. 터미널 종류가 PowerShell인지 확인하세요. 다른 종류라면 터미널의 `+` 옆 메뉴에서 **PowerShell**을 선택합니다.

터미널에는 다음과 비슷한 경로가 표시됩니다.

```text
PS C:\Users\사용자이름\Downloads\lens_quality-main>
```

`PS`, 폴더 경로, 마지막 `>`는 터미널이 보여 주는 표시입니다. **명령어만 입력하고 Enter를 누르세요.** 명령은 편집기나 Python 파일 안이 아니라, 화면 아래 PowerShell 터미널에 입력합니다.

현재 위치와 파일을 확인하려면 다음 두 줄을 차례로 실행합니다.

```powershell
Get-Location
Get-ChildItem
```

파일 목록에 `gui.py`, `requirements.txt`, `data`가 보이면 맞는 위치입니다. 보이지 않으면 3단계에서 올바른 폴더를 다시 열고 새 터미널을 만드세요.

새로 설치한 Python 3.13을 확인합니다.

```powershell
py -3.13 --version
```

`Python 3.13.x`처럼 버전 번호가 나오면 다음 단계로 진행합니다. `py`는 설치된 Python 중 지정한 버전을 실행하는 명령입니다.

<details>
<summary>기존 Python 사용 방법: 이미 설치한 환경으로 실행하는 경우</summary>

기존에 이 프로젝트를 정상 실행한 Python이 있다면 다음 명령으로 버전을 확인할 수 있습니다.

```powershell
python --version
```

다음 5단계의 `py -3.13 -m venv .venv` 대신 **`python -m venv .venv`**를 한 번 실행합니다. 그 뒤의 설치·실행 명령은 동일합니다.

`py`와 `python`은 서로 다른 Python을 가리킬 수 있습니다. 가상환경을 만들 때 선택한 버전이 맞는지는 `.\.venv\Scripts\python.exe --version`으로 확인하세요.

</details>

### 5. 이 프로젝트의 가상환경 만들기

프로젝트 폴더에서 다음 명령을 실행합니다.

```powershell
py -3.13 -m venv .venv
```

명령이 끝나고 입력할 수 있는 상태로 돌아오면, 탐색기에 `.venv` 폴더가 생겼는지 확인합니다. 별도의 완료 문장이 출력되지 않아도 정상일 수 있습니다.

가상환경의 Python을 직접 실행해 확인합니다.

```powershell
.\.venv\Scripts\python.exe --version
```

이 가상환경은 프로젝트에서 사용할 Python과 라이브러리를 따로 관리하기 위한 공간입니다. 다른 컴퓨터의 `.venv`를 복사해 사용하지 않고 **사용할 컴퓨터에서 새로 만듭니다.** 프로젝트 폴더를 다른 위치로 옮긴 경우에도 가상환경을 다시 만드는 것이 좋습니다.

이후 명령은 항상 `.\.venv\Scripts\python.exe`를 직접 지정합니다. 따라서 `Activate.ps1`로 활성화하거나 PowerShell 실행 정책을 변경할 필요가 없고, 터미널 앞에 `(.venv)` 표시가 없어도 실행할 수 있습니다.

### 6. VSCode가 사용할 Python 선택하기

**가상환경을 만든 다음**, 아래 순서로 선택합니다.

1. `Ctrl + Shift + P`를 누릅니다.
2. **Python: Select Interpreter**를 검색해 선택합니다.
3. 현재 프로젝트의 **`.venv\Scripts\python.exe`**를 선택합니다.
4. 목록에 없다면 **Enter interpreter path(인터프리터 경로 입력)** → **Find(찾기)**를 선택하고 프로젝트 안의 `.venv` → `Scripts` → `python.exe`를 찾습니다.

이 설정은 VSCode에서 실행·디버깅할 때 사용할 Python을 지정합니다. 라이브러리를 설치하는 기능은 아니므로 다음 7단계도 진행해야 합니다.

명령 자체가 검색되지 않으면 2단계의 **Microsoft Python 확장**이 설치되어 있는지 확인하세요.

### 7. 필요한 라이브러리 설치하기

다음 명령을 PowerShell에 입력합니다. 처음 설치할 때는 인터넷 연결이 필요합니다.

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`requirements.txt`에 적힌 **pandas, openpyxl, PyQt6**와 관련 라이브러리를 가상환경 안에 설치합니다. 진행 중에는 기다렸다가 설치가 끝나고 새 명령을 입력할 수 있는 상태로 돌아오면 다음 확인 명령을 실행하세요.

```powershell
.\.venv\Scripts\python.exe -c "import pandas, openpyxl, PyQt6; print('Library imports OK')"
```

`Library imports OK`가 출력되면 세 라이브러리를 불러올 수 있다는 뜻입니다. 이는 설치 확인이며, 검사를 실행하거나 결과 파일을 생성하는 명령은 아닙니다. 오류가 나오면 아래 [문제 해결](#처음-실행할-때-자주-겪는-문제)을 먼저 확인하세요.

<details>
<summary>requirements.txt와 requirements-lock.txt 중 무엇을 설치하나요?</summary>

이 안내에서는 `requirements.txt`를 사용합니다. 기존 환경의 간접 의존성까지 기록된 버전으로 맞추려는 경우에는 위 설치 명령 **대신** 다음 명령을 사용합니다.

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-lock.txt
```

**둘 다 설치할 필요는 없습니다.** lock 파일을 사용해도 Python 자체가 설치되거나 Python 버전이 바뀌지는 않으며, 다른 Python·운영체제에서의 호환성까지 보장하는 것은 아닙니다.

</details>

### 8. GUI 실행하고 검사하기

라이브러리 설치를 마쳤다면 다음 명령을 실행합니다.

```powershell
.\.venv\Scripts\python.exe .\gui.py --output ".\results"
```

- `gui.py`: 버튼과 표가 있는 프로그램 창을 엽니다.
- `--output ".\results"`: **현재 터미널 폴더 안의 `results`**를 결과 저장 위치로 지정합니다.

창을 연 것만으로 검사가 시작되지는 않습니다. GUI에서 다음 순서대로 진행하세요.

1. **원본 Excel**이 프로젝트의 `data/lens_qualityCheck.xlsx`인지 확인합니다. 다른 파일이 선택되어 있다면 **파일 선택**으로 지정합니다.
2. **결과 저장 폴더**를 확인합니다. 다른 곳에 저장하려면 **폴더 선택**을 누릅니다.
3. **검사 실행**을 누르고 완료될 때까지 기다립니다.
4. **검사 결과** 탭에서 제품별 판정과 사유를 확인합니다. 제품 행을 선택하면 해당 제품의 상세 내용도 볼 수 있습니다.
5. **측정 데이터**, **검사 기준**, **항목 상세** 탭을 눌러 각각의 내용을 확인합니다.
6. **결과 폴더 열기**로 저장된 파일을 확인하거나 **결과 Excel 열기**로 Excel 파일을 엽니다.

제공된 샘플을 수정하지 않았다면 예상 결과는 **전체 30개 / 합격 14개 / 불합격 16개 / 판정불가 0개**입니다. 저장되는 파일은 바로 아래의 [생성되는 파일과 저장 위치](#생성되는-파일과-저장-위치)를 참고하세요.

**저장된 DB 열기**는 이미 만든 `lens_quality.db`를 읽어 보여 주는 기능입니다. 새 검사를 시작하는 버튼이 아닙니다. 새로 판정하려면 원본 Excel과 저장 폴더를 확인하고 **검사 실행**을 누릅니다.

SQL 문장이 들어 있는 파일도 필요하다면, 검사 완료 후 **SQL 내보내기**를 눌러 저장 위치를 선택합니다. 검사 실행만으로 `.sql` 파일이 생성되지는 않습니다.

### 다음부터 다시 실행하기

같은 컴퓨터의 같은 폴더를 계속 사용한다면 **가상환경 생성과 라이브러리 설치를 매번 반복할 필요가 없습니다.**

VSCode에서 프로젝트 폴더를 열고 새 PowerShell 터미널에서 다음 명령만 실행하세요.

```powershell
.\.venv\Scripts\python.exe .\gui.py --output ".\results"
```

GUI가 실행 중인 동안 해당 터미널은 프로그램 종료를 기다립니다. 다른 명령을 입력하려면 GUI 창을 닫거나 별도의 터미널을 여세요.

### 처음 실행할 때 자주 겪는 문제

| 보이는 현상 | 확인할 내용과 해결 방법 |
|---|---|
| `py`를 찾을 수 없거나 `No suitable Python runtime found`가 나옴 | 2단계에서 Python 3.13과 Launcher 설치 여부를 확인하고 VSCode를 다시 엽니다. 기존 Python을 사용할 때는 4단계의 접힌 안내를 참고합니다. |
| `python`을 입력하면 Microsoft Store가 열리거나 명령을 찾지 못함 | 이 안내의 `py -3.13 --version`으로 설치한 버전을 확인합니다. 설치되지 않았다면 2단계부터 진행합니다. |
| `requirements.txt` 또는 `gui.py`를 찾을 수 없음 | `Get-Location`, `Get-ChildItem`으로 현재 폴더를 확인합니다. 압축을 모두 푼 뒤 두 파일이 함께 있는 폴더를 VSCode에서 엽니다. |
| `.\.venv\Scripts\python.exe`를 찾을 수 없음 | 현재 폴더가 맞는지, 5단계에서 가상환경 생성이 완료됐는지 확인합니다. 다른 PC에서 가져온 가상환경은 새로 만들어야 합니다. |
| `No module named 'pandas'`, `'openpyxl'`, `'PyQt6'` | 7단계 설치 명령과 확인 명령을 실행합니다. GUI도 같은 `.\.venv\Scripts\python.exe`로 실행합니다. |
| `Python: Select Interpreter`가 보이지 않음 | VSCode에 게시자가 Microsoft인 Python 확장을 설치합니다. |
| `Lens.py`의 `BASE.parents[1]`에서 `IndexError`가 나옴 | 프로젝트가 드라이브 바로 아래에 있지 않은지 확인합니다. 다운로드·문서 폴더 안으로 옮긴 뒤 가상환경을 다시 만듭니다. |
| 원본 Excel을 찾지 못함 | GUI의 **파일 선택**으로 프로젝트의 `data/lens_qualityCheck.xlsx`를 지정합니다. |
| 결과 Excel 저장 시 `PermissionError` 또는 저장 실패가 나옴 | 열려 있는 결과 Excel을 닫고 저장 폴더에 쓰기 권한이 있는지 확인한 뒤 **Excel 저장 다시 시도**를 누릅니다. |
| **결과 Excel 열기**를 눌러도 파일이 열리지 않음 | **결과 폴더 열기**로 파일 생성 여부를 확인합니다. `.xlsx` 연결 프로그램이 없어도 GUI 안에서 결과를 조회할 수 있습니다. |
| `-h` 또는 `--help`만 입력하니 오류가 나옴 | 옵션만 입력하지 말고 아래처럼 Python과 파일명을 함께 입력합니다. |

설치가 완료된 뒤 GUI 명령의 도움말을 보려면 다음과 같이 실행합니다.

```powershell
.\.venv\Scripts\python.exe .\gui.py --help
```

도움말만 출력하고 종료하며 검사와 결과 파일 저장은 시작하지 않습니다.

## 생성되는 파일과 저장 위치

위 실행 예시에서는 검사 후 다음 두 파일이 만들어집니다.

```text
프로젝트 폴더/
└─ results/
   ├─ lens_inspection_result.xlsx
   └─ lens_quality.db
```

| 파일 | 저장 내용 |
|---|---|
| `lens_inspection_result.xlsx` | 검사결과·항목별상세·적용기준의 3개 시트 |
| `lens_quality.db` | 측정 데이터·검사 기준·제품별 결과·항목별 상세의 4개 테이블 |
| `lens_quality.sql` | **SQL 내보내기**를 별도로 실행했을 때 저장하는 구조와 데이터 |

같은 폴더에서 다시 검사하면 기존 DB의 4개 테이블과 결과 Excel을 최신 내용으로 갱신합니다. 이전 검사 이력을 누적하는 방식은 아닙니다.

DB 저장 후 Excel 저장만 실패할 수 있습니다. 오류가 표시되면 열려 있는 결과 Excel을 닫은 뒤 GUI의 **Excel 저장 다시 시도** 버튼을 누르세요. 저장할 파일 경로도 다시 선택할 수 있습니다. 이 기능은 해당 GUI 창에서 검사한 결과에 사용할 수 있으며, DB만 열어 조회한 상태에서는 사용할 수 없습니다.

## 사용한 기술

| 기술 | 프로젝트에서의 역할 |
|---|---|
| Python | 파일 읽기, 입력 검증, 조건문 판정, 저장 기능 연결 |
| pandas | Excel·SQL 데이터를 DataFrame 형태로 처리 |
| SQLite / SQL | 입력과 결과를 테이블로 저장하고 조회 |
| openpyxl | 결과 Excel 작성과 셀 색상 적용 |
| PyQt6 | 버튼·표·검색·필터가 있는 데스크톱 화면 구성 |
| HTML·CSS·JavaScript | 터미널 조회 결과를 브라우저용 표로 표시 |

`sqlite3`는 Python에 포함된 표준 라이브러리입니다. 별도 DB 서버를 설치하지 않고 사용합니다. 직접 사용하는 외부 라이브러리 버전은 [requirements.txt](requirements.txt)에 정리했습니다.

## 주요 파일 안내

| 파일 | 역할 |
|---|---|
| [gui.py](gui.py) | GUI를 여는 시작 파일. 사용자 동작과 기능을 연결 |
| [Lens.py](Lens.py) | Excel 입력 검증, 검사, DB·결과 Excel 저장 |
| [table_model.py](table_model.py) | GUI 표의 데이터·색상과 검색·필터·정렬 처리 |
| [view_db.py](view_db.py) | 터미널 도움말, DB 조회, HTML 생성, SQL 문장 확인 |
| [export_sql.py](export_sql.py) | DB 전체를 SQL 파일로 내보내기 |
| [data/lens_qualityCheck.xlsx](data/lens_qualityCheck.xlsx) | 샘플 측정값과 검사 기준 |
| [requirements.txt](requirements.txt) | 직접 사용하는 라이브러리와 버전 |
| [requirements-lock.txt](requirements-lock.txt) | 간접 의존성까지 포함한 라이브러리 버전 |

## 더 자세히 알아보기

- [데이터 흐름](docs/data-flow.md): 입력부터 검사·저장·조회·내보내기까지의 흐름도
- [데이터·DB 명세서](docs/database-spec.md): 컬럼·단위·기준값·판정 규칙·오류 처리
- [테이블 생성 SQL](docs/schema.sql): 데이터가 제외된 4개 테이블의 CREATE TABLE 정의

<details>
<summary><strong>터미널 명령과 도움말 보기</strong></summary>

### GUI에서 선택한 경로와 터미널의 기본 경로

`--output`을 생략한 GUI와 터미널의 `Lens.py`, `view_db.py`, `export_sql.py`는 코드에 지정된 기본 경로를 사용합니다. 기본 결과 폴더는 프로젝트 폴더 기준으로 다음 위치입니다.

```text
../../resources/lens_quality/outputs/results/
```

앞의 GUI 실행 예시에서 지정한 `.\results`와는 다른 위치입니다. **GUI에서 저장 폴더를 바꿔도 터미널 명령의 기본 경로가 바뀌지는 않습니다.** GUI에서 별도로 저장한 DB는 GUI의 **저장된 DB 열기**로 선택해 조회·내보내기하세요.

### 기본 경로로 검사·조회·내보내기

아래 명령은 프로젝트 폴더의 PowerShell에서 실행합니다.

```powershell
# 기본 경로로 검사하고 DB와 결과 Excel 저장
.\.venv\Scripts\python.exe .\Lens.py

# 기본 DB의 결과를 브라우저용 HTML로 조회
.\.venv\Scripts\python.exe .\view_db.py 결과

# 기본 DB의 기준을 터미널에서 조회
.\.venv\Scripts\python.exe .\view_db.py 기준 --터미널

# 결과 테이블의 CREATE TABLE 문장 확인
.\.venv\Scripts\python.exe .\view_db.py create 결과

# 기본 DB 전체를 SQL 파일로 내보내기
.\.venv\Scripts\python.exe .\export_sql.py
```

`view_db.py`의 일반 조회는 기본 결과 폴더의 `조회화면` 아래에 HTML을 생성합니다. `--터미널`은 터미널 출력만, `--열지않기`는 HTML 저장만 수행합니다. `create`, `insert`, `sql`은 SQL 문장을 보여 주는 명령이며 해당 문장을 실행하지 않습니다.

`export_sql.py`는 기본 DB와 같은 폴더에 `lens_quality.sql`을 저장합니다. 검사 실행만으로 HTML이나 SQL 파일이 만들어지지는 않습니다. DB를 갱신한 뒤에는 조회·내보내기를 다시 실행해야 최신 내용이 반영됩니다.

### 도움말

프로젝트 전체 사용법은 다음 명령으로 확인합니다.

```powershell
.\.venv\Scripts\python.exe .\view_db.py --help
```

`view_db.py`는 인수 생략, `help`, `도움말`, `-h`, `--help`의 다섯 방식으로 같은 도움말을 보여 줍니다. 다른 실행 파일은 다음처럼 확인합니다.

```powershell
.\.venv\Scripts\python.exe .\gui.py --help
.\.venv\Scripts\python.exe .\Lens.py --help
.\.venv\Scripts\python.exe .\export_sql.py --help
```

도움말은 검사·DB 조회·결과 저장을 시작하지 않습니다. 필요한 라이브러리는 설치되어 있어야 합니다. `gui.py`, `Lens.py`, `export_sql.py`를 인수 없이 실행하면 각각 GUI 열기, 검사, SQL 내보내기가 실행됩니다.

</details>
