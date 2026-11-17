# Windows 실행 파일 만들기

이 배포판은 프로그램의 동작과 제작 아이디어를 공유하는 데모입니다.
받는 사람은 ZIP을 모두 풀고 `LensQuality.exe`를 더블클릭합니다.
Python·VSCode·DB 서버 설치는 필요하지 않습니다.

## 만드는 사람의 PowerShell 명령

프로젝트 폴더에서 Windows 64비트 Python 3.13을 사용합니다.
기존 `.venv`와 별도로 **처음 한 번** 빌드 환경을 만듭니다.

```powershell
py -3.13 -m venv .venv-build
.\.venv-build\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv-build\Scripts\python.exe .\build_exe.py
```

`py -3.13`을 찾지 못하면 먼저 Python 3.13 64비트를 설치합니다.
환경을 만든 뒤 다음 빌드부터는 마지막 명령만 실행하면 됩니다.
라이브러리 목록이 바뀌었을 때는 설치 명령도 다시 실행합니다.
PowerShell 활성화나 실행 정책 변경은 필요하지 않습니다.

빌드는 `gui.py`를 시작 파일로 사용하고, 일반 import로 연결된 검사·표 표시·조회·내보내기 모듈도 수집합니다.
샘플 `data/lens_qualityCheck.xlsx`를 `_internal/data`에 포함하며 원본은 수정하지 않습니다.
다른 프로그램의 동명 DLL이 섞이지 않도록 빌드 하위 프로세스의 PATH를 프로젝트 Python과 Windows 기본 폴더로 제한합니다. PC 전체의 환경 변수는 변경하지 않습니다.

## 만들어지는 파일

```text
dist/<빌드날짜-시간>/
├─ LensQuality/
│  ├─ LensQuality.exe
│  ├─ _internal/                 # Python, Qt, 라이브러리, 샘플 Excel
│  ├─ 사용방법.txt
│  ├─ BUILD-INFO.json             # 빌드 버전, 소스·샘플·exe 해시
│  ├─ THIRD_PARTY_NOTICES.txt
│  ├─ licenses/
│  └─ source/lens_quality-source.zip
├─ LensQuality-windows-x64.zip
└─ SHA256SUMS.txt
```

이전 배포물을 지우지 않도록 매번 별도 폴더를 만듭니다.
공유할 파일은 **LensQuality-windows-x64.zip**입니다. `.exe`만 전달하지 않습니다.
`build/<빌드날짜-시간>/pyinstaller.log`에는 상세 빌드 로그가 있습니다.
`.venv-build`, `build`, `dist`는 Git에서 제외합니다. ZIP은 필요할 때 GitHub Release의 첨부 파일로 직접 올립니다.

## 경로와 기존 실행 방식

- `.exe`: 기본 결과는 `Path.home() / "LensQuality" / "results"`에 저장합니다.
- `python gui.py`, `python Lens.py`: 기존 개발용 기본 경로를 유지합니다.
- GUI에서 결과 폴더를 변경할 수 있으며, 창을 여는 것만으로 검사는 실행되지 않습니다.
- 포함된 샘플은 읽기만 합니다. 수정한 데이터를 체험하려면 사본을 만든 뒤 GUI에서 선택하세요.
- Python 실행 검증과 `.exe` 실행 검증은 별개입니다.

## 빌드 뒤 확인

ZIP을 프로젝트 밖의 한글·공백 포함 폴더에 풀어 실행합니다.
샘플 검사 결과 30/14/16/0, Excel·DB 저장, 검색·필터·상세 조회, SQL 내보내기와 저장된 DB 열기를 확인합니다.
최종적으로 Python이 설치되지 않은 Windows PC에서도 실행해 확인하세요.
다른 OS나 모든 Windows 버전에서의 실행까지 보장하는 빌드는 아닙니다.

PyInstaller는 Python 실행기와 라이브러리를 함께 수집합니다.
공식 문서: https://pyinstaller.org/en/stable/operating-mode.html
