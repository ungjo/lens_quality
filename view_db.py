"""명령별 DB 조회. 기본은 브라우저 표, --터미널은 텍스트 출력입니다."""
import argparse
from contextlib import closing
import os
from datetime import datetime
from html import escape
import sqlite3
import webbrowser

import pandas as pd
from Lens import DB_FILE, INPUT, OUTPUT

TABLES = {"측정": "measurements", "기준": "criteria",
          "결과": "inspection_results", "상세": "inspection_details"}

COMMAND_HELP = r"""프로젝트 폴더의 PowerShell에서 명령을 한 줄씩 실행하세요.
예시의 python은 pandas, openpyxl, PyQt6가 설치된 프로젝트 Python입니다.
가상환경을 활성화하지 않았다면 python 대신 .\.venv\Scripts\python.exe를 사용하세요.
  .\.venv\Scripts\python.exe .\view_db.py --help

[1] GUI 실행

  python .\gui.py
    창을 열고 사용자 입력을 기다립니다. 자동으로 검사하지 않습니다.
    원본 Excel과 결과 저장 폴더를 선택한 뒤 [검사 실행]을 누르세요.
    검사와 DB·결과 Excel 저장은 Lens.py의 run_inspection()이 처리합니다.

  python .\gui.py --input ".\data\lens_qualityCheck.xlsx" --output ".\results"
    원본 Excel과 결과 폴더의 초기값을 지정하여 창을 엽니다.
    --output은 폴더 경로이며, 검사는 [검사 실행]을 눌러야 시작합니다.

  GUI에서 할 수 있는 일:
    - 네 테이블 조회, 검색·정렬, 판정·소재 필터, 제품별 상세·사유 확인
    - 합격 연두색, 불합격 분홍색, 불합격 원인 측정값 회색 표시
    - SQL 구조·데이터 보기와 전체 SQL 내보내기
    - 결과 Excel 열기, 다른 이름으로 저장, 저장 실패 시 재시도

  [저장된 DB 열기]는 기존 판정 결과를 읽기 전용으로 조회합니다.
  이후 [검사 실행]은 상단의 원본 Excel을 읽어 새로 판정합니다.
  DB 조회만 한 상태에서는 Excel 저장·열기 버튼을 사용할 수 없습니다.
  GUI 표는 앱 안에 표시하며 HTML 파일을 만들지 않습니다.
  GUI에서 선택한 경로는 아래 터미널 명령의 기본 경로를 바꾸지 않습니다.

[2] 터미널에서 검사 실행 및 결과 저장

  python .\Lens.py
    기본 원본 Excel을 읽고 합격·불합격·판정불가를 판정합니다.
    DB와 결과 Excel을 생성하거나 기존 결과를 갱신합니다.
    .sql 파일은 자동 생성하지 않습니다. [6]의 내보내기를 실행하세요.
    DB 저장 후 Excel 저장에 실패하면 오류 사유를 출력하고 종료합니다.
    GUI에서 실행한 검사는 Excel 저장 실패 시 재검사 없이 저장만 재시도할 수 있습니다.

[3] DB 데이터 조회

  python .\view_db.py 측정
    제품 정보와 원본 측정값을 확인합니다.
  python .\view_db.py 기준
    검사항목, 적용소재, 단위, 하한·목표치·상한을 확인합니다.
  python .\view_db.py 결과
    제품별 최종 판정과 판정사유를 확인합니다.
  python .\view_db.py 상세
    제품의 각 검사항목별 측정값, 기준, 판정, 사유를 확인합니다.
  python .\view_db.py 불합격
    최종 판정이 불합격인 제품만 확인합니다.
  python .\view_db.py 전체
    네 테이블의 모든 데이터를 확인합니다.
  python .\view_db.py 테이블
    DB에 있는 테이블 이름을 확인합니다.
  python .\view_db.py 구조
    각 테이블의 컬럼 이름, 자료형, 기본키 여부 등을 확인합니다.

  위 명령은 기본 DB를 읽기 전용으로 조회합니다. 새 검사는 하지 않습니다.
  기본적으로 HTML 파일을 만들고 연결된 브라우저 등으로 엽니다.

[4] 조회 결과의 표시 방법

  python .\view_db.py 결과 --터미널
    HTML을 만들지 않고 터미널에 표를 출력합니다.
    --text도 같은 옵션입니다.
  python .\view_db.py 결과 --열지않기
    HTML 파일만 저장하고 브라우저는 열지 않습니다.
    --no-open도 같은 옵션입니다.

[5] SQL 문장 확인

  python .\view_db.py create
    네 테이블의 CREATE TABLE 문장을 터미널에 출력합니다.
  python .\view_db.py insert
    현재 데이터를 복원할 수 있는 INSERT INTO 문장을 출력합니다.
  python .\view_db.py sql
    CREATE TABLE과 INSERT INTO 문장을 함께 출력합니다.

  특정 테이블만 확인하려면 뒤에 대상을 입력하세요.
  python .\view_db.py create 결과
  python .\view_db.py insert 기준
  python .\view_db.py sql 상세

  대상 이름(실제 테이블 이름도 입력할 수 있습니다):
    측정 → measurements
    기준 → criteria
    결과 → inspection_results
    상세 → inspection_details

  SQL 확인 명령은 문장을 출력하며 해당 SQL을 실행하지 않습니다.
  INSERT 문은 현재 데이터의 복원용이며 과거 실행 기록이 아닙니다.

[6] SQL 파일로 내보내기

  python .\export_sql.py
    기본 DB의 구조와 현재 데이터를 같은 폴더의 lens_quality.sql로 저장합니다.
    입력 데이터와 검사 결과가 모두 포함됩니다. 기존 SQL 파일은 갱신합니다.
    다른 경로의 DB를 내보내려면 GUI에서 해당 DB를 열고 [SQL 내보내기]를 누르세요.
    GUI에서도 SQL 파일은 [SQL 내보내기]를 실행할 때 저장합니다.

[7] 도움말

  python .\view_db.py --help
  python .\view_db.py -h
  python .\view_db.py
  python .\view_db.py help
  python .\view_db.py 도움말

  위 다섯 명령은 같은 프로젝트 전체 도움말을 표시하고 종료합니다.
  각 실행 파일의 도움말:
  python .\gui.py --help
  python .\Lens.py --help
  python .\export_sql.py --help

  각 파일의 --help는 -h로 줄여 쓸 수 있습니다.
  인수 생략/help/도움말 방식은 view_db.py에서만 지원합니다.
  gui.py는 인수 없이 실행하면 창을 열고, Lens.py는 검사,
  export_sql.py는 SQL 내보내기를 실행합니다.
  도움말은 DB가 없어도 볼 수 있으며 필요한 라이브러리는 설치되어 있어야 합니다.
  도움말은 검사·DB 조회·결과 파일 생성·GUI/브라우저 열기를 실행하지 않습니다.

[8] 파일별 역할과 이용 순서

  gui.py          : GUI 시작점. 버튼과 검사·조회·저장 기능을 연결합니다.
  Lens.py         : 검사 및 DB·결과 Excel 저장. GUI와 터미널이 함께 사용합니다.
  table_model.py  : GUI 표의 값·색상·정렬·필터 담당. 따로 실행할 필요가 없습니다.
  view_db.py      : 전체 도움말, 기본 DB의 HTML/터미널 조회와 SQL 확인.
  export_sql.py   : DB를 SQL 파일로 저장. GUI와 터미널이 함께 사용합니다.
  requirements.txt / requirements-lock.txt:
    필요한 라이브러리 목록입니다. 다른 컴퓨터에서는 둘 중 하나로 설치합니다.
    python -m pip install -r requirements.txt
    같은 버전의 전체 의존성을 재현하려면 requirements-lock.txt를 선택하세요.

  DB가 없으면 Lens.py 또는 GUI의 [검사 실행]으로 먼저 DB를 만드세요.
  기존 DB가 있다면 재검사 없이 조회·내보내기를 실행할 수 있습니다.
  view_db.py와 export_sql.py는 아래 기본 DB를 사용합니다.
  GUI에서 다른 폴더에 저장한 DB는 [저장된 DB 열기]로 선택하여 확인하세요.
  DB를 갱신하면 기존 HTML 조회와 SQL 내보내기도 다시 실행해야 합니다.
  다른 파일에서 import하는 것만으로 검사·내보내기나 GUI 실행을 시작하지 않습니다.
"""


def build_help_text():
    """명령 예시와 현재 설정의 경로를 안내합니다. 파일을 읽거나 만들지 않습니다."""
    return COMMAND_HELP + f"""
[9] 기본 파일 경로

  프로젝트 폴더:
    {INPUT.parent.parent}
  원본 Excel:
    {INPUT}
  DB (터미널 검사 또는 GUI 기본 결과 폴더):
    {DB_FILE}
  결과 Excel (터미널 검사 또는 GUI 기본 결과 폴더):
    {OUTPUT}
  HTML (view_db.py의 일반 조회):
    {DB_FILE.parent / "조회화면" / "<명령>.html"}
  SQL (export_sql.py의 기본 저장 위치):
    {DB_FILE.with_suffix(".sql")}

  GUI에서 경로를 변경했다면 실제 조회·저장 경로는 GUI 화면에서 확인하세요.
"""


def read_table(conn, table, failed_only=False):
    if table not in TABLES.values():
        raise ValueError("지원하지 않는 테이블입니다.")
    sql = f'SELECT * FROM "{table}"'
    params = ()
    if failed_only:
        sql += ' WHERE "판정" = ?'
        params = ("불합격",)
    sql += ' ORDER BY rowid' if table == "criteria" else ' ORDER BY "제품번호", rowid'
    return pd.read_sql_query(sql, conn, params=params)


def make_html(frames, command):
    sections = []
    for name, frame in frames:
        table = frame.to_html(index=False, escape=True, border=0, na_rep="", classes="grid")
        sections.append(f'<section><h2>{escape(name)}</h2><p class="count">{len(frame)}행 · {len(frame.columns)}열</p><div class="scroll">{table}</div></section>')
    return '''<!doctype html><html lang="ko"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>렌즈 검사 DB 조회</title>
<style>
*{box-sizing:border-box}body{margin:0;padding:24px;background:#f5f7fa;color:#182b40;font:15px "Malgun Gothic",sans-serif}
h1{margin:0 0 8px;font-size:25px}h2{font-size:17px;margin:0}p{line-height:1.6}.muted,.count{color:#566578;font-size:13px}
.toolbar{display:flex;gap:16px;align-items:center;flex-wrap:wrap;margin:18px 0}input{padding:11px;width:380px;max-width:100%;border:1px solid #aab8c8;border-radius:6px;font:inherit}
section{background:white;border:1px solid #d4dde7;border-radius:8px;padding:16px;margin:18px 0}.count{margin:6px 0 12px}
.scroll{max-height:68vh;overflow:auto;border:1px solid #cdd7e1}table{border-collapse:separate;border-spacing:0;min-width:100%;width:max-content}
th,td{padding:10px 14px;border-right:1px solid #dce3eb;border-bottom:1px solid #dce3eb;text-align:left;white-space:pre;vertical-align:middle}
th{position:sticky;top:0;z-index:2;background:#233c58;color:white;font-weight:600;white-space:nowrap}
tbody tr:nth-child(even){background:#f1f5f9}tbody tr:hover{background:#e9f2ff}td.pass{background:#d9ead3;color:#274e13}td.fail{background:#f4cccc;color:#8c2020}
</style><h1>렌즈 검사 DB 조회 · ''' + escape(command) + '''</h1><p class="muted">조회 시각: ''' + datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ''' · 읽기 전용 조회<br>DB를 갱신한 뒤에는 터미널에서 조회 명령을 다시 실행하세요.</p>
<div class="toolbar"><label for="search">표 안에서 검색</label><input id="search" type="search" placeholder="제품번호, 소재, 제품명 등 입력"><span class="muted">표 아래 가로 스크롤로 오른쪽 열까지 확인할 수 있습니다.</span></div>
''' + ''.join(sections) + '''
<script>
const sections=[...document.querySelectorAll('section')];
for(const section of sections){
 const headers=[...section.querySelectorAll('th')].map(x=>x.textContent.trim());
 for(const row of section.querySelectorAll('tbody tr'))for(let i=0;i<row.cells.length;i++){
  if(['판정','항목판정'].includes(headers[i])){const cell=row.cells[i];if(cell.textContent==='합격')cell.classList.add('pass');if(cell.textContent==='불합격')cell.classList.add('fail');}
 }
}
document.querySelector('#search').addEventListener('input',event=>{
 const query=event.target.value.trim().toLocaleLowerCase();
 for(const section of sections){const rows=[...section.querySelectorAll('tbody tr')];let visible=0;
  for(const row of rows){row.hidden=!row.textContent.toLocaleLowerCase().includes(query);if(!row.hidden)visible++;}
  section.querySelector('.count').textContent=`${visible} / ${rows.length}행 · ${section.querySelectorAll('th').length}열`;
 }
});
</script></html>'''


def show_sql(conn, command, table=None):
    """생성문 또는 현재 데이터의 복원용 입력문을 출력합니다."""
    names = [table] if table else list(TABLES.values())
    if command in ("create", "sql"):
        for name in names:
            row = conn.execute(
                "SELECT sql FROM sqlite_master WHERE type='table' AND name=?",
                (name,),
            ).fetchone()
            if row is None:
                raise sqlite3.OperationalError(f"테이블이 없습니다: {name}")
            print(row[0].rstrip(";") + ";\n")
    if command in ("insert", "sql"):
        print("-- 현재 저장된 데이터의 복원용 INSERT 문입니다. 과거 실행 기록이 아닙니다.")
        prefixes = tuple(f'INSERT INTO "{name}" ' for name in names)
        for statement in conn.iterdump():
            if statement.startswith(prefixes):
                print(statement)

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="렌즈 검사 프로젝트 - 명령어 안내",
        usage=r"python .\view_db.py [명령] [대상테이블] [옵션]",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=build_help_text(),
        add_help=False,
    )
    commands = parser.add_argument_group("명령과 대상")
    commands.add_argument(
        "command", nargs="?", metavar="명령",
        choices=[*TABLES, "불합격", "전체", "테이블", "구조",
                 "create", "insert", "sql", "help", "도움말", *TABLES.values()],
        help="조회·SQL 명령. 생략하거나 help/도움말을 입력하면 사용법 표시",
    )
    commands.add_argument(
        "table", nargs="?", metavar="대상테이블",
        choices=[*TABLES, *TABLES.values()],
        help="create/insert/sql의 대상: 측정·기준·결과·상세 (생략: 전체)",
    )
    options = parser.add_argument_group("옵션")
    options.add_argument(
        "-h", "--help", action="help", help="도움말을 표시하고 종료",
    )
    options.add_argument(
        "--터미널", "--text", dest="text", action="store_true",
        help="일반 조회 결과를 HTML 대신 터미널에 출력",
    )
    options.add_argument(
        "--열지않기", "--no-open", dest="no_open", action="store_true",
        help="일반 조회 결과를 HTML로 저장하고 자동으로 열지 않음",
    )
    args = parser.parse_args(argv)
    if args.table and args.command not in ("create", "insert", "sql"):
        parser.error("두 번째 테이블 인수는 create, insert, sql 명령에서만 사용합니다.")
    if args.command in (None, "help", "도움말"):
        parser.print_help()
        return
    if not DB_FILE.is_file():
        parser.exit(1,f"DB 파일이 없습니다: {DB_FILE}\nLens.py를 실행하거나 GUI의 기본 결과 폴더에서 검사를 실행하세요.\n")
    try:
        with closing(sqlite3.connect(DB_FILE.resolve().as_uri()+"?mode=ro",uri=True)) as conn:
            if args.command in ("create", "insert", "sql"):
                show_sql(conn, args.command, TABLES.get(args.table, args.table))
                return
            if args.command == "테이블":
                frames=[("테이블 목록",pd.read_sql_query("SELECT name AS 테이블명 FROM sqlite_master WHERE type='table' ORDER BY name",conn))]
            elif args.command == "구조":
                frames=[(table+" 열 구조",pd.read_sql_query(f'PRAGMA table_info("{table}")',conn)) for table in TABLES.values()]
            elif args.command == "전체":
                frames=[(table,read_table(conn,table)) for table in TABLES.values()]
            else:
                table=TABLES.get(args.command,args.command)
                if args.command == "불합격":table="inspection_results"
                frames=[(table,read_table(conn,table,args.command=="불합격"))]
    except (sqlite3.Error,pd.errors.DatabaseError) as exc:
        parser.exit(1,f"DB 조회 오류: {exc}\n")
    if args.text:
        for name,frame in frames:
            print(f"\n[{name}] {len(frame)}행 / {len(frame.columns)}열")
            with pd.option_context("display.unicode.east_asian_width",True):
                print(frame.to_string(index=False,max_rows=None,max_cols=None))
    else:
        output=DB_FILE.parent / "조회화면" / (args.command+".html")
        output.parent.mkdir(parents=True,exist_ok=True)
        output.write_text(make_html(frames,args.command),encoding="utf-8")
        print(f"조회 화면 저장: {output}")
        if not args.no_open:
            print(f"직접 열기: {output.resolve().as_uri()}")
            try:
                if os.name == "nt":
                    # Windows의 HTML 연결 프로그램으로 파일 자체를 엽니다.
                    os.startfile(str(output.resolve()))
                elif not webbrowser.open(output.resolve().as_uri()):
                    print("자동 열기 실패: 브라우저에서 Ctrl+O로 위 HTML 파일을 선택하세요.")
            except OSError as exc:
                print(f"자동 열기 실패: {exc}")
                print("브라우저에서 Ctrl+O를 누르고 위 HTML 파일을 선택하세요.")


if __name__ == "__main__":
    main()



