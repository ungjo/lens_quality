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

COMMAND_HELP = r"""[1] 검사 실행 및 결과 저장

  python .\Lens.py
    원본 Excel을 읽고 합격·불합격·판정불가를 판정합니다.
    DB와 결과 Excel을 생성하거나 기존 결과를 갱신합니다.

[2] DB 데이터 조회

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

  위 조회 명령은 기본적으로 HTML 파일을 만들고 브라우저로 엽니다.

[3] 조회 결과의 표시 방법

  python .\view_db.py 결과 --터미널
    HTML을 만들지 않고 터미널에 표를 출력합니다.
    --text도 같은 옵션입니다.
  python .\view_db.py 결과 --열지않기
    HTML 파일만 저장하고 브라우저는 열지 않습니다.
    --no-open도 같은 옵션입니다.

[4] SQL 문장 확인

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

[5] SQL 파일로 내보내기

  python .\export_sql.py
    네 테이블의 구조와 현재 데이터를 lens_quality.sql로 저장합니다.
    입력 데이터와 검사 결과가 모두 포함됩니다.

[6] 도움말

  python .\view_db.py --help
  python .\view_db.py -h
  python .\view_db.py
  python .\view_db.py help
  python .\view_db.py 도움말

  다섯 명령은 같은 도움말을 표시하고 종료합니다.
  DB가 없어도 도움말을 볼 수 있습니다.
  도움말은 검사·DB 조회·파일 생성·브라우저 열기를 실행하지 않습니다.

이용 순서:
  프로젝트 폴더에서 필요한 라이브러리가 설치된 Python으로 실행하세요.
  처음에는 Lens.py로 DB를 만든 뒤 조회·내보내기를 실행하세요.
  DB를 갱신하면 HTML 조회와 SQL 내보내기도 다시 실행해야 합니다.
"""


def build_help_text():
    """명령 예시와 현재 설정의 경로를 안내합니다. 파일을 읽거나 만들지 않습니다."""
    return COMMAND_HELP + f"""
[7] 현재 파일 경로

  원본 Excel:
    {INPUT}
  DB (Lens.py에서 생성):
    {DB_FILE}
  결과 Excel (Lens.py에서 생성):
    {OUTPUT}
  HTML (view_db.py에서 생성):
    {DB_FILE.parent / "조회화면" / "<명령>.html"}
  SQL (export_sql.py에서 생성):
    {DB_FILE.with_suffix(".sql")}
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
        parser.exit(1,f"DB 파일이 없습니다: {DB_FILE}\nLens.py를 먼저 실행하세요.\n")
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



