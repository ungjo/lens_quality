"""명령별 DB 조회. 기본은 브라우저 표, --터미널은 텍스트 출력입니다."""
import argparse
from contextlib import closing
import os
from datetime import datetime
from html import escape
import sqlite3
import webbrowser

import pandas as pd
from Lens import DB_FILE

TABLES = {"측정": "measurements", "기준": "criteria",
          "결과": "inspection_results", "상세": "inspection_details"}


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
    parser = argparse.ArgumentParser(description="DB 데이터를 열이 정렬된 브라우저 표로 조회합니다.")
    parser.add_argument("command", nargs="?", choices=[*TABLES,"불합격","전체","테이블","구조","create","insert","sql",*TABLES.values()])
    parser.add_argument("--터미널", "--text", dest="text", action="store_true", help="기존 터미널 텍스트 출력")
    parser.add_argument("--열지않기", "--no-open", dest="no_open", action="store_true", help="HTML 파일만 저장")
    parser.add_argument("table", nargs="?", choices=[*TABLES, *TABLES.values()], help="create/insert/sql 조회 대상을 선택 (생략하면 네 테이블 전체)")
    args = parser.parse_args(argv)
    if args.table and args.command not in ("create", "insert", "sql"):
        parser.error("두 번째 테이블 인수는 create, insert, sql 명령에서만 사용합니다.")
    if args.command is None:
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



