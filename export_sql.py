import sqlite3
from Lens import DB_FILE

if not DB_FILE.is_file():
    raise FileNotFoundError("Lens.py를 먼저 실행하세요.")

output = DB_FILE.with_suffix(".sql")

with sqlite3.connect(
    DB_FILE.resolve().as_uri() + "?mode=ro",
    uri=True,
) as conn:
    with output.open("w", encoding="utf-8") as file:
        for statement in conn.iterdump():
            file.write(statement + "\n")

print(f"SQL 파일 저장: {output}")