"""SQLite DB를 SQL 파일로 내보냅니다. import만 하면 내보내기는 실행하지 않습니다."""

from contextlib import closing
from pathlib import Path
import sqlite3
from tempfile import TemporaryDirectory


def export_sql(db_path=None, output_path=None) -> Path:
    """DB의 테이블 구조와 데이터를 내보내고 저장한 파일의 경로를 반환합니다.

    경로를 생략하면 Lens.py의 기본 DB를 읽어 같은 폴더의 .sql 파일에 저장합니다.
    GUI에서는 검사 결과의 DB 경로와 원하는 저장 경로를 전달할 수 있습니다.
    오류는 호출한 곳으로 전달하며, 이 함수에서는 터미널 출력이나 창을 띄우지 않습니다.
    """
    # 기본 경로가 필요할 때만 Lens를 불러옵니다.
    # 따라서 import export_sql만 할 때는 pandas 등도 불러오지 않습니다.
    if db_path is None:
        from Lens import DB_FILE

        db_path = DB_FILE

    db_path = Path(db_path).resolve()
    if not db_path.is_file():
        raise FileNotFoundError(f"DB 파일을 찾을 수 없습니다: {db_path}\nLens.py를 먼저 실행하세요.")

    output = Path(output_path).resolve() if output_path is not None else db_path.with_suffix(".sql")
    if output == db_path or (output.exists() and output.samefile(db_path)):
        raise ValueError("SQL 파일은 원본 DB와 다른 경로에 저장해야 합니다.")

    output.parent.mkdir(parents=True, exist_ok=True)
    # 임시 파일을 완성한 뒤 교체하므로, 저장에 실패해도 기존 SQL은 보존합니다.
    with TemporaryDirectory(prefix=".lens_sql_", dir=output.parent) as temp_dir:
        temporary_file = Path(temp_dir) / output.name
        # GUI 작업에서도 연결을 만든 함수 안에서 사용하고 반드시 닫습니다.
        with closing(sqlite3.connect(db_path.as_uri() + "?mode=ro", uri=True)) as conn:
            # 내보내는 동안 같은 시점의 DB 내용을 읽습니다.
            conn.execute("BEGIN")
            with temporary_file.open("w", encoding="utf-8") as file:
                for statement in conn.iterdump():
                    file.write(statement + "\n")
        temporary_file.replace(output)

    return output


def main():
    """PowerShell에서 직접 실행하면 기본 DB를 내보내고 저장 경로를 출력합니다."""
    output = export_sql()
    print(f"SQL 파일 저장: {output}")


if __name__ == "__main__":
    main()
