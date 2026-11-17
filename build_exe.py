r"""Windows 배포 폴더와 ZIP 생성. 실행: python .\build_exe.py"""

from datetime import datetime, timezone
import argparse
import hashlib
from importlib import metadata
import json
import os
from pathlib import Path
import platform
import shutil
import struct
import subprocess
import sys
from zipfile import ZIP_DEFLATED, ZipFile


BASE = Path(__file__).resolve().parent
APP_NAME = "LensQuality"
REPOSITORY = "https://github.com/ungjo/lens_quality"


def sha256(path):
    """큰 실행 파일도 한 번에 메모리에 올리지 않고 해시를 계산합니다."""
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def zip_folder(folder, destination):
    """실행 파일뿐 아니라 _internal까지 폴더 전체를 ZIP에 포함합니다."""
    with ZipFile(destination, "x", ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(folder.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(folder.parent).as_posix())
    with ZipFile(destination) as archive:
        if archive.testzip() is not None:
            raise RuntimeError("배포 ZIP의 무결성 검사에 실패했습니다.")


def copy_licenses(package):
    """빌드 환경에 설치된 패키지의 라이선스 원문과 출처를 함께 보관합니다."""
    records = []
    for distribution in sorted(metadata.distributions(), key=lambda d: d.metadata["Name"].lower()):
        name = distribution.metadata["Name"]
        target = package / "licenses" / name
        target.mkdir(parents=True, exist_ok=True)
        # METADATA에는 프로젝트 주소, 버전, 라이선스 선언이 들어 있습니다.
        (target / "METADATA.txt").write_text(distribution.read_text("METADATA") or "", encoding="utf-8")
        for entry in distribution.files or ():
            if any(part.lower().startswith(("license", "licence", "copying", "notice")) for part in entry.parts):
                original = Path(distribution.locate_file(entry))
                if original.is_file() and ".." not in entry.parts:
                    destination = target / Path(*entry.parts)
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(original, destination)
        records.append({"name": name, "version": distribution.version})
    # Python 자체의 라이선스도 실행기와 함께 제공합니다.
    for name in ("LICENSE.txt", "LICENSE"):
        original = Path(sys.base_prefix) / name
        if original.is_file():
            shutil.copy2(original, package / "licenses" / "Python-LICENSE.txt")
            break
    return records


def source_archive(package):
    """가상환경·기존 결과·임시 파일을 제외한 제작 소스를 함께 제공합니다."""
    names = (
        "Lens.py", "gui.py", "table_model.py", "view_db.py", "export_sql.py",
        "build_exe.py", "requirements.txt", "requirements-lock.txt", "requirements-build.txt",
        "README.md", ".gitignore", "LICENSE",
    )
    paths = [BASE / name for name in names if (BASE / name).is_file()]
    paths += [BASE / "data" / "lens_qualityCheck.xlsx"]
    paths += [path for path in (BASE / "docs").rglob("*") if path.is_file()]
    destination = package / "source"
    destination.mkdir()
    with ZipFile(destination / "lens_quality-source.zip", "x", ZIP_DEFLATED) as archive:
        for path in sorted(paths):
            archive.write(path, (Path("lens_quality") / path.relative_to(BASE)).as_posix())


def main(argv=None):
    # 도움말만 요청한 경우에는 빌드와 파일 생성을 시작하지 않습니다.
    parser = argparse.ArgumentParser(description="샘플 Excel을 포함한 Windows 실행 파일과 배포 ZIP을 만듭니다.")
    parser.parse_args(argv)
    if sys.platform != "win32" or struct.calcsize("P") != 8:
        raise SystemExit("이 빌드는 Windows 64비트 Python에서 실행하세요.")
    if sys.version_info < (3, 11):
        raise SystemExit("빌드 스크립트는 Python 3.11 이상이 필요합니다. 안내는 3.13 기준입니다.")
    metadata.version("PyInstaller")  # 도구가 없으면 출력 폴더를 만들기 전에 실패합니다.
    subprocess.run([sys.executable, "-m", "pip", "check"], check=True)
    sample = BASE / "data" / "lens_qualityCheck.xlsx"
    sample_hash = sha256(sample)
    # 기존 배포물을 지우지 않도록 빌드마다 별도 폴더를 만듭니다.
    build_id = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    work = BASE / "build" / build_id
    output = BASE / "dist" / build_id
    work.mkdir(parents=True)
    output.mkdir(parents=True)
    command = [
        sys.executable, "-m", "PyInstaller", "--onedir", "--windowed", "--noupx",
        "--name", APP_NAME, "--specpath", str(work), "--workpath", str(work / "pyinstaller"),
        "--distpath", str(output), "--add-data", f"{sample}:data", str(BASE / "gui.py"),
    ]
    log = work / "pyinstaller.log"
    print(f"BUILD LOG: {log}", flush=True)
    # 다른 프로그램의 PATH에 있는 동명 DLL(ICU 등)이 배포물에 섞이지 않게 합니다.
    # Qt와 NumPy의 자체 DLL은 PyInstaller의 패키지별 hook이 찾아 포함합니다.
    build_environment = os.environ.copy()
    windows = Path(os.environ["SystemRoot"])
    build_environment["PATH"] = os.pathsep.join(map(str, (
        Path(sys.executable).parent, Path(sys.base_prefix), windows / "System32", windows,
    )))
    for name in ("PYTHONHOME", "PYTHONPATH", "QT_PLUGIN_PATH", "QT_QPA_PLATFORM_PLUGIN_PATH"):
        build_environment.pop(name, None)
    # 상세 출력은 파일에 남겨 빌드 오류가 나면 원인을 확인할 수 있게 합니다.
    with log.open("w", encoding="utf-8") as stream:
        result = subprocess.run(command, cwd=BASE, env=build_environment, stdout=stream, stderr=subprocess.STDOUT)
    if result.returncode:
        raise SystemExit(f"PyInstaller 실패. 로그를 확인하세요: {log}")
    package = output / APP_NAME
    executable = package / f"{APP_NAME}.exe"
    bundled_sample = package / "_internal" / "data" / sample.name
    if not executable.is_file() or sha256(bundled_sample) != sample_hash:
        raise RuntimeError("실행 파일 또는 포함된 샘플 Excel 확인에 실패했습니다.")
    shutil.copy2(BASE / "docs" / "exe-quickstart.txt", package / "사용방법.txt")
    shutil.copy2(BASE / "docs" / "third-party-notices.txt", package / "THIRD_PARTY_NOTICES.txt")
    if (BASE / "LICENSE").is_file():
        shutil.copy2(BASE / "LICENSE", package / "LICENSE")
    versions = copy_licenses(package)
    source_archive(package)
    info = {
        "built_at_utc": datetime.now(timezone.utc).isoformat(), "python": sys.version,
        "platform": platform.platform(), "repository": REPOSITORY, "format": "PyInstaller onedir, Windows x64",
        "input_sha256": sample_hash, "exe_sha256": sha256(executable), "packages": versions,
        "source_hashes": {name: sha256(BASE / name) for name in ("Lens.py", "gui.py", "table_model.py", "view_db.py", "export_sql.py", "build_exe.py")},
    }
    (package / "BUILD-INFO.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    archive = output / "LensQuality-windows-x64.zip"
    zip_folder(package, archive)
    (output / "SHA256SUMS.txt").write_text(f"{sha256(archive)}  {archive.name}\n", encoding="ascii")
    print(f"EXE: {executable}\nZIP: {archive}\nSHA256: {sha256(archive)}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
