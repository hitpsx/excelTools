"""PyInstaller 打包脚本。

执行：
    python build.py

产物：
    dist/ExcelTools.exe  (单文件,无控制台窗口)
"""
import os
import shutil
import subprocess
import sys


APP_NAME = "ExcelTools"
ENTRY = "main.py"
DIST_DIR = "dist"
BUILD_DIR = "build"


def clean() -> None:
    for d in (DIST_DIR, BUILD_DIR):
        if os.path.isdir(d):
            print(f"[clean] removing {d}")
            shutil.rmtree(d, ignore_errors=True)
    spec = f"{APP_NAME}.spec"
    if os.path.isfile(spec):
        os.remove(spec)


def build() -> None:
    clean()
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        APP_NAME,
        ENTRY,
    ]
    print("[build]", " ".join(cmd))
    subprocess.check_call(cmd)
    print(f"\n[done] 产物: {os.path.join(DIST_DIR, APP_NAME + ('.exe' if os.name == 'nt' else ''))}")


if __name__ == "__main__":
    build()
