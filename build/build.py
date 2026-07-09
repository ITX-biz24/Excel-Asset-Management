# -*- coding: utf-8 -*-
"""
ビルドオーケストレータ
======================
構造生成(openpyxl) → VBA注入/仕上げ(Excel COM) を順に実行し、
Excel-Asset-Management.xlsm を生成する。

    python build/build.py
"""
import os
import sys
import subprocess

BUILD_DIR = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable


def run(script):
    print(f"\n===== {script} =====")
    r = subprocess.run([PY, os.path.join(BUILD_DIR, script)], cwd=BUILD_DIR)
    if r.returncode != 0:
        print(f"[build] {script} が失敗しました (rc={r.returncode})")
        sys.exit(r.returncode)


def main():
    run("build_structure.py")
    run("inject_vba.py")
    print("\n[build] 完了: Excel-Asset-Management.xlsm を生成しました")


if __name__ == "__main__":
    main()
