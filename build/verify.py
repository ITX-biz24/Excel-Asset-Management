# -*- coding: utf-8 -*-
"""
構造検証スクリプト（win32com）
==============================
生成した .xlsx / .xlsm を Excel COM で開き、フル再計算してから
主要セル（口座残高・KPI）とエラーセル数を報告する。破損検知も兼ねる。

使い方:
    python verify.py [path]
"""
import os
import sys
import io
import win32com.client as win32

BUILD_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT = os.path.join(BUILD_DIR, "_structure.xlsx")
REPORT = os.path.join(BUILD_DIR, "verify_out.txt")

XL_CELLTYPE_FORMULAS = -4123

# レポート行を蓄積し、UTF-8ファイルへ保存（コンソールcp932の文字化け回避）
_lines = []
def report(msg=""):
    _lines.append(str(msg))


def find_sheet_with_table(wb, table_name):
    for s in wb.Worksheets:
        for lo in s.ListObjects:
            if lo.Name == table_name:
                return s
    return None


def verify(path):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        report(f"RESULT: ERROR - not found: {path}")
        return 1
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    rc = 0
    try:
        wb = xl.Workbooks.Open(path)
        xl.CalculateFullRebuild()

        report("=== シート一覧 ===")
        for s in wb.Worksheets:
            report(f"  - {s.Name}")

        m = find_sheet_with_table(wb, "T_Account")
        if m is not None:
            report("=== 口座現在残高（T_Account）===")
            r = 6
            while True:
                nm = m.Range(f"B{r}").Text
                if not nm:
                    break
                report(f"  {nm:<8} 初期={m.Range(f'C{r}').Text:>10}  現在={m.Range(f'E{r}').Text}")
                r += 1

        # KPI（ダッシュボードがある場合）
        try:
            dash = wb.Worksheets("Dashboard")
            report("=== ダッシュボードKPI ===")
            for label, addr in [("総資産", None)]:
                pass
        except Exception:
            pass

        report("=== エラーセル走査（数式セル）===")
        err_total = 0
        for s in wb.Worksheets:
            cnt = 0
            try:
                fcells = s.UsedRange.SpecialCells(XL_CELLTYPE_FORMULAS)
                for c in fcells:
                    t = str(c.Text)
                    if t.startswith("#"):
                        cnt += 1
            except Exception:
                pass
            if cnt:
                report(f"  {s.Name}: {cnt}")
                err_total += cnt
        report(f"  エラー合計 = {err_total}")
        rc = 0 if err_total == 0 else 2
        wb.Close(False)
        report(f"RESULT: {'OK' if rc == 0 else 'HAS_ERRORS'} errors={err_total}")
    except Exception as e:
        report(f"RESULT: ERROR - {e}")
        rc = 1
    finally:
        xl.Quit()
    io.open(REPORT, "w", encoding="utf-8").write("\n".join(_lines))
    print("VERIFY_DONE rc=%d (see verify_out.txt)"%rc)
    for ln in _lines:
        if ln.startswith("RESULT"):
            print(ln.encode("ascii","replace").decode("ascii"))
    return rc


if __name__ == "__main__":
    p = sys.argv[1] if len(sys.argv) > 1 else DEFAULT
    sys.exit(verify(p))
