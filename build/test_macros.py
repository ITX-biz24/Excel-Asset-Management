# -*- coding: utf-8 -*-
"""
マクロ動作テスト（win32com）
============================
生成した .xlsm を開き、MsgBox を含まない関数を Application.Run で実行して
VBAのコンパイル・ロジック・書込経路を検証する。結果は UTF-8 ファイルに出力。

    python test_macros.py
"""
import os
import io
import sys
import datetime
import win32com.client as win32

BUILD_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(BUILD_DIR)
XLSM = os.path.join(REPO_DIR, "Excel-Asset-Management.xlsm")
REPORT = os.path.join(BUILD_DIR, "test_out.txt")

_lines = []
def log(m=""): _lines.append(str(m))


def txn_count(wb):
    """T_Txn の日付が入っている行数を数える。"""
    for s in wb.Worksheets:
        for lo in s.ListObjects:
            if lo.Name == "T_Txn":
                n = 0
                body = lo.DataBodyRange
                for i in range(1, lo.ListRows.Count + 1):
                    if str(body.Cells(i, 1).Value or "").strip():
                        n += 1
                return n
    return -1


def main():
    if not os.path.exists(XLSM):
        print("ERROR: .xlsm がありません。先に build.py を実行してください。")
        return 1
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        xl.AutomationSecurity = 1  # msoAutomationSecurityLow（マクロ有効）
    except Exception:
        pass
    rc = 0
    passed = failed = 0

    def check(name, got, want):
        nonlocal passed, failed
        ok = (got == want)
        log(f"  [{'OK' if ok else 'NG'}] {name}: got={got!r} want={want!r}")
        if ok: passed += 1
        else: failed += 1

    try:
        wb = xl.Workbooks.Open(XLSM)
        xl.CalculateFullRebuild()

        log("=== 関数テスト（コンパイル＋ロジック）===")
        # コンパイルはRun時に走る。ここが通れば全モジュール構文OK。
        ym = xl.Run("TargetYM")
        check("TargetYM", str(ym), "2026-07")

        total = xl.Run("TotalAssets")
        check("TotalAssets", round(float(total)), 148300)

        # 残高不足の検出（現金24400 < 1億）
        msg_over = xl.Run("ValidateTxn", "2026-07-10", "支出", "現金", "",
                          100000000, "食費", True)
        log(f"  [{'OK' if msg_over else 'NG'}] ValidateTxn(残高不足): '{msg_over}'")
        passed += 1 if msg_over else 0
        failed += 0 if msg_over else 1

        # 正常な支出は "" が返る
        msg_ok = xl.Run("ValidateTxn", "2026-07-10", "支出", "現金", "",
                        1000, "食費", True)
        check("ValidateTxn(正常)", str(msg_ok), "")

        # 同一口座振替は弾く
        msg_same = xl.Run("ValidateTxn", "2026-07-10", "振替", "現金", "現金",
                          1000, "", True)
        log(f"  [{'OK' if msg_same else 'NG'}] ValidateTxn(同一口座): '{msg_same}'")
        passed += 1 if msg_same else 0
        failed += 0 if msg_same else 1

        # --- 書込/反映テスト（製品と同じ内部呼び出し経路）---
        # 外部 Application.Run で多引数Subを直接呼ぶとマーシャリングで
        # 落ちるため、内部から呼ぶ薄いラッパを注入して検証する。
        log("=== 書込/反映テスト（内部呼び出し経路）===")
        helper = (
            'Function T_AddOne() As String\n'
            '  gSilent = True\n'
            '  AddTransaction CDate("2026-07-11"), "支出", "現金", "", 500, "食費", "現金", "テスト"\n'
            '  gSilent = False\n'
            '  T_AddOne = "ok"\n'
            'End Function\n'
            'Function T_ReflectFixed() As String\n'
            '  gSilent = True\n'
            '  ReflectFixedItems\n'
            '  gSilent = False\n'
            '  T_ReflectFixed = "ok"\n'
            'End Function\n'
            'Function T_Snapshot() As String\n'
            '  gSilent = True\n'
            '  SaveMonthEndSnapshot "2026-07"\n'
            '  gSilent = False\n'
            '  T_Snapshot = "ok"\n'
            'End Function\n'
        ).replace("\n", "\r\n")
        comp = wb.VBProject.VBComponents.Add(1)
        comp.Name = "modTestHelper"
        comp.CodeModule.AddFromString(helper)

        before = txn_count(wb)
        cash_before = float(xl.Run("AccountBalance", "現金"))
        xl.Run("T_AddOne")
        xl.CalculateFullRebuild()
        after = txn_count(wb)
        cash_after = float(xl.Run("AccountBalance", "現金"))
        check("AddTransaction 行数+1", after, before + 1)
        check("AddTransaction 現金残高-500", round(cash_before - cash_after), 500)

        # 固定収支反映（有効ON=給与/家賃/電気代/携帯代 の4件）
        before2 = txn_count(wb)
        xl.Run("T_ReflectFixed")
        xl.CalculateFullRebuild()
        after2 = txn_count(wb)
        check("ReflectFixedItems +4件", after2 - before2, 4)

        # 二重反映防止（もう一度実行しても増えない）
        xl.Run("T_ReflectFixed")
        xl.CalculateFullRebuild()
        check("固定 二重反映防止", txn_count(wb) - after2, 0)

        # スナップショット保存
        r = xl.Run("T_Snapshot")
        check("SaveMonthEndSnapshot 実行", str(r), "ok")

        wb.Close(False)   # 変更を破棄（原本は無傷）

        log("")
        log(f"RESULT: passed={passed} failed={failed}")
        rc = 0 if failed == 0 else 3
    except Exception as e:
        log(f"RESULT: ERROR - {e}")
        rc = 1
    finally:
        xl.Quit()

    io.open(REPORT, "w", encoding="utf-8").write("\n".join(_lines))
    print(f"TEST_DONE rc={rc} (see test_out.txt)")
    for ln in _lines:
        if ln.startswith("RESULT"):
            print(ln.encode("ascii", "replace").decode("ascii"))
    return rc


if __name__ == "__main__":
    sys.exit(main())
