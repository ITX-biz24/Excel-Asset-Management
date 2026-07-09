# Excel Asset Management v2

![License](https://img.shields.io/github/license/ITX-biz24/Excel-Asset-Management)
![Last Commit](https://img.shields.io/github/last-commit/ITX-biz24/Excel-Asset-Management)
![Release](https://img.shields.io/github/v/release/ITX-biz24/Excel-Asset-Management)
![Issues](https://img.shields.io/github/issues/ITX-biz24/Excel-Asset-Management)

「一生使える個人資産管理システム」を目指した、Excel + VBA の資産管理アプリです。
単一の取引テーブルを唯一の真実とし、口座はマスタで自由に追加できます。ダッシュボード・
グラフ・固定収支/サブスク反映・月末スナップショット・翌月ロールオーバー・バックアップ・
数式保護などをワンクリックで扱えます。

A personal asset-management system built with Microsoft Excel and VBA.

---

## 📸 Screenshot

![Dashboard](images/main.png)

---

## 🇯🇵 日本語

### 主な機能

- **ダッシュボード**: 総資産・現金/銀行/QR残高・今月の収入/支出/収支・貯蓄率をカード表示。
  資産割合（円）・カテゴリ別支出（棒）・月別収支（縦棒）・資産推移（折線）の4グラフ。
- **取引履歴**: 1行=1取引のテーブル。収入/支出/振替を種類で管理し、入力規則と
  条件付き書式で入力ミスを防止。残高は口座に依存しない2本の `SUMIFS` で自動計算。
- **固定収支 / サブスク**: 毎月の定期収支・定期課金を管理。「今月へ反映」ボタンで
  取引履歴へ一括追加（メモの反映マークで二重反映を防止）。
- **レポート**: 月次サマリ・年間履歴・カテゴリ/支払方法ランキング。
- **スナップショット**: 月末資産を記録し資産推移グラフの元データにする。
- **マスタ / 設定**: 口座・カテゴリ・支払方法・対象年月などを一元管理。**口座は行を
  足すだけで追加**でき、数式は一切変更不要。
- **残高の強制プリセット**: マスタの「初期残高」に実残高を入力し「残高を強制セット」
  ボタンを押すと、取引履歴を消去して **現在残高＝初期残高** に強制リセット（過去履歴を
  参照しない）。デモ値から実運用へ切り替えるときに使う。

### VBA（マクロ）

翌月ロールオーバー / 固定・サブスク反映 / 月末スナップショット保存 / バックアップ /
入力チェック / 数式保護 / ワンクリック更新 / 取引入力フォーム / 残高の強制プリセット。

### 使い方

1. `Excel-Asset-Management.xlsm` を開く。
2. マクロを有効化する（初回は「コンテンツの有効化」）。
3. マスタ・設定で口座や初期残高・対象年月を調整する。
4. 取引履歴に日々の取引を入力（またはダッシュボードの「取引を追加」ボタン）。
5. ダッシュボードで残高・収支・グラフを確認する。

### 動作環境

- Windows + Microsoft Excel（VBA 有効）
- 表示言語: 日本語 / 通貨: 日本円（¥）

---

## 🇺🇸 English

Open `Excel-Asset-Management.xlsm`, enable macros, and start tracking. Accounts are managed in
the master sheet (add a row to add an account — no formula changes needed). To force real
balances, enter them in the master's "初期残高" column and click "残高を強制セット": the
transaction history is wiped and each current balance is reset to its preset (no history is
referenced). The dashboard shows total assets, monthly income/expense/net, savings rate, and
four charts. VBA provides month rollover, fixed/subscription posting, month-end snapshots,
backup, input checks, formula protection, one-click refresh, a transaction-entry form, and the
forced balance preset. Japanese UI, JPY currency.

---

## 🛠 開発者向け（ビルド）

本ブックは **ソースから再現生成** できます（保守性のためレイアウト・数式・VBA を
コードで一元管理）。

```text
build/
  build_structure.py   # openpyxl: シート・数式・入力規則・条件付き書式・グラフ・名前付き範囲
  inject_vba.py        # Excel COM: VBA注入・フォーム生成・ボタン配置・グラフのダーク化・.xlsm保存
  build.py             # 上記を順に実行するオーケストレータ
  verify.py            # 生成物を開いて再計算し残高・エラーを検証
  test_macros.py       # マクロを内部呼び出し経路で自動テスト
vba/                   # VBA ソース（*.bas / *.cls）。Git で差分管理
docs/SPEC.md           # 設計仕様書（問題点・設計・数式・VBA・リスク・改善案）
```

前提: Windows + Excel + Python(openpyxl, pywin32)。「VBAプロジェクトオブジェクトモデルへの
アクセスを信頼」(AccessVBOM=1) を有効化しておくこと。

```bash
python build/build.py     # → Excel-Asset-Management.xlsm を生成
python build/verify.py    # 構造検証（残高・エラー）
python build/test_macros.py  # マクロ動作テスト
```

詳細な設計は [docs/SPEC.md](docs/SPEC.md) を参照。

---

## 📄 License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
