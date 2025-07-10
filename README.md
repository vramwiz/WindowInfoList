# WindowInfoList - ウィンドウ一覧取得＆アクティブ化ツール

Windows上に表示されているすべてのウィンドウ（トップレベル）を一覧表示する Delphi アプリです。  
リストには各ウィンドウのタイトルとアイコンが表示され、クリックすることでそのウィンドウを前面に表示します。

---

## 🧩 主な機能

- 現在表示中のトップレベルウィンドウ一覧を取得
- 各ウィンドウのタイトルとアイコンを `TListView` に表示
- リスト上でアイテムをクリックすると該当ウィンドウをアクティブ化

---

## 📷 スクリーンショット

（※ここに画像を貼ってください）

---

## 🚀 使い方

1. アプリを起動すると、自動的に現在のウィンドウ一覧が取得されます。
2. リストに表示されたウィンドウタイトルをクリックすると、そのウィンドウが前面に表示されます。

---

## 🛠 開発環境・依存

- Delphi 10.x 以降（確認済み）
- Win32 API：`EnumWindows`, `GetWindowText`, `SendMessage`, `SetForegroundWindow`, など
- 使用ユニット：`WindowInfoList.pas`, `MainForm.pas`

---

## 📦 同梱ファイル

| ファイル名 | 説明 |
|------------|------|
| `WindowInfoList.pas` | ウィンドウ一覧を取得するクラスユニット |
| `MainForm.pas/.dfm` | UIを構成するフォーム。リストビューと連動 |

---

## 📘 ライセンス

MIT License（予定）  
ご自由にご利用・改変・再配布いただけます。

---

## 🌐 English Summary

**WindowInfoList** is a small Delphi application that displays all top-level windows currently shown on the system.

- Lists each window's icon and title in a `TListView`
- Clicking an item brings the corresponding window to the front (activates it)
- Built with Delphi and pure Win32 API

---
