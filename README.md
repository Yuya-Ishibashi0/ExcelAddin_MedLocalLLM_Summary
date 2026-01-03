# Excel Local LLM

Excelの選択範囲をローカルAI（Ollama）で処理し、サイドバーに表示・必要ならセルへ書き込みます。

## 使い方（非エンジニア向け）

### 初回だけ: セットアップ

1. `ExcelLocalLLM-Setup.exe` がある場合は右クリック → 「管理者として実行」
2. 無い場合は `installer\install-all.cmd` を右クリック → 「管理者として実行」
3. Python / Ollama の導入と、Excelアドインの登録を行います

※ インターネット接続があれば Python / Ollama / モデルを自動で取得します。初回は時間がかかります。

### 毎回: 起動

1. 通常はExcel起動時に自動で起動します（タスク登録済み）
2. 手動で起動する場合は PowerShell で `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\start.ps1`

### Excelで使う

1. Excel → 挿入 → アドイン → 共有フォルダー から「Excel Local LLM」を開く
2. セル/範囲を選択 → 「コンテキスト更新」
3. 指示を入力して送信 → 結果を確認 → 必要なら書き込み

## アンインストール

1. `installer\uninstall.cmd` を右クリック → 「管理者として実行」
2. アプリ本体のフォルダ（このフォルダ）を削除

※ Python / Ollama は残ります。不要なら Windows の「アプリと機能」から削除してください。

## フォルダ構成（ざっくり）

- `add-in/`：Excelの画面とマニフェスト
- `app/`：ローカルAPI（`server.py`）と設定（`settings.json`）
- `installer/`：オフライン配布/インストール用
- `scripts/`：起動・自動起動の補助
- `docs/`：詳細設定や開発向けの説明

## 詳細設定・開発者向け

`docs/advanced.md` を参照してください。
