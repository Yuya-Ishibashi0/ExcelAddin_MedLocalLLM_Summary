# 詳細設定・開発者向け

## 手動起動（開発用）

```bash
cd app
python -m pip install -r requirements.txt
python -m uvicorn server:app --host 127.0.0.1 --port 8787
```

```bash
cd add-in
python -m http.server 3000
```

## 環境変数（任意）

- `OLLAMA_MODEL`：使用モデル（例 `gemma3:4b`）
- `OLLAMA_URL`：OllamaのURL（既定 `http://127.0.0.1:11434`）
- `OLLAMA_TIMEOUT`：タイムアウト秒（既定 `120`）
- `OLLAMA_NUM_PREDICT`：生成トークン上限
- `OLLAMA_NUM_CTX`：入力コンテキスト長
- `OLLAMA_NUM_THREAD`：推論スレッド数
- `OLLAMA_NUM_BATCH`：バッチサイズ
- `OLLAMA_PRESET`：既定プリセット（`fast`/`balanced`/`long`）
- `OLLAMA_TEMPERATURE`：温度
- `OLLAMA_TOP_K`：top-k
- `OLLAMA_TOP_P`：top-p
- `OLLAMA_REPEAT_PENALTY`：反復ペナルティ
- `OLLAMA_REPEAT_LAST_N`：反復判定の直近トークン数
- `OLLAMA_PRESENCE_PENALTY`：出現ペナルティ
- `OLLAMA_FREQUENCY_PENALTY`：頻度ペナルティ
- `OLLAMA_TFS_Z`：TFS
- `OLLAMA_TYPICAL_P`：typical-p
- `OLLAMA_MIROSTAT`：mirostat（0/1/2）
- `OLLAMA_MIROSTAT_TAU`：mirostat_tau
- `OLLAMA_MIROSTAT_ETA`：mirostat_eta
- `OLLAMA_PENALIZE_NEWLINE`：改行ペナルティ（true/false）
- `OLLAMA_SEED`：乱数シード
- `OLLAMA_STOP`：停止トークン（例: `###,END` または JSON配列）

## `settings.json`（環境変数が面倒な場合）

`app/settings.json` を編集すると、環境変数なしで設定できます。
環境変数が設定されている場合はそちらが優先されます。

主なキー:

- `ollama_url` / `ollama_model` / `system_prompt` / `inference_prompt`
- `inference_num_predict` / `inference_num_ctx`
- `num_predict` / `num_ctx` / `num_thread` / `num_batch`
- `temperature` / `top_k` / `top_p`
- `repeat_penalty` / `repeat_last_n`
- `presence_penalty` / `frequency_penalty`
- `tfs_z` / `typical_p`
- `mirostat` / `mirostat_tau` / `mirostat_eta`
- `penalize_newline` / `seed` / `stop`

※ 変更後は `uvicorn` を再起動してください。

### プリセット例

- `fast`：短文・速度重視（`num_ctx=1024` / `num_predict=256`）
- `balanced`：通常（`num_ctx=2048` / `num_predict=512`）
- `long`：長文・文脈重視（`num_ctx=4096` / `num_predict=1024`）

## よくある調整ポイント

1. 出力が途中で切れる → `OLLAMA_NUM_PREDICT` を大きくする
2. 入力が長すぎる → `OLLAMA_NUM_CTX` を調整する
3. 履歴要約のON/OFF → `add-in/taskpane.js` の先頭定数を編集
4. プリセット内容変更 → `app/server.py` の `get_preset_options` を編集

## 自動起動（ログオン時）

`scripts/start.ps1` をタスクスケジューラの「ログオン時」に登録すると、自動起動できます。
Excel起動に合わせたい場合は `scripts/excel-watch.ps1` を使います。

## 配布（オンライン/オフライン）

### GitHubから試す（オンライン）

1. プロジェクト一式をダウンロード（zipなど）
2. `installer/install-all.cmd` を右クリック → 管理者として実行

※ Python / Ollama / モデルを自動で取得します。時間がかかる場合があります。
※ 使うモデルを変えたい場合は、先に `app/settings.json` の `ollama_model` を編集してください。

### オフライン配布

1. プロジェクト一式を配布（zipなど）
2. `installer/assets` に以下を置く
   - `python-<version>-amd64.exe`（x64）
   - `python-<version>-arm64.exe`（ARM64）
   - `OllamaSetup.exe`
3. `installer/install-all.cmd` を右クリック → 管理者として実行

※ オフライン環境ではモデルの自動取得はできません。
※ 別途、オンライン環境で `ollama pull <model>` が必要です。

### IExpress で EXE を作る方法（GUI）

installer\build-iexpress.cmd -RunIExpress

1. `installer/build-iexpress.cmd` を実行（オンライン版。`installer/assets` は含めません）
2. Win+R → `iexpress`（管理者として実行）
3. `installer/build/iexpress` の `payload.zip` と `run.cmd` を追加
4. 出力先は `installer/dist/ExcelLocalLLM-Setup.exe`

※ オフライン版にしたい場合は `installer/build-iexpress.ps1 -IncludeAssets` で作成します（サイズが大きくなります）。

### アンインストール

`installer\\uninstall.cmd` を管理者で実行すると、
アドイン登録と自動起動設定を削除します。

※ Python / Ollama は削除しません。不要なら Windows の「アプリと機能」から削除してください。
