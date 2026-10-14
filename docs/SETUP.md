# エクセルAI — セットアップガイド（チーム配布用）

## 概要

100% VBA の AI 作業エージェント。Excel だけでファイル操作・CSV 結合・レポート生成などを、  
指定した作業フォルダ（サンドボックス）内だけで実行します。

## 必要環境

- Windows 10/11
- Microsoft Excel（マクロ有効 `.xlsm`）
- インターネット（クラウド API 利用時）
- **Ollama は不要**（ローカル LLM は LM Studio 等を推奨）

## 初回セットアップ（管理者 / 配布者）

1. `ExcelAI.xlsm` を任意の場所に配置（共有フォルダ可）
2. Excel で開く → 「コンテンツの有効化」をクリック
3. `設定` シートでプロバイダー、APIキー、Base URL、モデル、Workspaceを入力（下表参照）
4. `ホーム` シートの「モデル更新」で接続先のモデルを取得
5. `ホーム` シートの「接続テスト」ボタンで API 疎通確認
6. `ホーム` シートのRequest欄に依頼を書き、実行ボタンまたは Ctrl+Enter で起動

## 設定シート（B列）

| 行 | 項目 | 例 |
|----|------|-----|
| 2 | Connection type | `OpenAI` または `Anthropic` |
| 3 | API Key | 選択した接続方式のキー（ローカル LLM なら空欄可） |
| 4 | Base URL | `https://api.openai.com/v1` または `https://api.anthropic.com` |
| 5 | Model | 「モデル一覧を更新」で取得したモデルID |
| 6 | Workspace | このPCの作業フォルダ。`workspace` はブックと同じ場所 |
| 7 | AutoRun | `はい` または `いいえ`（初期値は `いいえ`） |
| 8 | Confirm danger | `はい` または `いいえ`（初期値は `はい`） |
| 9 | Max steps | `12` など。1回の依頼で許可する最大ステップ数 |
| 10 | Anthropic version | Anthropic互換APIのバージョン（通常 `2023-06-01`） |
| 11 | Max tokens | 1応答あたりの出力上限。既定 `4096`。長いファイル書き込みが切れる場合のみ増加 |

WorkspaceはPCごとに異なるため、配布先では各ユーザーが `設定!B6` に自分のフォルダを入力してください。絶対パスも使えます。APIキーはこのブックの外へ送信・公開しないでください。

### プロバイダ例

**クラウド（OpenAI 互換）**
```
Base URL: https://api.openai.com/v1
Model:    gpt-4o-mini
API Key:  （取得したキー）
```

**Anthropic 互換**
```
Base URL: https://api.anthropic.com
Model:    （モデル一覧から選択）
API Key:  （取得したキー）
Anthropic version: 2023-06-01
```

**9router（ローカル互換ゲートウェイの例）**
```
Base URL: http://127.0.0.1:20128/v1
Model:    （モデル一覧から選択）
API Key:  9routerで発行したキー
```

**OpenCode Zen（OpenAI互換）**
```
Connection type: OpenAI
Base URL: https://opencode.ai/zen/v1
Model:    deepseek-v4-flash-free など（モデル一覧から選択）
API Key:  OpenCode Zenで発行したキー
```

OpenCode Zenはモデルごとに利用条件と料金が異なります。無料モデルを含め、機密情報を送る前にOpenCode Zenの最新の利用条件を確認してください。

**ローカル LLM（LM Studio — Ollama 不要）**
```
1. LM Studio をインストール（https://lmstudio.ai）
2. モデルをダウンロードして Load
3. Local Server を起動（デフォルト http://localhost:1234）
4. Excel 設定:
   Base URL: http://localhost:1234/v1
   Model:    （LM Studio に表示されているモデル ID）
   API Key:  （空欄で OK）
```

**その他ローカル（GPT4All / llama.cpp server 等）**
- OpenAI 互換 `/v1/chat/completions` があれば Base URL を差し替えるだけ

## VBA モジュールのインポート（ビルド手順）

`.xlsm` を自分で作る場合:

1. 新規ブックを作成 → `.xlsm` で保存
2. `Alt+F11` で VBA エディタを開く
3. `ファイル` → `ファイルのインポート` で `vba/` 内の `.bas` をすべてインポート
4. `frmAgent.frm` をインポート
5. `ThisWorkbook` に `vba/ThisWorkbook.cls` のコードを貼り付け
6. シート `設定` `ホーム` `モデル一覧` `チャットログ` `README` を作成（`docs/sheets.md` 参照）
7. `ホーム` シートにRequest欄とボタンを配置してマクロを割り当て

## チーム配布時の注意

- **マクロ署名**: 社内 CA で `.xlsm` に署名すると「信頼できる発行元」として配布しやすい
- **信頼できる場所**: ネットワーク共有ではなく、各 PC のローカルまたは社内信頼フォルダへ
- **API Key**: 共有シートに書かず、各自の設定シート B3 に入力させる
- **ショートカット**: `Ctrl+Enter` で依頼実行、`Ctrl+Shift+E` でホームへ戻る
- **作業場所**: 各自の設定シート B6 に入力。チーム共通なら `\\server\share\ExcelAI-workspace\ユーザー名` 形式を推奨
- **バージョン管理**: `設定!A1` にバージョン番号を記載（更新時の確認用）
- **商用利用**: 会社・法人・団体で使う場合は、配布前に商用ライセンス契約を締結する
- **Microsoftとの関係**: ExcelAIはMicrosoftの公式製品・公式アドインではない

## セキュリティ

- 作業場所外のパスは **すべて拒否**（PathGuard）
- `delete_file` は「危険操作の確認=はい」時のみ実行前にダイアログ
- `web_fetch` は公開 `http/https` のGET専用。読み取り専用で毎回の確認は行わず、ローカル・プライベートIP、認証情報付きURL、リダイレクトを拒否し、応答を500KB・出力を50000文字までに制限
- `inspect_excel` はExcelバイナリをテキストとして読まず、読み取り専用でシート構造を調べる
- Web取得結果は外部の未信頼データとして扱い、取得内容の指示・コード・リンクを自動実行しない
- URLは外部へ送信されるため、APIキー・個人情報・ブック内容をURLやクエリに含めない
- 全ツール実行は `チャットログ` シートに記録
- 自動実行 OFF 時は、書き込み・削除ツールの依頼ごとに一度だけ内容を表示して確認。読み取りとWeb取得は確認なし

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| マクロが動かない | ファイルを信頼できる場所に移動、マクロを有効化 |
| API 接続エラー | Connection type と Base URL の組み合わせ、キー、対象モデル、ローカルサーバーの起動状態を確認 |
| ローカル LLM が遅い | 小さいモデルに変更、maxSteps を下げる |
| パス拒否エラー | 作業場所配下のパスだけ指定（`..` 不可） |
