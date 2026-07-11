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
3. `設定` シートで各項目を入力（下表参照）
4. `操作` シートの「接続テスト」ボタンで API 疎通確認
5. `操作` シートの「エージェントを開く」でチャット UI を起動

## 設定シート（B列）

| 行 | 項目 | 例 |
|----|------|-----|
| 2 | API Key | `sk-...`（ローカル LLM なら空欄可） |
| 3 | Base URL | `https://api.openai.com/v1` |
| 4 | Model | `gpt-4o-mini` |
| 5 | ローカル LLM パス | `C:\Users\xxx\AppData\Local\LM Studio` |
| 6 | 作業場所 | `C:\Users\xxx\Downloads\エクセルAI\workspace` |
| 7 | 自動実行 | `はい` または `いいえ` |
| 8 | 危険操作の確認 | `はい` または `いいえ` |
| 9 | 最大ステップ数 | `20` |

### プロバイダ例

**クラウド（OpenAI 互換）**
```
Base URL: https://api.openai.com/v1
Model:    gpt-4o-mini
API Key:  （取得したキー）
```

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
6. シート `設定` `操作` `チャットログ` を作成（`docs/sheets.md` 参照）
7. `操作` シートにボタンを配置してマクロを割り当て

## チーム配布時の注意

- **マクロ署名**: 社内 CA で `.xlsm` に署名すると「信頼できる発行元」として配布しやすい
- **信頼できる場所**: ネットワーク共有ではなく、各 PC のローカルまたは社内信頼フォルダへ
- **API Key**: 共有シートに書かず、各自の設定シート B2 に入力させる
- **作業場所**: チーム共通なら `\\server\share\ExcelAI-workspace\ユーザー名` 形式を推奨
- **バージョン管理**: `設定!A1` にバージョン番号を記載（更新時の確認用）

## セキュリティ

- 作業場所外のパスは **すべて拒否**（PathGuard）
- `delete_file` は「危険操作の確認=はい」時のみ実行前にダイアログ
- 全ツール実行は `チャットログ` シートに記録
- 自動実行 OFF 時は、ツール実行前に内容を表示して確認

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| マクロが動かない | ファイルを信頼できる場所に移動、マクロを有効化 |
| API 接続エラー | Base URL 末尾に `/v1` があるか、Local Server が起動中か確認 |
| ローカル LLM が遅い | 小さいモデルに変更、maxSteps を下げる |
| パス拒否エラー | 作業場所配下のパスだけ指定（`..` 不可） |
