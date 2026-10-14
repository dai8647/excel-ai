# エクセルAI（開発コードネーム）

会社で Excel しか使えない環境向けの **100% VBA AI エージェント**。

OpenClaude の「ツールループ + プロバイダ切替 + サンドボックス」の考え方を Excel/VBA で再実装しています。

## 利用条件・商用利用

このリポジトリは、個人・教育・非営利の利用を無償とするSource-Available型のライセンスです。会社・法人・団体の社内利用、顧客向け業務、有償サービスへの組み込み、商用配布には、別途商用ライセンス契約が必要です。料金は利用人数、配布範囲、サポート、Enterprise機能に応じて個別見積もりとします。

- [LICENSE](LICENSE) — 個人・非商用利用の条件
- [COMMERCIAL_LICENSE.md](COMMERCIAL_LICENSE.md) — 商用契約のドラフト
- [NOTICE.md](NOTICE.md) — Microsoft等との非提携表明
- [TRADEMARKS.md](TRADEMARKS.md) — 商標・ブランド方針

このプロジェクトはMicrosoft Corporationの製品、公式アドイン、提携サービスではありません。API利用料、Microsoft 365/Excelの契約料、導入作業費は別途必要です。

## 機能

- ホームシート型のチャット画面（依頼→実行→進行表示→結果）
- 会話の連続性 — 履歴を自動保存し、「さっきの続き」「前回のファイルを直して」が通じる。ブック再起動後も継続。「新しい会話」でリセット
- 実行中の停止ボタン（現在のツール完了後に安全に停止、履歴は保持）
- クラウド API / ローカル LLM（LM Studio 等、Ollama 不要）
- 作業場所サンドボックス内のファイル操作（作成/コピー/移動/削除/フォルダ作成/再帰リスト）
- 計算式付きExcelファイルとグラフの生成、既存ブックへのシート追加・値/数式書き込み（`update_excel`）
- 成果物を既定アプリで開いて見せる `open_file`
- 既存xlsx/xlsm/xlsb/xlsの構造・シート・数式の読み取り
- 公開Webページ・APIレスポンスの取得（読み取り専用 `web_fetch`）
- 自動実行 ON/OFF（設定で切替）
- **スキル** — 「スキル」シートに1行書くだけで手順定義（prompt型）や自作VBA関数（vba型）をツールとして追加できる（再ビルド不要）
- **プラグイン** — `vba/plugins/Plugin_*.bas` を置くだけでツール一式を追加・配布できる（[docs/SKILLS_PLUGINS.md](docs/SKILLS_PLUGINS.md)）
- ホーム画面、Ctrl+Enter実行、処理進行状況の表示
- 操作ログ（監査用）

## フォルダ構成

```
エクセルAI/
├── ExcelAI.xlsm          ← 配布物（build-xlsm.ps1 で生成）
├── .excelai-session.json ← 会話履歴（自動生成・配布前に削除）
├── vba/                  ← VBA ソース（インポート用）
│   └── plugins/          ← プラグイン（Plugin_*.bas を置くだけ）
├── scripts/              ← build-xlsm.ps1（xlsm 生成）・ selftest.ps1
├── workspace/            ← デフォルト作業場所（AI の作業範囲）
└── docs/                 ← セットアップ・シート構成
```

## クイックスタート

1. [docs/SETUP.md](docs/SETUP.md) を読む
2. `powershell -File scripts/build-xlsm.ps1` で `ExcelAI.xlsm` を生成（または `vba/` を手動インポート）
3. `設定` シートで API / 作業場所を設定
4. `ホーム` → Request欄に依頼を書き、実行ボタンまたは `Ctrl+Enter`

完成後は `Application.Run "SelfTest_Run"`（PowerShell からでも可）で自己テストが実行できます。

## デフォルト作業場所

```
workspace
```

設定シート B6 に同じパスを入れてください。

Per-PC setup: enter each user's local folder path in `設定!B6`. An absolute path is supported; `workspace` means a folder beside the workbook.
