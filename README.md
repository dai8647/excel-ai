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

- チャット UI（UserForm）
- クラウド API / ローカル LLM（LM Studio 等、Ollama 不要）
- 作業場所サンドボックス内のファイル操作
- 計算式付きExcelファイルとグラフの生成
- 自動実行 ON/OFF（設定で切替）
- 操作ログ（監査用）

## フォルダ構成

```
エクセルAI/
├── ExcelAI.xlsm          ← 配布物（要ビルド or 手動作成）
├── vba/                  ← VBA ソース（インポート用）
├── workspace/            ← デフォルト作業場所（AI の作業範囲）
├── templates/            ← xlsm テンプレート等
└── docs/                 ← セットアップ・シート構成
```

## クイックスタート

1. [docs/SETUP.md](docs/SETUP.md) を読む
2. `vba/` のモジュールを `.xlsm` にインポート
3. `設定` シートで API / 作業場所を設定
4. `操作` → 「エージェントを開く」

## デフォルト作業場所

```
workspace
```

設定シート B6 に同じパスを入れてください。

Per-PC setup: enter each user's local folder path in `設定!B6`. An absolute path is supported; `workspace` means a folder beside the workbook.
