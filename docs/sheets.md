# シート構成

## 設定（Sheet1 → 名前を「設定」に変更）

A列: ラベル / B列: 値

```
A1: ExcelAI  v1.2.0
A2: Connection type   → OpenAI / Anthropic
A3: API Key
A4: Base URL
A5: Model
A6: Workspace
A7: AutoRun           → はい / いいえ
A8: Confirm danger    → はい / いいえ
A9: Max steps         → 12 など
A10: Anthropic version → 2023-06-01
A11: Max tokens       → 4096 など（1応答あたりの出力上限）
```

名前付き範囲（推奨）:
- `cfgProvider`   = 設定!$B$2
- `cfgApiKey`     = 設定!$B$3
- `cfgBaseUrl`    = 設定!$B$4
- `cfgModel`      = 設定!$B$5
- `cfgWorkspace`  = 設定!$B$6
- `cfgAutoRun`    = 設定!$B$7
- `cfgConfirmDanger` = 設定!$B$8
- `cfgMaxSteps`   = 設定!$B$9
- `cfgAnthropicVersion` = 設定!$B$10

## ホーム（Sheet3 → 「ホーム」）

ボタンとマクロの対応:
- 「実行 Ctrl+Enter」 → `RunSheetPrompt`
- 「モデル更新」 → `RefreshModels`
- 「接続テスト」 → `TestConnection`
- 「設定を開く」 → `OpenSettings`
- 「Workspace」 → `OpenWorkspaceFolder`
- 「フォルダ選択」 → `PickWorkspaceFolder`
- 「ログを開く」 → `OpenChatLog`
- 「安全チェック」 → `WorkspaceSafetyCheck`
- 「停止」 → `Agent_RequestStop`
- 「新しい会話」 → `Session_Reset`

## モデル一覧（「モデル一覧」）

接続先から取得したモデルIDを表示します。A列がモデルID、B列が取得時の接続方式です。`設定` シートのConnection typeを切り替えた後に「モデル一覧を更新」を押してください。

## チャットログ（Sheet4 → 「チャットログ」）

| A: 日時 | B: 種別 | C: 内容 |
|---------|---------|---------|
| user / assistant / tool / tool_result / progress / system / error |

## 会話セッション（シートではない）

会話履歴はブックと同じ場所の隠しファイル `.excelai-session.json` に保存される。
再読み込み・ブック再起動後も自動で継続。「新しい会話」(`Session_Reset`) で削除。
ツール呼び出し途中で終わっている末尾メッセージは復元時に破棄される。
内容は依頼・ツール結果・ファイル内容をそのまま含むため、配布・共有前に削除すること。
