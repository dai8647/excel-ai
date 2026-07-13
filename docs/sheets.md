# シート構成

## 設定（Sheet1 → 名前を「設定」に変更）

A列: ラベル / B列: 値

```
A1: ExcelAI  v1.1.0
A2: Connection type   → OpenAI / Anthropic
A3: API Key
A4: Base URL
A5: Model
A6: Workspace
A7: AutoRun           → はい / いいえ
A8: Confirm danger    → はい / いいえ
A9: Max steps         → 12 など
A10: Anthropic version → 2023-06-01
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

## 操作（Sheet2 → 「操作」）

ボタンとマクロの対応:
- 「モデル一覧を更新」 → `RefreshModels`
- 「接続テスト」     → `TestConnection`
- 「エージェントを開く」 → `ShowAgent`
- 「Requestパネルを開く」 → `ShowAgent`
- 「Requestを実行」 → `RunSheetPrompt`
- 「Workspaceを選ぶ」 → `PickWorkspaceFolder`
- 「Workspaceフォルダを開く」 → `OpenWorkspaceFolder`
- 「ログをクリア」     → `ClearChatLog`
- 「Workspace安全確認」 → `WorkspaceSafetyCheck`

## モデル一覧（「モデル一覧」）

接続先から取得したモデルIDを表示します。A列がモデルID、B列が取得時の接続方式です。`設定` シートのConnection typeを切り替えた後に「モデル一覧を更新」を押してください。

## チャットログ（Sheet3 → 「チャットログ」）

| A: 日時 | B: 種別 | C: 内容 |
|---------|---------|---------|
| user / assistant / tool / system / error |
