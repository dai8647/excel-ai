# シート構成

## 設定（Sheet1 → 名前を「設定」に変更）

A列: ラベル / B列: 値

```
A1: エクセルAI  v1.0.0
A2: API Key
A3: Base URL
A4: Model
A5: ローカル LLM パス
A6: 作業場所
A7: 自動実行          → はい / いいえ
A8: 危険操作の確認    → はい / いいえ
A9: 最大ステップ数    → 20
```

名前付き範囲（推奨）:
- `cfgApiKey`     = 設定!$B$2
- `cfgBaseUrl`    = 設定!$B$3
- `cfgModel`      = 設定!$B$4
- `cfgLlmPath`    = 設定!$B$5
- `cfgWorkspace`  = 設定!$B$6
- `cfgAutoRun`    = 設定!$B$7
- `cfgConfirmDanger` = 設定!$B$8
- `cfgMaxSteps`   = 設定!$B$9

## 操作（Sheet2 → 「操作」）

ボタンとマクロの対応:
- 「接続テスト」     → `TestConnection`
- 「エージェントを開く」 → `ShowAgent`
- 「作業場所を開く」   → `OpenWorkspaceFolder`
- 「ログをクリア」     → `ClearChatLog`

## チャットログ（Sheet3 → 「チャットログ」）

| A: 日時 | B: 種別 | C: 内容 |
|---------|---------|---------|
| user / assistant / tool / system / error |
