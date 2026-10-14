# スキルとプラグイン

ExcelAI はエージェントの機能を後から増やせます。方法は2つです。

| 方式 | 誰が | 再ビルド | 用途 |
|---|---|---|---|
| スキル(スキルシート) | ユーザー | 不要 | 手順の定義・自作VBA関数のツール化 |
| プラグイン(.bas) | 開発者 | 要(または実行時取込+登録) | ツール一式の外付け配布 |

## スキル(シート編集で追加)

「スキル」シートの行として登録します。列:

| 列 | 内容 |
|---|---|
| A スキル名 | 半角英数とハイフン等が推奨。VBAツールの名前 `vba_<名前>` になります(空白と記号は `_` に変換) |
| B 種別 | `prompt` または `vba` |
| C 説明 | モデルがいつ使うか判断するための一文 |
| D 本文/VBA関数 | prompt型: 手順テキスト。vba型: 呼び出す関数名 |
| E パラメータJSON | vba型のみ。省略時は `{"input":{"type":"string"}}` |
| F 書込確認 | `yes` にすると実行前に確認ダイアログ(ファイルを書き換えるスキルは yes に) |
| G 有効 | `yes` で有効化 |

### prompt型スキル

本文がシステムプロンプトに注入され、モデルが依頼に応じて手順に従います。
例: 「月次レポート」=「数値表から①前月比②最多③注意点を順に算出して表にする」

### vba型スキル

このブック内の標準モジュールの Public 関数をツールとして公開します。関数の形:

```vba
Public Function my_tool(ByVal argsJson As String) As String
    ' argsJson はツールの引数JSON。JsonGetArg(argsJson, "path") で取得できる
    my_tool = "結果テキスト"
End Function
```

- 引数: 1つのString(JSON)。戻り値: String(モデルに返る結果)。
- 実行は `Application.Run` 経由。関数が見つからない・エラーの場合はエラー文言が結果になります。
- ファイルを書くなら F列=yes にしてください。

## プラグイン(.bas を置くだけ)

`vba\plugins\` に `Plugin_<名前>.bas` を置いてビルドすると、自動的にimport+登録されます。
(実行時に追加する場合は Alt+F11 で .bas を取り込み → マクロ `Plugins_Register` で登録)

### 必須メンバー

```vba
Attribute VB_Name = "Plugin_MyPlugin"
Option Explicit

Public Function Plugin_Name() As String          ' 短いID(例: "my")
    Plugin_Name = "my"
End Function

Public Function Plugin_ToolNames() As String     ' 持つツール名(小文字, カンマ区切り)
    Plugin_ToolNames = "my_tool"
End Function

Public Function Plugin_ToolsOpenAI() As String   ' 定義(ToolDefヘルパー使用, []は付けない)
    Plugin_ToolsOpenAI = ToolDef("my_tool", "説明", "{""param"":{""type"":""string""}}")
End Function

Public Function Plugin_ToolsAnthropic() As String
    Plugin_ToolsAnthropic = AnthropicToolDef("my_tool", "説明", "{""param"":{""type"":""string""}}")
End Function

Public Function Plugin_Execute(ByVal toolName As String, ByVal argsJson As String) As String
    ' argsJson から JsonGetArg で引数を取り、結果文字列を返す
    Plugin_Execute = "実行しました: " & JsonGetArg(argsJson, "param")
End Function
```

### 任意メンバー

- `Plugin_WriteTools() As String` … 書き込み確認ゲートを通すツール名(カンマ区切り)。ファイルを作る・消すツールは必ず登録
- `Plugin_OnBeforeWrite(ByVal targetPath As String)` … どの書き込みツールが承認された直後(書き込み前)に呼ばれるフック。スナップショット等に使用

### 約束事

- ツール名は**小文字**・英数と `_`・内蔵ツール/スキルと重複禁止(プラグイン名を接頭辞にすると安全: `my_tool`)
- モジュールは `Plugin_` で始める(自動登録の目印)
- シート「プラグイン」で有効/無効を切り替えられる(隠しシート。右クリック→「表示」で出せる)
- 壊れたプラグインは自己テストが失敗を捕まえる。Plugin_Echo.bas が動く最小実装の見本
