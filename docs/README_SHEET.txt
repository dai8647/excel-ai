SECTION|まずこの順番で使う
STEP|1|設定シートを確認|設定シートでプロバイダー、APIキー、ベースURL、モデル、ワークスペースを確認します。
STEP|2|モデルを更新|操作シートの「1. Refresh models」を押すと、接続先のモデル一覧を取得します。
STEP|3|接続を確認|操作シートの「2. Test connection」を押し、接続OKが表示されることを確認します。
STEP|4|AIに依頼|操作シートのRequest欄に依頼を書き、「4. Run request」を押します。

SECTION|設定シート
ITEM|設定!B2|Connection type|OpenAI または Anthropic を選びます。9routerはOpenAI互換とAnthropic互換の両方に対応します。
ITEM|設定!B3|API Key|9routerの有効なキーを入力します。キーはこのブックの外へ送信・公開しないでください。
ITEM|設定!B4|Base URL|9routerは http://127.0.0.1:20128/v1。OpenAI互換は /v1 を含め、Anthropic互換は自動調整します。
ITEM|設定!B5|Model|モデル一覧から使うモデルIDを選びます。最初は軽量モデルで接続確認するのがおすすめです。
ITEM|設定!B6|Workspace|AIが読み書きできる作業フォルダーです。安全のため、このフォルダー内だけを対象にします。
ITEM|設定!B7|AutoRun|noを推奨します。書き込み・移動・削除などの操作は確認を挟みます。
ITEM|設定!B8|Confirm danger|yesを推奨します。危険な操作の前に確認ダイアログを表示します。
ITEM|設定!B9|Max steps|1回の依頼でAIが進める最大ステップ数です。通常は12で十分です。

SECTION|操作シートのボタン
ITEM|Refresh models|モデル一覧を接続先から再取得します。
ITEM|Test connection|現在のプロバイダー・モデルで疎通を確認します。
ITEM|Open request panel|入力フォームが使える環境ではAI入力パネルを開きます。
ITEM|Run request|Request欄の内容をAIへ送り、必要なら安全確認を行いながら処理します。
ITEM|Open workspace|現在のWorkspaceフォルダーをエクスプローラーで開きます。
ITEM|Clear chat log|チャットログを消去します。必要な記録は先に保存してください。

SECTION|安全ルール
ITEM|Workspace限定|ファイル操作はWorkspace内だけに制限されます。親フォルダー参照（..）は拒否します。
ITEM|確認|AutoRun=no、Confirm danger=yesの組み合わせを推奨します。
ITEM|APIキー|キーをGitHubやチャットログへ貼り付けないでください。公開リポジトリには含めません。

SECTION|困ったとき
ITEM|接続エラー|9routerが起動しているか、Base URLが http://127.0.0.1:20128/v1 か、APIキーが有効か確認します。
ITEM|モデルがない|設定シートの「Refresh models」を実行し、モデル一覧からモデルIDを選び直します。
ITEM|処理が止まる|Max stepsを確認し、チャットログの最後のerror行を確認します。
NOTE|このREADMEはExcelAI.xlsmの生成時に自動で作成されます。設定を変更した場合は、次回ビルド時に内容が更新されます。
