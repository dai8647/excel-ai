SECTION|まずこの順番で使う
STEP|1|設定シートを確認|設定シートでプロバイダー、APIキー、ベースURL、モデル、ワークスペースを確認します。配布版では各PCで入力してください。
STEP|2|モデルを更新|操作シートの「1. Refresh models」を押すと、接続先のモデル一覧を取得します。
STEP|3|接続を確認|操作シートの「2. Test connection」を押し、接続OKが表示されることを確認します。
STEP|4|AIに依頼|操作シートのRequest欄に依頼を書き、「4. Run request」を押します。

SECTION|設定シート
ITEM|設定!B2|Connection type|OpenAI または Anthropic を選びます。9routerはOpenAI互換とAnthropic互換の両方に対応します。
ITEM|設定!B3|API Key|選択したプロバイダーのAPIキーを入力します。キーはこのブックの外へ送信・公開しないでください。
ITEM|設定!B4|Base URL|接続先のベースURLを入力します。OpenAI互換は通常 /v1 を含め、Anthropic互換は互換仕様に合わせます。9routerは http://127.0.0.1:20128/v1 です。
ITEM|設定!B5|Model|モデル一覧から使うモデルIDを選びます。最初は軽量モデルで接続確認するのがおすすめです。
ITEM|設定!B6|Workspace|AIが読み書きできる作業フォルダーです。安全のため、このフォルダー内だけを対象にします。
ITEM|設定!B7|AutoRun|noを推奨します。書き込み・移動・削除などの操作は確認を挟みます。
ITEM|設定!B8|Confirm danger|yesを推奨します。危険な操作の前に確認ダイアログを表示します。
ITEM|設定!B9|Max steps|1回の依頼でAIが進める最大ステップ数です。通常は12で十分です。
ITEM|設定!B10|Anthropic version|Anthropic互換APIを使う場合のバージョンです。通常は2023-06-01のままにします。
NOTE|配布時の注意|生成された配布ブックにはAPIキーを含めません。各ユーザーが自分のキーを設定!B3に入力し、Workspaceを設定!B6で選択してください。
NOTE|OpenCode Zen|OpenAI互換として接続できます。Base URLは https://opencode.ai/zen/v1、モデルはRefresh modelsで取得してください。

SECTION|操作シートのボタン
ITEM|Refresh models|モデル一覧を接続先から再取得します。
ITEM|Test connection|現在のプロバイダー・モデルで疎通を確認します。
ITEM|Open request panel|入力フォームが使える環境ではAI入力パネルを開きます。
ITEM|Run request|Request欄の内容をAIへ送り、必要なら安全確認を行いながら処理します。
ITEM|Open workspace|現在のWorkspaceフォルダーをエクスプローラーで開きます。
ITEM|Choose workspace|このPCで使うWorkspaceフォルダーを選択し、設定!B6に保存します。
ITEM|Clear chat log|チャットログを消去します。必要な記録は先に保存してください。
ITEM|Workspace safety|Workspaceの実体パスと、親フォルダー・外部パスの拒否状態を確認します。

SECTION|AIが使えるExcel生成ツール
ITEM|create_excel|CSV形式のデータから、数式・見出し・罫線・列幅を含むExcelファイルを作成します。
ITEM|create_chart|既存のExcelファイルに、縦棒・横棒・折れ線・円グラフを追加します。

SECTION|AIが使えるWeb調査ツール
ITEM|web_fetch|公開されたhttp/https URLをGETし、HTML・JSON・テキストを最大50000文字まで参照します。実行前に確認を表示します。
ITEM|web_fetchの制限|ローカル・プライベートIP、認証情報付きURL、リダイレクトは拒否します。取得内容は外部の未信頼データとして扱います。

SECTION|安全ルール
ITEM|Workspace限定|ファイル操作はWorkspace内だけに制限されます。親フォルダー参照（..）は拒否します。
ITEM|確認|AutoRun=no、Confirm danger=yesの組み合わせを推奨します。
ITEM|APIキー|キーをGitHubやチャットログへ貼り付けないでください。公開リポジトリには含めません。
ITEM|Web取得|URLは外部へ送信されます。APIキー、個人情報、ブック内容をURLやクエリに含めないでください。

SECTION|利用条件
ITEM|個人利用|個人・教育・非営利の利用はLICENSEの範囲で無償です。
ITEM|商用利用|会社・法人・団体の社内業務、顧客向け業務、有償サービスへの組み込み、商用配布は商用ライセンス契約が必要です。
ITEM|非提携|ExcelAIはMicrosoftの公式製品・公式アドインではありません。

SECTION|困ったとき
ITEM|接続エラー|9routerが起動しているか、Base URLが http://127.0.0.1:20128/v1 か、APIキーが有効か確認します。
ITEM|モデルがない|設定シートの「Refresh models」を実行し、モデル一覧からモデルIDを選び直します。
ITEM|処理が止まる|Max stepsを確認し、チャットログの最後のerror行を確認します。
NOTE|このREADMEはExcelAI.xlsmの生成時に自動で作成されます。設定を変更した場合は、次回ビルド時に内容が更新されます。
