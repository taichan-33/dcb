# LINE Bot を使った現金出納簿アプリ

![5ca26bb2bcfa7f8afafc70cf0aace378d487a83865ffd7c805fe83 82880308](https://github.com/taichan-33/dcb/assets/151983276/60c3d52b-6e06-45fd-8549-574383cd4e14)

このコードは、LINE Bot と Google スプレッドシートを連携させて出納簿をつけるためのアプリケーションです。

ユーザーは LINE アプリを通じて項目、収入、支出、備考を入力し、月ごとの収支計算を行うことができます。

データは Google スプレッドシートに自動的に保存され、年ごとに新しいシートが作成されます。

# 機能

・項目、収入、支出、備考の入力：ユーザーは LINE アプリを通じて家計簿の各項目を入力できます。

・月ごとの収支計算：ユーザーは特定の月の収支計算を要求することができ、その月の収入と支出の合計が表示されます。

・年ごとのスプレッドシートの自動作成：新しい年のデータが入力されると、自動的に新しいシートが作成されます。

・ユーザーごとのメッセージ履歴の保存：各ユーザーのメッセージ履歴が保存され、前回のメッセージに基づいて適切な処理が行われます。

# 設定

１．LINE Developers にアクセスし、ログインまたは新規登録を行います。

新しいプロバイダーを作成し、プロバイダー名を入力します。


２．プロバイダー内で新しい Messaging API チャンネルを作成します。

チャンネル名とチャンネル説明を入力し、大業種と小業種を選択します。


３．チャンネル基本設定で、チャンネルアクセストークン（長期）を発行し、控えておきます。

このトークンは後で GAS コード内で使用します。


４．Messaging API 設定で、Webhook 送信を有効にし、Webhook URL を設定します。この URL は後ほど GAS プロジェクトのデプロイ後に取得します。

５．Google スプレッドシートを新規作成し、「原本」というシートを作成します。

このシートをテンプレートとして使用し、新しい年のシートを作成する際にコピーします。

６．スプレッドシート内で「ツール」>「スクリプトエディタ」を選択し、GAS プロジェクトを開きます。

７．コードエディタ内に添付のコードをコピーし、LINE_ACCESS_TOKEN 変数に控えておいたチャンネルアクセストークンを設定します。

８．GAS プロジェクトをデプロイします。「公開」>「デプロイ済みのウェブアプリケーション」を選択し、「新規作成」をクリックします。

説明を入力し、アクセスを「全員（匿名ユーザーを含む）」に設定し、「デプロイ」をクリックします。


９．デプロイ後、ウェブアプリケーションのURL（Webhook URL）が表示されるので、これをコピーします。


１０．コピーした Webhook URL を LINE Developers の Messaging API 設定の Webhook URL に設定し、「更新」をクリックします。

以上でセットアップは完了です。Bot が正常に動作するようになります。

# 使い方

１．LINE アプリで作成した Bot と友達になります。「検索」から Bot のQRコードまたはIDを入力して友達追加します。

２．Bot に以下のキーワードを送信することで、対応する操作が行われます。

・項目：項目を入力するためのプロンプトが表示されます。ユーザーは項目名を入力し、送信します。

・収入：収入金額を入力するためのプロンプトが表示されます。ユーザーは金額を入力し、送信します。

・支出：支出金額を入力するためのプロンプトが表示されます。ユーザーは金額を入力し、送信します。

・備考：備考を入力するためのプロンプトが表示されます。ユーザーは備考を入力し、送信します。

・月収支計算：指定した月の収支計算結果が表示されます。ユーザーは月を数字で入力し、送信します。

３．入力されたデータは自動的にスプレッドシートに保存されます。

新しい年のデータが入力されると、新しいシートが作成されます。

# コード説明

・```doPost 関数```：LINE Bot からのリクエストを処理するメイン関数です。

以下の処理を行います。

１.リクエストから必要なデータ（ユーザーメッセージ、リプライトークン、ユーザーID）を取得します。

２.現在のスプレッドシートを取得し、現在の年に対応するシートを取得または作成します。

３.ユーザーのメッセージ履歴を取得し、新しいメッセージを追加します。

４.ユーザーのメッセージに基づいて適切な処理を行います（項目、収入、支出、備考の入力、月収支計算）。

５.処理結果をユーザーに返信します。

・```createQuickReply``` 関数：クイックリプライのメッセージオブジェクトを作成します。ユーザーがタップできるボタンを表示するために使用します。

・```getLastEntryRow``` 関数：スプレッドシートの最後の入力行を取得します。新しいデータを追加する際に使用します。

・```createNewShee``` 関数：新しい年のシートを作成します。テンプレートシートをコピーして新しいシートを作成します。

・```getMonthlyBalance``` 関数：指定した月の収支計算を行います。その月の収入と支出のデータを取得し、合計を計算します。

このアプリケーションを使用することで、ユーザーは LINE アプリを通じて簡単に家計簿をつけることができます。

データは自動的にスプレッドシートに保存され、月ごとの収支計算も行えるため、財務管理に役立ちます。また、年ごとに新しいシートが作成されるため、長期的な財務データの管理も容易です。

コードはモジュール化されており、各関数は特定の機能を担当しています。

これにより、コードの可読性と保守性が向上しています。

また、ユーザーごとのメッセージ履歴を保存することで、ユーザーの入力状況に応じた適切な処理を行うことができます。

# ライセンス

このプロジェクトは、MIT ライセンスの下で公開されています。

以下の条件に従う限り、自由に使用、複製、変更、配布することができます。

ライセンス表示：このソフトウェアを使用する際は、上記の著作権表示とこの許諾表示を、ソフトウェアのすべての複製または重要な部分に記載する必要があります。

無保証：作者または著作権者は、ソフトウェアに関して何らの保証も行いません。

ソフトウェアの使用に起因するいかなる損害についても、作者または著作権者は責任を負いません。

詳細については、LICENSE ファイルを参照してください。
