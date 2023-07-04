# SPO_Auth
## 環境
* windows11
* Microsoft365
* SeleniumBasic
* chrome

## 概要
SharePoint Online にログインするための自動化スクリプトです。
- 指定されたメールアドレスに送信されたワンタイムコードを自動的に取得します。
- 取得したワンタイムコードを使用して、SharePoint Online にログインします。


## 導入手順
* SPO_Auth.basをExcel VBAにインポート
* ThisOutlookSession.cls の中身をコピーしてThisOutlookSessionにペースト

  ![image](https://github.com/kuma6082/SPO_Auth/assets/89393398/368f0543-46e0-4b5f-a346-99e23788bb83)

## 使用方法

1. ダウンロードフォルダのパス、SharePoint Online の URL、メールアドレス、保存フォルダのパスを正しく設定してください。

2. Excel VBAでSPO_Authを実行します。

3. Outlook に送信されたワンタイムコードを含むメールを受信すると、スクリプトが自動的にワンタイムコードを取得し、SharePoint Online にログインします。

4. ログインが成功すると、`Sample` 関数内の希望の処理が実行されます。この部分を必要に応じて変更してください。

## 注意
- このスクリプトが正しく機能するには、特定の構成とソフトウェア (Microsoft Excel、Outlook、Google Chrome) に依存していることに注意してください。
- スクリプトが確実に動作するために、安定したインターネット接続があることを確認してください。
- このスクリプトは、対話するソフトウェア (Excel、Outlook、Chrome など) の将来のバージョンで動作するように更新または変更が必要になる場合があります。
- このスクリプトは、SharePoint Online の利用規約に従って責任を持って使用してください。
