# PhraseXross Bot Setup Guide

## 注意: トンネリング時の認証エラー（Azure Bot 側）

Azure Bot（Bot Service）からローカル環境へトンネリング接続すると、Azure Bot 側で次のエラーが表示されることがあります。

```
There was an error sending this message to your bot: HTTP status code Unauthorized
```

### 理由
- チャネル→ボット間はトークン検証が前提で、ローカル匿名（未認証）とは整合しません。
- Single-tenant/ユーザー割当 MSI のアプリタイプは、Dev Tunnels と Emulator では非サポートです。
  - 参考: [Azure Bot Service Quickstart Registration](https://learn.microsoft.com/en-us/azure/bot-service/bot-service-quickstart-registration?view=azure-bot-service-4.0&utm_source=chatgpt.com&tabs=userassigned)

### 対処
- ローカル以外（Web/Teams）から試す場合は、Azure App Service（Web Apps）等にデプロイしてください。
- 必要な環境変数（`MicrosoftAppId`、`MicrosoftAppPassword`、`MicrosoftAppTenantId`、`MicrosoftAppType`）を正しく構成してください。
- 詳細は本 README の「Azure へデプロイ」セクションを参照してください。

---

この手順に従って設定を行うことで、PhraseXross Bot を正常に動作させることができます。


## 必要な環境変数
Web Apps に以下の環境変数を設定してください:

- `MicrosoftAppId`: Azure Bot によって生成されたアプリケーション ID。
- `MicrosoftAppPassword`: Azure Bot で作成したシークレット。
- `MicrosoftAppType`: `SingleTenant` を指定。
- `MicrosoftAppTenantId`: Azure AD (Entra ID) のテナント ID。

## Azure Bot の設定
1. **メッセージング エンドポイント**:
   - Azure Bot の設定で、メッセージング エンドポイントを正しく設定してください。
   - 例: `https://<your-webapp-name>.azurewebsites.net/api/messages`

2. **ボットタイプ**:
   - `SingleTenant` を選択してください。

3. **シークレット**:
   - Azure Bot のアプリ登録でシークレットを作成し、`MicrosoftAppPassword` として使用します。

## Entra ID (Azure AD) の注意点
- Azure Bot を作成すると、自動的にアプリ登録が行われます。
- 現在の設定では、リダイレクト URI の設定は不要です。
- API 許可の追加も不要です。

## デプロイ後の確認
- Web Apps にデプロイ後、Bot Framework Emulator または Microsoft Teams を使用して応答を確認してください。

---

## 設定方法（appsettings.* ではなく .env / 環境変数へ移行）

このプロジェクトでは機密値を `appsettings.*.json` に保持せず、以下の優先順で読み込みます:

1. プロセス / コンテナ環境変数 (Azure Web Apps / docker -e)
2. 実行ディレクトリの `.env` ファイル（存在する場合のみ）

`.env.example` をコピーして `.env` を作成し、必要なキーを設定してください。`.env` は `.gitignore` 済みです。

主要キー一覧:

- Bot: `MicrosoftAppId`, `MicrosoftAppPassword`, `MicrosoftAppTenantId`, `MicrosoftAppType`
- Semantic Kernel (Azure OpenAI): `ENABLE_SK`, `AOAI_ENDPOINT`, `AOAI_API_KEY`, `AOAI_DEPLOYMENT`, `AOAI_API_VERSION`
- OneDrive / Graph: `OneDriveClientId`, `OneDriveTenantId` （後方互換: 旧 `ONEDRIVE_CLIENT_ID`, `ONEDRIVE_TENANT_ID` も可）

Docker 実行例（PowerShell）:

```powershell
docker run --rm -p 8080:8080 `
   -e ENABLE_SK=true `
   -e AOAI_ENDPOINT=https://<your-aoai>.openai.azure.com/ `
   -e AOAI_API_KEY=<key> `
   -e AOAI_DEPLOYMENT=gpt-4o `
   -e MicrosoftAppId=<bot-app-id> `
   -e MicrosoftAppPassword=<bot-secret> `
   phrasexross:latest
```

ローカル開発では `.env` に書くだけで `Program.cs` の簡易ローダが起動時に取り込みます。

## Semantic Kernel (Azure OpenAI) 連携

このボットはオプションで Semantic Kernel を使ったAI応答が可能です。無効時は従来のエコー応答になります。

### 有効化フラグ
- `ENABLE_SK`: `true` で有効化（既定は無効）。

### 必要な環境変数（AOAI）
- `AOAI_ENDPOINT`: 例 `https://<your-aoai>.openai.azure.com/`
- `AOAI_API_KEY`: Azure OpenAI のAPIキー
- `AOAI_DEPLOYMENT`: 使用するデプロイ名（gpt-4o/4.1/3.5などのデプロイ）

上記3つが揃っていない場合は、ENABLE_SK=true でも自動的にSK登録はスキップされ、エコー応答にフォールバックします（起動ログにWARNを出力）。

### ローカル（dotnet run）
PowerShell の例:

```
$env:ENABLE_SK = "true"
$env:AOAI_ENDPOINT = "https://<your-aoai>.openai.azure.com/"
$env:AOAI_API_KEY = "<your-key>"
$env:AOAI_DEPLOYMENT = "<your-deployment>"
dotnet run --project ./PhraseXross
```

### コンテナ（Docker）
`-e` で必要な環境変数を渡します。

```
docker run -it --rm -p 5006:8080 \
   -e ENABLE_SK=true \
   -e AOAI_ENDPOINT=https://<your-aoai>.openai.azure.com/ \
   -e AOAI_API_KEY=<your-key> \
   -e AOAI_DEPLOYMENT=<your-deployment> \
   phrasexross:latest
```

Bot Framework Emulator を使う場合、コンテナでは `serviceUrl` が localhost にならないように、Dev Tunnels などを使ってパブリックHTTPSのURLにする必要があります（本README先頭の注意を参照）。

### Azure Web Apps
アプリケーション設定に以下を追加してください:

- `ENABLE_SK` = `true`
- `AOAI_ENDPOINT` = `https://<your-aoai>.openai.azure.com/`
- `AOAI_API_KEY` = `<your-key>`
- `AOAI_DEPLOYMENT` = `<your-deployment>`

Azure上では `MicrosoftAppId/MicrosoftAppPassword/MicrosoftAppTenantId/MicrosoftAppType` も既に設定済みであることを前提とします。

### 動作確認
1. 起動ログで `[SK] Semantic Kernel を登録しました` が出力されていることを確認。
2. Emulator（またはTeams）でメッセージを送信。
3. AI応答が返らない場合、起動時のWARNログや `AOAI_*` 値の設定漏れ/誤りを確認してください。

