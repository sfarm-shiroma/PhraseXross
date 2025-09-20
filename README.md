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

