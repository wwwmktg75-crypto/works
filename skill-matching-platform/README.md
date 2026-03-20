# skill-matching-platform

Supabase中心で最短立ち上げする、スキルマッチングMVP用の最小API実装です。

## できること
- メンバー登録 (`POST /api/members`)
- 公開メンバー検索 (`GET /api/members`)
- 問い合わせ送信 + メール同報 (`POST /api/inquiries`)
- Stripe Checkout発行 (`POST /api/payments/checkout`)
- Stripe Webhook受信と注文記録 (`POST /api/payments/webhook`)

## セットアップ
1. 依存関係をインストール
   - `npm install`
2. 環境変数を作成
   - `.env.example` を `.env` にコピーして値を設定
3. 起動
   - `npm run dev`

## Supabase準備
- `docs/supabase-schema.sql` をSupabase SQL Editorで実行
- `.env` に `SUPABASE_URL` / `SUPABASE_SERVICE_ROLE_KEY` を設定

## Stripe準備
- Secret Key と Webhook Secret を `.env` に設定
- Webhook URL 例: `https://your-api-domain/api/payments/webhook`
- 受信イベント: `checkout.session.completed`

## Vercelデプロイ
- `vercel.json` を同梱済み（Express APIをServerlessとして実行）
- Vercelの環境変数に `.env.example` の全項目を設定
- デプロイ後のAPI例: `https://<project>.vercel.app/api/members`

## 機密情報の取り扱い
- `SUPABASE_DB_PASSWORD` は `.env` と Vercel Environment Variables にのみ設定する
- `README.md` やソースコードに秘密情報を平文で書かない
- `.env` はGitにコミットしない（`.gitignore` 対象のまま運用する）
- パスワードは1Password / Bitwardenなどのパスワードマネージャで管理する

## 想定フロント
- Softr / Studio 側から本APIを叩く構成
- 初期段階はSupabaseフォーム相当 + Make連携でも運用可能
