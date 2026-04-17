# Bottle Mail — 引き継ぎ書

作成日: 2026-04-16  
ブランチ: `claude/nice-dhawan`

---

## アプリ概要

**Bottle Mail** — プロジェクトURLをAIが分析し、個性の近い人へ自動でボトルメールを届けるマッチングアプリ。  
本番URL: https://bottle-mail.vercel.app

---

## ディレクトリ構成

```
bottle-mail/
├── public/index.html       # シングルページフロントエンド（全UI・JS込み）
├── src/app.js              # Express APIサーバー（Vercel Serverless）
├── scripts/seed-dummies.js # ダミーユーザー8人を本番DBに登録するスクリプト
├── docs/supabase-schema.sql# DBスキーマ（Supabaseで実行済み）
├── vercel.json             # Vercelデプロイ設定
├── package.json            # type: "module", ES Modules
└── .env                    # ローカル用（Vercelには環境変数として設定済み）

bottle-mail-mcp/
├── index.js                # Claude Desktop用MCPサーバー
└── package.json
```

---

## 技術スタック

| レイヤー | 使用技術 |
|---------|---------|
| フロントエンド | Vanilla HTML/CSS/JS（SPA）、ocean-theme |
| バックエンド | Node.js + Express（ES Modules） |
| デプロイ | Vercel Serverless Functions |
| DB | Supabase（PostgreSQL） |
| AI（分析） | Claude Sonnet 4.6 — URL→個性・タグ・名前候補生成 |
| AI（マッチング） | Claude Haiku 4.5 — ユーザー間のマッチスコア計算 |
| MCP | @modelcontextprotocol/sdk — Claude Desktop連携 |

---

## 環境変数（Vercel + ローカル `.env` に設定済み）

```
SUPABASE_URL=https://adsqrasuxrqnvwqacmit.supabase.co
SUPABASE_SERVICE_ROLE_KEY=eyJhbGci...（Supabaseダッシュボードで確認）
ANTHROPIC_API_KEY=sk-ant-api03-...（Anthropic Consoleで確認）
```

---

## DBスキーマ（Supabase）

### `bottle_users`
| カラム | 型 | 説明 |
|--------|-----|------|
| id | uuid PK | ユーザーID |
| email | text unique | ログイン用メール |
| character_name | text | AIが提案したキャラクター名 |
| personality_summary | text | AI生成の個性説明文 |
| personality_tags | text[] | 個性タグ配列 |
| project_urls | text[] | 登録したプロジェクトURL |
| social_links | jsonb | SNSリンク |
| created_at | timestamptz | |

### `bottles`
| カラム | 型 | 説明 |
|--------|-----|------|
| id | uuid PK | |
| sender_id | uuid FK | 送信者 |
| recipient_id | uuid FK | 受信者 |
| match_score | float | AIマッチスコア（0–100） |
| match_reason | text | マッチ理由テキスト |
| status | text | unread / opened / replied |
| created_at | timestamptz | |

### `bottle_messages`
| カラム | 型 | 説明 |
|--------|-----|------|
| id | uuid PK | |
| bottle_id | uuid FK | |
| sender_id | uuid FK | |
| content | text | メッセージ本文 |
| created_at | timestamptz | |

---

## APIエンドポイント一覧

| Method | Path | 説明 |
|--------|------|------|
| POST | `/api/analyze` | URLを分析して個性・名前候補を返す |
| POST | `/api/users` | 新規ユーザー登録（登録後にマッチング非同期実行） |
| POST | `/api/login` | メールアドレスでログイン |
| GET | `/api/users/me` | 自分のプロフィール取得 |
| GET | `/api/bottles` | 受信ボトル一覧 |
| GET | `/api/bottles/sent` | 送信ボトル一覧 |
| POST | `/api/bottles/:id/open` | ボトルを開封（status → opened） |
| GET | `/api/messages/:bottleId` | チャット履歴取得 |
| POST | `/api/messages` | メッセージ送信 |
| GET | `/` | index.html を返す |

**認証**: `x-user-id` リクエストヘッダーにユーザーIDを付与（JWTなし）

---

## フロントエンド画面フロー

```
hero → register（URL入力）→ analyzing（ローディング）→ name（名前選択）→ complete
                                                                         ↓
login ─────────────────────────────────────────────────────────→ dashboard
                                                                  （受信/送信タブ）
                                                                         ↓
                                                                  bottle（詳細）→ chat
```

---

## 実装済み機能

- [x] URLからAIが個性・タグ・キャラクター名を生成
- [x] メールアドレスでのログイン（UUID不要）
- [x] ダッシュボード：受信/送信済みタブ
- [x] ボトル開封・チャット
- [x] 👋 手を振るボタン
- [x] IME対応Enterキー送信（日本語入力中は送信しない）
- [x] 二重送信防止（isSendingフラグ）
- [x] 多言語対応（日本語/英語）— ナビに `JA/EN` 切替ボタン
- [x] ダミーユーザー8人（本番DBに登録済み）
- [x] Claude Desktop用MCPサーバー（`bottle-mail-mcp/index.js`）

---

## MCPサーバー（Claude Desktop連携）

### セットアップ

`~/Library/Application Support/Claude/claude_desktop_config.json` に以下を追加済み：

```json
"mcpServers": {
  "bottle-mail": {
    "command": "node",
    "args": ["/Users/akiko/Desktop/claudecode/.claude/worktrees/nice-dhawan/bottle-mail-mcp/index.js"],
    "env": {
      "BOTTLE_MAIL_EMAIL": "あなたのメールアドレス@example.com"  ← 要変更
    }
  }
}
```

Claude Desktopを再起動すると有効になる。

### 利用可能ツール

| ツール | 機能 |
|--------|------|
| `bottle_inbox` | 受信ボトル一覧 |
| `bottle_sent` | 送信ボトル一覧 |
| `bottle_open` | ボトルを開封 |
| `bottle_reply` | 返信・メッセージ送信 |
| `bottle_chat` | チャット履歴確認 |
| `bottle_profile` | 自分のプロフィール確認 |
| `bottle_analyze` | URLを分析（登録前） |
| `bottle_register` | 新規ユーザー登録 |

---

## 未完了 / 要対応タスク

### 1. **デプロイ（最重要）**
多言語対応のコードをVercelにpushしていない。以下の手順でデプロイ：

```bash
# node_modules を git から除外してから push
# .gitignore に node_modules/ が追加済みなので:
git rm -r --cached bottle-mail-mcp/node_modules/
git add .gitignore bottle-mail/public/index.html bottle-mail-mcp/index.js bottle-mail-mcp/package.json
git commit -m "fix: exclude node_modules from git"
git push origin claude/nice-dhawan
```

その後 Vercel ダッシュボード or `npx vercel --prod` でデプロイ。  
または GitHub → Vercel の自動デプロイが設定されていれば push だけでOK。

### 2. **認証の強化**
現在 `x-user-id` ヘッダーのみで認証。本番では JWT or Supabase Auth への移行を検討。

### 3. **メール通知**
「マッチした人のボトルが届いたら通知」とUIに書かれているが未実装。Resend / SendGrid などを使って実際にメール送信する。

### 4. **MCPサーバーの `z` (zod) 依存**
`bottle-mail-mcp/index.js` 内で `import { z } from "zod"` しているが、`package.json` の dependencies に zod が含まれていない可能性がある。`npm install zod` を実行して確認。

---

## ローカル開発

```bash
cd bottle-mail
cp .env.example .env   # または .env を手動作成（上記環境変数を設定）
npm install
npm run dev            # http://localhost:3000 で起動
```

---

## 既知のはまりポイント

| 問題 | 解決策 |
|------|--------|
| `ANTHROPIC_API_KEY not found` on Vercel | `dotenv.config({ path: new URL("../.env", import.meta.url).pathname })` で解決済み |
| Vercel登録タイムアウト | マッチング処理を `res.json()` の後に非同期で実行するよう修正済み |
| `vercel.json` の `functions` + `builds` 競合 | `builds` のみに統一済み |
| 改行入りAPIキー | `printf 'key' \| npx vercel env add` で設定すること（heredocは改行が入る） |
| GitHub push認証 | PAT をリモートURLに含めて push: `git remote set-url origin https://TOKEN@github.com/...` |
