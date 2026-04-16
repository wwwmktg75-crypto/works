#!/usr/bin/env node
/**
 * Bottle Mail MCP Server
 * Claude Desktopからボトルメールを操作できるMCPサーバー
 *
 * 環境変数:
 *   BOTTLE_MAIL_EMAIL    - ログイン用メールアドレス
 *   BOTTLE_MAIL_API_BASE - APIベースURL（省略時: https://bottle-mail.vercel.app）
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

const API_BASE = process.env.BOTTLE_MAIL_API_BASE || "https://bottle-mail.vercel.app";
const USER_EMAIL = process.env.BOTTLE_MAIL_EMAIL || "";

let cachedUserId = null;
let cachedUserName = null;

// ---------- ユーティリティ ----------

async function apiGet(path, userId) {
  const headers = { "Content-Type": "application/json" };
  if (userId) headers["x-user-id"] = userId;
  const res = await fetch(`${API_BASE}${path}`, { headers });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
  return data;
}

async function apiPost(path, body, userId) {
  const headers = { "Content-Type": "application/json" };
  if (userId) headers["x-user-id"] = userId;
  const res = await fetch(`${API_BASE}${path}`, {
    method: "POST",
    headers,
    body: JSON.stringify(body)
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
  return data;
}

async function getMyUserId(emailOverride) {
  if (cachedUserId) return { id: cachedUserId, name: cachedUserName };
  const email = emailOverride || USER_EMAIL;
  if (!email) throw new Error("メールアドレスが設定されていません。BOTTLE_MAIL_EMAIL 環境変数を設定してください。");
  const data = await apiPost("/api/login", { email });
  cachedUserId = data.user.id;
  cachedUserName = data.user.character_name;
  return { id: cachedUserId, name: cachedUserName };
}

function formatBottle(bottle, perspective) {
  const other = perspective === "inbox" ? bottle.sender : bottle.recipient;
  const name = other?.character_name || "不明";
  const tags = (other?.personality_tags || []).join(", ");
  const score = bottle.match_score ? `マッチ度: ${bottle.match_score}点` : "";
  const reason = bottle.match_reason ? `理由: ${bottle.match_reason}` : "";
  const status = bottle.status === "unread" ? "📬 未開封" : bottle.status === "opened" ? "💌 開封済み" : "💬 返信あり";
  return [
    `ID: ${bottle.id}`,
    `相手: ${name}`,
    tags ? `タグ: ${tags}` : null,
    score,
    reason,
    status,
    `受信日: ${new Date(bottle.created_at).toLocaleString("ja-JP")}`
  ].filter(Boolean).join("\n");
}

// ---------- MCPサーバー設定 ----------

const server = new McpServer({
  name: "bottle-mail",
  version: "0.1.0"
});

// 受信ボトル一覧
server.tool(
  "bottle_inbox",
  "受信したボトルメールの一覧を確認します。マッチした相手から届いたボトルを表示します。",
  {
    email: z.string().optional().describe("メールアドレス（省略時は環境変数 BOTTLE_MAIL_EMAIL を使用）")
  },
  async ({ email }) => {
    const { id, name } = await getMyUserId(email);
    const data = await apiGet("/api/bottles", id);
    const bottles = data.bottles || [];
    if (!bottles.length) {
      return { content: [{ type: "text", text: `${name} さんへの受信ボトルはまだありません。` }] };
    }
    const lines = [`📬 ${name} さんの受信ボトル（${bottles.length}件）\n`];
    for (const b of bottles) {
      lines.push("---");
      lines.push(formatBottle(b, "inbox"));
    }
    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// 送信ボトル一覧
server.tool(
  "bottle_sent",
  "自分が送ったボトルメールの一覧を確認します。どの相手にマッチして届いたか確認できます。",
  {
    email: z.string().optional().describe("メールアドレス（省略時は環境変数 BOTTLE_MAIL_EMAIL を使用）")
  },
  async ({ email }) => {
    const { id, name } = await getMyUserId(email);
    const data = await apiGet("/api/bottles/sent", id);
    const bottles = data.bottles || [];
    if (!bottles.length) {
      return { content: [{ type: "text", text: `${name} さんの送信ボトルはまだありません。` }] };
    }
    const lines = [`📤 ${name} さんの送信ボトル（${bottles.length}件）\n`];
    for (const b of bottles) {
      lines.push("---");
      lines.push(formatBottle(b, "sent"));
    }
    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// ボトルを開封
server.tool(
  "bottle_open",
  "ボトルを開封して送信者の詳細プロフィールとマッチ理由を確認します。",
  {
    bottle_id: z.string().describe("開封するボトルのID（bottle_inbox で確認できます）"),
    email: z.string().optional().describe("メールアドレス（省略時は環境変数 BOTTLE_MAIL_EMAIL を使用）")
  },
  async ({ bottle_id, email }) => {
    const { id } = await getMyUserId(email);
    const data = await apiPost(`/api/bottles/${bottle_id}/open`, {}, id);
    const bottle = data.bottle;
    const sender = bottle.sender;
    const lines = [
      `💌 ボトルを開封しました！`,
      ``,
      `送信者: ${sender.character_name}`,
      `個性: ${sender.personality_summary || "（情報なし）"}`,
      `タグ: ${(sender.personality_tags || []).join(", ")}`,
      ``,
      `マッチ度: ${bottle.match_score}点`,
      `マッチ理由: ${bottle.match_reason}`,
      ``,
      `このボトルに返信するには bottle_reply を使ってください。`,
      `bottle_id: ${bottle_id}`
    ];
    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// ボトルに返信・メッセージ送信
server.tool(
  "bottle_reply",
  "ボトルに返信またはチャットを続けます。開封済みのボトルに対してメッセージを送れます。",
  {
    bottle_id: z.string().describe("返信するボトルのID"),
    message: z.string().describe("送信するメッセージ内容"),
    email: z.string().optional().describe("メールアドレス（省略時は環境変数 BOTTLE_MAIL_EMAIL を使用）")
  },
  async ({ bottle_id, message, email }) => {
    const { id, name } = await getMyUserId(email);
    await apiPost("/api/messages", { bottle_id, content: message }, id);
    return { content: [{ type: "text", text: `✉️ ${name} さんがメッセージを送信しました。\n\n「${message}」` }] };
  }
);

// チャット履歴取得
server.tool(
  "bottle_chat",
  "ボトルのチャット履歴を取得します。",
  {
    bottle_id: z.string().describe("チャット履歴を見るボトルのID"),
    email: z.string().optional().describe("メールアドレス（省略時は環境変数 BOTTLE_MAIL_EMAIL を使用）")
  },
  async ({ bottle_id, email }) => {
    const { id } = await getMyUserId(email);
    const data = await apiGet(`/api/messages/${bottle_id}`, id);
    const messages = data.messages || [];
    if (!messages.length) {
      return { content: [{ type: "text", text: "まだメッセージはありません。" }] };
    }
    const lines = [`💬 チャット履歴（${messages.length}件）\n`];
    for (const m of messages) {
      const sender = m.sender?.character_name || "不明";
      const time = new Date(m.created_at).toLocaleString("ja-JP");
      lines.push(`[${time}] ${sender}:`);
      lines.push(m.content);
      lines.push("");
    }
    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// 自分のプロフィール確認
server.tool(
  "bottle_profile",
  "自分のBottle Mailプロフィール（キャラクター名・個性・タグ）を確認します。",
  {
    email: z.string().optional().describe("メールアドレス（省略時は環境変数 BOTTLE_MAIL_EMAIL を使用）")
  },
  async ({ email }) => {
    const { id } = await getMyUserId(email);
    const data = await apiGet("/api/users/me", id);
    const user = data.user;
    const lines = [
      `🌊 あなたのBottle Mailプロフィール`,
      ``,
      `キャラクター名: ${user.character_name}`,
      `個性: ${user.personality_summary || "（未設定）"}`,
      `タグ: ${(user.personality_tags || []).join(", ")}`,
      `登録日: ${new Date(user.created_at).toLocaleString("ja-JP")}`,
      `ユーザーID: ${user.id}`
    ];
    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// URLを分析してキャラクター候補を生成
server.tool(
  "bottle_analyze",
  "プロジェクトURLを分析してあなたの個性・タグ・キャラクター名候補を生成します。新規登録前の分析に使います。",
  {
    project_urls: z.array(z.string()).min(1).max(5).describe("分析するプロジェクト/アプリのURL（最大5件）"),
    twitter: z.string().optional().describe("TwitterのURL（任意）"),
    github: z.string().optional().describe("GitHubのURL（任意）"),
    instagram: z.string().optional().describe("InstagramのURL（任意）")
  },
  async ({ project_urls, twitter, github, instagram }) => {
    const social_links = {};
    if (twitter) social_links.twitter = twitter;
    if (github) social_links.github = github;
    if (instagram) social_links.instagram = instagram;

    const data = await apiPost("/api/analyze", { project_urls, social_links });

    const lines = [
      `🔍 分析結果`,
      ``,
      `個性: ${data.personality_summary}`,
      `タグ: ${(data.personality_tags || []).join(", ")}`,
      ``,
      `キャラクター名候補:`
    ];
    for (const s of (data.name_suggestions || [])) {
      lines.push(`  ・${s.name}  — ${s.reason}`);
    }
    lines.push(``, `登録するには bottle_register ツールを使ってください。`);
    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// 新規ユーザー登録
server.tool(
  "bottle_register",
  "Bottle Mailに新規ユーザーとして登録します。bottle_analyze で取得した情報を使って登録できます。",
  {
    email: z.string().email().describe("登録するメールアドレス"),
    character_name: z.string().describe("選んだキャラクター名"),
    personality_summary: z.string().optional().describe("個性の説明（bottle_analyze の結果を貼り付け）"),
    personality_tags: z.array(z.string()).optional().describe("個性タグ（bottle_analyze の結果を貼り付け）"),
    project_urls: z.array(z.string()).optional().describe("プロジェクトURL一覧")
  },
  async ({ email, character_name, personality_summary, personality_tags, project_urls }) => {
    const data = await apiPost("/api/users", {
      email,
      character_name,
      personality_summary,
      personality_tags,
      project_urls: project_urls || [],
      social_links: {}
    });
    const user = data.user;
    const lines = [
      `✅ 登録完了！`,
      ``,
      `キャラクター名: ${user.character_name}`,
      `ユーザーID: ${user.id}`,
      ``,
      `AIがバックグラウンドで他のユーザーとのマッチングを行っています。`,
      `しばらくしてから bottle_inbox で届いたボトルを確認してみてください！`
    ];
    // 登録後はキャッシュを更新
    cachedUserId = user.id;
    cachedUserName = user.character_name;
    return { content: [{ type: "text", text: lines.join("\n") }] };
  }
);

// ---------- サーバー起動 ----------

const transport = new StdioServerTransport();
await server.connect(transport);
