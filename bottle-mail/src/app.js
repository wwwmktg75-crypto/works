import cors from "cors";
import dotenv from "dotenv";
import express from "express";
import { createClient } from "@supabase/supabase-js";
import Anthropic from "@anthropic-ai/sdk";
import path from "path";
import { fileURLToPath } from "url";

dotenv.config({ path: new URL("../.env", import.meta.url).pathname });

const app = express();
const port = Number(process.env.PORT || 8788);
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "../public")));

// ---------- クライアント初期化 ----------

const getSupabase = () => {
  if (!process.env.SUPABASE_URL || !process.env.SUPABASE_SERVICE_ROLE_KEY) return null;
  return createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);
};

const getAnthropic = () => {
  if (!process.env.ANTHROPIC_API_KEY) return null;
  return new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
};

// ---------- ユーティリティ ----------

async function fetchPageText(url) {
  try {
    const res = await fetch(url, {
      headers: { "User-Agent": "Mozilla/5.0 (compatible; BottleMail/1.0)" },
      signal: AbortSignal.timeout(5000)
    });
    const html = await res.text();
    // タグ除去・空白圧縮（簡易）
    return html
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 3000);
  } catch {
    return "";
  }
}

async function analyzeWithClaude(anthropic, urlTexts, socialLinks) {
  const urlSummaries = urlTexts
    .map((t, i) => `=== URL${i + 1} ===\n${t.url}\n${t.text}`)
    .join("\n\n");

  const socialSummary = Object.entries(socialLinks || {})
    .filter(([, v]) => v)
    .map(([k, v]) => `${k}: ${v}`)
    .join("\n");

  const prompt = `
あなたはユーザーのプロジェクトやSNSからその人の個性を読み取るAIです。

以下の情報を分析してください：

${urlSummaries}
${socialSummary ? `\n=== SNSリンク ===\n${socialSummary}` : ""}

次のJSONを返してください（説明文なし、JSONのみ）：
{
  "personality_summary": "この人の個性・価値観・得意なことを200字以内で表現（日本語）",
  "personality_tags": ["タグ1", "タグ2", "タグ3", "タグ4", "タグ5"],
  "name_suggestions": [
    { "name": "キャラクター名1", "reason": "この名前にした理由（30字以内）" },
    { "name": "キャラクター名2", "reason": "この名前にした理由（30字以内）" },
    { "name": "キャラクター名3", "reason": "この名前にした理由（30字以内）" }
  ]
}

キャラクター名は詩的・ユニークな日本語名にしてください（例：「深海の設計者」「夜明けを待つコード」「静かな炎の建築家」）。
`;

  const message = await anthropic.messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 1024,
    messages: [{ role: "user", content: prompt }]
  });

  const text = message.content[0].text.trim();
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error("Claude応答のパースに失敗しました");
  return JSON.parse(jsonMatch[0]);
}

async function calcMatchScore(anthropic, userA, userB) {
  const prompt = `
2人のユーザーのマッチ度を判定してください。

【ユーザーA】
個性: ${userA.personality_summary}
タグ: ${(userA.personality_tags || []).join(", ")}

【ユーザーB】
個性: ${userB.personality_summary}
タグ: ${(userB.personality_tags || []).join(", ")}

次のJSONのみ返してください：
{
  "score": 0から100の数値（高いほど相性◎）,
  "reason": "マッチした理由を50字以内で（日本語）"
}
`;

  const message = await anthropic.messages.create({
    model: "claude-haiku-4-5-20251001",
    max_tokens: 256,
    messages: [{ role: "user", content: prompt }]
  });

  const text = message.content[0].text.trim();
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) return { score: 50, reason: "AIがマッチを感じました" };
  return JSON.parse(jsonMatch[0]);
}

// ---------- API ----------

app.get("/health", (_req, res) => res.json({
  status: "ok",
  hasSupabase: !!process.env.SUPABASE_URL,
  hasAnthropic: !!process.env.ANTHROPIC_API_KEY,
  node: process.version
}));

// URLを解析してAIが個性・名前候補を返す
app.post("/api/analyze", async (req, res) => {
  try {
    const anthropic = getAnthropic();
    if (!anthropic) return res.status(400).json({ error: "ANTHROPIC_API_KEY が未設定です" });

    const { project_urls, social_links } = req.body;
    if (!project_urls?.length) {
      return res.status(400).json({ error: "project_urls を1つ以上入力してください" });
    }

    // 並行してページ取得
    const urlTexts = await Promise.all(
      project_urls.slice(0, 5).map(async (url) => ({
        url,
        text: await fetchPageText(url)
      }))
    );

    const analysis = await analyzeWithClaude(anthropic, urlTexts, social_links);
    return res.json(analysis);
  } catch (err) {
    return res.status(500).json({ error: "分析に失敗しました", detail: err.message, cause: err.cause?.message });
  }
});

// ユーザー登録 + 自動マッチング
app.post("/api/users", async (req, res) => {
  try {
    const supabase = getSupabase();
    if (!supabase) return res.status(400).json({ error: "Supabase設定が未完了です" });

    const anthropic = getAnthropic();

    const { character_name, personality_summary, personality_tags, project_urls, social_links, email } = req.body;
    if (!character_name || !email) {
      return res.status(400).json({ error: "character_name と email は必須です" });
    }

    const { data: newUser, error: insertErr } = await supabase
      .from("bottle_users")
      .insert({
        character_name,
        personality_summary: personality_summary || "",
        personality_tags: personality_tags || [],
        project_urls: project_urls || [],
        social_links: social_links || {},
        email
      })
      .select()
      .single();

    if (insertErr) throw new Error(insertErr.message);

    // 先にレスポンスを返してからマッチングを実行（タイムアウト回避）
    res.status(201).json({ user: newUser });

    if (anthropic && personality_summary) {
      runMatching(supabase, anthropic, newUser).catch(console.error);
    }
  } catch (err) {
    return res.status(500).json({ error: "登録に失敗しました", detail: err.message });
  }
});

// マッチング処理（登録後バックグラウンド実行）
async function runMatching(supabase, anthropic, newUser) {
  if (!newUser.personality_summary) return;

  const { data: others } = await supabase
    .from("bottle_users")
    .select("*")
    .neq("id", newUser.id)
    .not("personality_summary", "is", null);

  if (!others?.length) return;

  // 既存ボトルを一括取得
  const { data: existingBottles } = await supabase
    .from("bottles")
    .select("sender_id, recipient_id")
    .or(`sender_id.eq.${newUser.id},recipient_id.eq.${newUser.id}`);

  const alreadyMatched = new Set(
    (existingBottles || []).map(b =>
      b.sender_id === newUser.id ? b.recipient_id : b.sender_id
    )
  );

  const targets = others.filter(o => !alreadyMatched.has(o.id));
  if (!targets.length) return;

  // 全員と並列でスコア計算
  const results = await Promise.all(
    targets.map(async (other) => {
      const match = await calcMatchScore(anthropic, newUser, other);
      return { other, match };
    })
  );

  // 60点以上のみボトル送信
  const toInsert = results
    .filter(({ match }) => match.score >= 60)
    .flatMap(({ other, match }) => [
      { sender_id: newUser.id, recipient_id: other.id, match_score: match.score, match_reason: match.reason, status: "unread" },
      { sender_id: other.id, recipient_id: newUser.id, match_score: match.score, match_reason: match.reason, status: "unread" }
    ]);

  if (toInsert.length) {
    await supabase.from("bottles").insert(toInsert);
  }
}

// 送ったボトル一覧
app.get("/api/bottles/sent", async (req, res) => {
  try {
    const supabase = getSupabase();
    if (!supabase) return res.status(400).json({ error: "Supabase設定が未完了です" });

    const userId = req.headers["x-user-id"];
    if (!userId) return res.status(401).json({ error: "x-user-id ヘッダーが必要です" });

    const { data, error } = await supabase
      .from("bottles")
      .select("*, recipient:recipient_id(id, character_name, personality_tags)")
      .eq("sender_id", userId)
      .order("created_at", { ascending: false });

    if (error) throw new Error(error.message);
    return res.json({ bottles: data });
  } catch (err) {
    return res.status(500).json({ error: "取得に失敗しました", detail: err.message });
  }
});

// 受け取ったボトル一覧
app.get("/api/bottles", async (req, res) => {
  try {
    const supabase = getSupabase();
    if (!supabase) return res.status(400).json({ error: "Supabase設定が未完了です" });

    const userId = req.headers["x-user-id"];
    if (!userId) return res.status(401).json({ error: "x-user-id ヘッダーが必要です" });

    const { data, error } = await supabase
      .from("bottles")
      .select("*, sender:sender_id(id, character_name, personality_tags)")
      .eq("recipient_id", userId)
      .order("created_at", { ascending: false });

    if (error) throw new Error(error.message);
    return res.json({ bottles: data });
  } catch (err) {
    return res.status(500).json({ error: "取得に失敗しました", detail: err.message });
  }
});

// ボトルを開封
app.post("/api/bottles/:bottleId/open", async (req, res) => {
  try {
    const supabase = getSupabase();
    if (!supabase) return res.status(400).json({ error: "Supabase設定が未完了です" });

    const userId = req.headers["x-user-id"];
    const { bottleId } = req.params;

    const { data, error } = await supabase
      .from("bottles")
      .update({ status: "opened" })
      .eq("id", bottleId)
      .eq("recipient_id", userId)
      .select("*, sender:sender_id(id, character_name, personality_summary, personality_tags)")
      .single();

    if (error) throw new Error(error.message);
    return res.json({ bottle: data });
  } catch (err) {
    return res.status(500).json({ error: "開封に失敗しました", detail: err.message });
  }
});

// 同じユーザー2人のペアに対応する bottles 行は2件（送信方向が逆）ある。
// メッセージはどちらか一方の bottle_id に紐づくため、スレッド表示では両方の id を束ねる。
async function getBottleIdsForPair(supabase, senderId, recipientId) {
  const [{ data: rowsAB }, { data: rowsBA }] = await Promise.all([
    supabase.from("bottles").select("id").eq("sender_id", senderId).eq("recipient_id", recipientId),
    supabase.from("bottles").select("id").eq("sender_id", recipientId).eq("recipient_id", senderId)
  ]);
  const ids = [...new Set([...(rowsAB || []), ...(rowsBA || [])].map((r) => r.id))];
  return ids;
}

// メッセージ送信（ボトルに返信 or チャット続行）
app.post("/api/messages", async (req, res) => {
  try {
    const supabase = getSupabase();
    if (!supabase) return res.status(400).json({ error: "Supabase設定が未完了です" });

    const userId = req.headers["x-user-id"];
    const { bottle_id, content } = req.body;
    if (!bottle_id || !content?.trim()) {
      return res.status(400).json({ error: "bottle_id と content は必須です" });
    }
    if (!userId) return res.status(401).json({ error: "x-user-id ヘッダーが必要です" });

    const { data: bottleRow, error: bottleErr } = await supabase
      .from("bottles")
      .select("sender_id, recipient_id")
      .eq("id", bottle_id)
      .single();

    if (bottleErr || !bottleRow) {
      return res.status(404).json({ error: "ボトルが見つかりません" });
    }
    if (bottleRow.sender_id !== userId && bottleRow.recipient_id !== userId) {
      return res.status(403).json({ error: "このボトルにメッセージを送る権限がありません" });
    }

    // ボトルをreplied状態に
    await supabase
      .from("bottles")
      .update({ status: "replied" })
      .eq("id", bottle_id);

    // 逆方向のボトルも replied に
    const { data: thisBottle } = await supabase
      .from("bottles")
      .select("sender_id, recipient_id")
      .eq("id", bottle_id)
      .single();

    if (thisBottle) {
      await supabase
        .from("bottles")
        .update({ status: "replied" })
        .eq("sender_id", thisBottle.recipient_id)
        .eq("recipient_id", thisBottle.sender_id);
    }

    const { data: msg, error } = await supabase
      .from("bottle_messages")
      .insert({ bottle_id, sender_id: userId, content: content.trim() })
      .select()
      .single();

    if (error) throw new Error(error.message);
    return res.status(201).json({ message: msg });
  } catch (err) {
    return res.status(500).json({ error: "送信に失敗しました", detail: err.message });
  }
});

// チャット履歴取得（マッチ時の双方向ボトル行で共有された1スレッドとして返す）
app.get("/api/messages/:bottleId", async (req, res) => {
  try {
    const supabase = getSupabase();
    if (!supabase) return res.status(400).json({ error: "Supabase設定が未完了です" });

    const userId = req.headers["x-user-id"];
    if (!userId) return res.status(401).json({ error: "x-user-id ヘッダーが必要です" });

    const { bottleId } = req.params;

    const { data: bottleRow, error: bottleErr } = await supabase
      .from("bottles")
      .select("sender_id, recipient_id")
      .eq("id", bottleId)
      .single();

    if (bottleErr || !bottleRow) {
      return res.status(404).json({ error: "ボトルが見つかりません" });
    }
    if (bottleRow.sender_id !== userId && bottleRow.recipient_id !== userId) {
      return res.status(403).json({ error: "このチャットを表示する権限がありません" });
    }

    const pairIds = await getBottleIdsForPair(supabase, bottleRow.sender_id, bottleRow.recipient_id);
    if (!pairIds.length) {
      return res.json({ messages: [] });
    }

    const { data, error } = await supabase
      .from("bottle_messages")
      .select("*, sender:sender_id(id, character_name)")
      .in("bottle_id", pairIds)
      .order("created_at", { ascending: true });

    if (error) throw new Error(error.message);
    return res.json({ messages: data });
  } catch (err) {
    return res.status(500).json({ error: "取得に失敗しました", detail: err.message });
  }
});

// メールアドレスでログイン（IDを返す）
app.post("/api/login", async (req, res) => {
  try {
    const supabase = getSupabase();
    if (!supabase) return res.status(400).json({ error: "Supabase設定が未完了です" });

    const { email } = req.body;
    if (!email) return res.status(400).json({ error: "email は必須です" });

    const { data, error } = await supabase
      .from("bottle_users")
      .select("id, character_name, personality_tags")
      .eq("email", email.trim().toLowerCase())
      .single();

    if (error || !data) return res.status(404).json({ error: "メールアドレスが見つかりません" });
    return res.json({ user: data });
  } catch (err) {
    return res.status(500).json({ error: "ログインに失敗しました", detail: err.message });
  }
});

// ユーザー情報取得
app.get("/api/users/me", async (req, res) => {
  try {
    const supabase = getSupabase();
    if (!supabase) return res.status(400).json({ error: "Supabase設定が未完了です" });

    const userId = req.headers["x-user-id"];
    if (!userId) return res.status(401).json({ error: "x-user-id ヘッダーが必要です" });

    const { data, error } = await supabase
      .from("bottle_users")
      .select("id, character_name, personality_summary, personality_tags, project_urls, created_at")
      .eq("id", userId)
      .single();

    if (error) throw new Error(error.message);
    return res.json({ user: data });
  } catch (err) {
    return res.status(500).json({ error: "取得に失敗しました", detail: err.message });
  }
});

if (!process.env.VERCEL) {
  app.listen(port, () => {
    console.log(`🌊 bottle-mail running on http://localhost:${port}`);
  });
}

export default app;
