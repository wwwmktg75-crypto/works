import cors from "cors";
import dotenv from "dotenv";
import express from "express";
import { createClient } from "@supabase/supabase-js";
import nodemailer from "nodemailer";
import Stripe from "stripe";

dotenv.config();

const app = express();
const port = Number(process.env.PORT || 8787);

const stripe = process.env.STRIPE_SECRET_KEY
  ? new Stripe(process.env.STRIPE_SECRET_KEY, {
      apiVersion: "2024-06-20"
    })
  : null;

const tableMembers = process.env.SUPABASE_TABLE_MEMBERS || "members";
const tableInquiries = process.env.SUPABASE_TABLE_INQUIRIES || "inquiries";
const tableOrders = process.env.SUPABASE_TABLE_ORDERS || "orders";

const smtpEnabled =
  !!process.env.SMTP_HOST &&
  !!process.env.SMTP_USER &&
  !!process.env.SMTP_PASS &&
  !!process.env.OPERATIONS_EMAIL;

const mailer = smtpEnabled
  ? nodemailer.createTransport({
      host: process.env.SMTP_HOST,
      port: Number(process.env.SMTP_PORT || 587),
      secure: process.env.SMTP_SECURE === "true",
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS
      }
    })
  : null;

const getSupabaseClient = () => {
  if (!process.env.SUPABASE_URL || !process.env.SUPABASE_SERVICE_ROLE_KEY) {
    return null;
  }
  return createClient(
    process.env.SUPABASE_URL,
    process.env.SUPABASE_SERVICE_ROLE_KEY
  );
};

app.use(cors());

app.post(
  "/api/payments/webhook",
  express.raw({ type: "application/json" }),
  async (req, res) => {
    try {
      const signature = req.headers["stripe-signature"];
      if (!signature || !process.env.STRIPE_WEBHOOK_SECRET) {
        return res.status(400).json({ error: "Webhook設定が不正です" });
      }

      if (!stripe) {
        return res.status(400).json({ error: "Stripe秘密鍵が未設定です" });
      }
      const supabaseClient = getSupabaseClient();
      if (!supabaseClient) {
        return res.status(400).json({ error: "Supabase設定が未完了です" });
      }

      const event = stripe.webhooks.constructEvent(
        req.body,
        signature,
        process.env.STRIPE_WEBHOOK_SECRET
      );

      if (event.type === "checkout.session.completed") {
        const session = event.data.object;
        const { error } = await supabaseClient.from(tableOrders).insert({
          paymentIntentId: session.payment_intent || "",
          checkoutSessionId: session.id,
          memberRecordId: session.metadata?.memberRecordId || "",
          serviceTitle: session.metadata?.serviceTitle || "",
          amount: (session.amount_total || 0) / 100,
          currency: session.currency || "jpy",
          payerEmail: session.customer_details?.email || "",
          status: "paid"
        });
        if (error) {
          throw new Error(error.message);
        }
      }

      return res.status(200).json({ received: true });
    } catch (error) {
      return res.status(400).json({
        error: "Webhook処理に失敗しました",
        detail: error.message
      });
    }
  }
);

app.use(express.json());

app.get("/health", (_req, res) => {
  res.json({ status: "ok" });
});

app.get("/api/members", async (req, res) => {
  try {
    const supabaseClient = getSupabaseClient();
    if (!supabaseClient) {
      return res.status(400).json({ error: "Supabase設定が未完了です" });
    }

    const keyword = req.query.keyword?.toString().trim();
    const minRate = req.query.minRate ? Number(req.query.minRate) : 0;
    const maxRate = req.query.maxRate ? Number(req.query.maxRate) : 9999999;

    const { data, error } = await supabaseClient
      .from(tableMembers)
      .select("*")
      .eq("isApproved", true);
    if (error) {
      throw new Error(error.message);
    }

    const members = (data || [])
      .filter((member) => {
        const title = (member.displayTitle || "").toLowerCase();
        const skills = (member.skillTags || "").toLowerCase();
        const rate = Number(member.hourlyRate || 0);
        const hit = keyword
          ? title.includes(keyword.toLowerCase()) ||
            skills.includes(keyword.toLowerCase())
          : true;
        return hit && rate >= minRate && rate <= maxRate;
      });

    return res.json({ members });
  } catch (error) {
    return res.status(500).json({
      error: "メンバー取得に失敗しました",
      detail: error.message
    });
  }
});

app.post("/api/members", async (req, res) => {
  try {
    const supabaseClient = getSupabaseClient();
    if (!supabaseClient) {
      return res.status(400).json({ error: "Supabase設定が未完了です" });
    }

    const {
      displayName,
      displayTitle,
      skillTags,
      hourlyRate,
      profileText,
      portfolioUrl,
      email
    } = req.body;

    if (!displayName || !skillTags || !email) {
      return res.status(400).json({
        error: "displayName, skillTags, email は必須です"
      });
    }

    const { data, error } = await supabaseClient
      .from(tableMembers)
      .insert({
        displayName,
        displayTitle: displayTitle || "",
        skillTags,
        hourlyRate: Number(hourlyRate || 0),
        profileText: profileText || "",
        portfolioUrl: portfolioUrl || "",
        email,
        isApproved: false
      })
      .select("id")
      .single();
    if (error) {
      throw new Error(error.message);
    }

    return res.status(201).json({
      message: "登録を受け付けました。運営承認後に公開されます。",
      memberId: data.id
    });
  } catch (error) {
    return res.status(500).json({
      error: "メンバー登録に失敗しました",
      detail: error.message
    });
  }
});

app.post("/api/inquiries", async (req, res) => {
  try {
    const supabaseClient = getSupabaseClient();
    if (!supabaseClient) {
      return res.status(400).json({ error: "Supabase設定が未完了です" });
    }

    const { memberRecordId, clientName, clientEmail, message } = req.body;
    if (!memberRecordId || !clientName || !clientEmail || !message) {
      return res.status(400).json({
        error: "memberRecordId, clientName, clientEmail, message は必須です"
      });
    }

    const { data: member, error: memberError } = await supabaseClient
      .from(tableMembers)
      .select("id, displayName, email")
      .eq("id", memberRecordId)
      .single();
    if (memberError) {
      throw new Error(memberError.message);
    }

    const memberEmail = member.email;
    if (!memberEmail) {
      return res.status(400).json({
        error: "対象メンバーのメールアドレスが未登録です"
      });
    }

    const { data: inquiry, error: inquiryError } = await supabaseClient
      .from(tableInquiries)
      .insert({
        memberRecordId,
        clientName,
        clientEmail,
        message,
        status: "new"
      })
      .select("id")
      .single();
    if (inquiryError) {
      throw new Error(inquiryError.message);
    }

    if (mailer) {
      const to = [memberEmail, process.env.OPERATIONS_EMAIL].join(",");
      await mailer.sendMail({
        from: process.env.SMTP_USER,
        to,
        subject: `【新規相談】${member.displayName || "メンバー"}宛`,
        text: `依頼主: ${clientName}\nメール: ${clientEmail}\n\n相談内容:\n${message}`
      });
    }

    return res.status(201).json({
      message: "問い合わせを送信しました",
      inquiryId: inquiry.id
    });
  } catch (error) {
    return res.status(500).json({
      error: "問い合わせ送信に失敗しました",
      detail: error.message
    });
  }
});

app.post("/api/payments/checkout", async (req, res) => {
  try {
    const { memberRecordId, serviceTitle, amount, buyerEmail } = req.body;
    if (!memberRecordId || !serviceTitle || !amount) {
      return res.status(400).json({
        error: "memberRecordId, serviceTitle, amount は必須です"
      });
    }

    const safeAmount = Number(amount);
    if (!Number.isFinite(safeAmount) || safeAmount <= 0) {
      return res.status(400).json({ error: "amount は正の数値にしてください" });
    }

    const platformFeeRate = Number(process.env.STRIPE_PLATFORM_FEE_RATE || 0.1);
    const platformFee = Math.round(safeAmount * platformFeeRate);

    if (!stripe) {
      return res.status(400).json({ error: "Stripe秘密鍵が未設定です" });
    }

    const session = await stripe.checkout.sessions.create({
      mode: "payment",
      customer_email: buyerEmail || undefined,
      line_items: [
        {
          price_data: {
            currency: "jpy",
            product_data: {
              name: serviceTitle
            },
            unit_amount: Math.round(safeAmount)
          },
          quantity: 1
        }
      ],
      metadata: {
        memberRecordId,
        serviceTitle,
        platformFee: String(platformFee)
      },
      success_url:
        process.env.STRIPE_SUCCESS_URL || `${process.env.APP_BASE_URL}/success`,
      cancel_url:
        process.env.STRIPE_CANCEL_URL || `${process.env.APP_BASE_URL}/cancel`
    });

    return res.status(201).json({
      checkoutUrl: session.url,
      checkoutSessionId: session.id
    });
  } catch (error) {
    return res.status(500).json({
      error: "Stripe決済セッションの作成に失敗しました",
      detail: error.message
    });
  }
});

if (!process.env.VERCEL) {
  app.listen(port, () => {
    console.log(`skill-matching-platform API running on :${port}`);
  });
}

export default app;
