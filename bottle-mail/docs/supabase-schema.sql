-- ボトルメール マッチングアプリ Supabase スキーマ

-- ユーザーテーブル
CREATE TABLE bottle_users (
  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  character_name  TEXT NOT NULL,
  personality_summary TEXT,
  personality_tags    TEXT[] DEFAULT '{}',
  project_urls        TEXT[] DEFAULT '{}',
  social_links        JSONB  DEFAULT '{}',
  email               TEXT NOT NULL UNIQUE,
  created_at          TIMESTAMPTZ DEFAULT NOW()
);

-- ボトルテーブル（マッチ通知）
CREATE TABLE bottles (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  sender_id     UUID REFERENCES bottle_users(id) ON DELETE CASCADE,
  recipient_id  UUID REFERENCES bottle_users(id) ON DELETE CASCADE,
  match_score   FLOAT,
  match_reason  TEXT,
  status        TEXT DEFAULT 'unread',  -- unread | opened | replied | ignored
  created_at    TIMESTAMPTZ DEFAULT NOW(),
  UNIQUE(sender_id, recipient_id)
);

-- メッセージテーブル（チャット）
CREATE TABLE bottle_messages (
  id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  bottle_id   UUID REFERENCES bottles(id) ON DELETE CASCADE,
  sender_id   UUID REFERENCES bottle_users(id) ON DELETE CASCADE,
  content     TEXT NOT NULL,
  created_at  TIMESTAMPTZ DEFAULT NOW()
);

-- インデックス
CREATE INDEX idx_bottles_recipient ON bottles(recipient_id);
CREATE INDEX idx_bottles_sender    ON bottles(sender_id);
CREATE INDEX idx_messages_bottle   ON bottle_messages(bottle_id);
CREATE INDEX idx_messages_created  ON bottle_messages(created_at);
