/**
 * ブラウザ → このAPI → GAS ウェブアプリ（CORS回避用プロキシ）
 * 環境変数:
 *   GAS_WEBAPP_URL … GAS デプロイ後の https://script.google.com/macros/s/.../exec
 *   DISCORD_WEBHOOK_URL …（任意）Discord サーバー設定 → 連携サービス → Webhooks → URL
 */

function readJsonBody(req) {
  if (req.body != null && typeof req.body === 'object' && !Buffer.isBuffer(req.body)) {
    return Promise.resolve(req.body);
  }
  if (typeof req.body === 'string') {
    try {
      return Promise.resolve(JSON.parse(req.body));
    } catch (e) {
      return Promise.resolve(null);
    }
  }
  return new Promise(function (resolve, reject) {
    var raw = '';
    req.on('data', function (chunk) {
      raw += chunk;
    });
    req.on('end', function () {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch (e) {
        resolve(null);
      }
    });
    req.on('error', reject);
  });
}

function occupationJa(value) {
  var m = {
    company: '会社員',
    entrepreneur: '経営者・起業家',
    freelance: 'フリーランス',
    student: '学生',
    public: '公務員',
    other: 'その他'
  };
  return m[value] || value || '（未選択）';
}

function truncateField(text, maxLen) {
  if (text == null || text === '') {
    return '（なし）';
  }
  var s = String(text);
  if (s.length <= maxLen) {
    return s;
  }
  return s.slice(0, maxLen - 1) + '…';
}

/**
 * 応募通知（失敗しても応募自体は成功扱い。ログのみ）
 */
async function notifyDiscord_(body) {
  var url = process.env.DISCORD_WEBHOOK_URL;
  if (!url || url.indexOf('https://discord.com/api/webhooks/') !== 0) {
    return;
  }

  var embed = {
    title: '新規お申し込み（AI-YRooze-in-FUKUOKA）',
    color: 0x5865f2,
    fields: [
      { name: 'お名前', value: truncateField(body.fullName, 256), inline: true },
      { name: '職業', value: truncateField(occupationJa(body.occupation), 256), inline: true },
      {
        name: 'AI活用について',
        value: truncateField(body.aiUsage, 1024)
      },
      {
        name: 'メッセージ',
        value: truncateField(body.message, 1024)
      }
    ],
    timestamp: new Date().toISOString()
  };

  var payload = {
    username: 'AI交流会フォーム',
    embeds: [embed]
  };

  var r = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify(payload)
  });

  if (!r.ok) {
    var t = await r.text();
    console.error('[Discord] 通知失敗', r.status, t);
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Method not allowed' });
    return;
  }

  var gasUrl = process.env.GAS_WEBAPP_URL;
  if (!gasUrl || gasUrl.indexOf('https://script.google.com/') !== 0) {
    res.status(503).json({
      ok: false,
      error: 'GAS_WEBAPP_URL が未設定です。Vercel の環境変数を確認してください。'
    });
    return;
  }

  var body = await readJsonBody(req);
  if (body == null) {
    res.status(400).json({ ok: false, error: 'Invalid JSON' });
    return;
  }

  try {
    var r = await fetch(gasUrl, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: JSON.stringify(body)
    });

    var text = await r.text();
    var parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      parsed = { raw: text };
    }

    if (!r.ok) {
      res.status(502).json({
        ok: false,
        error: 'GAS からの応答が異常です',
        status: r.status,
        gas: parsed
      });
      return;
    }

    if (parsed && parsed.ok === false) {
      res.status(502).json({
        ok: false,
        error: parsed.error || 'スプレッドシートへの書き込みに失敗しました',
        gas: parsed
      });
      return;
    }

    try {
      await notifyDiscord_(body);
    } catch (discordErr) {
      console.error('[Discord] notifyDiscord_ error', discordErr);
    }

    res.status(200).json({ ok: true, gas: parsed });
  } catch (err) {
    res.status(500).json({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  }
};
