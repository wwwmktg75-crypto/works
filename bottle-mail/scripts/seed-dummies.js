/**
 * ダミーユーザーをBottle Mailに登録するスクリプト
 * Usage: node scripts/seed-dummies.js [API_BASE_URL]
 */

const API_BASE = process.argv[2] || "https://bottle-mail.vercel.app";

const dummies = [
  {
    character_name: "深海で地図を描く人",
    email: "dummy-deepmap@bottlemail.dev",
    personality_summary: "誰も行かない場所に面白さを見つけるUXリサーチャー。人の行動を観察して静かに設計に落とし込む。言葉より図で考えるタイプで、ノートはいつも余白だらけ。夜中に突然プロトタイプを作り始める。",
    personality_tags: ["UXリサーチ", "観察派", "図解思考", "夜型", "静かな行動力"],
    project_urls: ["https://example.com/ux-research"],
    social_links: {}
  },
  {
    character_name: "朝4時の実験者",
    email: "dummy-earlybird@bottlemail.dev",
    personality_summary: "機械学習エンジニアとして働きながら、早朝の2時間だけ個人プロダクトを作り続ける。スタートアップ的な速度感が好きで、完璧より動くものを優先する。最近はLLMを使ったツールづくりにハマっている。",
    personality_tags: ["機械学習", "早起き", "個人開発", "スピード重視", "LLM活用"],
    project_urls: ["https://example.com/ml-tools"],
    social_links: {}
  },
  {
    character_name: "余白を売るデザイナー",
    email: "dummy-whitespace@bottlemail.dev",
    personality_summary: "フリーランスのグラフィックデザイナー。「引き算のデザイン」を信条に、クライアントを説得し続ける日々。ブランディングからUI、時には空間デザインまで手がける。コーヒーと銭湯が生きる活力。",
    personality_tags: ["グラフィックデザイン", "ブランディング", "引き算美学", "フリーランス", "銭湯愛好家"],
    project_urls: ["https://example.com/design-portfolio"],
    social_links: {}
  },
  {
    character_name: "コードを詩に変える人",
    email: "dummy-poet@bottlemail.dev",
    personality_summary: "Webエンジニアだが根っこは文学部出身。変数名や関数名に異常なこだわりを持ち、コードレビューで詩的な命名を提案して煙たがられる。個人ブログで技術と哲学を混ぜた記事を書いている。",
    personality_tags: ["Webエンジニア", "言語オタク", "哲学好き", "命名こだわり派", "ブログ書き"],
    project_urls: ["https://example.com/poetic-code"],
    social_links: {}
  },
  {
    character_name: "地方をプロダクトにする人",
    email: "dummy-localtech@bottlemail.dev",
    personality_summary: "東京から島根にUターンし、地域課題をテクノロジーで解くプロダクトマネージャー。農業×IoTのスタートアップを立ち上げ中。「東京に出なくていい選択肢を作りたい」が口癖。焚き火しながらリモート会議をやる。",
    personality_tags: ["地方創生", "プロダクトマネジメント", "農業テック", "UIターン", "焚き火"],
    project_urls: ["https://example.com/rural-tech"],
    social_links: {}
  },
  {
    character_name: "音と数字の翻訳者",
    email: "dummy-sounddata@bottlemail.dev",
    personality_summary: "音楽プロデューサー兼データサイエンティスト。音楽のストリーミングデータを分析してトレンドを予測する研究をしている。DJとして週末はクラブで働き、平日は大学院でアルゴリズムを書く二重生活。",
    personality_tags: ["音楽プロデュース", "データサイエンス", "DJ", "大学院生", "二重生活"],
    project_urls: ["https://example.com/music-data"],
    social_links: {}
  },
  {
    character_name: "失敗ログを公開する人",
    email: "dummy-faillog@bottlemail.dev",
    personality_summary: "連続起業家（3回失敗）。失敗の経緯や意思決定をすべてnoteに書き続けている。「失敗は公開資産」という考えで、今は4度目のチャレンジ中。SaaSのPMFを探しながら、同じ境遇の仲間を探している。",
    personality_tags: ["連続起業家", "失敗公開派", "SaaS", "PMF探索中", "透明性重視"],
    project_urls: ["https://example.com/fail-startup"],
    social_links: {}
  },
  {
    character_name: "境界線を引かない建築家",
    email: "dummy-architect@bottlemail.dev",
    personality_summary: "建築とソフトウェアの両方を学んだレアな経歴を持つ。空間設計の考え方をアプリのIA（情報設計）に応用している。建物を見るように画面を見る癖があり、「動線」「素材感」という言葉を多用する。",
    personality_tags: ["建築×IT", "情報設計", "デュアルバックグラウンド", "動線思考", "素材感重視"],
    project_urls: ["https://example.com/arch-software"],
    social_links: {}
  }
];

async function registerUser(user) {
  const res = await fetch(`${API_BASE}/api/users`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(user)
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "登録失敗");
  return data.user;
}

async function main() {
  console.log(`\n🌊 Bottle Mail ダミーユーザー登録\n対象: ${API_BASE}\n`);

  for (const dummy of dummies) {
    process.stdout.write(`  登録中: ${dummy.character_name} ... `);
    try {
      const user = await registerUser(dummy);
      console.log(`✅ ${user.id.slice(0, 8)}...`);
    } catch (err) {
      if (err.message.includes("duplicate") || err.message.includes("unique")) {
        console.log("⏭️  すでに登録済み");
      } else {
        console.log(`❌ ${err.message}`);
      }
    }
    // レート制限対策で少し待つ
    await new Promise(r => setTimeout(r, 1000));
  }

  console.log("\n✨ 完了！ https://bottle-mail.vercel.app で確認してください\n");
}

main().catch(console.error);
