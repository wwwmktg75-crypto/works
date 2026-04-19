# Vercel で AIキャリアコーチ LP をデプロイする手順

## 1. Vercel にアクセス・ログイン

1. ブラウザで **https://vercel.com** を開く
2. **Sign Up** または **Log In** をクリック
3. **Continue with GitHub** を選ぶ（GitHub アカウントでログインすると、リポジトリ連携が簡単）

---

## 2. 新規プロジェクトの作成（リポジトリのインポート）

1. ダッシュボードで **「Add New...」** → **「Project」** をクリック
2. **Import Git Repository** の一覧から **`wwwmktg75-crypto/works`** を探す
   - 表示されない場合は **「Import Third-Party Git Repository」** に  
     `https://github.com/wwwmktg75-crypto/works` を貼って **Import**
3. リポジトリを選択して **「Import」** をクリック

---

## 3. 重要：Root Directory の設定

LP はリポジトリの **サブフォルダ** に入っているため、ここを必ず設定します。

1. **Configure Project** 画面で **「Root Directory」** の右側 **「Edit」** をクリック
2. 入力欄に **`ai-career-coach-lp`** と入力
3. **「Continue」** をクリックして確定

これで、Vercel は `ai-career-coach-lp` フォルダだけをサイトのルートとしてデプロイします。

---

## 4. ビルド設定（この LP ではほぼそのままでOK）

- **Framework Preset**: 何も選ばない（または **Other**）でOK
- **Build Command**: 空のままでOK（静的 HTML のためビルド不要）
- **Output Directory**: 空、または **`.`** のままでOK
- **Install Command**: 空でOK

そのまま **「Deploy」** をクリック。

---

## 5. デプロイ完了まで

- 数十秒〜1分ほどでデプロイが完了します
- 完了すると **「Visit」** または **「Go to Dashboard」** が表示されます
- **「Visit」** をクリックすると、LP の公開 URL（例: `https://works-xxxx.vercel.app`）が開きます

---

## 6. よく使う操作（ダッシュボード）

### デプロイ履歴を見る

- プロジェクト名をクリック → **「Deployments」** タブ  
  - 過去のデプロイ一覧と、各デプロイの URL・ログを確認できます

### 再デプロイする

- **「Deployments」** で対象のデプロイの **「⋯」** → **「Redeploy」**
- または GitHub に `main` をプッシュすると、自動で再デプロイされます（Git 連携時）

### 本番 URL を確認する

- プロジェクトの **「Settings」** → **「Domains」**  
  - デフォルトの `*.vercel.app` の URL が表示されます

### カスタムドメインを付ける（任意）

1. プロジェクトの **「Settings」** → **「Domains」**
2. **「Add」** でドメイン（例: `lp.example.com`）を入力
3. 表示される CNAME や A レコードを、ドメイン管理画面（お名前.com 等）で設定

---

## 7. トラブルシューティング

| 現象 | 対処 |
|------|------|
| 404 や真っ白なページ | **Root Directory** が `ai-career-coach-lp` になっているか確認し、保存して **Redeploy** |
| スタイルやJSが効かない | `index.html` の `href="styles.css"` / `src="main.js"` が相対パス（`./` なし）になっているか確認（現在の構成で問題なし） |
| デプロイが失敗する | **Deployments** の該当デプロイを開き **「Building」** のログを確認。ビルドコマンドを空にしているか確認 |

---

## 8. まとめ（最低限の手順）

1. **vercel.com** で GitHub ログイン
2. **Add New → Project** で `works` リポジトリを **Import**
3. **Root Directory** を **`ai-career-coach-lp`** に設定
4. **Deploy** をクリック
5. 表示された **Visit** の URL で LP を確認

以降は、GitHub の `main` にプッシュするたびに自動で再デプロイされます。
