# PerioBox風 4枚構成PowerPoint生成 MVP

歯科医療論文PDFをアップロードすると、PDF本文を抽出し、OpenAI APIで日本語に構造化要約して、4枚構成のPowerPoint（`.pptx`）を生成するStreamlitアプリです。

## 機能

- PDFアップロード
- PDF本文抽出
- OpenAI APIによる3段階処理
  - PDF本文からセクション別抽出
  - 原文に忠実な日本語構造化要約
  - PerioBox式4枚スライド要約
- 4:3、PerioBox見本に寄せた背景・余白のPowerPoint生成
- PowerPoint上で後から編集しやすいテキストボックス・図形構造
- 研究論文の2枚目・3枚目は、本文情報をもとにPowerPoint図形・表で作れる範囲の図表を自動作成
- PDF抽出本文をページ表記つきで全文翻訳したMarkdownを出力
- Supabase AuthによるGoogleログイン
- `allowed_users` に登録されたメールアドレスだけが利用可能
- `admin` ユーザーだけが許可メールを追加・削除できる管理画面

## スライド構成

1. タイトル、著者、ジャーナル、発表年、Clinical Question、結論
2. 研究デザイン、対象、選定条件、研究内容を補完する図・表、評価項目
3. 主な結果、数値データ、統計的有意差、主要結果の根拠表
4. 臨床的選択ポイント2つ、臨床家向けコメント、限界

## セットアップ

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Supabase認証のセットアップ

このアプリは少人数モデレーター用の内部ツールとして、Supabase AuthのGoogleログインと許可メール制限を使います。

1. Supabaseで新規プロジェクトを作成します。
2. Supabase SQL Editorで [`supabase_schema.sql`](supabase_schema.sql) を実行します。
3. SQLファイル末尾のコメントを参考に、自分のGoogleアカウントを最初の `admin` として追加します。

```sql
insert into public.allowed_users (email, role)
values ('your-email@example.com', 'admin')
on conflict (email) do update set role = excluded.role;
```

4. Supabaseの Authentication > Providers で Google を有効にします。
5. Google Cloud側でOAuthクライアントを作成し、SupabaseにClient ID / Client Secretを設定します。
6. Supabaseの Authentication > URL Configuration にリダイレクトURLを追加します。

ローカル開発:

```text
http://localhost:8501/
```

本番:

```text
https://your-app.example.com/
```

7. 環境変数を設定します。`.env.example` も参考にしてください。

```bash
export SUPABASE_URL="https://xxxxxxxxxxxx.supabase.co"
export SUPABASE_ANON_KEY="your-supabase-anon-key"
export APP_BASE_URL="http://localhost:8501/"
```

OpenAI APIキーを設定します。

本番運用では、OpenAI APIキーは管理者がサーバー側のSecrets/環境変数に設定してください。モデレーターにはAPIキー入力欄を表示しません。

```bash
export OPENAI_API_KEY="sk-..."
```

必要に応じてモデルを変更できます。

```bash
export OPENAI_MODEL="gpt-4o-mini"
```

## 起動

```bash
streamlit run app.py
```

ブラウザで表示されたStreamlit画面からPDFをアップロードし、「PowerPointを生成」を押してください。

## モデレーターへの共有

デプロイと共有の詳しい手順は [`DEPLOYMENT.md`](DEPLOYMENT.md) を参照してください。

最初はStreamlit Community Cloudに置き、Supabaseのadmin画面からモデレーター2名のGoogleメールアドレスを `moderator` として追加する運用を推奨します。

## Netlifyデプロイについて

Netlifyは静的サイト・フロントエンド向けのホスティングです。このリポジトリの本体はPythonのStreamlitアプリなので、そのままNetlifyだけにデプロイして常時起動することはできません。

Netlify前提で運用する場合は、次の構成にしてください。

- Netlify: Googleログイン済みユーザー向けのフロントエンド
- Supabase Auth: Googleログインと `allowed_users` による権限管理
- Python実行環境: Streamlit Community Cloud、Render、Railway、Cloud Runなどでこのアプリを実行
- Netlify環境変数: `SUPABASE_URL`, `SUPABASE_ANON_KEY`, `APP_BASE_URL`

現在の実装では認証・権限チェックを [`auth.py`](auth.py) に分離しているため、将来Netlify用のReact/Next.jsフロントエンドへ切り出す場合も、許可メール制限の考え方をそのまま移植できます。

## 注意

- スキャンPDFなど、テキスト層がないPDFは本文抽出できません。OCR済みPDFを使用してください。
- 論文本文が非常に長い場合、MVPでは冒頭と末尾を中心に最大60,000文字へ丸めてOpenAI APIへ送信します。
- 研究論文の2枚目・3枚目では、本文情報をもとにPowerPoint図形・表で作れる範囲の図表を自動作成します。論文図表の完全な再現ではないため、必要に応じて原著本文で確認してください。
- 日本語訳MarkdownはPDFから抽出できた本文の全文翻訳です。PDF抽出の崩れやスキャンPDFでは精度が落ちるため、重要箇所は原著PDFで確認してください。
- 医療判断の最終根拠として使う前に、必ず原著論文の本文・図表・統計値を確認してください。
