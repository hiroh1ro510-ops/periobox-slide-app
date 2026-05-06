# PerioBoxモデレーター向け運用・共有ガイド

このアプリは、PerioBoxスライド作成を担当する少人数モデレーター向けの内部ツールとして運用します。

おすすめは、まず **Streamlit Community Cloud** に置いて、Supabase AuthのGoogleログインと `allowed_users` で利用者を制限する方法です。

## 推奨構成

- アプリ本体: Streamlit Community Cloud
- ログイン: Supabase Auth Google
- 利用制限: Supabase `allowed_users`
- OpenAI APIキー: アプリのSecrets/環境変数で管理
- モデレーター: Googleログインのみ。APIキーは見せない

## 1. Supabaseを設定する

1. Supabaseでプロジェクトを作成します。
2. SQL Editorで `supabase_schema.sql` を実行します。
3. 最初のadminとして先生のGoogleメールアドレスを追加します。

```sql
insert into public.allowed_users (email, role)
values ('your-email@example.com', 'admin')
on conflict (email) do update set role = excluded.role;
```

4. Authentication > Providers > Google を有効にします。
5. Google Cloud ConsoleでOAuthクライアントを作成し、Client ID / Client SecretをSupabaseに設定します。
6. Authentication > URL Configuration に、後で発行されるStreamlitアプリURLを追加します。

ローカル検証用:

```text
http://localhost:8501/
```

本番用:

```text
https://your-app-name.streamlit.app/
```

## 2. Streamlit Community Cloudにデプロイする

1. このプロジェクトをGitHubリポジトリに置きます。
2. Streamlit Community Cloudにログインします。
3. New appを押します。
4. Repositoryを選びます。
5. Main file pathに `app.py` を指定します。
6. Advanced settingsまたはSecretsに以下を設定します。

```toml
OPENAI_API_KEY = "sk-..."
OPENAI_MODEL = "gpt-4o-mini"
SUPABASE_URL = "https://xxxxxxxxxxxx.supabase.co"
SUPABASE_ANON_KEY = "your-supabase-anon-key"
APP_BASE_URL = "https://your-app-name.streamlit.app/"
```

7. Deployを押します。
8. 発行されたURLをSupabaseのRedirect URLsにも追加します。

## 3. モデレーター2名に共有する

1. 先生がアプリにGoogleログインします。
2. サイドバーの「管理画面: 許可ユーザー」を開きます。
3. モデレーター2名のGoogleメールアドレスを追加します。
4. roleは通常 `moderator` にします。
5. アプリURLを2名に共有します。

共有文の例:

```text
PerioBoxスライド生成アプリのURLです。
Googleログインで入ってください。

URL:
https://your-app-name.streamlit.app/

登録済みのGoogleメールアドレスでのみ利用できます。
PDFをアップロードすると、PowerPointと全文日本語訳Markdownが生成されます。
```

## 4. Render / Railway / Cloud Runに置く場合

Streamlit Community Cloudで負荷や安定性に不安が出たら、次の移行先が候補です。

- Render: 比較的簡単。Web Serviceとして `streamlit run app.py --server.port $PORT --server.address 0.0.0.0` を起動
- Railway: 環境変数管理が簡単。小規模検証向き
- Google Cloud Run: 最も本格運用向き。Docker化が必要

いずれの場合も設定する環境変数は同じです。

```text
OPENAI_API_KEY
OPENAI_MODEL
SUPABASE_URL
SUPABASE_ANON_KEY
APP_BASE_URL
```

`APP_BASE_URL` は公開URLに合わせ、SupabaseのRedirect URLsにも同じURLを登録してください。

## 5. Netlifyを使う場合

Netlify単体では、このPython Streamlitアプリをそのまま常時起動する用途には向きません。

Netlifyを使う場合は、将来的に次の構成にします。

- Netlify: ログイン済みユーザー向けポータルまたはReact/Next.jsフロントエンド
- Supabase Auth: Googleログイン
- Python API: Render / Railway / Cloud Run
- PPTX生成: Python API側

まずはStreamlit Cloudでモデレーター運用を開始し、実際の使われ方が固まってからNetlifyフロントエンド化するのが安全です。

## 6. 運用上の注意

- OpenAI APIキーはモデレーターに共有しません。
- モデレーターを外す場合は、admin画面でメールアドレスを削除します。
- `admin` は最小人数にしてください。
- 生成物は医療判断の最終資料ではなく、必ず原著PDFで確認してください。
- PDFや論文データの取り扱いは、所属組織・出版社・学会のルールに従ってください。
