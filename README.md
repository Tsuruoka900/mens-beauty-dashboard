# メンズビューティ販売分析ダッシュボード

## デプロイ手順（Streamlit Community Cloud）

### 1. GitHubリポジトリ作成
```bash
git init
git add .
git commit -m "initial commit"
git remote add origin https://github.com/<ユーザー名>/<リポジトリ名>.git
git push -u origin main
```

### 2. Streamlit Community Cloudでデプロイ
1. https://share.streamlit.io にアクセス
2. "New app" → GitHubリポジトリを選択
3. Branch: `main` / Main file path: `app.py`
4. "Deploy" をクリック

## ローカル実行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 入力ファイル仕様

| # | 種類 | 形式 | エンコード |
|---|------|------|-----------|
| ① | IDPOS CSV | CSV | cp932 |
| ② | SRI Excel | xlsx | — |
| ③ | マスタ CSV | CSV | cp932 |

## 期の定義

| 期 | 期間 |
|----|------|
| 46期 | 2024年7月〜2025年6月 |
| 47期 | 2025年7月〜2026年6月 |
