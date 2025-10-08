# デプロイメント手順

## 推奨デプロイメントプラットフォーム

Streamlitアプリケーションは継続的に実行されるサーバーが必要なため、以下のプラットフォームが推奨されます：

### 1. Streamlit Cloud（推奨）
- **URL**: https://share.streamlit.io/
- **特徴**: Streamlit専用のホスティングサービス
- **料金**: 無料プランあり
- **手順**:
  1. GitHubリポジトリにコードをプッシュ
  2. Streamlit Cloudでリポジトリを選択
  3. 自動デプロイ

### 2. Heroku
- **URL**: https://heroku.com/
- **特徴**: 汎用クラウドプラットフォーム
- **料金**: 有料（無料プランは廃止）
- **手順**:
  1. Heroku CLIをインストール
  2. `heroku create your-app-name`
  3. `git push heroku main`

### 3. Railway
- **URL**: https://railway.app/
- **特徴**: モダンなデプロイメントプラットフォーム
- **料金**: 無料プランあり
- **手順**:
  1. GitHubリポジトリを接続
  2. 自動デプロイ

### 4. Render
- **URL**: https://render.com/
- **特徴**: シンプルなデプロイメント
- **料金**: 無料プランあり
- **手順**:
  1. GitHubリポジトリを接続
  2. Web Serviceを選択
  3. 設定を入力

## 現在のディレクトリ構成でのデプロイ

現在のディレクトリ構成は以下のようになっています：

```
/Users/yuji/t_kento/
├── app.py              # メインアプリケーション
├── requirements.txt    # 依存関係
├── Procfile           # Heroku用設定
├── runtime.txt        # Pythonバージョン指定
├── netlify.toml       # Netlify用設定（Streamlitには非推奨）
├── README.md          # 使用方法
└── venv/              # 仮想環境（デプロイ時は不要）
```

## 各プラットフォームでの設定

### Streamlit Cloud
- 追加設定不要
- GitHubリポジトリを接続するだけ

### Heroku
- `Procfile`が使用される
- `runtime.txt`でPythonバージョンを指定

### Railway
- `Procfile`が使用される
- 自動的にPython環境を検出

### Render
- `Procfile`が使用される
- 環境変数で設定可能

## 注意事項

- **Netlifyは非推奨**: Netlifyは静的サイト用のため、Streamlitアプリケーションには適していません
- **継続実行**: Streamlitアプリケーションは継続的に実行される必要があります
- **メモリ使用量**: 大容量のエクセルファイルを処理する場合、メモリ制限に注意してください

## 推奨デプロイメント手順

1. **Streamlit Cloud**を使用することを強く推奨します
2. GitHubリポジトリにコードをプッシュ
3. Streamlit Cloudでリポジトリを選択
4. 自動デプロイ完了
