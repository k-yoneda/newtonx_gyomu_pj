# NewtonX エージェント開発キット (NewtonX ADK)

**NewtonX（ニュートンX）のAIアシスタントと簡単にやり取りできるPythonライブラリです**

このライブラリを使うと、PythonプログラムからNewtonXのAIアシスタントとチャットしたり、ファイルを送ったりできます。

## 🚀 まずは使ってみよう

### ステップ1: ライブラリをインストール

```bash
pip install newtonx-adk
```

### ステップ2: 認証設定（PAT）

NewtonX のWEBアプリで発行した Personal Access Token（PAT）を使用します。

```bash
PYTHONPATH=./src python tools/setup_config.py
# Company Subdomain（例: seraku）と PAT を入力

# 設定確認（Authorizationヘッダーの有無を確認）
PYTHONPATH=./src python tools/check_config.py
```

#### セキュリティ上の注意（重要）

- ADKやスクリプトに PAT をハードコードしないでください。
- トークンは個人の権限に直結します。漏えいや共有は厳禁です。
- 認証情報は `~/.newtonx/tokens.json` に保存され、ADKが自動で読み込みます。

### ステップ3: 基本設定を確認する（任意）

```bash
# 現在の設定を表示（JSON）
PYTHONPATH=./src python tools/check_config.py

# 必要に応じてホスト名のみ再設定（/api と subdomain は自動導出）
newtonx-config set --host "seraku.newton-x.net"
```


### ステップ4: 実際に使ってみる

```python
from newtonx_adk import NewtonXClient, ConfigManager

# 設定を読み込んでクライアントを作成
config = ConfigManager()
client = NewtonXClient(config)

# 認証: PAT が設定されていれば追加操作は不要

# ユーザー情報を確認
user_info = client.get_user_info()
print(f"ログインユーザー: {user_info.get('name', 'N/A')}")

# アシスタント一覧を取得
assistants = client.get_assistants()
print(f"利用可能なアシスタント: {len(assistants)}個")

# チャットを開始
chat_uid = client.create_chat(
    assistant_uid=assistants[0]['uid'],  # 最初のアシスタントを選択
    title="テストチャット"
)

# メッセージを送信
response = client.send_message(
    chat_uid=chat_uid,
    message="こんにちは！"
)

print(f"アシスタントの返事: {response}")
```

## 📋 できること

### 🤖 AIアシスタントとのチャット
- 複数のアシスタントから選んでチャット
- ウェブ検索やナレッジ検索を有効にできる
- チャットのタイトルを変更可能

### 📁 ファイルの送信
- 画像ファイル（JPG、PNG、GIFなど）
- ドキュメント（PDF、Word、テキストなど）
- ファイルサイズは10MBまで

### 🔐 認証の考え方
- **PAT ベース**: Web アプリで発行した個人用トークン（PAT）を使用
- **Host-first**: サブドメインから `api_base_url` などを自動導出
- トークンはローカルに安全に保存（`~/.newtonx/config.json`, `~/.newtonx/tokens.json`）

## 🛠️ よくある使い方

### 1. アシスタント一覧を見る

```python
assistants = client.get_assistants()
for assistant in assistants:
    print(f"名前: {assistant['name']}")
    print(f"説明: {assistant.get('description', '説明なし')}")
    print("---")
```

### 2. チャットを始める

```python
# チャットを作成
chat_uid = client.create_chat(
    assistant_uid="アシスタントのID",
    title="私のチャット"
)

# メッセージを送る
response = client.send_message(
    chat_uid=chat_uid,
    message="今日の天気はどうですか？",
    web_search=True  # ウェブ検索を有効にする
)
```

### 3. ファイルを送る

```python
# 画像を送る
image_id = client.upload_image(
    chat_uid=chat_uid,
    file_path="写真.jpg"
)

# ドキュメントを送る
success = client.upload_document(
    chat_uid=chat_uid,
    file_path="資料.pdf"
)
```

## ❗ 困ったときは

### コマンドが動かない場合

```bash
# 代替の実行方法（環境によっては有効です）
python -m newtonx_adk.config_cli --help
python -m newtonx_adk.config_cli show
```

### 認証でエラーが出る場合

1. **認証設定**:
   - `PYTHONPATH=./src python tools/setup_config.py` を実行し、Company Subdomain と PAT を設定
   - `PYTHONPATH=./src python tools/check_config.py` で Authorization ヘッダーの有無を確認

2. **その他のエラー**:
   - ブラウザが開かない: 手動で認証URLを開いてください
   - 「無効なクライアント」エラー: 認証設定を確認してください
   - 404エラー: ホスト名に `/api` が付いているか確認してください
   - 403エラー: IP制限の可能性があります

### 設定を確認する

```bash
# 現在の設定を見る
newtonx-config show

# 設定をリセットする
newtonx-config reset

# 認証テストを実行する
python tools/test_user_specific_auth.py  # ユーザー固有認証
```

## 📚 もっと詳しく知りたい方へ

- **APIリファレンス**: [API_REFERENCE.md](docs/adk/API_REFERENCE.md)
- **使用ガイド**: [USAGE_GUIDE.md](docs/adk/USAGE_GUIDE.md)
- **実装例**: [adk_examples](adk_examples/) フォルダ

## 🔧 開発者向け情報

### エラーハンドリング

```python
from newtonx_adk import (
    NewtonXError,           # 基本エラー
    AuthenticationError,    # 認証エラー
    APIError,              # APIエラー
    ConfigurationError,    # 設定エラー
    FileUploadError,       # ファイルアップロードエラー
    ChatError              # チャットエラー
)

try:
    client.get_assistants()
except AuthenticationError:
    print("認証が必要です")
except APIError as e:
    print(f"APIエラー: {e}")
```

### 設定ファイルの場所を環境変数で指定（任意）

```bash
export NEWTONX_CONFIG_PATH="$HOME/.newtonx/config.json"
```

## 📄 ライセンス

MIT License

## 🆘 サポート

- 問題が起きたら、まずは `newtonx-config show` で設定を確認してください
- 認証テストを実行して問題を特定してください
- それでも解決しない場合は、エラーメッセージと一緒にご連絡ください

## 🔄 更新履歴
### v0.10.5 (最新)
- **バージョン更新**: パッケージバージョンを 0.10.5 に更新
- **User-Agent更新**: 内部HTTPの `User-Agent` を `NewtonX-ADK/0.10.5` に更新
- **ドキュメント整備**: ドキュメント内のバージョン表記を最新化

### v0.10.4
- **バージョン更新**: パッケージバージョンを 0.10.4 に更新
- **User-Agent更新**: 内部HTTPの `User-Agent` を `NewtonX-ADK/0.10.4` に更新
- **音声アップデート対応**: Gemini系アシスタントで利用できる音声ファイルのアップロードに対応
- **API仕様変更対応**: 2025.11に行われるアプリケーション基盤の変更に伴うAPI仕様変更への対応

### v0.10.3
- **オブジェクトID認証を廃止**: ユーザー固有認証のみをサポート
- **認証フローの簡素化**: よりシンプルで分かりやすい認証プロセス
- **コードの整理**: 不要な認証方式のコードを削除してメンテナンス性を向上

### v0.10.2
- **ユーザー固有認証を追加**: 各ユーザーが自分のEntraIDアカウントで認証可能
- **認証方式の選択**: ユーザー固有認証（推奨）とオブジェクトID認証（管理者用）をサポート
- **認証フローの改善**: より明確な認証メッセージとエラーハンドリング

---

**NewtonX ADKで、AIアシスタントとの新しい体験を始めましょう！** 🎉