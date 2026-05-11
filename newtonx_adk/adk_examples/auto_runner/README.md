# NewtonX Auto Runner Example

YAML設定ファイルに従って自動実行するサンプルです。CLI版の`cli_runner.py`に近い実装です。

## 機能

- **設定ファイルベース**: YAML設定ファイルによる自動実行
- **タスク定義**: 複数のタスクを順次実行
- **進捗表示**: Richライブラリによる美しい進捗表示
- **結果保存**: 実行結果の自動保存
- **ログ機能**: 詳細なログ出力

## インストール

```bash
cd auto_runner
pip install -r requirements.txt
```

## 使用方法

### 基本的な実行
```bash
python auto_runner.py
```

### 設定ファイルを指定
```bash
python auto_runner.py --config my_config.yml
```

### ドライラン（実際の実行は行わない）
```bash
python auto_runner.py --dry-run
```

### 詳細出力
```bash
python auto_runner.py --verbose
```

### 事前準備（認証設定: PAT）

```bash
PYTHONPATH=./src python tools/setup_config.py   # Company Subdomain と PAT を設定
PYTHONPATH=./src python tools/check_config.py   # Authorization ヘッダーの有無を確認
```

## 設定ファイル

`config.yml`ファイルで実行内容を定義します：

```yaml
project:
  name: "NewtonX Auto Runner Example"
  description: "設定ファイルベースの自動実行サンプル"
  version: "1.0.0"

# 自動実行タスク
tasks:
  - name: "初期化"
    type: "init"
    description: "プロジェクトの初期化"
    enabled: true
    
  - name: "ファイルアップロード"
    type: "upload"
    description: "サンプルファイルのアップロード"
    files:
      - "sample_document.txt"
    enabled: true
    
  - name: "チャット作成"
    type: "create_chat"
    description: "新しいチャットの作成"
    title: "Auto Runner Test Chat"
    assistant: "高速アシスタント"
    enabled: true
    
  - name: "メッセージ送信"
    type: "send_message"
    description: "テストメッセージの送信"
    message: "こんにちは！Auto Runnerのテストです。"
    web_search: false
    knowledge_search: false
    enabled: true
```

## 対応タスクタイプ

### init
プロジェクトの初期化を行います。

### upload
ファイルをアップロードします。
- `files`: アップロードするファイルのリスト

### create_chat
新しいチャットを作成します。
- `title`: チャットタイトル
- `assistant`: アシスタント名

### send_message
メッセージを送信します。
- `message`: 送信するメッセージ
- `web_search`: WEB検索の有無
- `knowledge_search`: ナレッジ検索の有無

### analyze_file
アップロードしたファイルを分析します。
- `message`: 分析指示メッセージ

## 出力

実行結果は以下の形式で保存されます：

- **Rich形式**: コンソールに美しく表示
- **JSON形式**: 構造化されたデータ
- **テキスト形式**: 読みやすいテキスト

## ファイル構成

```
auto_runner/
├── auto_runner.py        # メインアプリケーション
├── config.yml            # 設定ファイル
├── requirements.txt      # 依存関係
├── README.md            # このファイル
├── sample_document.txt  # サンプルドキュメント
└── results/             # 実行結果保存ディレクトリ
```

## 特徴

### 設定ファイルベース
- YAML形式の設定ファイル
- タスクの有効/無効切り替え
- 柔軟なパラメータ設定

### 進捗表示
- Richライブラリによる美しい表示
- リアルタイム進捗表示
- エラー状態の視覚化

### 結果管理
- 実行結果の自動保存
- 複数形式での出力
- タイムスタンプ付き保存

### エラーハンドリング
- 詳細なログ出力
- エラー状態の記録
- 部分的な失敗対応 