# NewtonX ADK Examples

NewtonX ADKの実装例集です。各実装例は独立したフォルダに配置され、独自の`requirements.txt`とドキュメントを持っています。

## ディレクトリ構造

```
cli/
├── newtonx_adk/          # ADK本体
│   ├── __init__.py
│   ├── client.py
│   ├── auth.py
│   ├── config.py
│   ├── exceptions.py
│   └── ...
└── adk_examples/         # 実装例（ADKと並列）
    ├── cli_example/      # CLIアプリケーション例
    ├── expense_processor/ # 経費処理例
    └── auto_runner/      # 自動実行例
```

## 実装例一覧

### 1. CLI Example (`cli_example/`)
タブキー補完機能付きのCLIアプリケーション実装例です。

**特徴:**
- ファイルパス入力時のタブキー補完
- ファイルアップロード機能
- リアルタイムチャット
- コマンド履歴保存

**使用方法:**
```bash
cd adk_examples/cli_example
pip install -r requirements.txt
python cli_example.py
```

### 2. Expense Processor Example (`expense_processor/`)
経費精算処理の実装例です。

**特徴:**
- 画像ファイル（PNG、JPEG）の自動処理
- AI分析による経費データ分類
- レポート生成（CSV形式）
- データ検証

**使用方法:**
```bash
cd adk_examples/expense_processor
pip install -r requirements.txt
python expense_processor_example.py input_dir output.csv
```

**サンプルデータ:**
- `sample_data/receipt1.png` - コンビニ領収書
- `sample_data/recept2.jpeg` - タクシー領収書

### 3. Auto Runner Example (`auto_runner/`)
YAML設定ファイルに従って自動実行するサンプルです。CLI版の`cli_runner.py`に近い実装です。

**特徴:**
- 設定ファイルベースの自動実行
- 複数タスクの順次実行
- 進捗表示と結果保存
- ログ機能

**使用方法:**
```bash
cd adk_examples/auto_runner
pip install -r requirements.txt
python auto_runner.py --config config.yml
```

## 共通特徴

### 独立した環境
各実装例は独立したフォルダに配置され、独自の依存関係を持っています：

```
adk_examples/
├── cli_example/
│   ├── cli_example.py
│   ├── requirements.txt
│   ├── README.md
│   └── upload_sample.txt
├── expense_processor/
│   ├── expense_processor_example.py
│   ├── requirements.txt
│   ├── README.md
│   └── sample_data/
│       ├── receipt1.png
│       └── recept2.jpeg
└── auto_runner/
    ├── auto_runner.py
    ├── config.yml
    ├── requirements.txt
    ├── README.md
    └── sample_document.txt
```

### 認証方式（公開クライアント / MSAL, PKCE優先）
- すべての実装例は公開クライアント（MSAL）を使用します。
- 実行前に以下でEntra ID設定を登録してください:
  ```bash
  newtonx-config configure
  # または
  python -m newtonx_adk.config_cli configure
  ```
- 実行時はサイレント→PKCE（ブラウザ）→デバイスコードの順に認証します。
- クライアントシークレットは不要です。

### 必要なモジュールの配置
各実装例に必要なモジュールは、そのフォルダ内の`requirements.txt`に記載されています。

### サンプルファイル
各実装例に必要なサンプルファイルは、そのフォルダ内に配置されています。

## 開発ガイドライン

### 新しい実装例の追加
1. `adk_examples/`フォルダ内に新しいフォルダを作成
2. `requirements.txt`を作成
3. `README.md`を作成
4. サンプルファイルを配置
5. メインのREADME.mdを更新

### 実装例の構造
```
adk_examples/example_name/
├── main_script.py      # メインスクリプト
├── requirements.txt    # 依存関係
├── README.md          # ドキュメント
├── config.yml         # 設定ファイル（必要に応じて）
└── sample_files/      # サンプルファイル（必要に応じて）
```

## トラブルシューティング

### ADKのインポートエラー
各実装例は、ADKのパスを自動的に設定します。エラーが発生した場合は、以下を確認してください：

1. ADKが正しくインストールされているか
2. `requirements.txt`の依存関係がインストールされているか
3. Pythonのバージョンが3.8以上か

### 認証エラー
NewtonX APIの認証が必要です。各実装例のREADME.mdを参照してください。

### ファイルアップロードエラー
- ファイルサイズ制限（10MB以下）
- サポートされているファイル形式を確認
- ネットワーク接続を確認

## テスト結果

### CLI Example
- ✅ 認証機能
- ✅ アシスタント一覧表示
- ✅ チャット作成・管理
- ✅ ファイルアップロード
- ✅ メッセージ送信

### Expense Processor
- ✅ 画像ファイルアップロード
- ✅ AI分析による経費情報抽出
- ✅ CSV形式での結果出力
- ✅ 複数ファイルの一括処理

### Auto Runner
- ✅ YAML設定ファイルの読み込み
- ✅ 複数タスクの順次実行
- ✅ 進捗表示と結果保存
- ✅ エラーハンドリング

## ライセンス

各実装例は、NewtonX ADKと同じライセンスに従います。 