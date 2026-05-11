# Expense Processor Example

NewtonX ADKを使用した経費精算処理の実装例です。

## 機能

- **CSV/Excelファイル処理**: 経費データの読み込み・処理
- **AI分析**: NewtonX APIを使用した経費データの分析
- **レポート生成**: 処理結果のレポート生成
- **データ検証**: 経費データの自動検証

## インストール

```bash
cd expense_processor
pip install -r requirements.txt
```

## 使用方法

```bash
python expense_processor_example.py input_dir output.csv
```

### 事前準備（認証設定）
内部トークン登録フローに統一しました。
```bash
python tools/setup_config.py
# 認証（PAT）: WEBアプリでPAT発行 → セットアップに入力

# クリップボードからの自動保存
PYTHONPATH=./src python tools/save_token_from_clipboard.py
```

## 特徴

### データ処理
- CSV/Excelファイルの読み込み
- 経費データの自動分類
- 異常値の検出
- 集計レポートの生成

### AI分析
- NewtonX APIを使用した経費分析
- カテゴリ別分類の自動化
- コスト最適化の提案

### レポート機能
- 月次・年次レポート
- 部門別集計
- 予算対実績分析

## ファイル構成

```
expense_processor/
├── expense_processor_example.py  # メインアプリケーション
├── requirements.txt              # 依存関係
├── README.md                    # このファイル
└── sample_data/                 # サンプルデータ（必要に応じて追加）
```

## 画像付きの領収書解析フロー（新）

画像ファイルは `send_message_with_images` を使うと、アップロードとメッセージ送信を一括で実行できます。

```python
response = client.send_message_with_images(
    chat_uid=chat_uid,
    message="この領収書から日付・金額・T番号有無・目的・支払い先をJSONで返して",
    image_file_paths=["/path/to/receipt.jpg"],
    knowledge_search=False,
    web_search=False,
)
```

`expense_processor_example.py` では、拡張子が画像（.jpg/.jpeg/.png/.gif）の場合はこの新メソッドを使い、PDFなどのドキュメントは従来通り `upload_document` → `send_message` の順で処理します。 