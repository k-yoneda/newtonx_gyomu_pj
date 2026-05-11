# NewtonX 認証設定ガイド（PAT 方式・最新版）

## 概要

NewtonX ADK は Personal Access Token（PAT）による簡易認証を採用しています。NewtonX の Web 版から PAT を発行し、セットアップツールで Host と PAT を入力するだけで完了します。

## 手順

### 1. NewtonX Web 版で PAT を発行

1. NewtonX の Web 版にログイン
2. 画面右上のアカウントから「アクセストークン」のメニューを開く
3. 「発行」ボタンを押して名前をつけてPATを新規発行
4. 表示された PAT をクリップボードにコピーボタンを使ってコピー

### 2. セットアップツールを実行し、Host と PAT を登録

```bash
cd newtonx_adk-main
python tools/setup_config.py
```

- 入力項目
  - Company Subdomain（例: seraku）
    - 入力すると Host は自動で `<subdomain>.newton-x.net` に設定されます
  - Personal Access Token（PAT）
    - Web 版で発行した PAT を貼り付け

実行後、設定ファイル（`~/.newtonx/config.json`）に Host と PAT が保存され、以後 ADK は `Authorization: Bearer <PAT>` を自動付与します。

### 3. 設定を確認（任意）

```bash
python tools/check_config.py
```

- 期待される出力（抜粋）
  - Host や CompanySubdomain が表示
  - Authorization に基づくヘッダー生成を確認できれば OK

## トラブルシューティング

1. PAT を紛失した
   - Web 版で再発行してください（既存の PAT は再表示不可の場合があります）
2. Authorization ヘッダーが付かない
   - `tools/setup_config.py` を再実行し、PAT を貼り付け直してください
3. ホスト名が分からない
   - 会社のサブドメイン（例: `seraku`）を入力すれば `seraku.newton-x.net` が自動で設定されます

## 参考

- ADK の設定保存先: `~/.newtonx/config.json`
- セットアップツール: `tools/setup_config.py`
- 設定確認ツール: `tools/check_config.py`
