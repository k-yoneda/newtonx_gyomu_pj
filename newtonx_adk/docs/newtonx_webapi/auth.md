# 認証仕様（Host + PAT / NewtonX Web 発行トークン）

## 概要
- NewtonX CLI/ADK は、NewtonX Web で発行した Personal Access Token（PAT）を使用します。
- Host-first 方針により、サブドメインから `api_base_url` と `company_subdomain` を自動導出します。

## 設定項目
- `host`（推奨）: 例 `seraku.newton-x.net`
- `personal_access_token`（必須）: NewtonX Web で発行した PAT
- `api_base_url`（任意）: 未指定時は `https://{host}/api`
- `company_subdomain`（任意）: 未指定時は `host` の第一ラベル
- `token_file`（任意）: 例 `~/.newtonx/tokens.json`

## フロー
1. セットアップ: `tools/setup_config.py` を実行し、Company Subdomain と PAT を登録
2. 設定確認: `tools/check_config.py` で `Authorization: Bearer <PAT>` 付与を確認
3. API 呼び出し: ADK が自動で `Authorization` ヘッダーを付与

## 参考
- PAT の発行は NewtonX Web の「アクセストークン」メニューから行います
- Host-first: `host` を設定すると `api_base_url` と `company_subdomain` を自動導出

## NewtonX API 呼び出し
- ADK は `Authorization: Bearer <PAT>` を自動付与します

## ログとセキュリティ
- PAT をログやリポジトリに残さない（ハードコード禁止）
- トークン/設定ファイルは OS 権限で保護（既定: `~/.newtonx/`）

## 例（CLI）
```bash
PYTHONPATH=./src python tools/setup_config.py   # Subdomain と PAT 登録
PYTHONPATH=./src python tools/check_config.py   # Authorization を確認
```

## 例（ADK）
```python
from newtonx_adk import ConfigManager, NewtonXClient

cm = ConfigManager()
client = NewtonXClient(cm)

print(client.get_user_info())
```

## エラーハンドリング
- `401 Unauthorized`: PAT を再登録（`tools/setup_config.py`）→ `tools/check_config.py` で確認
- `403 Forbidden`: 権限またはIP制限を確認
