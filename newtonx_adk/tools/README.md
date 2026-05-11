# NewtonX ADK Tools

NewtonX ADKの設定と管理を行うツール集です。

## ツール一覧

### 設定関連

- **`setup_config.py`** - 認証設定を行う
- **`check_config.py`** - 認証設定の確認とテスト
- **`clear_config.py`** - 認証設定をクリア
  (デバイスコード認証へ統一のため、関連ツールは削除されました)

## 使用方法

### 1. 認証設定

```bash
python tools/setup_config.py
```

初回使用時や設定を変更したい場合に実行します。

### 2. 設定確認

```bash
python tools/check_config.py
```

現在の認証設定を確認し、認証テストを実行します。

### 3. 設定クリア

```bash
python tools/clear_config.py
```

認証設定を完全にクリアします。次回使用時は新しく設定が必要です。

### 4. 認証について（PAT）

現在、認証は Personal Access Token（PAT）方式をサポートします。ブラウザ拡張は不要です。

セットアップと確認:
```bash
PYTHONPATH=./src python tools/setup_config.py
PYTHONPATH=./src python tools/check_config.py
```

### 5. 設定クリア

```bash
PYTHONPATH=./src python tools/clear_config.py
```

## Azure CLI 環境構築手順（azure_cli モード用）

NewtonX ADKの認証モードで `azure_cli` を選ぶ場合、ローカルに Azure CLI が必要です。

- macOS
  - Homebrew を使用
  ```bash
  brew update
  brew install azure-cli
  ```

- Windows
  - winget を使用
  ```powershell
  winget install Microsoft.AzureCLI
  ```
  - もしくはインストーラー（MSI）からインストール（Microsoft公式ドキュメント参照）

- Linux（Debian/Ubuntu系の簡易手順）
  ```bash
  curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash
  ```
  - その他ディストリビューションは公式手順を参照

### セットアップと確認
- バージョン確認
```bash
az version
```

- サインイン（標準）
```bash
az login
```

- 特定テナントでサインイン（推奨）
```bash
az login --tenant <YOUR_TENANT_ID>
```

- テナント/サブスクリプション選択（必要に応じて）
```bash
az account tenant list
az account set --tenant <YOUR_TENANT_ID>
az account show
```

- トークン取得テスト（疎通確認）
```bash
az account get-access-token --scope https://graph.microsoft.com/.default --output json
```

### 運用ヒント
- サインアウトする場合
```bash
az logout
```
- 複数アカウントが混在する場合は `az account clear` でキャッシュを整理
- 企業プロキシ環境では Azure CLI のプロキシ設定を合わせて構成してください

### ADK側の動作
- `azure_cli` モード時、ADKは内部的に `az account get-access-token --scope <scopes>` を実行してBearerを取得します
- NewtonX API 呼び出しには内部トークン＋Cookie＋XSRF も必要なため、初回は `auth_callback` の取り込み→ブートストラップが別途必要です

## 設定ファイル

- **設定ファイル**: `~/.newtonx/config.json`
- **トークンファイル**: `~/.newtonx/tokens.json`

## 認証フロー

1. **EntraID認証**: Microsoft Entra IDでユーザー認証
2. **NewtonXセッション確立**: 認証成功後、自動的にCookieセッションを確立
3. **API呼び出し**: Cookieセッション認証でNewtonX APIにアクセス