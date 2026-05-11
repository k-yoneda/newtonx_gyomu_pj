NewtonX エージェント開発キット (NewtonX ADK) 開発・運用知見集

# NewtonX ADK v0.10.5 技術文書

## 概要
NewtonX ADKは、NewtonX APIを簡単に利用できるPythonライブラリです。認証、チャット、アシスタント、ファイルアップロードなどの機能を提供します。

## アーキテクチャ

### 主要コンポーネント
1. **NewtonXClient** - メインクライアントクラス
2. **ConfigManager** - 設定管理クラス
3. **AuthManager** - 認証管理クラス
4. **例外クラス群** - エラーハンドリング

### ファイル構成
```
newtonx_adk/
├── __init__.py      # ADK初期化
├── client.py        # メインクライアント
├── auth.py          # 認証管理
├── config.py        # 設定管理
├── exceptions.py    # 例外クラス
├── setup.py         # インストール設定
├── pyproject.toml   # プロジェクト設定
├── requirements.txt # 依存関係
└── README.md        # ドキュメント
```

## ADKクラス・メソッド詳細

### 1. NewtonXClient クラス

#### 初期化
```python
class NewtonXClient:
    def __init__(self, config_manager: Optional[ConfigManager] = None):
        """
        NewtonXクライアントの初期化
        
        Args:
            config_manager: 設定管理クラスのインスタンス（オプション）
        """
```

#### 認証メソッド
```python
def authenticate(self) -> bool:
    """
    認証を実行
    
    Returns:
        bool: 認証成功時True、失敗時False
    """
```

#### ユーザー情報関連メソッド
```python
def get_user_info(self) -> Optional[Dict]:
    """
    ユーザー情報を取得
    
    Returns:
        Dict: ユーザー情報の辞書、失敗時None
    """

def get_assistants(self) -> List[Dict]:
    """
    アシスタント一覧を取得
    
    Returns:
        List[Dict]: アシスタント情報のリスト
    """
```

#### チャット関連メソッド
```python
def create_chat(self, assistant_uid: str, title: Optional[str] = None, 
                folder_uid: Optional[str] = None) -> Optional[str]:
    """
    チャットを作成
    
    Args:
        assistant_uid: アシスタントのUID
        title: チャットタイトル（オプション）
        folder_uid: フォルダUID（オプション）
    
    Returns:
        str: 作成されたチャットのUID、失敗時None
    """

def send_message(self, chat_uid: str, message: str, knowledge_search: bool = False, 
                web_search: bool = True, image_ids: Optional[List[str]] = None, 
                document_ids: Optional[List[str]] = None, parent_order: int = 0) -> Optional[str]:
    """
    メッセージを送信
    
    Args:
        chat_uid: チャットのUID
        message: 送信メッセージ
        knowledge_search: ナレッジ検索の有効化
        web_search: ウェブ検索の有効化
        image_ids: 画像IDリスト（オプション）
        document_ids: ドキュメントIDリスト（オプション）
        parent_order: 親メッセージの順序（デフォルト: 0）
    
    Returns:
        str: アシスタントの応答、失敗時None
    """

def get_chat(self, chat_uid: str) -> Optional[Dict]:
    """
    チャット情報を取得
    
    Args:
        chat_uid: チャットのUID
    
    Returns:
        Dict: チャット情報の辞書、失敗時None
    """

def delete_chat(self, chat_uid: str) -> bool:
    """
    チャットを削除
    
    Args:
        chat_uid: チャットのUID
    
    Returns:
        bool: 削除成功時True、失敗時False
    """

def update_chat_title(self, chat_uid: str, title: str) -> bool:
    """
    チャットタイトルを更新
    
    Args:
        chat_uid: チャットのUID
        title: 新しいタイトル
    
    Returns:
        bool: 更新成功時True、失敗時False
    """
```

#### ファイルアップロード関連メソッド
```python
def upload_image(self, chat_uid: str, file_path: str, 
                file_name: Optional[str] = None) -> Optional[str]:
    """
    画像をアップロード
    
    Args:
        chat_uid: チャットのUID
        file_path: 画像ファイルのパス
        file_name: ファイル名（オプション）
    
    Returns:
        str: 画像ID、失敗時None
        
    Raises:
        FileUploadError: アップロード失敗時
    """

def upload_document(self, chat_uid: str, file_path: str, 
                   file_name: Optional[str] = None) -> Optional[str]:
    """
    ドキュメントをアップロード
    
    Args:
        chat_uid: チャットのUID
        file_path: ドキュメントファイルのパス
        file_name: ファイル名（オプション）
    
    Returns:
        str: ドキュメントID、失敗時None
        
    Raises:
        FileUploadError: アップロード失敗時
    """
```

#### フォルダ関連メソッド
```python
def create_folder(self, name: str) -> Optional[str]:
    """
    フォルダを作成
    
    Args:
        name: フォルダ名
    
    Returns:
        str: 作成されたフォルダのUID、失敗時None
    """

def get_folders(self) -> List[Dict]:
    """
    フォルダ一覧を取得
    
    Returns:
        List[Dict]: フォルダ情報のリスト
    """

def move_chat_to_folder(self, chat_uid: str, folder_uid: str) -> bool:
    """
    チャットをフォルダに移動
    
    Args:
        chat_uid: チャットのUID
        folder_uid: フォルダのUID
    
    Returns:
        bool: 移動成功時True、失敗時False
    """

def get_folder_chats(self, folder_uid: str) -> List[Dict]:
    """
    フォルダ内のチャット一覧を取得
    
    Args:
        folder_uid: フォルダのUID
    
    Returns:
        List[Dict]: チャット情報のリスト
    """
```

#### システム情報関連メソッド
```python
def get_model_status(self) -> Dict[str, bool]:
    """
    モデルステータスを取得
    
    Returns:
        Dict[str, bool]: モデルステータスの辞書
    """

def get_company_info(self) -> Optional[Dict]:
    """
    会社情報を取得
    
    Returns:
        Dict: 会社情報の辞書、失敗時None
    """
```

### 2. ConfigManager クラス

#### 初期化
```python
class ConfigManager:
    def __init__(self, config_file: Optional[str] = None):
        """
        設定管理クラスの初期化
        
        Args:
            config_file: 設定ファイルのパス（オプション）
        """
```

#### 設定関連メソッド
```python
def get_config(self) -> ADKConfig:
    """
    設定を取得
    
    Returns:
        ADKConfig: 設定オブジェクト
    """

def set_credentials(self, client_id: str, client_secret: str) -> None:
    """
    認証情報を設定
    
    Args:
        client_id: クライアントID
        client_secret: クライアントシークレット
    """

def save_config(self) -> None:
    """
    設定を保存
    """

def load_config(self) -> None:
    """
    設定を読み込み
    """
```

### 3. AuthManager クラス

#### 初期化
```python
class AuthManager:
    def __init__(self, config_manager: ConfigManager):
        """
        認証管理クラスの初期化
        
        Args:
            config_manager: 設定管理クラスのインスタンス
        """
```

#### 認証関連メソッド
```python
def authenticate(self) -> bool:
    """
    手動認証を実行
    
    Returns:
        bool: 認証成功時True、失敗時False
    """

def authenticate_auto(self) -> bool:
    """
    自動認証を実行（保存されたトークンを使用）
    
    Returns:
        bool: 認証成功時True、失敗時False
    """

def is_authenticated(self) -> bool:
    """
    認証状態を確認
    
    Returns:
        bool: 認証済みの場合True
    """

def refresh_token(self) -> bool:
    """
    トークンを更新
    
    Returns:
        bool: 更新成功時True、失敗時False
    """

def _get_headers(self) -> Dict[str, str]:
    """
    認証ヘッダーを取得
    
    Returns:
        Dict[str, str]: 認証ヘッダーの辞書
    """
```

### 4. 例外クラス

#### NewtonXError
```python
class NewtonXError(Exception):
    """NewtonX ADKの基本例外クラス"""
```

#### AuthenticationError
```python
class AuthenticationError(NewtonXError):
    """認証関連の例外"""
```

#### APIError
```python
class APIError(NewtonXError):
    """API通信関連の例外"""
```

#### ConfigurationError
```python
class ConfigurationError(NewtonXError):
    """設定関連の例外"""
```

#### FileUploadError
```python
class FileUploadError(NewtonXError):
    """ファイルアップロード関連の例外"""
```

#### ChatError
```python
class ChatError(NewtonXError):
    """チャット関連の例外"""
```

## サンプルアプリケーション詳細

### 1. Auto Runner サンプル

#### ファイル構成
```
adk_examples/auto_runner/
├── auto_runner.py      # メインスクリプト
├── config.yml          # 設定ファイル
├── sample_document.txt # サンプルドキュメント
├── requirements.txt    # 依存関係
└── README.md          # ドキュメント
```

#### AutoRunner クラス

##### 初期化
```python
class AutoRunner:
    def __init__(self, config_path: str):
        """
        初期化
        
        Args:
            config_path: 設定ファイルのパス
        """
        self.console = Console()
        self.config_path = config_path
        self.config = self._load_config()
        self.setup_logging()
        
        # ADKクライアントの初期化
        self.config_manager = ConfigManager()
        self.auth_manager = AuthManager(self.config_manager)
        self.client = NewtonXClient(self.config_manager)
        
        # 実行結果の保存
        self.results = []
        self.current_chat_uid = None
        self.uploaded_files = []  # アップロードしたファイルの情報を保存
```

##### メイン実行メソッド
```python
def run(self):
    """
    メイン実行
    
    Returns:
        bool: 実行成功時True、失敗時False
    """
```

##### タスク実行メソッド
```python
def _execute_task(self, task: Dict[str, Any]) -> Any:
    """
    タスクを実行
    
    Args:
        task: タスク情報の辞書
    
    Returns:
        Any: タスク実行結果
    """
```

##### 初期化タスク
```python
def _task_init(self, task: Dict[str, Any]) -> Dict[str, Any]:
    """
    初期化タスク
    
    Args:
        task: タスク情報
    
    Returns:
        Dict[str, Any]: 初期化結果
    """
```

##### ファイルアップロードタスク
```python
def _task_upload(self, task: Dict[str, Any]) -> Dict[str, Any]:
    """
    ファイルアップロードタスク
    
    Args:
        task: タスク情報
    
    Returns:
        Dict[str, Any]: アップロード結果
    """
```

##### チャット作成タスク
```python
def _task_create_chat(self, task: Dict[str, Any]) -> Dict[str, Any]:
    """
    チャット作成タスク
    
    Args:
        task: タスク情報
    
    Returns:
        Dict[str, Any]: チャット作成結果
    """
```

##### メッセージ送信タスク
```python
def _task_send_message(self, task: Dict[str, Any]) -> Dict[str, Any]:
    """
    メッセージ送信タスク
    
    Args:
        task: タスク情報
    
    Returns:
        Dict[str, Any]: メッセージ送信結果
    """
```

##### ファイル分析タスク
```python
def _task_analyze_file(self, task: Dict[str, Any]) -> Dict[str, Any]:
    """
    ファイル分析タスク
    
    Args:
        task: タスク情報
    
    Returns:
        Dict[str, Any]: ファイル分析結果
    """
```

#### 設定ファイル (config.yml)
```yaml
# プロジェクト設定
project:
  name: "NewtonX Auto Runner Example"
  description: "設定ファイルベースの自動実行サンプル"
  version: "1.0.0"

# 認証設定
auth:
  auto_login: true
  session_timeout: 3600

# チャット設定
chat:
  default_assistant: "高速アシスタント"
  search_types:
    - "WEB検索（最新情報）"
    - "ナレッジ検索（登録ファイル）"
    - "両方（制限あり）"
    - "検索なし（基本対話）"

# ファイルアップロード設定
upload:
  max_file_size: 10485760  # 10MB
  allowed_extensions:
    images: [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"]
    documents: [".pdf", ".docx", ".txt", ".pptx", ".xls", ".xlsx", ".md"]

# 自動実行タスク
tasks:
  - name: "初期化"
    type: "init"
    description: "プロジェクトの初期化"
    enabled: true
    
  - name: "チャット作成"
    type: "create_chat"
    description: "新しいチャットの作成"
    title: "Auto Runner Test Chat"
    assistant: "高速アシスタント"
    enabled: true
    
  - name: "ファイルアップロード"
    type: "upload"
    description: "サンプルファイルのアップロード"
    files:
      - "sample_document.txt"
    enabled: true
    
  - name: "ファイル分析"
    type: "analyze_file"
    description: "アップロードしたファイルの分析"
    message: "このファイルの中身について教えてください。"
    enabled: true
    
  - name: "メッセージ送信"
    type: "send_message"
    description: "テストメッセージの送信"
    message: "こんにちは！Auto Runnerのテストです。"
    web_search: false
    knowledge_search: false
    enabled: true

# ログ設定
logging:
  level: "INFO"
  file: "auto_runner.log"
  max_size: 10485760  # 10MB
  backup_count: 5

# 出力設定
output:
  format: "rich"  # rich, json, plain
  save_results: true
  output_dir: "results"
```

### 2. CLI Example サンプル

#### ファイル構成
```
adk_examples/cli_example/
├── cli_example.py     # メインスクリプト
├── upload_sample.txt  # アップロードサンプル
├── requirements.txt   # 依存関係
└── README.md         # ドキュメント
```

#### CLIExample クラス

##### 初期化
```python
class CLIExample:
    def __init__(self):
        """
        初期化
        """
        self.console = Console()
        self.config_manager = ConfigManager()
        self.auth_manager = AuthManager(self.config_manager)
        self.client = NewtonXClient(self.config_manager)
        self.current_chat_uid = None
```

##### メイン実行メソッド
```python
def run(self):
    """
    メイン実行
    
    Returns:
        bool: 実行成功時True、失敗時False
    """
```

##### 認証メソッド
```python
def authenticate(self) -> bool:
    """
    認証を実行
    
    Returns:
        bool: 認証成功時True、失敗時False
    """
```

##### チャット作成メソッド
```python
def create_chat(self) -> bool:
    """
    チャットを作成
    
    Returns:
        bool: 作成成功時True、失敗時False
    """
```

##### ファイルアップロードメソッド
```python
def upload_file(self, file_path: str) -> bool:
    """
    ファイルをアップロード
    
    Args:
        file_path: ファイルパス
    
    Returns:
        bool: アップロード成功時True、失敗時False
    """
```

##### 対話的チャットメソッド
```python
def interactive_chat(self):
    """
    対話的チャットを開始
    """
```

##### メッセージ送信メソッド
```python
def send_message(self, message: str, web_search: bool = True, 
                knowledge_search: bool = False) -> Optional[str]:
    """
    メッセージを送信
    
    Args:
        message: 送信メッセージ
        web_search: ウェブ検索の有効化
        knowledge_search: ナレッジ検索の有効化
    
    Returns:
        str: アシスタントの応答、失敗時None
    """
```

### 3. Expense Processor サンプル

#### ファイル構成
```
adk_examples/expense_processor/
├── expense_processor_example.py  # メインスクリプト
├── sample_data/                  # サンプルデータ
│   ├── receipt1.png
│   └── receipt2.jpeg
├── requirements.txt              # 依存関係
└── README.md                    # ドキュメント
```

#### ExpenseProcessor クラス

##### 初期化
```python
class ExpenseProcessor:
    def __init__(self):
        """
        初期化
        """
        self.console = Console()
        self.config_manager = ConfigManager()
        self.auth_manager = AuthManager(self.config_manager)
        self.client = NewtonXClient(self.config_manager)
        self.current_chat_uid = None
        self.processed_expenses = []
```

##### メイン実行メソッド
```python
def run(self):
    """
    メイン実行
    
    Returns:
        bool: 実行成功時True、失敗時False
    """
```

##### 認証メソッド
```python
def authenticate(self) -> bool:
    """
    認証を実行
    
    Returns:
        bool: 認証成功時True、失敗時False
    """
```

##### チャット作成メソッド
```python
def create_chat(self) -> bool:
    """
    チャットを作成
    
    Returns:
        bool: 作成成功時True、失敗時False
    """
```

##### レシート処理メソッド
```python
def process_receipt(self, receipt_path: str) -> Dict[str, Any]:
    """
    レシートを処理
    
    Args:
        receipt_path: レシート画像のパス
    
    Returns:
        Dict[str, Any]: 処理結果
    """
```

##### 経費データ抽出メソッド
```python
def extract_expense_data(self, receipt_path: str) -> Dict[str, Any]:
    """
    経費データを抽出
    
    Args:
        receipt_path: レシート画像のパス
    
    Returns:
        Dict[str, Any]: 抽出された経費データ
    """
```

##### 結果保存メソッド
```python
def save_results(self, results: List[Dict[str, Any]], 
                output_file: str = "expense_results.json"):
    """
    結果を保存
    
    Args:
        results: 処理結果のリスト
        output_file: 出力ファイル名
    """
```

## 認証システム

### OAuth 2.0認証フロー（PKCE）
- クライアントIDとテナントIDを使用（シークレット不要）
- アクセストークンの取得・更新
- セッション管理と永続化

### 認証情報の取得方法

#### 1. NewtonX管理画面での認証情報取得
1. NewtonX管理画面にログイン
2. 「API設定」または「開発者設定」セクションに移動
3. 「クライアントID」と「クライアントシークレット」を確認
4. 必要に応じて新しい認証情報を生成

#### 2. 認証情報の設定方法（PKCE）
```python
from newtonx_adk import ConfigManager

# 設定管理クラスを初期化
config_manager = ConfigManager()

# 認証情報を設定
config_manager.set_credentials(
    client_id="your_client_id",      # 管理画面で取得したクライアントID
    tenant_id="your_tenant_id"       # 管理画面で取得したテナントID
)

# 設定を保存（次回以降の自動認証用）
config_manager.save_config()
```

#### 3. 環境変数での認証情報設定（PKCE）
```bash
# 環境変数を設定
export NEWTONX_CLIENT_ID="your_client_id"
export NEWTONX_TENANT_ID="your_tenant_id"
```

```python
import os
from newtonx_adk import ConfigManager

config_manager = ConfigManager()

# 環境変数から認証情報を取得
client_id = os.getenv("NEWTONX_CLIENT_ID")
tenant_id = os.getenv("NEWTONX_TENANT_ID")

if client_id and tenant_id:
    config_manager.set_credentials(client_id, tenant_id)
else:
    print("環境変数NEWTONX_CLIENT_IDとNEWTONX_TENANT_IDが設定されていません")
```

#### 4. 設定ファイルでの認証情報管理
```python
# config.json ファイルを作成
{
    "auth": {
        "client_id": "your_client_id",
        # client_secret は不要
    },
    "api": {
        "base_url": "https://api.newtonx.com",
        "timeout": 30
    }
}
```

### 認証の実行

#### 1. 基本的な認証
```python
from newtonx_adk import NewtonXClient, ConfigManager

# 設定管理クラスを初期化
config_manager = ConfigManager()

# 認証情報を設定
config_manager.set_credentials(
    client_id="your_client_id",
    tenant_id="your_tenant_id"
)

# クライアントを初期化
client = NewtonXClient(config_manager)

# 認証を実行
if client.authenticate():
    print("認証が成功しました！")
else:
    print("認証に失敗しました。")
```

#### 2. 自動認証（保存されたトークンを使用）
```python
from newtonx_adk import NewtonXClient, ConfigManager

# 設定管理クラスを初期化（保存された設定を読み込み）
config_manager = ConfigManager()

# クライアントを初期化
client = NewtonXClient(config_manager)

# 自動認証を試行（保存されたトークンを使用）
if client.authenticate():
    print("自動認証が成功しました！")
else:
    print("自動認証に失敗しました。手動認証が必要です。")
```

#### 3. 認証状態の確認
```python
from newtonx_adk import AuthManager, ConfigManager

config_manager = ConfigManager()
auth_manager = AuthManager(config_manager)

# 認証状態を確認
if auth_manager.is_authenticated():
    print("認証済みです")
else:
    print("未認証です")
```

### 認証エラーの対処法

#### 1. 認証情報が無効な場合
**症状**: AuthenticationErrorが発生
**原因**: クライアントIDまたはシークレットが間違っている
**解決方法**:
```python
# 認証情報を再設定
config_manager.set_credentials(
    client_id="正しいクライアントID",
    tenant_id="正しいテナントID"
)

# 設定を保存
config_manager.save_config()

# 再認証
client.authenticate()
```

#### 2. トークンが期限切れの場合
**症状**: 401 Unauthorizedエラー
**原因**: アクセストークンの有効期限が切れている
**解決方法**:
```python
# トークンを更新
auth_manager.refresh_token()

# または再認証
client.authenticate()
```

#### 3. ネットワークエラーの場合
**症状**: APIErrorが発生
**原因**: ネットワーク接続の問題
**解決方法**:
```python
import time

# リトライ処理
max_retries = 3
for attempt in range(max_retries):
    try:
        if client.authenticate():
            print("認証成功")
            break
        else:
            print(f"認証失敗 (試行 {attempt + 1}/{max_retries})")
    except Exception as e:
        print(f"認証エラー: {e}")
        if attempt < max_retries - 1:
            time.sleep(2)  # 2秒待機
```

### 認証トークンの管理

#### 1. トークンの保存場所
- **デフォルト**: `~/.newtonx/config.json`
- **環境変数**: `NEWTONX_CONFIG_PATH`で指定可能

#### 2. トークンの有効期限
- **アクセストークン**: 通常1時間
- **リフレッシュトークン**: 通常30日
- **自動更新**: ADKが自動的にトークンを更新

#### 3. トークンの手動管理
```python
from newtonx_adk import AuthManager, ConfigManager

config_manager = ConfigManager()
auth_manager = AuthManager(config_manager)

# 現在のトークンを確認
current_token = auth_manager._get_headers().get('Authorization')
print(f"現在のトークン: {current_token}")

# トークンを手動で更新
if auth_manager.refresh_token():
    print("トークン更新成功")
else:
    print("トークン更新失敗")
```

### 認証のベストプラクティス

#### 1. 認証情報の安全な管理
```python
# 推奨: 環境変数を使用
import os
client_id = os.getenv("NEWTONX_CLIENT_ID")
client_secret = os.getenv("NEWTONX_CLIENT_SECRET")

# 非推奨: コード内に直接記述
client_id = "your_client_id"  # セキュリティリスク
```

#### 2. 認証エラーの適切な処理
```python
from newtonx_adk import AuthenticationError, APIError

try:
    if client.authenticate():
        print("認証成功")
    else:
        print("認証失敗")
except AuthenticationError as e:
    print(f"認証エラー: {e}")
    # 認証情報の再設定を促す
except APIError as e:
    print(f"API通信エラー: {e}")
    # ネットワーク接続の確認を促す
```

#### 3. 認証状態の定期的な確認
```python
import time
from threading import Timer

def check_auth_status():
    if not auth_manager.is_authenticated():
        print("認証が切れています。再認証を実行します。")
        client.authenticate()

# 30分ごとに認証状態を確認
Timer(1800, check_auth_status).start()
```

### 認証のデバッグ

#### 1. 詳細ログの有効化
```python
import logging

# ログレベルを設定
logging.basicConfig(level=logging.DEBUG)

# ADKのログを有効化
logger = logging.getLogger('newtonx_adk')
logger.setLevel(logging.DEBUG)
```

#### 2. 認証フローの詳細確認
```python
# 認証プロセスの詳細を確認
auth_manager.authenticate()

# 認証ヘッダーを確認
headers = auth_manager._get_headers()
print(f"認証ヘッダー: {headers}")

# トークン情報を確認
token_info = auth_manager.get_token_info()
print(f"トークン情報: {token_info}")
```

## API機能

### チャット機能
- **チャット作成**: `create_chat(assistant_uid, title)`
- **メッセージ送信**: `send_message(chat_uid, message, web_search, knowledge_search)`
- **チャット削除**: `delete_chat(chat_uid)`
- **チャットタイトル更新**: `update_chat_title(chat_uid, title)`

### ファイルアップロード
- **画像アップロード**: `upload_image(chat_uid, file_path)`
- **ドキュメントアップロード**: `upload_document(chat_uid, file_path)`

**重要**: NewtonX APIはファイルアップロード時にuidを返さない場合があります。代わりにチャット内でファイルを自動認識する仕組みになっています。

### アシスタント管理
- **アシスタント一覧取得**: `get_assistants()`
- **ユーザー情報取得**: `get_user_info()`

### フォルダ管理
- **フォルダ作成**: `create_folder(name)`
- **フォルダ一覧取得**: `get_folders()`
- **チャット移動**: `move_chat_to_folder(chat_uid, folder_uid)`

## エラーハンドリング

### 例外クラス
- **NewtonXError**: 基本例外クラス
- **AuthenticationError**: 認証エラー
- **APIError**: API通信エラー
- **ConfigurationError**: 設定エラー
- **FileUploadError**: ファイルアップロードエラー
- **ChatError**: チャット関連エラー

### エラー処理の例
```python
try:
    response = client.send_message(chat_uid, message)
except AuthenticationError as e:
    print(f"認証エラー: {e}")
except APIError as e:
    print(f"APIエラー: {e}")
```

## 開発・運用のベストプラクティス

### 1. 認証管理
- 認証情報は環境変数または設定ファイルで管理
- トークンの有効期限を監視
- 認証エラー時の再認証処理を実装

### 2. エラーハンドリング
- 適切な例外クラスを使用
- ログ出力によるデバッグ情報の記録
- ユーザーフレンドリーなエラーメッセージ

### 3. ファイルアップロード
- ファイルサイズ制限の確認（10MB）
- 対応ファイル形式の確認
- アップロード失敗時のリトライ処理

### 4. パフォーマンス
- 適切なタイムアウト設定
- 大量リクエスト時の制御
- キャッシュ機能の活用

## トラブルシューティング

### よくある問題と解決方法

#### 1. 認証エラー
**症状**: AuthenticationErrorが発生
**原因**: 認証情報が無効または期限切れ
**解決方法**: 
- 認証情報の再設定
- トークンの再取得

#### 2. ファイルアップロードエラー
**症状**: FileUploadErrorが発生
**原因**: ファイルサイズ超過、形式不対応
**解決方法**:
- ファイルサイズの確認（10MB以下）
- 対応形式の確認

#### 3. API通信エラー
**症状**: APIErrorが発生
**原因**: ネットワーク問題、API制限
**解決方法**:
- ネットワーク接続の確認
- リクエスト頻度の調整

#### 4. ファイル認識の問題
**症状**: アップロードしたファイルが認識されない
**原因**: NewtonX APIの仕様
**解決方法**:
- ファイルアップロードは成功しているが、uidが取得できない場合がある
- チャット内で自動認識されるため、通常は問題なし

## バージョン情報

### 現在のバージョン: v0.10.5
- 基本機能の実装完了
- 3つのサンプルアプリケーション
- エラーハンドリングの充実
- ドキュメントの整備

### 今後の開発予定
- より多くのAPI機能の追加
- パフォーマンスの最適化
- テストカバレッジの向上
- ドキュメントの拡充

## 技術仕様

### 対応Pythonバージョン
- Python 3.8以上

### 主要依存関係
- requests>=2.25.0
- dataclasses>=0.6 (Python 3.7未満)

### API制限
- ファイルサイズ: 10MB以下
- 対応画像形式: jpg, jpeg, png, gif, bmp, webp
- 対応ドキュメント形式: pdf, docx, txt, pptx, xls, xlsx, md

## 配布・インストール

### 社内配布用パッケージ
- ファイル名: `newtonx_adk-main.zip`
- サイズ: 約2.4MB
- 内容: ADK + サンプルアプリケーション

### インストール方法
```bash

# ADKインストール
cd newtonx_adk-main
pip install -e .

# 依存関係インストール
pip install -r requirements.txt
```

## サポート・問い合わせ

### 技術サポート
- 本ADKに関する技術的なサポートは、NewtonXの[NewtonX ADK](https://seraku.newton-x.net/aichat/chat/6dffd4c7-f632-4e58-95c0-b0feaa58e396) アシスタントを活用ください。
- NewtonX ADK開発を手伝っていただける方を募集しています。詳しくは各部のAI SHIFT担当経由でNewtonX課までお問い合わせください。

### 注意事項
- このADKは社内配布用です
- 外部への公開は禁止されています
- 認証情報は適切に管理してください

---

このドキュメントは、NewtonX ADK v0.10.5の開発・運用に関する知見をまとめたものです。技術的な質問や問題解決の際に参照してください。 