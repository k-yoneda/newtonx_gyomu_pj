#!/usr/bin/env python3
"""
NewtonX ADKを使用したCLI版の例

ADKを使用してCLIアプリケーションを構築する例を示します。
"""

import sys
import os
import time
# Windows対応: readline フォールバック
try:
    import readline  # type: ignore
except Exception:
    try:
        import pyreadline3 as readline  # type: ignore
    except Exception:
        readline = None  # type: ignore
import glob
from typing import Optional, List, Dict
from rich.console import Console
from rich.prompt import Prompt
from rich.table import Table
from rich.panel import Panel

# ADKのインポートパス設定
# 常にローカルのsrcディレクトリを優先して読み込むように設定
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
grandparent_dir = os.path.dirname(parent_dir)

# ADKディレクトリの絶対パスを追加（src配下）
adk_path = os.path.join(grandparent_dir, 'src')
if adk_path not in sys.path:
    sys.path.insert(0, adk_path)

print(f"ADK path added: {adk_path}")

try:
    from newtonx_adk import NewtonXClient, ConfigManager, AuthManager, NewtonXError, AuthenticationError, APIError
    print(f"Loaded newtonx_adk from: {sys.modules['newtonx_adk'].__file__}")
except ImportError as e:
    print(f"ADKのインポートに失敗しました: {e}")
    print("ADKが正しくインストールされているか、srcディレクトリが存在するか確認してください。")
    sys.exit(1)


class FilePathCompleter:
    """ファイルパス補完クラス"""
    
    def __init__(self):
        self.matches = []
        # アップロード可能なファイル形式
        self.uploadable_extensions = {
            # 画像ファイル
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp',
            # ドキュメントファイル
            '.pdf', '.docx', '.txt', '.pptx', '.xls', '.xlsx', '.md',
            # 音声ファイル
            '.wav', '.mp3', '.aiff', '.aac', '.ogg', '.flac'
        }
    
    def complete(self, text, state):
        """ファイルパス補完を実行"""
        if state == 0:
            # 初回呼び出し時にマッチするファイルを検索
            self.matches = self._get_matches(text)
        
        if state < len(self.matches):
            return self.matches[state]
        else:
            return None
    
    def _get_matches(self, text):
        """マッチするファイルパスを取得"""
        if not text:
            # 空の場合は現在のディレクトリのファイルを返す（アップロード可能なファイルを優先）
            files = os.listdir('.')
            uploadable = []
            others = []
            
            for f in files:
                if os.path.isdir(f):
                    others.append(f + '/')
                else:
                    ext = os.path.splitext(f)[1].lower()
                    if ext in self.uploadable_extensions:
                        uploadable.append(f)
                    else:
                        others.append(f)
            
            return uploadable + others
        
        # パスの分割
        dirname = os.path.dirname(text) or '.'
        basename = os.path.basename(text)
        
        try:
            # ディレクトリ内のファイルを検索
            files = os.listdir(dirname)
            uploadable = []
            others = []
            
            for file in files:
                if file.startswith(basename):
                    full_path = os.path.join(dirname, file)
                    if os.path.isdir(full_path):
                        others.append(full_path + '/')
                    else:
                        ext = os.path.splitext(file)[1].lower()
                        if ext in self.uploadable_extensions:
                            uploadable.append(full_path)
                        else:
                            others.append(full_path)
            
            return uploadable + others
        except (OSError, FileNotFoundError):
            return []


class NewtonXCLIExample:
    """NewtonX CLI例"""
    
    def __init__(self):
        """初期化"""
        self.console = Console()
        self.config_manager = ConfigManager()
        self.auth_manager = AuthManager(self.config_manager)
        self.client = NewtonXClient(self.config_manager)
        
        # ファイルパス補完を設定
        self._setup_file_completion()
    
    def _setup_file_completion(self):
        """ファイルパス補完を設定"""
        try:
            if readline is None:
                return
            # readlineの設定
            readline.set_completer_delims(" \t\n`!@#$%^&*()=+[{}]\\|;:'\",<>?")
            readline.set_completer(FilePathCompleter().complete)
            readline.parse_and_bind('tab: complete')
            
            # 履歴ファイルの設定
            history_file = os.path.expanduser('~/.newtonx_cli_history')
            try:
                readline.read_history_file(history_file)
            except FileNotFoundError:
                pass
            
            # 終了時に履歴を保存
            import atexit
            atexit.register(readline.write_history_file, history_file)
            
        except Exception as e:
            # readlineが利用できない場合は無視
            pass
    
    def _normalize_path(self, path_input: str) -> str:
        """ユーザー入力のパスを正規化（引用符除去・~や環境変数展開・絶対化）"""
        s = path_input.strip()
        # 先頭末尾の引用符を除去（'...'/"...")
        if (len(s) >= 2) and ((s[0] == s[-1]) and s[0] in ('"', "'")):
            s = s[1:-1]
        # 環境変数/チルダ展開
        s = os.path.expandvars(os.path.expanduser(s))
        # 絶対パス化
        if not os.path.isabs(s):
            s = os.path.abspath(s)
        # 正規化
        return os.path.normpath(s)
    
    def run(self):
        """CLIアプリケーションを実行"""
        try:
            # スパルタンX風のNewtonXアスキーアート
            self._show_newtonx_ascii_art()
            
            self.console.print("[bold blue]NewtonX CLI Example v1.0.0[/bold blue]")
            self.console.print("初期化中...")
            
            # 認証チェック
            if not self._authenticate():
                self.console.print("[red]認証に失敗しました。[/red]")
                return
            
            # メインループ
            self._main_loop()
            
        except KeyboardInterrupt:
            self.console.print("\n[yellow]中断されました。[/yellow]")
        except Exception as e:
            self.console.print(f"[red]エラーが発生しました: {e}[/red]")
    
    def _show_newtonx_ascii_art(self):
        """NewtonXアスキーアートを表示"""
        ascii_art = """
[bold red]
    ███╗   ██╗███████╗██╗    ██╗████████╗ ██████╗ ███╗   ██╗██╗    ██╗
    ████╗  ██║██╔════╝██║    ██║╚══██╔══╝██╔═══██╗████╗  ██║╚██╗  ██╔╝
    ██╔██╗ ██║█████╗  ██║ █╗ ██║   ██║   ██║   ██║██╔██╗ ██║  ╚███╔╝
    ██║╚██╗██║██╔══╝  ██║███╗██║   ██║   ██║   ██║██║╚██╗██║ ██║  ██╗
    ██║ ╚████║███████╗╚███╔███╔╝   ██║   ╚██████╔╝██║ ╚████║██╔╝   ██╗
    ╚═╝  ╚═══╝╚══════╝ ╚══╝╚══╝    ╚═╝    ╚═════╝ ╚═╝  ╚═══╝╚═╝    ╚═╝
[/bold red]

[bold yellow]╔════════════════════════════════════════════════════════════════════════╗[/bold yellow]
[bold yellow]║                    🚀 NEWTONX CLI EXAMPLE v1.0.0 🚀                    ║[/bold yellow]
[bold yellow]║                                                                        ║[/bold yellow]
[bold yellow]║                    AI-Powered Research Assistant                       ║[/bold yellow]
[bold yellow]║                                                                        ║[/bold yellow]
[bold yellow]║        Features:                                                       ║[/bold yellow]
[bold yellow]║        • Chat Management          • Assistant Management               ║[/bold yellow]
[bold yellow]║        • File Upload Support      • Message Sending                    ║[/bold yellow]
[bold yellow]║        • Browser-based Auth       • Session Management                 ║[/bold yellow]
[bold yellow]╚════════════════════════════════════════════════════════════════════════╝[/bold yellow]

[bold cyan]━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━[/bold cyan]
"""
        self.console.print(ascii_art)
        
        # アスキーアートを少し長く表示するための一時停止
        time.sleep(2)  # 2秒間表示
        
        # Enterキーを押すまで待つ
        self.console.print("\n[bold green]Enterキーを押して続行...[/bold green]")
        input()
        
        # 画面をクリアして次の表示に移行（OS別の方法）
        os.system('clear' if os.name == 'posix' else 'cls')
    
    def _authenticate(self) -> bool:
        """認証を実行（新しい認証フロー: EntraID + Cookie）"""
        try:
            # 認証状態を確認
            if self.auth_manager.is_authenticated():
                self.console.print("[green]認証済みです。[/green]")
                return True

            # 設定を確認
            config = self.config_manager.get_config()
            if not config.client_id or not config.tenant_id:
                self.console.print("[red]認証設定が不完全です。[/red]")
                self.console.print("[yellow]以下のコマンドで認証設定を行ってください:[/yellow]")
                self.console.print("[cyan]python tools/setup_config.py[/cyan]")
                return False

            # 新しい認証フローを実行（EntraID: PKCE/Device）
            self.console.print("[cyan]Entra ID 認証を開始します...[/cyan]")
            try:
                token = self.auth_manager.get_access_token(mode=config.auth_mode)
                if not token:
                    self.console.print("[red]Entra ID 認証に失敗しました。[/red]")
                    return False
                self.console.print("[green]Entra ID 認証が完了しました。[/green]")
                self.console.print("[dim]NewtonXの内部トークン登録に進みます。[/dim]")
            except KeyboardInterrupt:
                self.console.print("[yellow]認証が中断されました。[/yellow]")
                return False

            # NewtonX ログインページを開く（ユーザが社内SSOを完了できるように）
            try:
                import webbrowser
                web_base = (config.api_base_url or 'https://seraku.newton-x.net/api').rsplit('/api', 1)[0]
                sub = config.company_subdomain or ''
                login_url = f"{web_base}/account/redirectadlogin" + (f"?subdomain={sub}" if sub else '')
                webbrowser.open(login_url)
                self.console.print(f"[cyan]NewtonXログインページを開きました: {login_url}[/cyan]")
            except Exception:
                pass

            # auth_callback URL もしくは内部トークンの貼り付け
            self.console.print("\n[bold]auth_callback のURL、または内部トークン(例: 55058|xxxx...)を貼り付けてください。[/bold]")
            pasted = input("Paste URL or token: ").strip()

            # 抽出処理（set_internal_token.py と同等）
            import re
            def _extract_token(text: str):
                m = re.search(r"access_token=([^&\s]+)", text)
                if m:
                    return m.group(1)
                if '|' in text and ' ' not in text:
                    return text
                return None

            internal_token = _extract_token(pasted)
            if not internal_token:
                self.console.print("[red]内部トークンが抽出できませんでした。操作をやり直してください。[/red]")
                return False

            # 保存→ブートストラップ
            self.auth_manager.set_newtonx_access_token(internal_token)
            self.console.print("[cyan]内部トークンを保存しました。セッションをブートストラップします...[/cyan]")
            ok = self.auth_manager.bootstrap_session()
            if not ok:
                self.console.print("[yellow]ブートストラップに失敗しましたが、続行します。必要に応じて再試行されます。[/yellow]")

            # 完了
            self.console.print("[green]NewtonX認証の初期化が完了しました。[/green]")
            return True

        except Exception as e:
            self.console.print(f"[red]認証エラー: {e}[/red]")
            return False
    
    def _main_loop(self):
        """メインループ"""
        while True:
            try:
                # メインメニューを表示
                menu_items = [
                    "アシスタント一覧表示",
                    "チャット一覧表示",
                    "新規チャット作成",
                    "メッセージ送信",
                    "音声付きメッセージ送信",
                    "ファイルアップロード",
                    "アシスタントへのナレッジ登録",
                    "アシスタントナレッジ削除",
                    "設定",
                    "終了"
                ]
                choice = self._show_main_menu(menu_items)
                
                if choice == "1":
                    self._show_assistants()
                elif choice == "2":
                    self._show_chats()
                elif choice == "3":
                    self._create_chat()
                elif choice == "4":
                    self._send_message()
                elif choice == "5":
                    self._send_voice_message()
                elif choice == "6":
                    self._upload_file()
                elif choice == "7":
                    self._upload_assistant_knowledge()
                elif choice == "8":
                    self._delete_assistant_knowledge()
                elif choice == "9":
                    self._show_settings()
                elif choice == "10":
                    self.console.print("[yellow]終了します。[/yellow]")
                    break
                else:
                    self.console.print(f"[red]無効な選択です: {choice}[/red]")
                    
            except KeyboardInterrupt:
                self.console.print("[yellow]中断されました。[/yellow]")
                break
            except Exception as e:
                self.console.print(f"[red]エラーが発生しました: {e}[/red]")
    
    def _show_main_menu(self, menu_items: List[str]) -> str:
        """メインメニューを表示"""
        self.console.print("\n[bold cyan]メインメニュー[/bold cyan]")
        for i, item in enumerate(menu_items, 1):
            self.console.print(f"{i}. {item}")
        
        return Prompt.ask(
            "\n[bold green]選択してください[/bold green]",
            choices=[str(i) for i in range(1, len(menu_items) + 1)]
        )
    
    def _show_assistants(self):
        """アシスタント一覧を表示"""
        try:
            self.console.print("\n[bold]アシスタント一覧:[/bold]")
            assistants = self.client.get_assistants()
            
            if not assistants:
                self.console.print("[yellow]アシスタントが見つかりません。[/yellow]")
                return
            
            table = Table(title="アシスタント一覧")
            table.add_column("名前", style="cyan")
            table.add_column("ID", style="green")
            table.add_column("説明", style="white")
            
            for assistant in assistants:
                table.add_row(
                    assistant.get('name', 'N/A'),
                    assistant.get('uid', 'N/A'),
                    assistant.get('description', 'N/A')
                )
            
            self.console.print(table)
            
        except Exception as e:
            self.console.print(f"[red]アシスタント一覧取得エラー: {e}[/red]")
    
    def _show_chats(self):
        """チャット一覧を表示"""
        try:
            self.console.print("\n[bold]チャット一覧:[/bold]")
            chats = self.client.get_chats()
            
            if not chats:
                self.console.print("[yellow]チャットが見つかりません。[/yellow]")
                return
            
            table = Table(title="チャット一覧")
            table.add_column("タイトル", style="cyan")
            table.add_column("ID", style="green")
            table.add_column("アシスタント", style="white")
            table.add_column("作成日", style="yellow")
            
            for chat in chats:
                # アシスタント名を取得
                assistant_name = "N/A"
                if 'assistant' in chat and isinstance(chat['assistant'], dict):
                    assistant_name = chat['assistant'].get('name', 'N/A')
                
                table.add_row(
                    chat.get('title', 'N/A'),
                    chat.get('id', 'N/A'),
                    assistant_name,
                    chat.get('created_at', 'N/A')
                )
            
            self.console.print(table)
            
        except Exception as e:
            self.console.print(f"[red]チャット一覧取得エラー: {e}[/red]")
    
    def _create_chat(self):
        """新規チャットを作成"""
        try:
            # アシスタント一覧を取得
            assistants = self.client.get_assistants()
            if not assistants:
                self.console.print("[red]利用可能なアシスタントがありません。[/red]")
                return
            
            # アシスタント選択
            self.console.print("\n[bold]アシスタントを選択してください:[/bold]")
            for i, assistant in enumerate(assistants, 1):
                self.console.print(f"{i}. {assistant.get('name', 'N/A')}")
            
            choice = Prompt.ask("選択", choices=[str(i) for i in range(1, len(assistants) + 1)])
            selected_assistant = assistants[int(choice) - 1]
            
            # チャットタイトル入力
            title = Prompt.ask("チャットタイトルを入力してください")
            
            # チャット作成
            chat_uid = self.client.create_chat(
                assistant_uid=selected_assistant['uid'],
                title=title
            )
            
            if chat_uid:
                self.console.print(f"[green]チャットが作成されました！[/green]")
                self.console.print(f"チャットID: {chat_uid}")
            else:
                self.console.print("[red]チャットの作成に失敗しました。[/red]")
                
        except Exception as e:
            self.console.print(f"[red]チャット作成エラー: {e}[/red]")
    
    def _send_message(self):
        """メッセージを送信"""
        try:
            # チャット一覧を取得
            chats = self.client.get_chats()
            if not chats:
                self.console.print("[red]チャットが見つかりません。[/red]")
                return
            
            # チャット選択
            self.console.print("\n[bold]チャットを選択してください:[/bold]")
            for i, chat in enumerate(chats, 1):
                self.console.print(f"{i}. {chat.get('title', 'N/A')}")
            
            choice = Prompt.ask("選択", choices=[str(i) for i in range(1, len(chats) + 1)])
            selected_chat = chats[int(choice) - 1]
            
            # チャットUIDを取得（複数の可能性を試す）
            chat_uid = selected_chat.get('id') or selected_chat.get('uid') or selected_chat.get('chat_uid')
            
            if not chat_uid:
                self.console.print("[red]チャットUIDが見つかりません。[/red]")
                return
            
            # メッセージ入力
            message = Prompt.ask("メッセージを入力してください")
            
            # 検索設定
            web_search = Prompt.ask("ウェブ検索を使用しますか？", choices=["y", "n"], default="y") == "y"
            
            # ファイルがアップロードされている場合はナレッジ検索を推奨
            knowledge_search_default = "y"  # ファイル参照のためデフォルトを有効に
            knowledge_search = Prompt.ask("ナレッジ検索（アップロードしたファイルを参照）を使用しますか？", 
                                        choices=["y", "n"], default=knowledge_search_default) == "y"
            
            # 最新のメッセージ順序を取得して parent_order に設定（文脈維持）
            parent_order = 0
            try:
                chat_detail = self.client.get_chat(chat_uid)
                if chat_detail and 'messages' in chat_detail and chat_detail['messages']:
                    last_msg = chat_detail['messages'][-1]
                    parent_order = last_msg.get('chat_order', 0)
            except Exception:
                pass

            # メッセージ送信
            response = self.client.send_message(
                chat_uid=chat_uid,
                message=message,
                web_search=web_search,
                knowledge_search=knowledge_search,
                parent_order=parent_order
            )
            
            if response:
                self.console.print(f"\n[bold]アシスタントの応答:[/bold]")
                self.console.print(Panel(response, title="応答"))
            else:
                self.console.print("[red]メッセージの送信に失敗しました。[/red]")
                
        except Exception as e:
            self.console.print(f"[red]メッセージ送信エラー: {e}[/red]")
            import traceback
            traceback.print_exc()
    
    def _send_voice_message(self):
        """音声付きメッセージを送信"""
        try:
            # チャット一覧を取得
            chats = self.client.get_chats()
            if not chats:
                self.console.print("[red]チャットが見つかりません。[/red]")
                return
            
            # チャット選択
            self.console.print("\n[bold]チャットを選択してください:[/bold]")
            for i, chat in enumerate(chats, 1):
                self.console.print(f"{i}. {chat.get('title', 'N/A')}")
            
            choice = Prompt.ask("選択", choices=[str(i) for i in range(1, len(chats) + 1)])
            selected_chat = chats[int(choice) - 1]
            
            # チャットUIDを取得（複数の可能性を試す）
            chat_uid = selected_chat.get('id') or selected_chat.get('uid') or selected_chat.get('chat_uid')
            
            if not chat_uid:
                self.console.print("[red]チャットUIDが見つかりません。[/red]")
                return
            
            # メッセージ入力
            message = Prompt.ask("メッセージを入力してください")
            
            # 検索設定
            web_search = Prompt.ask("ウェブ検索を使用しますか？", choices=["y", "n"], default="y") == "y"
            knowledge_search = Prompt.ask("ナレッジ検索（アップロードしたファイルを参照）を使用しますか？", 
                                        choices=["y", "n"], default="n") == "y"
            
            # 音声ファイルパス入力
            self.console.print("\n[bold cyan]音声ファイルパスを入力してください[/bold cyan]")
            self.console.print("[yellow]ヒント: タブキーでファイル名を補完できます[/yellow]")
            self.console.print("[yellow]対応形式: .wav, .mp3, .aiff, .aac, .ogg, .flac[/yellow]")
            self.console.print("[yellow]最大サイズ: 15MB[/yellow]")
            
            # 現在のディレクトリの音声ファイルを表示
            self._show_available_audio_files()
            
            file_path_input = input("音声ファイルパス: ")
            
            # 入力が空の場合はキャンセル
            if not file_path_input.strip():
                self.console.print("[yellow]ファイルパスが入力されませんでした。キャンセルします。[/yellow]")
                return
            
            # 入力のパスを正規化（引用符除去・展開・絶対化）
            file_path = self._normalize_path(file_path_input)
            
            # ファイル存在チェック
            if not os.path.exists(file_path):
                self.console.print(f"[red]ファイルが存在しません: {file_path}[/red]")
                self.console.print(f"[yellow]現在のディレクトリ: {os.getcwd()}[/yellow]")
                return
            
            # ファイル拡張子チェック
            file_ext = os.path.splitext(file_path)[1].lower()
            allowed_extensions = {'.wav', '.mp3', '.aiff', '.aac', '.ogg', '.flac'}
            if file_ext not in allowed_extensions:
                self.console.print(f"[red]未対応のファイル形式です: {file_ext}[/red]")
                self.console.print(f"[yellow]対応形式: {', '.join(allowed_extensions)}[/yellow]")
                return
            
            # ファイルサイズチェック（15MB）
            file_size = os.path.getsize(file_path)
            max_size = 15 * 1024 * 1024  # 15MB
            if file_size > max_size:
                self.console.print(f"[red]ファイルサイズが大きすぎます（15MB以下にしてください）。[/red]")
                self.console.print(f"[yellow]現在のサイズ: {file_size / 1024 / 1024:.2f} MB[/yellow]")
                return
            
            # ファイル情報表示
            self.console.print(f"\nファイル情報:")
            self.console.print(f"  パス: {file_path}")
            self.console.print(f"  サイズ: {file_size / 1024:.1f} KB")
            self.console.print(f"  形式: {file_ext}")
            
            # 音声付きメッセージ送信
            self.console.print("\n音声付きメッセージを送信中...")
            try:
                # 最新のメッセージ順序を取得して parent_order に設定（文脈維持）
                parent_order = 0
                try:
                    chat_detail = self.client.get_chat(chat_uid)
                    if chat_detail and 'messages' in chat_detail and chat_detail['messages']:
                        last_msg = chat_detail['messages'][-1]
                        parent_order = last_msg.get('chat_order', 0)
                except Exception:
                    pass

                response = self.client.send_message(
                    chat_uid=chat_uid,
                    message=message,
                    web_search=web_search,
                    knowledge_search=knowledge_search,
                    audio_file_path=file_path,
                    parent_order=parent_order
                )
                
                if response:
                    self.console.print(f"\n[bold]アシスタントの応答:[/bold]")
                    self.console.print(Panel(response, title="応答"))
                else:
                    self.console.print("[red]メッセージの送信に失敗しました。[/red]")
            except Exception as e:
                self.console.print(f"[red]音声付きメッセージ送信エラー: {e}[/red]")
                import traceback
                traceback.print_exc()
                
        except Exception as e:
            self.console.print(f"[red]音声付きメッセージ送信処理エラー: {e}[/red]")
    
    def _upload_file(self):
        """ファイルをアップロード"""
        try:
            # チャット一覧を取得
            chats = self.client.get_chats()
            if not chats:
                self.console.print("[red]チャットが見つかりません。[/red]")
                return
            
            # チャット選択
            self.console.print("\n[bold]チャットを選択してください:[/bold]")
            for i, chat in enumerate(chats, 1):
                self.console.print(f"{i}. {chat.get('title', 'N/A')}")
            
            choice = Prompt.ask("選択", choices=[str(i) for i in range(1, len(chats) + 1)])
            selected_chat = chats[int(choice) - 1]
            
            # チャットUIDを取得（複数の可能性を試す）
            chat_uid = selected_chat.get('id') or selected_chat.get('uid') or selected_chat.get('chat_uid')
            
            if not chat_uid:
                self.console.print("[red]チャットUIDが見つかりません。[/red]")
                return
            
            # ファイルパス入力（相対パス対応）
            self.console.print("\n[bold cyan]ファイルパスを入力してください[/bold cyan]")
            self.console.print("[yellow]ヒント: タブキーでファイル名を補完できます[/yellow]")
            self.console.print("[yellow]例: ./test_upload.txt または ../documents/file.pdf[/yellow]")
            
            # 現在のディレクトリのアップロード可能なファイルを表示
            self._show_available_files()
            
            file_path_input = input("ファイルパス: ")
            
            # 入力が空の場合はキャンセル
            if not file_path_input.strip():
                self.console.print("[yellow]ファイルパスが入力されませんでした。キャンセルします。[/yellow]")
                return
            
            # 入力のパスを正規化（引用符除去・展開・絶対化）
            file_path = self._normalize_path(file_path_input)
            
            # ファイル存在チェック
            if not os.path.exists(file_path):
                self.console.print(f"[red]ファイルが存在しません: {file_path}[/red]")
                self.console.print(f"[yellow]現在のディレクトリ: {os.getcwd()}[/yellow]")
                return
            
            # ファイルサイズの確認
            file_size = os.path.getsize(file_path)
            if file_size > 50 * 1024 * 1024:  # 50MB制限
                self.console.print("[red]ファイルサイズが大きすぎます（50MB以下にしてください）。[/red]")
                return
            
            # ファイルタイプ判定
            file_ext = os.path.splitext(file_path)[1].lower()
            is_image = file_ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
            
            self.console.print(f"ファイル情報:")
            self.console.print(f"  パス: {file_path}")
            self.console.print(f"  サイズ: {file_size / 1024:.1f} KB")
            self.console.print(f"  タイプ: {'画像' if is_image else 'ドキュメント'}")
            
            if is_image:
                # 画像ファイルのアップロード
                self.console.print("画像をアップロード中...")
                try:
                    image_id = self.client.upload_image(
                        chat_uid=chat_uid,
                        file_path=file_path
                    )
                    if image_id:
                        self.console.print(f"[green]画像がアップロードされました！[/green]")
                        self.console.print(f"画像ID: {image_id}")
                    else:
                        self.console.print("[red]画像のアップロードに失敗しました。[/red]")
                except Exception as e:
                    self.console.print(f"[red]画像アップロードエラー: {e}[/red]")
            else:
                # ドキュメントファイルのアップロード
                self.console.print("ドキュメントをアップロード中...")
                try:
                    success = self.client.upload_document(
                        chat_uid=chat_uid,
                        file_path=file_path
                    )
                    if success:
                        self.console.print("[green]ドキュメントがアップロードされました！[/green]")
                    else:
                        self.console.print("[red]ドキュメントのアップロードに失敗しました。[/red]")
                except Exception as e:
                    self.console.print(f"[red]ドキュメントアップロードエラー: {e}[/red]")
                    
        except Exception as e:
            self.console.print(f"[red]ファイルアップロードエラー: {e}[/red]")
    
    def _upload_assistant_knowledge(self):
        """アシスタントにナレッジファイルを登録（RAG）"""
        try:
            assistants = self.client.get_assistants()
            if not assistants:
                self.console.print("[red]利用可能なアシスタントがありません。[/red]")
                return
            
            # アシスタント選択
            self.console.print("\n[bold]アシスタントを選択してください:[/bold]")
            for i, assistant in enumerate(assistants, 1):
                self.console.print(f"{i}. {assistant.get('name', 'N/A')}")
            choice = Prompt.ask("選択", choices=[str(i) for i in range(1, len(assistants) + 1)])
            selected = assistants[int(choice) - 1]
            assistant_uid = selected.get('uid') or selected.get('id')
            if not assistant_uid:
                self.console.print("[red]アシスタントUIDが取得できません。[/red]")
                return
            
            # ファイルパス入力
            self.console.print("\n[bold cyan]登録するファイルパスを入力してください（.docx/.xls/.xlsx/.pptx/.pdf/.txt、16MB以下）[/bold cyan]")
            self._show_available_files()
            file_path_input = input("ファイルパス: ").strip()
            if not file_path_input:
                self.console.print("[yellow]キャンセルしました。[/yellow]")
                return
            # 入力のパスを正規化（引用符除去・展開・絶対化）
            file_path = self._normalize_path(file_path_input)
            
            try:
                file_uid = self.client.add_assistant_knowledge(assistant_uid, file_path)
                self.console.print("[green]ナレッジの登録が完了しました。[/green]")
                if file_uid:
                    self.console.print(f"ファイルID: {file_uid}")
                else:
                    self.console.print("[yellow]APIからファイルIDが返らない環境です。削除時はIDの指定が必要です。[/yellow]")
            except Exception as e:
                self.console.print(f"[red]ナレッジ登録エラー: {e}[/red]")
        except Exception as e:
            self.console.print(f"[red]ナレッジ登録処理エラー: {e}[/red]")

    def _delete_assistant_knowledge(self):
        """アシスタントのナレッジを削除"""
        try:
            assistants = self.client.get_assistants()
            if not assistants:
                self.console.print("[red]利用可能なアシスタントがありません。[/red]")
                return
            
            self.console.print("\n[bold]アシスタントを選択してください:[/bold]")
            for i, assistant in enumerate(assistants, 1):
                self.console.print(f"{i}. {assistant.get('name', 'N/A')}")
            choice = Prompt.ask("選択", choices=[str(i) for i in range(1, len(assistants) + 1)])
            selected = assistants[int(choice) - 1]
            assistant_uid = selected.get('uid') or selected.get('id')
            if not assistant_uid:
                self.console.print("[red]アシスタントUIDが取得できません。[/red]")
                return
            
            # ナレッジ一覧を取得
            try:
                files = self.client.get_assistant_knowledge(assistant_uid)
            except Exception as e:
                self.console.print(f"[red]ナレッジ一覧取得エラー: {e}[/red]")
                return

            if not files:
                self.console.print("[yellow]このアシスタントにはナレッジが登録されていません。[/yellow]")
                return

            # 一覧表示
            table = Table(title="登録ナレッジ")
            table.add_column("No.", style="yellow")
            table.add_column("ID", style="green")
            table.add_column("名前", style="cyan")
            table.add_column("更新日", style="white")
            for idx, f in enumerate(files, 1):
                fid = f.get('uid') or f.get('id') or f.get('file_uid') or f.get('file_id') or 'N/A'
                name = f.get('name') or f.get('file_name') or f.get('title') or 'N/A'
                updated = f.get('updated_at') or f.get('modified_at') or f.get('created_at') or 'N/A'
                table.add_row(str(idx), str(fid), str(name), str(updated))
            self.console.print(table)

            # 削除対象選択
            max_choice = str(len(files))
            sel = Prompt.ask("削除するNo.を選択", choices=[str(i) for i in range(1, len(files)+1)])
            target = files[int(sel)-1]
            file_uid = target.get('uid') or target.get('id') or target.get('file_uid') or target.get('file_id')
            if not file_uid:
                self.console.print("[red]選択した項目のファイルIDが取得できません。[/red]")
                return
            ok = self.client.delete_assistant_knowledge(assistant_uid, str(file_uid))
            self.console.print("[green]削除に成功しました。[/green]" if ok else "[red]削除に失敗しました。[/red]")
        except Exception as e:
            self.console.print(f"[red]ナレッジ削除処理エラー: {e}[/red]")

    def _show_settings(self):
        """設定を表示"""
        try:
            config = self.config_manager.get_config()
            
            self.console.print("\n[bold]現在の設定:[/bold]")
            settings_table = Table(title="設定情報")
            settings_table.add_column("項目", style="cyan")
            settings_table.add_column("値", style="white")
            
            settings_table.add_row("API Base URL", config.api_base_url)
            settings_table.add_row("Company Subdomain", config.company_subdomain or "未設定")
            settings_table.add_row("Client ID", config.client_id or "未設定")
            settings_table.add_row("Tenant ID", config.tenant_id or "未設定")
            settings_table.add_row("Authority", config.authority or "未設定")
            settings_table.add_row("Auth Mode", config.auth_mode)
            settings_table.add_row("Use PKCE", str(config.use_pkce))
            settings_table.add_row("Token File", config.token_file)
            settings_table.add_row("Timeout", str(config.timeout))
            settings_table.add_row("Max Retries", str(config.max_retries))
            
            self.console.print(settings_table)
            
            # 認証状態も表示
            self.console.print("\n[bold]認証状態:[/bold]")
            auth_table = Table(title="認証情報")
            auth_table.add_column("項目", style="cyan")
            auth_table.add_column("値", style="white")
            
            auth_table.add_row("認証済み", "はい" if self.auth_manager.is_authenticated() else "いいえ")
            auth_table.add_row("セッションCookie数", str(len(self.auth_manager.session_cookies)))
            auth_table.add_row("トークン情報", "あり" if self.auth_manager.tokens else "なし")
            
            self.console.print(auth_table)
            
        except Exception as e:
            self.console.print(f"[red]設定表示エラー: {e}[/red]")

    def _show_available_files(self):
        """現在のディレクトリのアップロード可能なファイルを表示"""
        try:
            files = os.listdir('.')
            uploadable_files = []
            
            for file in files:
                if os.path.isfile(file):
                    ext = os.path.splitext(file)[1].lower()
                    if ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.pdf', '.docx', '.txt', '.pptx', '.xls', '.xlsx', '.md', '.wav', '.mp3', '.aiff', '.aac', '.ogg', '.flac']:
                        uploadable_files.append(file)
            
            if uploadable_files:
                self.console.print("\n[bold green]利用可能なファイル:[/bold green]")
                for file in sorted(uploadable_files):
                    ext = os.path.splitext(file)[1].lower()
                    if ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']:
                        self.console.print(f"  📷 {file} (画像)")
                    elif ext in ['.wav', '.mp3', '.aiff', '.aac', '.ogg', '.flac']:
                        self.console.print(f"  🎵 {file} (音声)")
                    else:
                        self.console.print(f"  📄 {file} (ドキュメント)")
                self.console.print("")
        except Exception as e:
            # エラーが発生しても処理を続行
            pass

    def _show_available_audio_files(self):
        """現在のディレクトリの音声ファイルを表示"""
        try:
            files = os.listdir('.')
            audio_files = []
            audio_extensions = {'.wav', '.mp3', '.aiff', '.aac', '.ogg', '.flac'}
            
            for file in files:
                if os.path.isfile(file):
                    ext = os.path.splitext(file)[1].lower()
                    if ext in audio_extensions:
                        audio_files.append(file)
            
            if audio_files:
                self.console.print("\n[bold green]利用可能な音声ファイル:[/bold green]")
                for file in sorted(audio_files):
                    file_size = os.path.getsize(file)
                    self.console.print(f"  🎵 {file} ({file_size / 1024:.1f} KB)")
                self.console.print("")
            else:
                self.console.print("[yellow]現在のディレクトリに音声ファイルが見つかりません。[/yellow]\n")
        except Exception as e:
            # エラーが発生しても処理を続行
            pass


def main():
    """メイン関数"""
    cli = NewtonXCLIExample()
    cli.run()


if __name__ == "__main__":
    main() 