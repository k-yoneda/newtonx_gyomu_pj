#!/usr/bin/env python3
"""
NewtonX Auto Runner Example

YAML設定ファイルに従って自動実行するサンプル
CLI版のcli_runner.pyに近い実装
"""

import sys
import os
import yaml
import argparse
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional
from datetime import datetime

# ADKのインポートを試行
try:
    from newtonx_adk import NewtonXClient, ConfigManager, AuthManager, NewtonXError, AuthenticationError, APIError
except ImportError as e:
    # ADKのパスを追加して再試行
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(current_dir)
    grandparent_dir = os.path.dirname(parent_dir)
    
    # ADKディレクトリの絶対パスを追加（並列階層）
    adk_path = os.path.join(grandparent_dir, 'newtonx_adk')
    if adk_path not in sys.path:
        sys.path.insert(0, adk_path)
    
    # デバッグ用
    print(f"ADK path added: {adk_path}")
    
    try:
        # 直接ADKモジュールからインポート
        from client import NewtonXClient
        from config import ConfigManager
        from auth import AuthManager
        from exceptions import NewtonXError, AuthenticationError, APIError
    except ImportError as e2:
        print(f"ADKのインポートに失敗しました: {e2}")
        print("ADKが正しくインストールされているか確認してください。")
        sys.exit(1)

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.prompt import Confirm
import click


class AutoRunner:
    """NewtonX Auto Runner メインクラス"""
    
    def __init__(self, config_path: str):
        """初期化"""
        self.console = Console()
        self.config_path = config_path
        self.config = self._load_config()
        self.setup_logging()
        
        # ADKクライアントの初期化
        self.config_manager = ConfigManager()
        self.auth_manager = AuthManager(self.config_manager)
        self.client = NewtonXClient(self.config_manager)
        # 既定のアシスタント名（設定未指定時のフォールバック含む）
        self.default_assistant_name: str = (
            (self.config.get('chat') or {}).get('default_assistant')
            or '高速アシスタント（GPT-4o mini）'
        )
        # アップロード後の待機秒数（メッセージ送信前に待機）
        try:
            self.wait_after_upload_seconds: int = int((self.config.get('chat') or {}).get('wait_after_upload_seconds', 0))
        except Exception:
            self.wait_after_upload_seconds = 0
        
        # 実行結果の保存
        self.results = []
        self.current_chat_uid = None
        self.uploaded_files = []  # アップロードしたファイルの情報を保存
        
    def _load_config(self) -> Dict[str, Any]:
        """設定ファイルを読み込み"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            self.console.print(f"[green]設定ファイルを読み込みました: {self.config_path}[/green]")
            return config
        except Exception as e:
            self.console.print(f"[red]設定ファイルの読み込みに失敗しました: {e}[/red]")
            sys.exit(1)
    
    def setup_logging(self):
        """ログ設定"""
        log_config = self.config.get('logging', {})
        log_level = getattr(logging, log_config.get('level', 'INFO'))
        log_file = log_config.get('file', 'auto_runner.log')
        
        logging.basicConfig(
            level=log_level,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def run(self):
        """メイン実行"""
        try:
            self.console.print(Panel.fit(
                f"[bold blue]{self.config['project']['name']}[/bold blue]\n"
                f"[dim]{self.config['project']['description']}[/dim]\n"
                f"Version: {self.config['project']['version']}",
                title="NewtonX Auto Runner"
            ))
            
            # 認証
            if not self._authenticate():
                self.console.print("[red]認証に失敗しました。[/red]")
                return False
            
            # タスク実行
            tasks = self.config.get('tasks', [])
            enabled_tasks = [task for task in tasks if task.get('enabled', True)]
            
            if not enabled_tasks:
                self.console.print("[yellow]実行可能なタスクがありません。[/yellow]")
                return True
            
            self.console.print(f"\n[bold]実行するタスク: {len(enabled_tasks)}個[/bold]")
            
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                console=self.console
            ) as progress:
                for i, task in enumerate(enabled_tasks, 1):
                    task_id = progress.add_task(
                        f"[{i}/{len(enabled_tasks)}] {task['name']}", 
                        total=None
                    )
                    
                    try:
                        result = self._execute_task(task)
                        self.results.append({
                            'task': task['name'],
                            'status': 'success',
                            'result': result,
                            'timestamp': datetime.now().isoformat()
                        })
                        progress.update(task_id, description=f"[green]✓ {task['name']}[/green]")
                        
                    except Exception as e:
                        self.logger.error(f"タスク実行エラー: {task['name']} - {e}")
                        self.results.append({
                            'task': task['name'],
                            'status': 'error',
                            'error': str(e),
                            'timestamp': datetime.now().isoformat()
                        })
                        progress.update(task_id, description=f"[red]✗ {task['name']}[/red]")
            
            # 結果表示
            self._show_results()
            
            # 結果保存
            if self.config.get('output', {}).get('save_results', True):
                self._save_results()
            
            return True
            
        except Exception as e:
            self.logger.error(f"実行エラー: {e}")
            self.console.print(f"[red]実行エラー: {e}[/red]")
            return False
    
    def _authenticate(self) -> bool:
        """認証を実行（内部トークン優先・MSALはインタラクティブのみ）"""
        try:
            # 既に内部トークン or Cookie があれば完了
            if self.auth_manager.is_authenticated():
                self.console.print("[green]認証済みです。[/green]")
                # 認証済みでもBearerが無い（Cookieのみ）の場合は内部トークンの貼り付けを促す
                try:
                    headers = self.auth_manager._get_headers()  # type: ignore[attr-defined]
                    if 'Authorization' not in headers:
                        self.console.print("[yellow]Bearerトークンが見つからないため、内部トークンの貼り付けを促します。[/yellow]")
                        pasted = input("PATを貼り付けてください（WEBアプリで発行）: ").strip()
                        import re as _re
                        def _extract_token(text: str):
                            m = _re.search(r"access_token=([^&\s]+)", text)
                            if m:
                                return m.group(1)
                            if '|' in text and ' ' not in text:
                                return text
                            return None
                        internal_token = _extract_token(pasted)
                        if internal_token:
                            self.auth_manager.set_newtonx_access_token(internal_token)
                            self.console.print("[cyan]内部トークンを保存しました。セッションをブートストラップします...[/cyan]")
                            try:
                                self.auth_manager.bootstrap_session()
                            except Exception:
                                pass
                except Exception:
                    pass
                return True

            # 設定確認
            config = self.config_manager.get_config()
            if not config.client_id or not config.tenant_id:
                self.console.print("[red]認証設定が不完全です。[/red]")
                self.console.print("[yellow]以下のコマンドで認証設定を行ってください:[/yellow]")
                self.console.print("[cyan]python tools/setup_config.py[/cyan]")
                return False

            # PAT方式に統一（ブラウザ拡張不要）
            try:
                import webbrowser
                web_base = (config.api_base_url or 'https://seraku.newton-x.net/api').rsplit('/api', 1)[0]
                sub = config.company_subdomain or ''
                login_url = f"{web_base}/account/redirectadlogin" + (f"?subdomain={sub}" if sub else '')
                webbrowser.open(login_url)
                self.console.print(f"[cyan]NewtonXログインページを開きました: {login_url}[/cyan]")
            except Exception:
                pass

            # クリップボードコピー済みの内部トークン（またはURL）を貼り付け
            pasted = input("PATを貼り付けてください（WEBアプリで発行）: ").strip()
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
                self.console.print("[red]内部トークンが抽出できませんでした。[/red]")
                return False

            # 保存→ブートストラップ
            self.auth_manager.set_newtonx_access_token(internal_token)
            self.console.print("[cyan]内部トークンを保存しました。セッションをブートストラップします...[/cyan]")
            ok = self.auth_manager.bootstrap_session()
            if not ok:
                self.console.print("[yellow]ブートストラップに失敗しましたが続行します。[/yellow]")

            return True

        except Exception as e:
            self.console.print(f"[red]認証エラー: {e}[/red]")
            return False
    
    def _execute_task(self, task: Dict[str, Any]) -> Any:
        """タスクを実行"""
        task_type = task['type']
        task_name = task['name']
        
        self.logger.info(f"タスク実行開始: {task_name} ({task_type})")
        
        if task_type == 'init':
            return self._task_init(task)
        elif task_type == 'upload':
            return self._task_upload(task)
        elif task_type == 'create_chat':
            return self._task_create_chat(task)
        elif task_type == 'send_message':
            return self._task_send_message(task)
        elif task_type == 'analyze_file':
            return self._task_analyze_file(task)
        else:
            raise ValueError(f"未知のタスクタイプ: {task_type}")
    
    def _task_init(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """初期化タスク"""
        self.console.print(f"  [dim]初期化: {task['description']}[/dim]")
        
        # アシスタント一覧取得
        assistants = self.client.get_assistants()
        
        return {
            'assistants_count': len(assistants),
            'assistants': [a['name'] for a in assistants[:5]]  # 最初の5つ
        }
    
    def _task_upload(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """ファイルアップロードタスク"""
        self.console.print(f"  [dim]ファイルアップロード: {task['description']}[/dim]")
        
        files = task.get('files', [])
        uploaded_files = []
        
        # チャットが必要な場合は作成
        if not self.current_chat_uid:
            chat_uid = self._create_default_chat()
            self.current_chat_uid = chat_uid
        
        for file_path in files:
            if os.path.exists(file_path):
                # ファイルアップロード
                file_ext = os.path.splitext(file_path)[1].lower()
                if file_ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']:
                    result = self.client.upload_image(self.current_chat_uid, file_path)
                    uploaded_files.append({
                        'file': file_path,
                        'type': 'image',
                        'result': result,
                        'image_id': result  # uidを保存
                    })
                else:
                    result = self.client.upload_document(self.current_chat_uid, file_path)
                    uploaded_files.append({
                        'file': file_path,
                        'type': 'document',
                        'result': result,
                        'document_id': result  # uidを保存
                    })
            else:
                self.console.print(f"  [yellow]ファイルが見つかりません: {file_path}[/yellow]")
        
        # アップロードしたファイルの情報を保存
        self.uploaded_files.extend(uploaded_files)
        
        # デバッグ用：アップロード結果をログ出力
        for file_info in uploaded_files:
            if file_info.get('document_id'):
                self.logger.info(f"ドキュメントアップロード成功: {file_info['file']} -> uid: {file_info['document_id']}")
            elif file_info.get('image_id'):
                self.logger.info(f"画像アップロード成功: {file_info['file']} -> uid: {file_info['image_id']}")
            else:
                self.logger.warning(f"ファイルアップロード成功だがuidが取得できません: {file_info['file']}")
        
        return {
            'uploaded_files': uploaded_files,
            'total_files': len(files),
            'chat_uid': self.current_chat_uid
        }
    
    def _task_create_chat(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """チャット作成タスク"""
        self.console.print(f"  [dim]チャット作成: {task['description']}[/dim]")
        
        title = task.get('title', 'Auto Runner Chat')
        preferred_name = task.get('assistant') or self.default_assistant_name

        assistant_uid, chosen_name = self._resolve_assistant_uid(preferred_name)

        if assistant_uid:
            chat_uid = self.client.create_chat(assistant_uid, title)
            self.current_chat_uid = chat_uid
            
            return {
                'chat_uid': chat_uid,
                'title': title,
                'assistant': chosen_name or preferred_name
            }
        else:
            raise ValueError("アシスタントが見つかりません")
    
    def _task_send_message(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """メッセージ送信タスク"""
        self.console.print(f"  [dim]メッセージ送信: {task['description']}[/dim]")
        
        message = task.get('message', '')
        web_search = task.get('web_search', False)
        knowledge_search = task.get('knowledge_search', False)
        
        if not self.current_chat_uid:
            chat_uid = self._create_default_chat()
            self.current_chat_uid = chat_uid
        
        # アップロード済みファイルがある場合、送信前に待機（インデックス反映の猶予）
        if self.uploaded_files and self.wait_after_upload_seconds > 0:
            try:
                self.console.print(f"  [dim]{self.wait_after_upload_seconds}秒待機してから送信します（アップロード反映待ち）[/dim]")
                import time as _t
                _t.sleep(self.wait_after_upload_seconds)
            except Exception:
                pass

        response = self.client.send_message(
            self.current_chat_uid,
            message,
            web_search=web_search,
            knowledge_search=knowledge_search
        )
        
        return {
            'message': message,
            'response': response,
            'web_search': web_search,
            'knowledge_search': knowledge_search
        }
    
    def _task_analyze_file(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """ファイル分析タスク（WEB版の送信ペイロードに準拠）"""
        self.console.print(f"  [dim]ファイル分析: {task['description']}[/dim]")
        
        # WEB版と同等のユーザーメッセージのみを送る
        message = task.get('message', 'このファイルの中身について教えてください。')
        
        if not self.current_chat_uid:
            chat_uid = self._create_default_chat()
            self.current_chat_uid = chat_uid
        
        # ファイル本文はメッセージへ埋め込まない（添付済みファイルをサーバ側で参照）
        # document_ids/image_ids も明示送信しない（チャットに紐づく添付をサーバ側が解決）
        response = self.client.send_message(
            self.current_chat_uid,
            message,
            web_search=False,
            knowledge_search=False,
        )
        
        return {
            'analysis_message': message,
            'response': response,
            'web_search': False,
            'knowledge_search': False,
            'uploaded_files': self.uploaded_files
        }
    
    def _create_default_chat(self) -> str:
        """デフォルトチャットを作成"""
        uid, _name = self._resolve_assistant_uid(self.default_assistant_name)
        if not uid:
            assistants = self.client.get_assistants()
            if not assistants:
                raise ValueError("アシスタントが見つかりません")
            uid = assistants[0]['uid']
        chat_uid = self.client.create_chat(uid, "Auto Runner Test Chat")
        return chat_uid

    # ===== アシスタント選択補助 =====
    def _resolve_assistant_uid(self, preferred_name: Optional[str]) -> tuple[Optional[str], Optional[str]]:
        """名前優先でアシスタントUIDを解決。見つからなければ近似→先頭を返す。"""
        try:
            assistants = self.client.get_assistants()
        except Exception as e:
            self.logger.error(f"アシスタント一覧取得エラー: {e}")
            return None, None

        if not assistants:
            return None, None

        # 1) 完全一致
        if preferred_name:
            for a in assistants:
                if a.get('name') == preferred_name:
                    return a.get('uid'), a.get('name')

        # 2) 代表パターンの近似一致（高速アシスタント / GPT-4o mini）
        keywords = []
        if preferred_name:
            pn = preferred_name.lower()
            if 'gpt-4o' in pn or 'mini' in pn or '高速' in pn:
                keywords = ['gpt-4o', 'mini', '高速']
        else:
            keywords = ['gpt-4o', 'mini', '高速']

        if keywords:
            for a in assistants:
                name_l = str(a.get('name', '')).lower()
                if all(k in name_l for k in keywords if k):
                    return a.get('uid'), a.get('name')
            for a in assistants:
                name_l = str(a.get('name', '')).lower()
                if any(k in name_l for k in keywords if k):
                    return a.get('uid'), a.get('name')

        # 3) 先頭
        a0 = assistants[0]
        return a0.get('uid'), a0.get('name')
    
    def _show_results(self):
        """結果を表示"""
        self.console.print("\n[bold]実行結果[/bold]")
        
        table = Table(title="タスク実行結果")
        table.add_column("タスク", style="cyan")
        table.add_column("ステータス", style="green")
        table.add_column("結果", style="white")
        
        for result in self.results:
            status_icon = "✓" if result['status'] == 'success' else "✗"
            status_color = "green" if result['status'] == 'success' else "red"
            
            table.add_row(
                result['task'],
                f"[{status_color}]{status_icon}[/{status_color}]",
                str(result.get('result', result.get('error', '')))[:50] + "..."
            )
        
        self.console.print(table)
    
    def _save_results(self):
        """結果を保存"""
        output_config = self.config.get('output', {})
        output_dir = output_config.get('output_dir', 'results')
        output_format = output_config.get('format', 'rich')
        
        # 出力ディレクトリを作成
        os.makedirs(output_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if output_format == 'json':
            import json
            output_file = os.path.join(output_dir, f"results_{timestamp}.json")
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(self.results, f, ensure_ascii=False, indent=2)
        else:
            output_file = os.path.join(output_dir, f"results_{timestamp}.txt")
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"NewtonX Auto Runner Results\n")
                f.write(f"Generated: {datetime.now().isoformat()}\n\n")
                
                for result in self.results:
                    f.write(f"Task: {result['task']}\n")
                    f.write(f"Status: {result['status']}\n")
                    f.write(f"Timestamp: {result['timestamp']}\n")
                    if 'result' in result:
                        f.write(f"Result: {result['result']}\n")
                    if 'error' in result:
                        f.write(f"Error: {result['error']}\n")
                    f.write("-" * 50 + "\n")
        
        self.console.print(f"[green]結果を保存しました: {output_file}[/green]")


@click.command()
@click.option('--config', '-c', default='config.yml', help='設定ファイルのパス')
@click.option('--dry-run', is_flag=True, help='ドライラン（実際の実行は行わない）')
@click.option('--verbose', '-v', is_flag=True, help='詳細出力')
def main(config, dry_run, verbose):
    """NewtonX Auto Runner Example"""
    if not os.path.exists(config):
        click.echo(f"設定ファイルが見つかりません: {config}")
        sys.exit(1)
    
    if dry_run:
        click.echo("ドライランモード: 実際の実行は行いません")
        return
    
    runner = AutoRunner(config)
    success = runner.run()
    
    if success:
        click.echo("実行が完了しました。")
    else:
        click.echo("実行中にエラーが発生しました。")
        sys.exit(1)


if __name__ == "__main__":
    main() 