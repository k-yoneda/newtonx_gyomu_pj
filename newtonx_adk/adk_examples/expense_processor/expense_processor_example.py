#!/usr/bin/env python3
"""
NewtonX ADKを使用した領収書解析アプリケーション

ADKを使用して領収書の自動解析を行うアプリケーションの例を示します。
"""

import os
import sys
import argparse
import csv
import json
from pathlib import Path
from typing import List, Dict, Optional
import logging
import re
import time

# ADKをインポート
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
    
    # 直接ADKモジュールからインポート
    from client import NewtonXClient
    from config import ConfigManager
    from auth import AuthManager
    from exceptions import NewtonXError, AuthenticationError, APIError


class ExpenseProcessorADK:
    """ADKを使用した経費精算支援クラス"""
    
    def __init__(self, config_path: str = None):
        """初期化"""
        # ADKの設定とクライアントを初期化
        self.config_manager = ConfigManager(config_path)
        self.auth_manager = AuthManager(self.config_manager)
        self.client = NewtonXClient(self.config_manager)
        
        # ロガーの初期化
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger('expense_processor_adk')
        
        # 認証チェック
        if not self._authenticate():
            raise Exception("認証に失敗しました。")
        
    def _authenticate(self) -> bool:
        """認証を実行（内部トークン優先・MSALはインタラクティブのみ）"""
        try:
            # 既に内部トークン or Cookie があれば完了
            is_auth = False
            # ダックタイピングで is_authenticated の有無に対応
            if hasattr(self.auth_manager, 'is_authenticated'):
                try:
                    is_auth = bool(self.auth_manager.is_authenticated())
                except Exception:
                    is_auth = False
            # 簡易fallback: get_access_token があれば試す
            if not is_auth and hasattr(self.auth_manager, 'get_access_token'):
                try:
                    token = self.auth_manager.get_access_token()
                    is_auth = bool(token)
                except Exception:
                    is_auth = False
            if is_auth:
                self.logger.info("認証済みです。")
                # 認証済みでもBearerが無い（Cookieのみ）の場合は内部トークンの入力を促す
                try:
                    headers = self.auth_manager._get_headers()  # type: ignore[attr-defined]
                    if 'Authorization' not in headers:
                        self.logger.info("Bearerトークンが見つからないため、内部トークンの貼り付けを促します。")
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
                            self.logger.info("内部トークンを保存しました。ブートストラップを実行します...")
                            try:
                                self.auth_manager.bootstrap_session()
                            except Exception:
                                pass
                except Exception:
                    pass
                return True

            # 設定を確認
            config = self.config_manager.get_config()
            if not config.client_id or not config.tenant_id:
                self.logger.error("認証設定が不完全です。")
                self.logger.error("以下のコマンドで認証設定を行ってください: python tools/setup_config.py")
                return False

            # PAT方式に統一（ブラウザ拡張不要）
            try:
                import webbrowser
                web_base = (config.api_base_url or 'https://seraku.newton-x.net/api').rsplit('/api', 1)[0]
                sub = config.company_subdomain or ''
                login_url = f"{web_base}/account/redirectadlogin" + (f"?subdomain={sub}" if sub else '')
                webbrowser.open(login_url)
                self.logger.info(f"NewtonXログインページを開きました: {login_url}")
            except Exception:
                pass

            # クリップボードでコピーされた内部トークン（またはURL）を貼り付けてもらう
            pasted = input("PATを貼り付けてください（WEBアプリで発行）: ").strip()
            def _extract_token(text: str) -> Optional[str]:
                m = re.search(r"access_token=([^&\s]+)", text)
                if m:
                    return m.group(1)
                if '|' in text and ' ' not in text:
                    return text
                return None
            internal_token = _extract_token(pasted)
            if not internal_token:
                self.logger.error("内部トークンが抽出できませんでした。")
                return False

            # 保存→ブートストラップ
            self.auth_manager.set_newtonx_access_token(internal_token)
            self.logger.info("内部トークンを保存しました。ブートストラップを実行します...")
            ok = self.auth_manager.bootstrap_session()
            if not ok:
                self.logger.warning("ブートストラップに失敗しましたが続行します。")

            return True

        except Exception as e:
            self.logger.error(f"認証エラー: {e}")
            return False
    
    def create_work_folder(self, folder_name: str) -> str:
        """作業用フォルダを作成"""
        try:
            folder_uid = self.client.create_folder(folder_name)
            if folder_uid:
                self.logger.info(f"作業フォルダを作成しました: {folder_name} (ID: {folder_uid})")
                return folder_uid
            else:
                raise Exception("フォルダ作成に失敗しました")
        except Exception as e:
            self.logger.error(f"フォルダ作成エラー: {e}")
            raise
    
    def upload_file(self, file_path: str, chat_uid: str) -> Optional[str]:
        """ファイルをアップロード"""
        try:
            file_path_obj = Path(file_path)
            file_extension = file_path_obj.suffix.lower()
            
            # ファイル形式に応じてアップロード方法を選択
            if file_extension in ['.jpg', '.jpeg', '.png', '.gif']:
                # 画像ファイルの場合はupload_imageメソッドを使用
                self.logger.info(f"画像ファイルをアップロード: {file_path}")
                
                try:
                    image_id = self.client.upload_image(chat_uid, file_path)
                    if image_id:
                        self.logger.info(f"画像ファイルをアップロードしました: {file_path} (ID: {image_id})")
                        return image_id
                    else:
                        self.logger.error(f"画像アップロードに失敗しました: {file_path}")
                        return None
                except Exception as e:
                    self.logger.error(f"画像アップロードエラー詳細: {e}")
                    return None
            else:
                # ドキュメントファイルの場合はupload_documentメソッドを使用
                self.logger.info(f"ドキュメントファイルをアップロード: {file_path}")
                success = self.client.upload_document(chat_uid, file_path)
                if success:
                    self.logger.info(f"ドキュメントファイルをアップロードしました: {file_path}")
                    return True
                else:
                    raise Exception("ドキュメントアップロードに失敗しました")
                    
        except Exception as e:
            self.logger.error(f"ファイルアップロードエラー: {e}")
            raise
    
    def extract_expense_info(self, chat_uid: str, message_response: Optional[str]) -> List[Dict]:
        """メッセージ応答から経費情報を抽出"""
        try:
            if message_response:
                return self._parse_response(message_response)
            self.logger.error("経費情報の抽出に失敗しました")
            return []
        except Exception as e:
            self.logger.error(f"経費情報抽出エラー: {e}")
            return []
    
    def _parse_response(self, response: str) -> List[Dict]:
        """レスポンスを解析"""
        try:
            # JSON形式のレスポンスを抽出
            import re
            
            # JSONブロックを検索
            json_pattern = r'\{[^{}]*\}'
            json_matches = re.findall(json_pattern, response)
            
            results = []
            for json_str in json_matches:
                try:
                    parsed = json.loads(json_str)
                    results.append(parsed)
                except json.JSONDecodeError:
                    self.logger.warning(f"JSON解析に失敗: {json_str}")
                    continue
            
            if not results:
                # JSONが見つからない場合は、レスポンス全体を解析
                try:
                    parsed = json.loads(response)
                    results.append(parsed)
                except json.JSONDecodeError:
                    self.logger.warning("レスポンス全体のJSON解析に失敗")
            
            return results
            
        except Exception as e:
            self.logger.error(f"レスポンス解析エラー: {e}")
            return []
    
    def process_files(self, input_dir: str, output_file: str):
        """ファイルを一括処理"""
        try:
            # アシスタント一覧を取得
            assistants = self.client.get_assistants()
            if not assistants:
                raise Exception("利用可能なアシスタントがありません")
            
            # アシスタント一覧を表示
            print("\n=== 利用可能なアシスタント一覧 ===")
            for i, a in enumerate(assistants, 1):
                print(f"  {i}. {a.get('name', '名前なし')} (UID: {a.get('uid', '不明')[:8]}...)")
            print("=" * 35)
            
            # ユーザーに選択させる
            while True:
                try:
                    choice = input(f"\n使用するアシスタントの番号を入力してください (1-{len(assistants)}): ").strip()
                    choice_num = int(choice)
                    if 1 <= choice_num <= len(assistants):
                        assistant = assistants[choice_num - 1]
                        self.logger.info(f"アシスタント '{assistant.get('name')}' を選択しました")
                        break
                    else:
                        print(f"1から{len(assistants)}の範囲で入力してください。")
                except ValueError:
                    print("数字を入力してください。")
            
            # フォルダ名を指定してチャットを作成（既存再利用/なければ作成）
            chat_uid = self.client.create_chat_in_folder_by_name(
                assistant_uid=assistant['uid'],
                folder_name="領収書読み取り",
                title=f"経費精算処理_{Path(input_dir).name}_{int(time.time())}"
            )
            
            if not chat_uid:
                raise Exception("チャットの作成に失敗しました")
            
            self.logger.info(f"チャットを作成しました: {chat_uid}")
            
            # 入力ディレクトリ内のファイルを処理
            input_path = Path(input_dir)
            if not input_path.exists():
                raise Exception(f"入力ディレクトリが存在しません: {input_dir}")
            
            all_results = []

            # 画像/ドキュメントに分類して先に収集
            image_exts = ['.jpg', '.jpeg', '.png', '.gif']
            doc_exts = ['.pdf']
            image_files = []
            doc_files = []

            for file_path in input_path.rglob('*'):
                if not file_path.is_file():
                    continue
                ext = file_path.suffix.lower()
                if ext in image_exts:
                    image_files.append(str(file_path))
                elif ext in doc_exts:
                    doc_files.append(str(file_path))

            prompt = """
            この領収書に記載されている日付（purchase_date）、金額（total_amount）、T番号の有無（has_t_number）、目的（purpose）、支払い先（vendor）をJSON形式で教えてください。領収書は複数のファイルで構成されている可能性がありますので、すべての領収書を読み取って教えてください。

            以下の形式で回答してください：
            {
                "purchase_date": "YYYY-MM-DD",
                "total_amount": 数値,
                "has_t_number": true/false,
                "purpose": "目的の説明",
                "vendor": "支払い先名"
            }
            """

            # 画像はまとめて送信
            if image_files:
                try:
                    self.logger.info(f"画像をまとめて送信します: {len(image_files)}件")
                    response = self.client.send_message_with_images(
                        chat_uid=chat_uid,
                        message=prompt,
                        image_file_paths=image_files,
                        knowledge_search=False,
                        web_search=False,
                    )
                    expense_info = self.extract_expense_info(chat_uid, response)

                    if expense_info:
                        if len(expense_info) == len(image_files):
                            for info, path in zip(expense_info, image_files):
                                p = Path(path)
                                info['file_name'] = p.name
                                info['file_path'] = str(p)
                                all_results.append(info)
                        else:
                            # 件数が一致しない場合はグループ情報として格納
                            grouped_names = [Path(p).name for p in image_files]
                            for info in expense_info:
                                info['file_group'] = ','.join(grouped_names)
                                info['file_count'] = len(grouped_names)
                                all_results.append(info)
                        self.logger.info(f"画像から経費情報を抽出しました: {len(expense_info)}件")
                    else:
                        self.logger.warning("画像からの経費情報抽出に失敗しました")
                except Exception as e:
                    self.logger.error(f"画像まとめ送信エラー: {e}")

            # ドキュメントは従来通り個別処理
            for doc_path in doc_files:
                try:
                    self.logger.info(f"ドキュメントを処理中: {doc_path}")
                    success = self.client.upload_document(chat_uid, doc_path)
                    if not success:
                        self.logger.error(f"ドキュメントアップロードに失敗: {doc_path}")
                        continue
                    response = self.client.send_message(
                        chat_uid=chat_uid,
                        message=prompt,
                        web_search=False,
                        knowledge_search=False,
                    )
                    expense_info = self.extract_expense_info(chat_uid, response)
                    if expense_info:
                        for info in expense_info:
                            p = Path(doc_path)
                            info['file_name'] = p.name
                            info['file_path'] = str(p)
                            all_results.append(info)
                        self.logger.info(f"ドキュメントから経費情報を抽出しました: {len(expense_info)}件")
                    else:
                        self.logger.warning(f"ドキュメントの経費情報抽出に失敗: {doc_path}")
                except Exception as e:
                    self.logger.error(f"ドキュメント処理エラー {doc_path}: {e}")
                    continue
            
            # 結果をCSVに出力
            if all_results:
                self._export_to_csv(all_results, output_file)
                self.logger.info(f"処理完了: {len(all_results)}件の経費情報を{output_file}に出力しました")
            else:
                self.logger.warning("抽出された経費情報がありません")
                
        except Exception as e:
            self.logger.error(f"ファイル処理エラー: {e}")
            raise
    
    def _export_to_csv(self, results: List[Dict], output_file: str):
        """結果をCSVに出力"""
        try:
            with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                if results:
                    fieldnames = results[0].keys()
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    
                    writer.writeheader()
                    for result in results:
                        writer.writerow(result)
                        
        except Exception as e:
            self.logger.error(f"CSV出力エラー: {e}")
            raise


def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(description='NewtonX ADKを使用した経費精算支援アプリケーション')
    parser.add_argument('input_dir', help='入力ディレクトリのパス')
    parser.add_argument('output_file', help='出力CSVファイルのパス')
    parser.add_argument('--config', help='設定ファイルのパス（オプション）')
    
    args = parser.parse_args()
    
    try:
        # 経費処理クラスを初期化
        processor = ExpenseProcessorADK(args.config)
        
        # ファイル処理を実行
        processor.process_files(args.input_dir, args.output_file)
        
        print(f"処理が完了しました。結果は {args.output_file} に出力されました。")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        sys.exit(1)


if __name__ == "__main__":
    import time
    main() 