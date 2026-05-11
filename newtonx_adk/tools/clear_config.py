#!/usr/bin/env python3
"""
NewtonX ADK 設定クリア（PAT）

~/.newtonx 配下の設定・トークン（含PAT）をクリアします。
"""

import sys
from pathlib import Path
import shutil

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src"))

from newtonx_adk.config import ConfigManager
from newtonx_adk.auth import AuthManager


def main() -> int:
    print("=== NewtonX ADK 設定クリア（PAT） ===")

    cfgm = ConfigManager()
    am = AuthManager(cfgm)

    confirm = input("設定とトークンをすべて削除します。よろしいですか？ (y/N): ").strip().lower()
    if confirm not in ("y", "yes"):
        print("キャンセルしました。")
        return 0

    # ログアウトでトークン/キャッシュファイルを削除
    try:
        am.logout()
        print("✓ トークン・キャッシュを削除しました")
    except Exception as e:
        print(f"⚠️ ログアウト中のエラー: {e}")

    # ~/.newtonx ディレクトリも削除
    root = Path("~/.newtonx").expanduser()
    if root.exists():
        try:
            shutil.rmtree(root)
            print("✓ ~/.newtonx ディレクトリを削除しました")
        except Exception as e:
            print(f"⚠️ ディレクトリ削除エラー: {e}")
    else:
        print("ℹ️ ~/.newtonx は存在しません")

    print("\n✅ クリア完了")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())