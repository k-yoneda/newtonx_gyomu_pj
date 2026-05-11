#!/usr/bin/env python3
"""
NewtonX ADK 設定確認（PAT）

Host と PAT の設定状態と、Authorization ヘッダー生成可否を確認します。
"""

import sys
from pathlib import Path

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src"))

from newtonx_adk.config import ConfigManager
from newtonx_adk.auth import AuthManager


def main() -> int:
    print("=== NewtonX ADK 設定確認（PAT） ===")
    cfgm = ConfigManager()
    cfg = cfgm.get_config()

    print("現在の設定:")
    print(f"  Host            : {cfg.host or '(未設定)'}")
    print(f"  API Base URL    : {cfg.api_base_url}")
    print(f"  CompanySubdomain: {cfg.company_subdomain or '(未設定)'}")
    print(f"  PAT             : {'(設定済み)' if (getattr(cfg, 'personal_access_token', '') or '').strip() else '(未設定)'}")

    am = AuthManager(cfgm)
    headers = am.get_headers()
    print("\nヘッダー確認:")
    auth = headers.get('Authorization', '')
    if auth:
        print("  Authorization: 設定あり")
        print("  ✓ PATに基づくヘッダー生成を確認しました")
        print("\n確認完了")
        return 0
    else:
        print("  Authorization: 設定なし")
        print("  ✗ PATが設定されていないか不正です。setup_config.py を再実行してください")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())