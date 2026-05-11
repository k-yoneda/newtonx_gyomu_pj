#!/usr/bin/env python3
"""
NewtonX ADK セットアップ（PAT方式）

必要な設定は以下の2つのみ:
- Host（例: seraku.newton-x.net）
- Personal Access Token（NewtonX WEBアプリで発行したADK用トークン）
"""

import sys
from pathlib import Path

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src"))

from newtonx_adk.config import ConfigManager
from newtonx_adk.auth import AuthManager


def main() -> int:
    print("=== NewtonX ADK セットアップ（PAT） ===")
    cfgm = ConfigManager()
    current = cfgm.get_config()

    # サブドメイン入力（ドメインは newton-x.net 固定）
    while True:
        default_sub = current.company_subdomain or (current.host.split('.')[0] if getattr(current, 'host', '') else '')
        sub_in = input(f"Company Subdomain（例: seraku）[{default_sub}]: ").strip()
        subdomain = sub_in or default_sub
        if subdomain:
            break
        print("Company Subdomain は必須です。例: seraku")

    # ホストは固定ドメインへ結合
    host = f"{subdomain}.newton-x.net"

    # PAT 入力
    while True:
        pat_in = input("Personal Access Token (PAT): ").strip()
        if pat_in:
            pat = pat_in
            break
        print("Personal Access Token は必須です。NewtonXのWEBアプリから発行してください。")

    # 設定反映（host-firstで api_base_url/company_subdomain が自動派生）
    print("\n設定を保存しています...")
    cfgm.update_config(host=host, personal_access_token=pat, auth_mode='none')

    updated = cfgm.get_config()
    print("\n更新された設定:")
    print(f"  Host            : {updated.host}")
    print(f"  API Base URL    : {updated.api_base_url}")
    print(f"  CompanySubdomain: {updated.company_subdomain}")
    print(f"  PAT             : {'(設定済み)'}")

    # ヘッダー簡易確認
    try:
        am = AuthManager(cfgm)
        headers = am.get_headers()
        auth_disp = (headers.get('Authorization') or '')
        if auth_disp:
            print("\nAuthorization ヘッダーを確認しました。設定は完了です。")
        else:
            print("\nAuthorization ヘッダーが生成されませんでした。PATを再確認してください。")
    except Exception as e:
        print(f"ヘッダー確認時のエラー: {e}")

    print("\n✅ セットアップ完了")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())