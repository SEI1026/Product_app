#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
テスト実行スクリプト

使用例:
    python run_tests.py              # 全テスト実行
    python run_tests.py --unit       # 単体テストのみ
    python run_tests.py --gui         # GUIテストのみ
    python run_tests.py --cov         # カバレッジ付きテスト
"""
import sys
import subprocess
import argparse


def run_tests(test_type=None, coverage=False, verbose=True):
    """テストを実行"""
    
    # pytestコマンドの基本構成
    cmd = ["python", "-m", "pytest"]
    
    # 詳細出力
    if verbose:
        cmd.append("-v")
    
    # テストタイプによる分岐
    if test_type == "unit":
        cmd.extend(["-m", "unit"])
    elif test_type == "gui":
        cmd.extend(["-m", "gui"])
    elif test_type == "integration":
        cmd.extend(["-m", "integration"])
    elif test_type == "slow":
        cmd.extend(["-m", "slow"])
    elif test_type == "not-slow":
        cmd.extend(["-m", "not slow"])
    
    # カバレッジ測定
    if coverage:
        cmd.extend([
            "--cov=.",
            "--cov-report=html",
            "--cov-report=term-missing",
            "--cov-exclude=tests/*"
        ])
    
    # テストディレクトリを指定
    cmd.append("tests/")
    
    print(f"実行コマンド: {' '.join(cmd)}")
    print("-" * 50)
    
    # テスト実行
    try:
        result = subprocess.run(cmd, check=False)
        return result.returncode
    except KeyboardInterrupt:
        print("\nテストが中断されました")
        return 1
    except Exception as e:
        print(f"テスト実行エラー: {e}")
        return 1


def main():
    """メイン処理"""
    parser = argparse.ArgumentParser(description="商品登録入力ツール テスト実行スクリプト")
    
    # テストタイプのオプション
    test_group = parser.add_mutually_exclusive_group()
    test_group.add_argument("--unit", action="store_const", const="unit", dest="test_type",
                           help="単体テストのみ実行")
    test_group.add_argument("--gui", action="store_const", const="gui", dest="test_type",
                           help="GUIテストのみ実行")
    test_group.add_argument("--integration", action="store_const", const="integration", dest="test_type",
                           help="統合テストのみ実行")
    test_group.add_argument("--slow", action="store_const", const="slow", dest="test_type",
                           help="時間のかかるテストのみ実行")
    test_group.add_argument("--fast", action="store_const", const="not-slow", dest="test_type",
                           help="高速なテストのみ実行")
    
    # その他のオプション
    parser.add_argument("--cov", action="store_true", help="カバレッジ測定を有効化")
    parser.add_argument("--quiet", action="store_true", help="詳細出力を無効化")
    
    args = parser.parse_args()
    
    # テスト実行
    return_code = run_tests(
        test_type=args.test_type,
        coverage=args.cov,
        verbose=not args.quiet
    )
    
    # 結果表示
    if return_code == 0:
        print("\n✅ 全てのテストが成功しました！")
    else:
        print(f"\n❌ テストが失敗しました (終了コード: {return_code})")
    
    return return_code


if __name__ == "__main__":
    sys.exit(main())
