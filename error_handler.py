"""
商品登録入力ツール - エラーハンドリング強化モジュール
"""
import logging
import traceback
from typing import Optional, Callable, Any
from functools import wraps


class ErrorHandler:
    """統一されたエラーハンドリングクラス"""
    
    @staticmethod
    def safe_execute(func: Callable, default_return: Any = None, 
                    error_msg: Optional[str] = None) -> Any:
        """
        関数を安全に実行し、エラー時はデフォルト値を返す
        """
        try:
            return func()
        except Exception as e:
            if error_msg:
                logging.error(f"{error_msg}: {e}")
            else:
                logging.error(f"関数実行エラー {func.__name__}: {e}")
            logging.debug(traceback.format_exc())
            return default_return
    
    @staticmethod
    def safe_method_call(obj: Any, method_name: str, *args, **kwargs) -> Any:
        """
        オブジェクトのメソッドを安全に呼び出す
        """
        try:
            if hasattr(obj, method_name):
                method = getattr(obj, method_name)
                if callable(method):
                    return method(*args, **kwargs)
            return None
        except Exception as e:
            logging.warning(f"メソッド {method_name} の呼び出しでエラー: {e}")
            return None
    
    @staticmethod
    def safe_file_operation(operation: Callable, filepath: str, 
                          error_msg: Optional[str] = None) -> Any:
        """
        ファイル操作を安全に実行する
        """
        try:
            return operation()
        except FileNotFoundError:
            logging.error(f"ファイルが見つかりません: {filepath}")
            return None
        except PermissionError:
            logging.error(f"ファイルアクセス権限がありません: {filepath}")
            return None
        except UnicodeDecodeError as e:
            logging.error(f"ファイルエンコーディングエラー {filepath}: {e}")
            return None
        except Exception as e:
            msg = error_msg or f"ファイル操作エラー {filepath}"
            logging.error(f"{msg}: {e}")
            return None


def safe_execution(default_return=None, error_msg=None):
    """
    関数を安全実行するデコレータ
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            return ErrorHandler.safe_execute(
                lambda: func(*args, **kwargs),
                default_return,
                error_msg
            )
        return wrapper
    return decorator


def log_performance(func):
    """
    関数の実行時間をログに記録するデコレータ
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        import time
        start_time = time.time()
        try:
            result = func(*args, **kwargs)
            execution_time = time.time() - start_time
            if execution_time > 1.0:  # 1秒以上の場合のみログ出力
                logging.info(f"{func.__name__} 実行時間: {execution_time:.2f}秒")
            return result
        except Exception as e:
            execution_time = time.time() - start_time
            logging.error(f"{func.__name__} エラー (実行時間: {execution_time:.2f}秒): {e}")
            raise
    return wrapper