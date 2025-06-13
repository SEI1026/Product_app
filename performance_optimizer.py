"""
商品登録入力ツール - パフォーマンス最適化モジュール
"""
import logging
import time
import gc
from typing import List, Dict, Any, Iterator
from contextlib import contextmanager


class PerformanceOptimizer:
    """パフォーマンス最適化ユーティリティ"""
    
    @staticmethod
    @contextmanager
    def memory_monitor(operation_name: str):
        """
        メモリ使用量をモニタリングするコンテキストマネージャ
        """
        try:
            import psutil
            process = psutil.Process()
            start_memory = process.memory_info().rss / 1024 / 1024  # MB
            
            yield
            
            end_memory = process.memory_info().rss / 1024 / 1024  # MB
            memory_diff = end_memory - start_memory
            
            if memory_diff > 50:  # 50MB以上増加した場合
                logging.warning(f"{operation_name}: メモリ使用量が {memory_diff:.1f}MB 増加しました")
            
        except ImportError:
            # psutilが利用できない場合はスキップ
            yield
        except Exception as e:
            logging.warning(f"メモリモニタリングエラー: {e}")
            yield
    
    @staticmethod
    def batch_process_data(data: List[Any], batch_size: int = 1000,
                          progress_callback=None) -> Iterator[List[Any]]:
        """
        大量データをバッチ処理する
        """
        total_batches = (len(data) + batch_size - 1) // batch_size
        
        for i in range(0, len(data), batch_size):
            batch = data[i:i + batch_size]
            
            if progress_callback and callable(progress_callback):
                batch_num = i // batch_size + 1
                progress_callback(batch_num, total_batches)
            
            yield batch
            
            # メモリ解放のためガベージコレクションを実行
            if i % (batch_size * 5) == 0:  # 5バッチごと
                gc.collect()
    
    @staticmethod
    def optimize_csv_reading(filepath: str, chunk_size: int = 10000):
        """
        CSVファイルの最適化された読み込み
        """
        import csv
        from utils import open_csv_file_with_fallback
        
        try:
            with open_csv_file_with_fallback(filepath, 'r') as (f, delimiter, encoding):
                reader = csv.DictReader(f, delimiter=delimiter)
                
                chunk = []
                for row_num, row in enumerate(reader, 1):
                    chunk.append(row)
                    
                    if len(chunk) >= chunk_size:
                        yield chunk
                        chunk = []
                        
                        # 定期的にガベージコレクション
                        if row_num % (chunk_size * 5) == 0:
                            gc.collect()
                
                # 残りのデータがある場合
                if chunk:
                    yield chunk
                    
        except Exception as e:
            logging.error(f"CSV読み込み最適化エラー {filepath}: {e}")
            raise
    
    @staticmethod
    def cache_manager(cache_dict: Dict[str, Any], max_size: int = 1000):
        """
        シンプルなLRUキャッシュマネージャ
        """
        if len(cache_dict) > max_size:
            # 古いキーを削除（簡易LRU）
            keys_to_remove = list(cache_dict.keys())[:-max_size//2]
            for key in keys_to_remove:
                cache_dict.pop(key, None)
            
            logging.info(f"キャッシュクリア: {len(keys_to_remove)}件のエントリを削除")
    
    @staticmethod
    def profile_function_execution(func, *args, **kwargs):
        """
        関数の実行プロファイリング
        """
        start_time = time.time()
        start_memory = 0
        
        try:
            import psutil
            process = psutil.Process()
            start_memory = process.memory_info().rss / 1024 / 1024
        except ImportError:
            pass
        
        try:
            result = func(*args, **kwargs)
            
            end_time = time.time()
            execution_time = end_time - start_time
            
            try:
                end_memory = process.memory_info().rss / 1024 / 1024
                memory_used = end_memory - start_memory
                
                logging.info(f"関数プロファイル {func.__name__}: "
                           f"実行時間={execution_time:.3f}s, "
                           f"メモリ使用量={memory_used:.1f}MB")
            except:
                logging.info(f"関数プロファイル {func.__name__}: "
                           f"実行時間={execution_time:.3f}s")
            
            return result
            
        except Exception as e:
            end_time = time.time()
            execution_time = end_time - start_time
            logging.error(f"関数実行エラー {func.__name__} "
                        f"(実行時間: {execution_time:.3f}s): {e}")
            raise


class DataProcessor:
    """大量データ処理用の最適化クラス"""
    
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback
        self._processed_count = 0
    
    def process_large_dataset(self, data: List[Dict], 
                            processor_func, batch_size: int = 1000):
        """
        大量データセットの効率的な処理
        """
        results = []
        total_items = len(data)
        
        with PerformanceOptimizer.memory_monitor(f"大量データ処理({total_items}件)"):
            for batch in PerformanceOptimizer.batch_process_data(data, batch_size):
                batch_results = []
                
                for item in batch:
                    try:
                        result = processor_func(item)
                        if result is not None:
                            batch_results.append(result)
                    except Exception as e:
                        logging.warning(f"データ処理エラー: {e}")
                        continue
                    
                    self._processed_count += 1
                    
                    # プログレス更新
                    if (self.progress_callback and 
                        self._processed_count % 100 == 0):
                        progress = (self._processed_count / total_items) * 100
                        self.progress_callback(progress)
                
                results.extend(batch_results)
                
                # バッチ処理後のメモリクリーンアップ
                del batch_results
                gc.collect()
        
        return results