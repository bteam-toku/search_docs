from abc import ABCMeta, abstractmethod
import pandas as pd

class AbstractSearchDocs(metaclass=ABCMeta):
    """ドキュメント検索抽象基底クラス
    """
    _pd_target : pd.DataFrame = None # 対象検索結果
    _pd_keyword : pd.DataFrame = None # キーワード検索結果

    def __init__(self) -> None:
        """コンストラクタ
        """
        pass

    def __del__(self) -> None:
        """デストラクタ
        """
        pass

    @abstractmethod
    def search_target(self, target_path:str, progress:bool=True) -> bool:
        """対象検索処理
        Args:
            target_path (str): 検索対象パス
        
        Returns:
            bool: True:成功, False:失敗
        """
        pass

    @abstractmethod
    def search_keyword(self, keywords:str, progress:bool=True) -> bool:
        """キーワード検索処理
        Args:
            keywords (str): 検索キーワード
            progress (bool): 進捗表示フラグ
        
        Returns:
            bool: True:成功, False:失敗
        """
        pass

    def get_target_list(self) -> pd.DataFrame:
        """対象検索結果を取得する
        Returns:
            pd.DataFrame: 対象検索結果のDataFrame
        """
        return self._pd_target

    def get_keyword_list(self) -> pd.DataFrame:
        """キーワード検索結果を取得する
        Returns:
            pd.DataFrame: キーワード検索結果のDataFrame
        """
        return self._pd_keyword
    
    def show_progress(self, ratio: int, task: str = "", status: str = ""):
        """進捗表示
        Args:
            ratio (int): 進捗率 (0-100)
            task (str): 現在のタスク名
            status (str): 現在の状態メッセージ
        """
        block_num = 50
        ratio = min(max(ratio, 0), 100)
        int_ratio = int(ratio * (block_num / 100))
        
        bar = '[' + '#' * int_ratio + '-' * (block_num - int_ratio) + ']'
        output = f"\r{bar} {ratio:3}% | {task} : {status}"
        print(f"{output:<100}", end="", flush=True)

        if ratio >= 100:
            print()    
