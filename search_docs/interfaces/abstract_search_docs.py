from abc import ABC, abstractmethod
import pandas as pd

class AbstractSearchDocs(ABC):
    """ドキュメント検索抽象基底クラス
    """
    # protected attributes
    _enable_progress: bool = True           # 進捗表示有無フラグ
    _doc_type: str = None                   # ドキュメントタイプ
    _extensions: list = []                  # 対応拡張子リスト
    _pd_element: pd.DataFrame = None        # ドキュメント要素検索結果データフレーム
    _pd_keyword: pd.DataFrame = None        # キーワード検索結果データフレーム

    #
    # constructor/destructor
    #
    def __init__(self, enable_progress: bool = True) -> None:
        """コンストラクタ
        """
        self._enable_progress = enable_progress
        
    def __del__(self) -> None:
        """デストラクタ
        """
        pass

    #
    # abstract public methods
    #
    @abstractmethod
    def search_element(self, target_path: str) -> bool:
        """ドキュメント要素検索処理
        Args:
            target_path (str): 検索対象パス
        
        Returns:
            bool: True:成功, False:失敗
        """
        pass

    @abstractmethod
    def search_keyword(self, keywords: list, enable_search_shapes: bool = False) -> bool:
        """キーワード検索処理

        このメソッドは、事前にsearch_elementメソッドが実行されていることを前提としています。

        Args:
            keywords (list): 検索キーワード
            enable_search_shapes (bool): 図形内テキスト検索有効フラグ

        Returns:
            bool: True:成功, False:失敗
        """
        pass

    def get_element_list(self) -> pd.DataFrame:
        """ドキュメント要素検索結果取得
        Returns:
            pd.DataFrame: ドキュメント要素検索結果データフレーム
        """
        return self._pd_element
    
    def get_keyword_list(self) -> pd.DataFrame:
        """キーワード検索結果取得
        Returns:
            pd.DataFrame: キーワード検索結果データフレーム
        """
        return self._pd_keyword
    
    def get_doc_type(self) -> str:
        """ドキュメントタイプ取得
        Returns:
            str: ドキュメントタイプ文字列
        """
        return self._doc_type