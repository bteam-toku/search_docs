from abc import ABCMeta, abstractmethod

class AbstractSearch(metaclass=ABCMeta):
    """ドキュメント検索抽象基底クラス
    """
    #
    # constructor/destructor
    #
    def __init__(self) -> None:
        """コンストラクタ
        """
        pass

    def __del__(self) -> None:
        """デストラクタ
        """
        pass

    #
    # abstract public methods
    #
    @abstractmethod
    def search(self, target_path:str, keywards:list=None, enable_search_shapes:bool=False) -> bool:
        """ドキュメント検索処理
        Args:
            target_path (str): 検索対象パス
            keywards (list): 検索キーワードリスト(Noneの場合は要素検索のみ実行)
            enable_search_shapes (bool): 図形内検索を有効にするかどうか
        
        Returns:
            bool: True:成功, False:失敗
        """
        pass

    @abstractmethod
    def save_results(self, output_path:str) -> bool:
        """検索結果保存処理
        Args:
            output_path (str): 出力パス
        
        Returns:
            bool: True:成功, False:失敗
        """
        pass
