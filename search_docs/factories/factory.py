from search_docs.interfaces import AbstractSearch
from search_docs.interfaces import AbstractSearchDocs
from search_docs.config import Config
from typing import Type, Optional, List
import importlib

class Factory:
    """検索ドキュメントファクトリークラス
    """
    _instance : Optional[AbstractSearch] = None
    _cached_type : Optional[str] = None

    #
    # コンストラクタ / デストラクタ
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
    # public methods
    #
    @classmethod
    def create(cls, adaptor_type_name: Optional[str] = None, config: Optional[Config] = None) -> AbstractSearch:
        """ 検索ドキュメントアダプター生成メソッド

        Args:
            adaptor_type_name (Optional[str], optional): アダプターの型名. デフォルトはNone.
        Returns:
            AbstractSearch: AbstractSearchオブジェクト
        """
        # 同じ型のアダプターがキャッシュされている場合はそれを返す（シングルトン）
        if cls._instance is not None and cls._cached_type == adaptor_type_name:
            return cls._instance

        config = config if config else Config()
        if adaptor_type_name is None:
            # デフォルトで必要なモジュールをインポート
            from search_docs.adaptors import DefaultSearchAdapter
            from search_docs.search_docs import DefaultSearchExcel
            # adaptor_type_nameが指定されていない場合はデフォルトのアダプターを使用
            # デフォルトのドキュメント検索クラスリストを作成
            default_search_docs: List[AbstractSearchDocs] = [DefaultSearchExcel(config.get("progress_display", True))]
            # デフォルトのアダプターを生成
            cls._instance = DefaultSearchAdapter(default_search_docs)
            cls._cached_type = adaptor_type_name
        else:
            # 指定された型名からアダプタークラスを動的にインポートして生成
            module_path, class_name = adaptor_type_name.rsplit('.', 1)
            module = importlib.import_module(module_path)
            adaptor_class = getattr(module, class_name)
            cls._instance = adaptor_class()
            cls._cached_type = adaptor_type_name

        # 生成したアダプターを返す
        return cls._instance
