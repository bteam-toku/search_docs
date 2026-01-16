from search_docs.interfaces import AbstractSearch
from search_docs.search_docs import AbstractSearchDocs
from typing import Optional, List, Type
import os

class DefaultSearchAdapter(AbstractSearch):
    """デフォルト検索アダプタークラス
    """
        #
        # protected attributes
    #
    _search_docs: Optional[List[AbstractSearchDocs]]    # 検索対象ドキュメント検索クラスのリスト

    def __init__(self, search_docs: Optional[List[AbstractSearchDocs]] = None) -> None:
        """コンストラクタ
        """
        self._search_docs = search_docs if search_docs else None

    #
    # public methods
    #
    def search(self, target_path:str, keywords:Optional[List[str]]=None, enable_search_shapes:bool=False) -> bool:
        """ドキュメント検索処理
        Args:
            target_path (str): 検索対象パス
            keywords (Optional[List[str]]): 検索キーワードリスト(Noneの場合は要素検索のみ実行)
            enable_search_shapes (bool): 図形内検索を有効にするかどうか
        
        Returns:
            bool: True:成功, False:失敗
        """
        # 検索対象ドキュメント検索クラスが設定されていない場合は失敗を返す
        if self._search_docs is None:
            return False
        
        # 検索対象ドキュメント検索クラスのリストをループ
        success = False
        for search_doc in self._search_docs:
            # ドキュメント要素検索処理を実行
            if not search_doc.search_element(target_path):
                continue
            
            # キーワード検索処理を実行
            if keywords is not None:
                if not search_doc.search_keyword(keywords, enable_search_shapes=enable_search_shapes):
                    continue
            
            # ひとつでも成功した場合は成功フラグをTrueに設定
            success = True

        # 復帰値を返す
        return success
    
    def save_results(self, output_path:str) -> bool:
        """検索結果保存処理
        Args:
            output_path (str): 出力パス（呼び出し元で作成すること）
        
        Returns:
            bool: True:成功, False:失敗
        """
        # 検索対象ドキュメント検索クラスが設定されていない場合は失敗を返す
        if self._search_docs is None:
            return False
        # 出力パスが存在しない場合は失敗を返す（呼び出し元で作成すること）
        if not os.path.exists(output_path):
            return False
        
        # 検索対象ドキュメント検索クラスのリストをループ
        success = False
        for search_doc in self._search_docs:
            # 検索結果保存処理を実行
            pd_element = search_doc.get_element_list()
            pd_keyword = search_doc.get_keyword_list()
            doc_type = search_doc.get_doc_type()
            # 検索結果を出力
            if pd_keyword is not None and not pd_keyword.empty:
                # キーワード検索結果が存在する場合は要素検索結果も含まれているのでキーワード検索結果を保存
                pd_keyword.to_csv(os.path.join(output_path, doc_type.lower()+'_search.csv'), encoding='utf-8-sig', index=False)
            elif pd_element is not None and not pd_element.empty:
                # 要素検索結果のみ存在する場合は要素検索結果を保存
                pd_element.to_csv(os.path.join(output_path, doc_type.lower()+'_search.csv'), encoding='utf-8-sig', index=False)
            else:
                continue
            # ひとつでも成功した場合は成功フラグをTrueに設定
            success = True

        # 復帰値を返す
        return success