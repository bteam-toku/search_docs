import os
import pathlib
import yaml
from typing import Any, Dict

class Config:
    """設定情報管理クラス
    """
    #
    # Constructor / Destructor
    #
    def __init__(self) -> None:
        """コンストラクタ
        """
        # ベースパスの初期化
        self._base_path = pathlib.Path.cwd()
        # 環境変数からDocker環境フラグの初期化
        self._is_docker = os.getenv('IS_DOCKER', 'false').lower() == 'true'
        # settings.yamlファイルパスの初期化
        if self._is_docker:
            # ローカル環境のssettings.yamlパスを優先的に使用する
            self._settings_file = pathlib.Path("/data/settings.yaml")
            if not os.path.exists(self._settings_file):
                # ローカル環境にsettings.yamlが存在しない場合はコンテナ内の設定を使用する
                self._settings_file = pathlib.Path("/app/settings.yaml")
        else:
            self._settings_file = self._base_path / "settings.yaml"
        # 設定データの読み込み
        self._config_data = self._load_settings()

    def __del__(self) -> None:
        """デストラクタ
        """
        pass

    #
    # public methods
    #
    def get(self, key: str, default=None):
        """設定値の取得

        Args:
            key (str): 設定キー
            default: デフォルト値（キーが存在しない場合に返される値）
        Returns:
            設定値またはデフォルト値
        """
        return self._config_data.get(key, default)
    
   
    def output_path(self) -> str:
        """出力パスの取得

        Returns:
            str: 出力パス
        """
        temp_path = self._config_data.get("output_path", "")
        if not temp_path:
            return temp_path
        
        if self._is_docker:
            # Docker環境の場合はそのまま返す
            return str(pathlib.Path('/data') / temp_path)
        else:
            # ローカル環境の場合はベースパスを考慮する
            temp_abs_path = self._base_path / temp_path
            if not os.path.exists(temp_abs_path):
                # 絶対パスが存在しない場合は作成する
                os.makedirs(temp_abs_path, exist_ok=True)
            # 絶対パスを返す
            return str(temp_abs_path)

    def keyword_path(self) -> str:
        """キーワードリストのファイルパスの取得

        Returns:
            str: キーワードリストのファイルパス
        """
        temp_path = self._config_data.get("keyword_path", "")
        if not temp_path:
            return temp_path
        
        if self._is_docker:
            # Docker環境の場合はそのまま返す
            return str(pathlib.Path('/data') / temp_path)
        else:
            # ローカル環境の場合はベースパスを考慮する
            return str(self._base_path / temp_path)

    def progress_display(self) -> bool:
        """進捗表示設定の取得

        Returns:
            bool: 進捗表示設定(True:表示, False:非表示)
        """
        return str(self._config_data.get("progress_display", "")).lower() != 'false'
    
    def shape_search(self) -> bool:
        """図形内検索設定の取得

        Returns:
            bool: 図形内検索設定(True:実行, False:非実行)
        """
        return str(self._config_data.get("shape_search", "")).lower() != 'false'

    #
    # protected methods
    #
    def _load_settings(self) -> Dict[str, Any]:
        """settings.yamlファイルの読み込み

        Returns:
            Dict[str, Any]: settings.yamlの内容を格納した辞書。settings.yamlが存在しない場合はデフォルト設定を返す。
        """
        # 一時的な辞書オブジェクトの作成
        temp_dict : Dict[str, Any] = {}
        # settings.yamlファイルの存在チェック
        if not os.path.exists(self._settings_file):
            # ファイルがない場合はデフォルト設定を返す
            temp_dict = {
                "output_path": "output",
                "keyword_path": "input/keywords.txt",
                "progress_display": True,
                "shape_search": True,
            }
        else:
            # settings.yamlファイルの読み込み
            with open(self._settings_file, 'r', encoding='utf-8') as f:
                temp_dict = yaml.safe_load(f)
        # 辞書オブジェクトを返す
        return temp_dict