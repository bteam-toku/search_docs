from .search_excel import SearchExcel
import os
import argparse
from datetime import datetime


def main():
    """メイン処理
    """
    parser = argparse.ArgumentParser(description='ドキュメント内の検索を行う')
    parser.add_argument('target_path', type=str, help='検索対象パス')
    parser.add_argument('--output_path', type=str, default='./output', help='出力先パスを指定（デフォルト:./output）')
    parser.add_argument('--keyword_list', type=str, default='./input/keyword.txt',  help='キーワードのリストを指定（デフォルト:./input/keyword.txt）')
    parser.add_argument('--no_shapes', action='store_true', help='図形内キーワード検索を無効にする')
    parser.add_argument('--no_progress', action='store_true', help='進捗表示を無効にする')
    args = parser.parse_args()

    # パラメータ設定
    target_path = os.path.abspath(args.target_path)
    output_path = os.path.abspath(args.output_path)
    keywords_list_path = os.path.abspath(args.keyword_list)
    is_search_shapes = not args.no_shapes
    is_progress = not args.no_progress

    # 検索対象パスの存在確認
    if not os.path.exists(target_path):
        print(f'検索対象パスが存在しません: {target_path}')
        exit()
    # 出力先パスの存在確認、なければ作成
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    # キーワードリストを作成
    keywords = []
    if os.path.exists(keywords_list_path):
        with open(keywords_list_path, 'r', encoding='utf-8') as f:
            for line in f:
                keyword = line.strip()
                #空行および先頭文字が//の場合はコメント行として無視
                if keyword != '' and not keyword.startswith('//'):
                    keywords.append(keyword)

    # EXEL検索クラスを初期化
    i_search_excel = SearchExcel()
    # 初期化処理を実行
    if i_search_excel.search_target(target_path, progress=is_progress):
        # キーワード検索を実行
        if len(keywords) > 0:
            # 図形内検索設定
            if is_search_shapes:
                i_search_excel.enable_search_shapes()
            # キーワード検索実行
            i_search_excel.search_keyword(keywords, progress=is_progress)
            # 検索結果を取得
            result_df = i_search_excel.get_keyword_list()
        else:
            result_df = i_search_excel.get_target_list()
        # 検索結果を出力
        result_csv_path = os.path.join(output_path, 'search_excel.csv')
        result_df.to_csv(result_csv_path, encoding='utf-8-sig', index=False)


if __name__ == "__main__":
    main()