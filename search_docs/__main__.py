from search_docs.interfaces import AbstractSearch
from search_docs.factories import Factory
from search_docs.config import Config
import os
import argparse
from datetime import datetime


def main():
    """メイン処理
    """
    parser = argparse.ArgumentParser(description='ドキュメント内の検索を行う')
    parser.add_argument('target_path', type=str, help='検索対象パス')
    parser.add_argument('--output_path', type=str, default='', help='出力先パスを指定（デフォルトは設定ファイルのoutput_path）')
    parser.add_argument('--keyword_list', type=str, default='',  help='キーワードのリストを指定（デフォルトは設定ファイルのkeyword_path）')
    args = parser.parse_args()

    # 設定ファイルの読み込み
    config = Config()
    # パラメータ設定
    target_path = os.path.abspath(args.target_path)
    output_path = os.path.abspath(args.output_path) if args.output_path else os.path.abspath(config.output_path())
    keywords_list_path = os.path.abspath(args.keyword_list) if args.keyword_list else os.path.abspath(config.keyword_path())

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

    # 検索ドキュメントアダプターの生成
    seacher = Factory.create(config=config)
    # ドキュメント検索処理の実行
    seacher.search(target_path=target_path, keywords=keywords, enable_search_shapes=config.shape_search())
    # 検索結果保存処理の実行
    seacher.save_results(output_path=output_path)

if __name__ == "__main__":
    main()