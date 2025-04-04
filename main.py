"""
このモジュールはパワーポイント資料を自動生成するためのメインプログラムです。
ユーザーが指定したテキストファイルからプレゼンテーション内容を読み取り、
AIエージェントを使用して、PowerPointスライドを生成するPythonコードを出力します。
画像や動画の挿入にも対応しています。

Usage:
    python main.py --file <テキストファイルのパス>

注意:
    サポートしているファイル形式は .txt, .docx, .md のみです。
"""

import argparse
import os
from langchain_openai import ChatOpenAI

from pptx_agent import PPTXAgent

def main():
    """
    メイン関数。コマンドライン引数を解析し、AIエージェントを実行します。
    """
    # コマンドライン引数のパーサーを作成
    parser = argparse.ArgumentParser(
        description="ユーザーがインプットしたテキストを基にスライドを生成するPythonファイルを出力します"
    )
    # "file"引数を追加
    parser.add_argument(
        "--file",
        type=str,
        help="プレゼンテーションの元になるテキストファイルのパス(.txt/.docx/.md)"
    )
    # コマンドライン引数を解析
    args = parser.parse_args()
    
    # ディレクトリの確認
    os.makedirs("workspace/input", exist_ok=True)
    os.makedirs("workspace/output", exist_ok=True)
    os.makedirs("workspace/input/images", exist_ok=True)
    
    # テキストの取得
    filepath = args.file
    if filepath.endswith((".txt", ".docx", ".md")):
        with open(filepath, "r") as f:
            user_request = f.read()
    else:
        raise ValueError("ファイル形式がサポートされていません")
    
    # ChatOpenAIモデルを初期化
    llm = ChatOpenAI(model="gpt-4o", temperature=0.0)
    # PPTXAgentを初期化
    agent = PPTXAgent(llm=llm)
    # エージェントを実行して最終的な出力を取得
    final_output = agent.run(user_request=user_request)
    # Python コードブロックが含まれている場合の処理
    if "```python" in final_output:
        final_output = final_output.split("```python\n")[-1].split("```")[0]
    # 出力をファイルに保存
    with open("workspace/output/create_pptx.py", "w") as f:
        f.write(final_output)
        
    print("生成されたPythonコードを workspace/output/create_pptx.py に保存しました。")
    print("実行するには以下のコマンドを使用してください：")
    print("python workspace/output/create_pptx.py")
    
    # 画像や動画ファイルの使用方法に関する情報を表示
    print("\n画像・動画の使用について:")
    print("プレゼンテーションで画像や動画を使用する場合は、workspace/input/images ディレクトリに")
    print("必要なファイルを配置し、テキスト内で以下のように指定してください：")
    print("[画像: 説明テキスト]")
    print("[動画: 説明テキスト]")
    
if __name__ == "__main__":
    main()
