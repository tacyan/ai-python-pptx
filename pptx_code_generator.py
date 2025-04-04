"""
このモジュールはスライド内容をPythonコードに変換するための機能を提供します。
生成されたスライド内容を、python-pptxを使用してPowerPointファイルを生成するためのコードに変換します。
画像や動画も挿入できるように拡張されています。

Classes:
    PPTXCodeGenerator: PowerPointスライド生成コードを生成するクラス
"""

import os
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

class PPTXCodeGenerator:
    """
    PowerPointスライド生成コードを生成するクラス

    Attributes:
        llm (ChatOpenAI): 対話型言語モデルのインスタンス
    """
    def __init__(self, llm: ChatOpenAI):
        """
        PPTXCodeGeneratorクラスの初期化

        Args:
            llm (ChatOpenAI): 対話型言語モデルのインスタンス
        """
        self.llm = llm
        
        # 画像アップロードディレクトリの確保
        os.makedirs("workspace/input/images", exist_ok=True)
        
    def run(self, slide_contents: str) -> str:
        """
        スライド内容からPowerPointスライド生成コードを生成する

        Args:
            slide_contents (str): 生成されたスライド内容

        Returns:
            str: PowerPointスライド生成のためのPythonコード
        """
        # プロンプトを定義
        prompt = ChatPromptTemplate.from_messages(
            [
                (
                    "system",
                    "python-pptxモジュールを用いてプレゼンテーション資料のスライドを自動生成する専門家です。"
                ),
                (
                    "human",
                    "以下のスライドの内容を生成するためのpythonコードを生成してください。\n\n"
                    "スライドの内容:\n{slide_contents}\n\n"
                    "以下のpptxファイルを読み込み、テンプレートとして使用してください。\n"
                    "workspace/input/template.pptx\n"
                    "テンプレートのレイアウト情報は以下を参照してください。記載にないレイアウト番号やプレースホルダー番号は決して使用しないでください。\n"
                    "    タイトルスライド: slide_layouts[2]\n"
                    "        placeholder_format.idx:\n"
                    "            0: 会社名など\n"
                    "            10: 発表タイトル\n"
                    "            11: サブタイトル・日付など\n"
                    "            12: 発表者名など\n"
                    "    一般スライド: slide_layouts[0]\n"
                    "        placeholder_format.idx:\n"
                    "            0: スライドタイトル\n"
                    "            1: 内容\n"
                    "作成したパワーポイントはworkspace/output内に出力されるようにしてください。\n\n"
                    "ルール:\n"
                    "- 【重要】必ずpython-pptxモジュールを使用したpythonコードのみを出力してください。\n"
                    "- 使用が許可されているのは、テキスト、図形、表、画像、動画です。\n"
                    "- テキスト以外の要素（図形、表、画像、動画）を使用してほしい箇所には、その旨が明記されています。\n"
                    "- 画像を挿入する場合は `workspace/input/images/` ディレクトリの画像ファイルを使用するコードを生成してください。\n"
                    "- 動画を挿入する場合も同様に `workspace/input/images/` ディレクトリのファイルを使用するコードを生成してください。\n"
                    "- '---next---' はスライド番号を進める合図です。このタイミングで新たなスライドを追加してください。\n\n"
                    "### 画像や動画の挿入方法について ###\n"
                    "以下のようなコード例を参考にしてください。\n"
                    "画像の挿入例:\n"
                    "```python\n"
                    "from pptx.util import Inches\n"
                    "slide = prs.slides.add_slide(prs.slide_layouts[0])\n"
                    "title = slide.shapes.title\n"
                    "title.text = \"画像の例\"\n"
                    "# 画像を追加\n"
                    "img_path = \"workspace/input/images/example.jpg\"\n"
                    "left = Inches(1)\n"
                    "top = Inches(2.5)\n"
                    "width = Inches(5)\n"
                    "slide.shapes.add_picture(img_path, left, top, width=width)\n"
                    "```\n"
                    "動画の挿入例:\n"
                    "```python\n"
                    "from pptx.util import Inches\n"
                    "slide = prs.slides.add_slide(prs.slide_layouts[0])\n"
                    "title = slide.shapes.title\n"
                    "title.text = \"動画の例\"\n"
                    "# 動画を追加\n"
                    "movie_path = \"workspace/input/images/example.mp4\"\n"
                    "left = Inches(2)\n"
                    "top = Inches(2.5)\n"
                    "width = Inches(5)\n"
                    "height = Inches(3)\n"
                    "slide.shapes.add_movie(movie_path, left, top, width, height)\n"
                    "```\n"
                )
            ]
        )
        # スライド生成のためのチェーンを作成
        chain = prompt | self.llm | StrOutputParser()
        # スライド生成のコードを生成
        return chain.invoke({"slide_contents": slide_contents}) 