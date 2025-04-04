"""
このモジュールはスライドの内容を生成するための機能を提供します。
ユーザーリクエストとストーリーを基に、プレゼンテーションのスライド内容を生成します。
画像や動画の挿入にも対応しています。

Classes:
    SlideContentsGenerator: スライドの内容を生成するクラス
"""

from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

class SlideContentsGenerator:
    """
    スライドの内容を生成するクラス

    Attributes:
        llm (ChatOpenAI): 対話型言語モデルのインスタンス
    """
    def __init__(self, llm: ChatOpenAI):
        """
        SlideContentsGeneratorクラスの初期化

        Args:
            llm (ChatOpenAI): 対話型言語モデルのインスタンス
        """
        self.llm = llm
        
    def run(self, user_request: str, story: str) -> str:
        """
        ユーザーリクエストとストーリーからスライドの内容を生成する

        Args:
            user_request (str): ユーザーからのリクエスト
            story (str): 生成されたストーリー

        Returns:
            str: 生成されたスライドの内容
        """
        # プロンプトを定義
        prompt = ChatPromptTemplate.from_messages(
            [
                (
                    "system",
                    "あなたは提供されたストーリーに基づいてプレゼンテーションの構成を作成する専門家です。"
                ),
                (
                    "human",
                    "以下のユーザーリクエストと生成されたストーリーに基づいて、プレゼンテーションのスライドの内容を作成してください。\n\n"
                    "ユーザーリクエスト: {user_request}\n\n"
                    "ストーリー:\n{story}\n\n"
                    "ルール:\n"
                    "- スライドの内容は、テキストベースで作成してください。\n"
                    "- 使用して良いのはテキスト、図形、表、画像、動画です。\n"
                    "- テキスト以外の要素（図形、表、画像、動画）を使用する場合は、その旨を明記してください。\n"
                    "- 画像を使用する場合は [画像: 説明] のフォーマットで記述し、説明には必要な画像の内容について具体的に書いてください。\n"
                    "- 動画を使用する場合は [動画: 説明] のフォーマットで記述し、説明には必要な動画の内容について具体的に書いてください。\n"
                    "- 図形を使用する場合は [図形: 説明] のフォーマットで記述してください。\n"
                    "- 表を使用する場合は [表: 説明] のフォーマットで記述し、その後に表の内容をテキストで記述してください。\n"
                    "- スライド番号を進める際は、'---next---' と記述してください。\n\n"
                    "例:\n"
                    "# AIエージェントの概要\n"
                    "- AIエージェントとは、特定のタスクを自律的に実行できるAIシステムです\n"
                    "- 主な特徴：\n"
                    "  - 自律性\n"
                    "  - 適応性\n"
                    "  - 目標指向\n"
                    "[図形: AIエージェントの主要コンポーネントを示す図。中央に「AIエージェント」、周囲に「知覚」「判断」「行動」「学習」と配置した円形図]\n\n"
                    "---next---\n\n"
                    "# AIエージェントの応用例\n"
                    "[画像: 様々な産業でのAIエージェント活用例を示す写真コラージュ。医療、金融、製造業などの分野を含む]\n"
                    "- 医療：診断支援、薬剤開発\n"
                    "- 金融：取引自動化、リスク分析\n"
                    "- カスタマーサービス：チャットボット\n"
                    "[動画: AIエージェントが自動運転車を操作する様子のデモンストレーション映像]"
                )
            ]
        )
        # スライド内容を生成するチェーンを作成
        chain = prompt | self.llm | StrOutputParser()
        # スライド内容を生成
        return chain.invoke({"user_request": user_request, "story": story}) 