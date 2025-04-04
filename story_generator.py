"""
このモジュールはプレゼンテーションのストーリーを生成するための機能を提供します。
ユーザーリクエストを受け取り、LLMを使用してプレゼンテーションの大まかなストーリーを作成します。

Classes:
    StoryGenerator: プレゼンテーションのストーリーを生成するクラス
"""

from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

class StoryGenerator:
    """
    プレゼンテーションのストーリーを生成するクラス

    Attributes:
        llm (ChatOpenAI): 対話型言語モデルのインスタンス
    """
    def __init__(self, llm: ChatOpenAI):
        """
        StoryGeneratorクラスの初期化

        Args:
            llm (ChatOpenAI): 対話型言語モデルのインスタンス
        """
        self.llm = llm
        
    def run(self, user_request: str) -> str:
        """
        ユーザーリクエストからプレゼンテーションのストーリーを生成する

        Args:
            user_request (str): ユーザーからのリクエスト

        Returns:
            str: 生成されたプレゼンテーションのストーリー
        """
        # プロンプトを定義
        prompt = ChatPromptTemplate.from_messages(
            [
                (
                    "system",
                    "あなたはプレゼンテーションのストーリーを作成する専門家です。"
                ),
                (
                    "human",
                    "以下のユーザーリクエストに基づいて、プレゼンテーションのストーリーを作成してください。\n\n"
                    "ユーザーの意図を理解し、その意図がオーディエンスにしっかりと伝わることを重視してください。\n\n"
                    "ユーザーリクエスト:\n{user_request}"
                )
            ]
        )
        # ストーリー作成のためのチェーンを作成
        chain = prompt | self.llm | StrOutputParser()
        # ストーリーを生成
        return chain.invoke({"user_request": user_request}) 