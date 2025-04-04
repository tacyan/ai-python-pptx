"""
このモジュールはプレゼンテーションのストーリーを評価するための機能を提供します。
生成されたストーリーが十分かどうかを評価し、Judgementオブジェクトを返します。

Classes:
    StoryEvaluator: プレゼンテーションのストーリーを評価するクラス
"""

from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

from datamodel import Judgement
    
class StoryEvaluator:
    """
    プレゼンテーションのストーリーを評価するクラス

    Attributes:
        llm (ChatOpenAI): 構造化出力をサポートする対話型言語モデルのインスタンス
    """
    def __init__(self, llm: ChatOpenAI):
        """
        StoryEvaluatorクラスの初期化

        Args:
            llm (ChatOpenAI): 対話型言語モデルのインスタンス
        """
        self.llm = llm.with_structured_output(Judgement)
        
    def run(self, user_request: str, story: str) -> Judgement:
        """
        ストーリーの十分性および適切性を評価する

        Args:
            user_request (str): ユーザーからのリクエスト
            story (str): 評価するストーリー

        Returns:
            Judgement: ストーリーの評価結果
        """
        # プロンプトを定義
        prompt = ChatPromptTemplate.from_messages(
            [
                (
                    "system",
                    "あなたはプレゼンテーションのストーリーの十分性および適切性を評価する専門家です。"
                ),
                (
                    "human",
                    "以下のユーザーリクエストと生成されたストーリーから、良いプレゼンテーション資料を作成するために十分で適切な情報が記載されているかどうかを判断してください。\n\n"
                    "ユーザーリクエスト: {user_request}\n\n"
                    "ストーリー:\n{story}"
                )
            ]
        )
        # ストーリーの十分性および適切性を評価するチェーンを作成
        chain = prompt | self.llm
        # 評価結果を返す
        return chain.invoke({"user_request": user_request, "story": story}) 