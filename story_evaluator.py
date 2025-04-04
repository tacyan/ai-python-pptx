"""
このモジュールはプレゼンテーションのストーリーを評価するための機能を提供します。
生成されたストーリーが十分かどうかを評価し、Judgementオブジェクトを返します。
OpenAIとGoogle Gemini両方のAPIに対応しています。

Classes:
    StoryEvaluator: プレゼンテーションのストーリーを評価するクラス
"""

import logging
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.language_models.chat_models import BaseChatModel

from datamodel import Judgement

# ロガーの設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Gemini APIの利用可能性をチェック
try:
    from langchain_google_genai import ChatGoogleGenerativeAI
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
    logger.warning("langchain_google_genai をインポートできませんでした。Google Gemini機能は無効化されます。")
    
class StoryEvaluator:
    """
    プレゼンテーションのストーリーを評価するクラス

    Attributes:
        llm (BaseChatModel): 構造化出力をサポートする対話型言語モデルのインスタンス
    """
    def __init__(self, llm: BaseChatModel):
        """
        StoryEvaluatorクラスの初期化

        Args:
            llm (BaseChatModel): 対話型言語モデルのインスタンス
        """
        self.llm = llm.with_structured_output(Judgement)
        # APIプロバイダーの検出（クラス名でチェック）
        self.api_provider = self._detect_api_provider(llm)
        logger.info(f"StoryEvaluator initialized with {self.api_provider} API")
        
    def _detect_api_provider(self, llm: BaseChatModel) -> str:
        """
        使用されているAPIプロバイダーを検出する
        
        Args:
            llm (BaseChatModel): 対話型言語モデルのインスタンス
            
        Returns:
            str: 検出されたAPIプロバイダー名 ("OpenAI" または "Google Gemini")
        """
        llm_class_str = str(llm.__class__).lower()
        
        if "openai" in llm_class_str:
            return "OpenAI"
        elif "google" in llm_class_str or "gemini" in llm_class_str:
            return "Google Gemini"
        else:
            logger.warning(f"不明なLLMタイプです: {llm.__class__.__name__}")
            return "Unknown"
        
    def run(self, user_request: str, story: str) -> Judgement:
        """
        ストーリーの十分性および適切性を評価する

        Args:
            user_request (str): ユーザーからのリクエスト
            story (str): 評価するストーリー

        Returns:
            Judgement: ストーリーの評価結果
        """
        try:
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
            judgement = chain.invoke({"user_request": user_request, "story": story})
            logger.info(f"{self.api_provider} APIでストーリー評価に成功しました")
            return judgement
        except Exception as e:
            logger.error(f"ストーリー評価中にエラーが発生しました: {e}")
            # エラー発生時は否定的な評価を返す（ストーリー生成のやり直しを促す）
            return Judgement(judge=False, reason="API呼び出し中にエラーが発生したため、ストーリーを再生成します。") 