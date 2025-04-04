"""
このモジュールはプレゼンテーションのストーリーを生成するための機能を提供します。
ユーザーリクエストを受け取り、LLMを使用してプレゼンテーションの大まかなストーリーを作成します。
APIエラーが発生した場合の適切なハンドリングも実装されています。
OpenAIとGoogle Gemini両方のAPIに対応しています。

Classes:
    StoryGenerator: プレゼンテーションのストーリーを生成するクラス
"""

import logging
from typing import Optional
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.language_models.chat_models import BaseChatModel

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

class StoryGenerator:
    """
    プレゼンテーションのストーリーを生成するクラス

    Attributes:
        llm (BaseChatModel): 対話型言語モデルのインスタンス
        max_retries (int): APIコール失敗時の最大再試行回数
    """
    def __init__(self, llm: BaseChatModel, max_retries: int = 2):
        """
        StoryGeneratorクラスの初期化

        Args:
            llm (BaseChatModel): 対話型言語モデルのインスタンス
            max_retries (int, optional): 最大再試行回数。デフォルトは2
        """
        self.llm = llm
        self.max_retries = max_retries
        # APIプロバイダーの検出（クラス名でチェック）
        self.api_provider = self._detect_api_provider(llm)
        logger.info(f"StoryGenerator initialized with {self.api_provider} API")
        
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
        
    def run(self, user_request: str) -> str:
        """
        ユーザーリクエストからプレゼンテーションのストーリーを生成する

        Args:
            user_request (str): ユーザーからのリクエスト

        Returns:
            str: 生成されたプレゼンテーションのストーリー
            
        Raises:
            Exception: APIエラーが発生し、再試行しても解決しない場合
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
        
        # エラーハンドリングを追加
        last_error = None
        for attempt in range(self.max_retries + 1):
            try:
                # ストーリーを生成
                result = chain.invoke({"user_request": user_request})
                logger.info(f"{self.api_provider} APIでストーリー生成に成功しました")
                return result
            except Exception as e:
                last_error = e
                error_msg = str(e).lower()
                
                # APIクォータ超過エラーまたはレート制限エラーの場合
                quota_errors = ["insufficient_quota", "quota exceeded", "rate_limit"]
                if any(err in error_msg for err in quota_errors):
                    logger.error(f"{self.api_provider} API制限エラー: {e}")
                    # クォータ超過は再試行しても解決しないので、すぐに例外を発生させる
                    raise e
                
                # レート制限エラーの場合は少し待機してから再試行
                if "rate" in error_msg and "limit" in error_msg:
                    logger.warning(f"レート制限エラーが発生しました（試行 {attempt+1}/{self.max_retries+1}）: {e}")
                    import time
                    time.sleep(2 ** attempt)  # 指数バックオフ
                    continue
                
                # その他のエラーの場合
                logger.warning(f"ストーリー生成中にエラーが発生しました（試行 {attempt+1}/{self.max_retries+1}）: {e}")
                
                # 最後の試行の場合はフォールバックメッセージを返す
                if attempt == self.max_retries:
                    logger.error("最大試行回数に達しました。フォールバックメッセージを返します。")
                    return """
プレゼンテーションの基本構成：

1. 導入部
   - トピックの紹介と背景情報
   - 聴衆の注目を集める導入

2. 主要ポイント
   - トピックの主要な側面を3-5点程度
   - 各ポイントは明確な見出しと簡潔な説明

3. データと分析
   - ポイントを裏付けるデータや事実
   - 簡潔なグラフや図表の提案

4. 結論
   - 主要ポイントのまとめ
   - 次のステップや行動の提案

5. 質疑応答
   - 想定される質問と回答
                    """
        
        # すべての試行が失敗した場合
        if last_error:
            raise last_error 