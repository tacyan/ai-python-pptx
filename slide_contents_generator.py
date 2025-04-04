"""
このモジュールはスライドの内容を生成するための機能を提供します。
ユーザーリクエストとストーリーを基に、プレゼンテーションのスライド内容を生成します。
画像や動画の挿入にも対応しています。
OpenAIとGoogle Gemini両方のAPIに対応しています。

Classes:
    SlideContentsGenerator: スライドの内容を生成するクラス
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

class SlideContentsGenerator:
    """
    スライドの内容を生成するクラス

    Attributes:
        llm (BaseChatModel): 対話型言語モデルのインスタンス
        max_retries (int): APIコール失敗時の最大再試行回数
    """
    def __init__(self, llm: BaseChatModel, max_retries: int = 2):
        """
        SlideContentsGeneratorクラスの初期化

        Args:
            llm (BaseChatModel): 対話型言語モデルのインスタンス
            max_retries (int, optional): 最大再試行回数。デフォルトは2
        """
        self.llm = llm
        self.max_retries = max_retries
        # APIプロバイダーの検出（クラス名でチェック）
        self.api_provider = self._detect_api_provider(llm)
        logger.info(f"SlideContentsGenerator initialized with {self.api_provider} API")
        
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
        
        # エラーハンドリングを追加
        last_error = None
        for attempt in range(self.max_retries + 1):
            try:
                # スライド内容を生成するチェーンを作成
                chain = prompt | self.llm | StrOutputParser()
                # スライド内容を生成
                result = chain.invoke({"user_request": user_request, "story": story})
                logger.info(f"{self.api_provider} APIでスライド内容生成に成功しました")
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
                logger.warning(f"スライド内容生成中にエラーが発生しました（試行 {attempt+1}/{self.max_retries+1}）: {e}")
                
                # 最後の試行の場合はフォールバックメッセージを返す
                if attempt == self.max_retries:
                    logger.error("最大試行回数に達しました。フォールバックメッセージを返します。")
                    return f"""
# プレゼンテーション資料

## このプレゼンテーションについて
- APIエラーが発生したため、基本的なスライド構成のみ生成されました
- ユーザーリクエスト: {user_request[:100]}...

---next---

# 目次
1. 導入
2. 主要ポイント
3. まとめ

---next---

# 導入
- このプレゼンテーションでは、{user_request[:50]}...について説明します

---next---

# 主要ポイント
- ポイント1: 詳細情報
- ポイント2: 詳細情報
- ポイント3: 詳細情報

---next---

# まとめ
- 主要ポイントの要約
- 次のステップ
                    """
        
        # すべての試行が失敗した場合
        if last_error:
            raise last_error 