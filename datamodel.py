"""
このモジュールはAIエージェントが使用するデータモデルを定義します。
状態管理および評価結果の構造化のためのPydanticモデルが含まれています。

Classes:
    Judgement: ストーリーの評価結果を表すデータモデル
    State: ワークフロー全体の状態を管理するデータモデル
"""

from langchain_core.pydantic_v1 import BaseModel, Field

class Judgement(BaseModel):
    """
    ストーリーの評価結果を表すデータモデル

    Attributes:
        judge (bool): ストーリーが十分かどうかの判定結果
        reason (str): ストーリーが十分かどうかの判定理由
    """
    judge: bool = Field(default=False, description="ストーリーが十分かどうかの判定結果")
    reason: str = Field(default="", description="ストーリーが十分かどうかの判定理由")

class State(BaseModel):
    """
    ステートを表すデータモデル

    Attributes:
        user_request (str): ユーザーからのリクエスト
        story (str): 生成されたストーリー
        iteration (int): ストーリー生成の反復回数
        current_judge (bool): ストーリーが十分かどうかの判定結果
        judgement_reason (str): ストーリーが十分かどうかの判定理由
        slide_contents (str): スライドの内容
        slide_gen_code (str): スライド生成のコード
    """
    user_request: str = Field(..., description="ユーザーからのリクエスト")
    story: str = Field(default="", description="生成されたストーリー")
    iteration: int = Field(default=0, description="ストーリー生成の反復回数")
    current_judge: bool = Field(default=False, description="ストーリーが十分かどうかの判定結果")
    judgement_reason: str = Field(default="", description="ストーリーが十分かどうかの判定理由")
    slide_contents: str = Field(default="", description="スライドの内容")
    slide_gen_code: str = Field(default="", description="スライド生成のコード") 