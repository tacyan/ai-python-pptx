"""
このモジュールはパワーポイント資料を自動生成するAIエージェントを実装します。
LangGraphを使ってワークフロー（グラフ）を定義し、複数のLLMを連携させて処理を実行します。

Classes:
    PPTXAgent: PowerPoint資料を自動生成するAIエージェント
"""

from typing import Any
import os
from langchain_openai import ChatOpenAI
from langgraph.graph import END, StateGraph

from datamodel import Judgement, State
from story_generator import StoryGenerator
from story_evaluator import StoryEvaluator
from slide_contents_generator import SlideContentsGenerator
from pptx_code_generator import PPTXCodeGenerator

class PPTXAgent:
    """
    PowerPoint資料を自動生成するAIエージェント

    Attributes:
        story_generator (StoryGenerator): ストーリー生成器
        story_evaluator (StoryEvaluator): ストーリー評価器
        slide_contents_generator (SlideContentsGenerator): スライド内容生成器
        pptx_code_generator (PPTXCodeGenerator): PPTXコード生成器
        graph (StateGraph): ワークフローを表すグラフ
    """
    def __init__(self, llm: ChatOpenAI):
        """
        PPTXAgentクラスの初期化

        Args:
            llm (ChatOpenAI): 対話型言語モデルのインスタンス
        """
        # 各種ジェネレーターの初期化
        self.story_generator = StoryGenerator(llm=llm)
        self.story_evaluator = StoryEvaluator(llm=llm)
        self.slide_contents_generator = SlideContentsGenerator(llm=llm)
        self.pptx_code_generator = PPTXCodeGenerator(llm=llm)
        
        # グラフの作成
        self.graph = self._create_graph()
        
    def _create_graph(self) -> StateGraph:
        """
        ワークフローグラフを作成する

        Returns:
            StateGraph: コンパイル済みのワークフローグラフ
        """
        # グラフの初期化
        workflow = StateGraph(State)
        
        # 各ノードの追加
        workflow.add_node("generate_story", self._generate_story)
        workflow.add_node("evaluate_story", self._evaluate_story)
        workflow.add_node("generate_slide_contents", self._generate_slide_contents)
        workflow.add_node("generate_pptx_code", self._generate_pptx_code)
        
        # エントリーポイントの設定
        workflow.set_entry_point("generate_story")
        
        # ノード間のエッジの追加
        workflow.add_edge("generate_story", "evaluate_story")
        workflow.add_conditional_edges(
            "evaluate_story",
            lambda state: not state.current_judge and state.iteration < 5,
            {True: "generate_story", False: "generate_slide_contents"}
        )
        workflow.add_edge("generate_slide_contents", "generate_pptx_code")
        workflow.add_edge("generate_pptx_code", END)
        
        # グラフのコンパイル
        return workflow.compile()
    
    def _generate_story(self, state: State) -> dict[str, Any]:
        """
        ストーリーを生成するノード処理

        Args:
            state (State): 現在の状態

        Returns:
            dict[str, Any]: 更新する状態の要素
        """
        # ストーリーの生成
        new_story: str = self.story_generator.run(state.user_request)
        return {
            "story": new_story,
            "iteration": state.iteration + 1
        }
        
    def _evaluate_story(self, state: State) -> dict[str, Any]:
        """
        ストーリーを評価するノード処理

        Args:
            state (State): 現在の状態

        Returns:
            dict[str, Any]: 更新する状態の要素
        """
        # ストーリーの評価
        judgement: Judgement = self.story_evaluator.run(state.user_request, state.story)
        return {
            "current_judge": judgement.judge,
            "judgement_reason": judgement.reason
        }
        
    def _generate_slide_contents(self, state: State) -> dict[str, Any]:
        """
        スライド内容を生成するノード処理

        Args:
            state (State): 現在の状態

        Returns:
            dict[str, Any]: 更新する状態の要素
        """
        # スライド内容の生成
        slide_contents: str = self.slide_contents_generator.run(state.user_request, state.story)
        return {"slide_contents": slide_contents}
    
    def _generate_pptx_code(self, state: State) -> dict[str, Any]:
        """
        PPTXコードを生成するノード処理

        Args:
            state (State): 現在の状態

        Returns:
            dict[str, Any]: 更新する状態の要素
        """
        # Python-pptxコードの生成
        pptx_code: str = self.pptx_code_generator.run(state.slide_contents)
        return {"slide_gen_code": pptx_code}
    
    def run(self, user_request: str) -> str:
        """
        AIエージェントを実行する

        Args:
            user_request (str): ユーザーからのリクエスト

        Returns:
            str: 生成されたPythonコード
        """
        # 初期状態の設定
        initial_state = State(user_request=user_request)
        
        # グラフ構造の可視化（Streamlit環境では不要のため削除）
        # 代わりにプロセス開始のログを出力
        os.makedirs("workspace/output", exist_ok=True)
        
        # グラフの実行
        final_state = self.graph.invoke(initial_state)
        # 最終的なPythonコードの取得
        return final_state["slide_gen_code"] 