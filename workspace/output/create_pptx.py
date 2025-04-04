#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PowerPointプレゼンテーション自動生成スクリプト
このスクリプトは完全に新規にPowerPointプレゼンテーションを作成します。
テンプレートファイルを使用せず、プレースホルダーにアクセスせず、
すべてのコンテンツをテキストボックスと図形で直接作成します。
タイムスタンプ付きファイル名で常に新規ファイルを生成します。

Features:
- テンプレート非依存の完全新規作成アプローチ
- プレースホルダーを使用しない堅牢な実装
- 包括的なエラーハンドリング
- タイムスタンプと乱数を用いた一意なファイル名生成
- 汎用性の高いプレゼンテーション生成モジュール

Author: AIアシスタント
Version: 2.0
Date: 2025-04-04
"""

import os
import sys
import datetime
import random
import subprocess
import importlib.util
import traceback

# 必要なパッケージのインストール確認と依存関係管理
def check_and_install_package(package_name):
    """
    パッケージがインストールされているか確認し、されていなければインストールする
    
    Args:
        package_name (str): インストールするパッケージ名
        
    Returns:
        bool: インストールが成功したかどうか
    """
    try:
        # パッケージが既にインポート可能か確認
        spec = importlib.util.find_spec(package_name.replace('-', '_').split('==')[0])
        if spec is None:
            print(f"{package_name} をインストールしています...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--no-cache-dir", package_name])
            print(f"{package_name} のインストールが完了しました")
            return True
        else:
            print(f"{package_name} は既にインストールされています")
            return True
    except Exception as e:
        print(f"警告: {package_name} のインストール中にエラーが発生しました: {e}")
        return False

# 必要なパッケージのインストール
required_packages = [
    "python-pptx==1.0.2",
    "Pillow"  # 画像処理に必要
]

# すべてのパッケージをインストール
all_installed = all(check_and_install_package(pkg) for pkg in required_packages)

if not all_installed:
    print("警告: 一部のパッケージのインストールに失敗しました。最低限の機能で続行します。")

# 必要なモジュールのインポート
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
    import datetime
    import os
    import random
except ImportError as e:
    print(f"重要なモジュールのインポートに失敗しました: {e}")
    print("プログラムを終了します。")
    sys.exit(1)

# 一意のファイル名を生成
def generate_unique_filename(prefix="Presentation", ext="pptx"):
    """
    タイムスタンプと乱数を用いた一意なファイル名を生成します
    
    Args:
        prefix (str): ファイル名の接頭辞
        ext (str): ファイル拡張子
        
    Returns:
        str: 生成されたファイルパス
    """
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    random_suffix = f"_{random.randint(1000, 9999)}"
    filename = f"{prefix}_{timestamp}{random_suffix}.{ext}"
    
    # 出力ディレクトリの存在確認と作成
    output_dir = os.path.join("workspace", "output")
    os.makedirs(output_dir, exist_ok=True)
    
    return os.path.join(output_dir, filename)

# テーマカラーの定義
class PresentationTheme:
    """プレゼンテーションのテーマカラーと書式設定を管理するクラス"""
    
    def __init__(self, primary_color=RGBColor(0, 112, 192), 
                 secondary_color=RGBColor(255, 192, 0),
                 accent_color=RGBColor(112, 48, 160),
                 background_color=RGBColor(255, 255, 255),
                 text_color=RGBColor(0, 0, 0)):
        """
        テーマカラーを初期化
        
        Args:
            primary_color (RGBColor): プライマリカラー
            secondary_color (RGBColor): セカンダリカラー
            accent_color (RGBColor): アクセントカラー
            background_color (RGBColor): 背景色
            text_color (RGBColor): テキスト色
        """
        self.primary = primary_color
        self.secondary = secondary_color
        self.accent = accent_color
        self.background = background_color
        self.text = text_color
        
        # フォントとサイズの定義
        self.title_font_size = Pt(44)
        self.subtitle_font_size = Pt(32)
        self.heading_font_size = Pt(28)
        self.body_font_size = Pt(20)
        self.footer_font_size = Pt(14)

# 汎用的なスライド作成関数
def add_simple_slide(prs, title_text, content_elements=None, theme=None):
    """
    シンプルなスライドを追加する汎用関数
    
    Args:
        prs (Presentation): プレゼンテーションオブジェクト
        title_text (str): スライドのタイトル
        content_elements (list): コンテンツ要素のリスト（テキスト、図形など）
        theme (PresentationTheme): プレゼンテーションのテーマ
        
    Returns:
        slide: 作成されたスライドオブジェクト
    """
    if theme is None:
        theme = PresentationTheme()
    
    # 最も基本的なレイアウトを使用
    try:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白レイアウト
    except:
        # インデックスエラーの場合は最初のレイアウトを使用
        slide = prs.slides.add_slide(prs.slide_layouts[0])
    
    # タイトルテキストボックスを追加
    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1.25))
    title_frame = title_shape.text_frame
    title_frame.text = title_text
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_frame.paragraphs[0].font.size = theme.title_font_size
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = theme.text
    
    # コンテンツ要素がある場合は追加
    if content_elements:
        content_shape = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
        content_frame = content_shape.text_frame
        content_frame.word_wrap = True
        
        # 最初のパラグラフを設定
        p = content_frame.paragraphs[0]
        p.text = content_elements[0] if content_elements else ""
        p.font.size = theme.body_font_size
        
        # 追加のコンテンツがあれば追加
        for i, text in enumerate(content_elements[1:], 1):
            p = content_frame.add_paragraph()
            p.text = text
            p.font.size = theme.body_font_size
            # 箇条書きスタイルの場合
            if text.startswith('•') or text.startswith('-'):
                p.level = 1
                
    return slide

# AIエージェントプレゼンテーションを生成
def create_ai_agent_presentation():
    """
    AIエージェントについてのプレゼンテーションを作成する
    
    Returns:
        str: 保存したファイルのパス
    """
    try:
        # 新しいプレゼンテーションを作成
        prs = Presentation()
        
        # テーマを設定
        theme = PresentationTheme(
            primary_color=RGBColor(0, 112, 192),  # 青
            secondary_color=RGBColor(112, 48, 160),  # 紫
            accent_color=RGBColor(255, 150, 0)  # オレンジ
        )
        
        # 1. タイトルスライド
        slide = prs.slides.add_slide(prs.slide_layouts[0])  # タイトルスライド
        
        # タイトル
        title_shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1.5))
        title_shape.text_frame.text = "AIエージェントの世界"
        title_frame = title_shape.text_frame
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        title_frame.paragraphs[0].font.size = Pt(44)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = theme.primary
        
        # サブタイトル
        subtitle_shape = slide.shapes.add_textbox(Inches(2), Inches(3), Inches(6), Inches(1))
        subtitle_shape.text_frame.text = "未来を変える自律型AI技術"
        subtitle_frame = subtitle_shape.text_frame
        subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        subtitle_frame.paragraphs[0].font.size = Pt(32)
        subtitle_frame.paragraphs[0].font.italic = True
        
        # 発表者情報
        author_shape = slide.shapes.add_textbox(Inches(2), Inches(5), Inches(6), Inches(1))
        author_frame = author_shape.text_frame
        author_frame.text = "発表者: AI研究チーム"
        author_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        author_frame.paragraphs[0].font.size = Pt(20)
        
        # 日付
        date_shape = slide.shapes.add_textbox(Inches(2), Inches(5.5), Inches(6), Inches(0.5))
        date_frame = date_shape.text_frame
        current_date = datetime.datetime.now().strftime("%Y年%m月%d日")
        date_frame.text = f"作成日: {current_date}"
        date_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        date_frame.paragraphs[0].font.size = Pt(16)
        
        # 2. 目次スライド
        toc_slide = add_simple_slide(
            prs, 
            "目次", 
            [
                "1. AIエージェントとは",
                "2. AIエージェントの主要な技術",
                "3. AIエージェントの応用分野",
                "4. 未来展望と課題",
                "5. まとめ"
            ],
            theme
        )
        
        # 3. AIエージェントとは
        definition_slide = add_simple_slide(
            prs,
            "AIエージェントとは",
            [
                "AIエージェントは自律的に行動し、環境と相互作用する人工知能システムです。",
                "",
                "• 自律性: 人間の指示なしに意思決定ができる",
                "• 適応性: 環境変化に応じて行動を調整できる",
                "• 目標指向: 特定の目標達成のために行動を最適化する",
                "• 学習能力: 経験から学び、パフォーマンスを向上させる"
            ],
            theme
        )
        
        # AIエージェントの図形を追加
        agent_circle = definition_slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(3), Inches(4.5), Inches(4), Inches(1.5)
        )
        agent_circle.fill.solid()
        agent_circle.fill.fore_color.rgb = theme.secondary
        agent_circle.line.color.rgb = theme.text
        
        # テキストを図形に追加
        agent_text = agent_circle.text_frame
        agent_text.text = "AIエージェント"
        agent_text.paragraphs[0].alignment = PP_ALIGN.CENTER
        agent_text.paragraphs[0].font.size = Pt(24)
        agent_text.paragraphs[0].font.bold = True
        agent_text.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        
        # 4. AIエージェントの主要技術
        tech_slide = add_simple_slide(
            prs,
            "AIエージェントの主要技術",
            [
                "現代のAIエージェントを支える主要な技術：",
                "",
                "• 機械学習・深層学習: パターン認識と予測モデル",
                "• 強化学習: 報酬に基づく意思決定の最適化",
                "• 自然言語処理: 人間の言語理解と生成",
                "• コンピュータビジョン: 視覚情報の認識と解釈",
                "• マルチエージェントシステム: 複数エージェントの協調行動"
            ],
            theme
        )
        
        # 5. 応用分野
        applications_slide = add_simple_slide(
            prs,
            "AIエージェントの応用分野",
            [],
            theme
        )
        
        # 応用分野の表を作成
        table_width = Inches(8)
        table_height = Inches(4)
        rows, cols = 4, 2
        table = applications_slide.shapes.add_table(
            rows, cols, Inches(1), Inches(2), table_width, table_height
        ).table
        
        # ヘッダーの設定
        cell = table.cell(0, 0)
        cell.text = "応用分野"
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.primary
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        cell.text_frame.paragraphs[0].font.bold = True
        
        cell = table.cell(0, 1)
        cell.text = "具体例"
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.primary
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        cell.text_frame.paragraphs[0].font.bold = True
        
        # データの設定
        table_data = [
            ["医療・ヘルスケア", "診断支援、個別化医療、患者モニタリング"],
            ["金融", "取引自動化、詐欺検出、リスク評価"],
            ["運輸・物流", "自動運転車、配送最適化、交通管理"]
        ]
        
        for i, (field, examples) in enumerate(table_data, 1):
            cell = table.cell(i, 0)
            cell.text = field
            cell.text_frame.paragraphs[0].font.bold = True
            
            cell = table.cell(i, 1)
            cell.text = examples
        
        # 6. 未来展望
        future_slide = add_simple_slide(
            prs,
            "未来展望と課題",
            [
                "AIエージェントの発展に伴う展望と課題：",
                "",
                "展望：",
                "• より高度な自律性と適応性の実現",
                "• 人間とAIの協調システムの発展",
                "• 社会全体への統合と生産性向上",
                "",
                "課題：",
                "• 倫理的・法的問題の解決",
                "• セキュリティとプライバシーの確保",
                "• 信頼性と透明性の向上"
            ],
            theme
        )
        
        # 7. まとめスライド
        summary_slide = add_simple_slide(
            prs,
            "まとめ",
            [
                "AIエージェントは現代技術の最前線に位置し、以下の価値を提供します：",
                "",
                "• 複雑な問題の自律的解決",
                "• 人間の能力の拡張と補完",
                "• 新たなビジネスモデルと価値創造",
                "",
                "持続可能な発展のためには、技術革新と社会的責任のバランスが不可欠です。"
            ],
            theme
        )
        
        # 8. お問い合わせスライド
        contact_slide = add_simple_slide(
            prs,
            "お問い合わせ",
            [
                "本プレゼンテーションに関するご質問やお問い合わせは下記までお願いします：",
                "",
                "Email: ai-team@example.com",
                "Web: https://www.ai-team-example.com",
                "電話: 03-1234-5678",
                "",
                "© 2025 AI研究チーム All Rights Reserved"
            ],
            theme
        )
        
        # ファイル名を生成して保存
        output_filename = generate_unique_filename("AI_agent_presentation")
        prs.save(output_filename)
        print(f"プレゼンテーションが正常に生成されました！保存先: {output_filename}")
        return output_filename
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        traceback.print_exc()
        
        # エラーが発生した場合、単純な代替プレゼンテーションを作成
        try:
            print("代替プレゼンテーションを作成します...")
            error_ppt_filename = generate_unique_filename("Error_Presentation")
            
            # 新しいプレゼンテーションを作成
            prs = Presentation()
            
            # エラースライドを追加
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            
            # エラー情報を表示
            title_shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1.5))
            title_frame = title_shape.text_frame
            title_frame.text = "エラーが発生しました"
            title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            title_frame.paragraphs[0].font.size = Pt(44)
            title_frame.paragraphs[0].font.bold = True
            title_frame.paragraphs[0].font.color.rgb = RGBColor(192, 0, 0)  # 赤色
            
            # エラーメッセージ
            error_shape = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(3))
            error_frame = error_shape.text_frame
            error_frame.text = f"プレゼンテーション生成中にエラーが発生しました:\n{str(e)}"
            error_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            error_frame.paragraphs[0].font.size = Pt(20)
            
            # 詳細情報
            detail_shape = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(8), Inches(1))
            detail_frame = detail_shape.text_frame
            detail_frame.text = "このエラープレゼンテーションは代替品として生成されました。"
            detail_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            detail_frame.paragraphs[0].font.size = Pt(16)
            detail_frame.paragraphs[0].font.italic = True
            
            # 日時情報
            now = datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")
            date_shape = slide.shapes.add_textbox(Inches(1), Inches(6), Inches(8), Inches(0.5))
            date_frame = date_shape.text_frame
            date_frame.text = f"生成日時: {now}"
            date_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            date_frame.paragraphs[0].font.size = Pt(14)
            
            # スライドの保存
            prs.save(error_ppt_filename)
            print(f"エラー用のプレゼンテーションが生成されました: {error_ppt_filename}")
            return error_ppt_filename
            
        except Exception as err:
            print(f"代替プレゼンテーションの作成中にエラーが発生しました: {err}")
            traceback.print_exc()
            return None

# スクリプトを直接実行した場合
if __name__ == "__main__":
    try:
        # AIエージェントのプレゼンテーションを作成
        output_file = create_ai_agent_presentation()
        
        if output_file:
            print(f"プレゼンテーションの生成が完了しました！")
            print(f"ファイルパス: {output_file}")
            sys.exit(0)
        else:
            print("プレゼンテーションの生成に失敗しました。")
            sys.exit(1)
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")
        traceback.print_exc()
        sys.exit(1)
