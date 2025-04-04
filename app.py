"""
このモジュールはパワーポイント資料を自動生成するStreamlitアプリケーションです。
ユーザーがアップロードしたテキストファイルからプレゼンテーション内容を読み取り、
AIエージェントを使用して、PowerPointスライドを生成するPythonコードを出力します。
画像や動画の挿入にも対応しています。

Usage:
    streamlit run app.py
"""

import os
import streamlit as st
import sys
import subprocess
import shutil

# 必要なライブラリが存在するか確認し、なければインストールする
def check_and_install_dependencies():
    """
    必要な依存関係をチェックし、不足している場合はインストールする
    """
    required_packages = [
        "langchain-core==0.3.0",
        "langchain-openai==0.2.0",
        "langgraph==0.2.22",
        "python-pptx==1.0.2"
    ]
    
    for package in required_packages:
        try:
            package_name = package.split("==")[0]
            __import__(package_name.replace("-", "_"))
        except ImportError:
            st.info(f"{package_name} をインストールしています...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            st.success(f"{package_name} のインストールが完了しました！")

# ディレクトリ構造を確認し、存在しない場合は作成
def ensure_directories():
    """
    必要なディレクトリ構造を確保する
    """
    os.makedirs("workspace/input", exist_ok=True)
    os.makedirs("workspace/output", exist_ok=True)
    os.makedirs("workspace/input/images", exist_ok=True)

# Streamlitアプリのメイン関数
def main():
    """
    Streamlitアプリのメイン関数
    """
    st.set_page_config(page_title="AIプレゼンテーション生成", page_icon="📊", layout="wide")
    
    # 依存関係のチェックとインストール
    check_and_install_dependencies()
    
    # モジュールのインポート
    try:
        from langchain_openai import ChatOpenAI
        from pptx_agent import PPTXAgent
    except ImportError:
        st.error("必要なライブラリをインポートできませんでした。アプリケーションを再起動してください。")
        return
    
    # ディレクトリの確認
    ensure_directories()
    
    # タイトルと説明
    st.title("AIプレゼンテーション自動生成アプリ")
    st.write("プレゼンテーションの内容を記述したテキストファイルをアップロードすると、AIがパワーポイントの資料を自動生成します。")
    
    # ファイルアップロード
    uploaded_file = st.file_uploader("テキストファイルをアップロード (.txt, .docx, .md)", type=["txt", "docx", "md"])
    
    # OpenAI APIキーの入力
    api_key = st.text_input("OpenAI APIキー", type="password")
    
    # モデル選択
    model = st.selectbox("使用するモデル", options=["gpt-4o", "gpt-3.5-turbo"], index=0)
    
    # テンプレートファイルの情報
    st.info("注意: このアプリケーションは `workspace/input/template.pptx` というパワーポイントテンプレートファイルを必要とします。")
    
    # テンプレートファイルのアップロード
    template_file = st.file_uploader("テンプレートファイルをアップロード (.pptx)", type=["pptx"])
    if template_file:
        with open("workspace/input/template.pptx", "wb") as f:
            f.write(template_file.getbuffer())
        st.success("テンプレートファイルがアップロードされました！")
    
    # 画像・動画ファイルのアップロード（複数可）
    st.subheader("画像・動画ファイルのアップロード（オプション）")
    st.write("プレゼンテーションに使用する画像や動画ファイルをアップロードできます。")
    
    # 既存の画像・動画ファイルを表示
    existing_media_files = os.listdir("workspace/input/images")
    if existing_media_files:
        st.write("現在アップロードされているファイル:")
        cols = st.columns(4)
        for i, file in enumerate(existing_media_files):
            file_path = f"workspace/input/images/{file}"
            file_type = file.split(".")[-1].lower()
            with cols[i % 4]:
                if file_type in ["jpg", "jpeg", "png", "gif"]:
                    st.image(file_path, caption=file, width=150)
                elif file_type in ["mp4", "mov", "avi"]:
                    st.video(file_path)
                else:
                    st.write(f"📁 {file}")
                
                # ファイル削除ボタン
                if st.button(f"削除: {file}", key=f"delete_{file}"):
                    os.remove(file_path)
                    st.success(f"{file} を削除しました。")
                    st.experimental_rerun()
    
    # 新しいファイルのアップロード
    uploaded_media_files = st.file_uploader("画像・動画ファイルをアップロード (.jpg, .jpeg, .png, .gif, .mp4, .mov)", 
                                          type=["jpg", "jpeg", "png", "gif", "mp4", "mov"], 
                                          accept_multiple_files=True)
    
    if uploaded_media_files:
        for uploaded_file in uploaded_media_files:
            file_path = f"workspace/input/images/{uploaded_file.name}"
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        
        st.success(f"{len(uploaded_media_files)}個のファイルがアップロードされました！")
        st.experimental_rerun()
    
    # 実行ボタン
    if st.button("プレゼンテーションを生成", disabled=not (uploaded_file and api_key)):
        if not os.path.exists("workspace/input/template.pptx"):
            st.error("テンプレートファイルが見つかりません。テンプレートファイルをアップロードしてください。")
            return
            
        with st.spinner("AIがプレゼンテーションを生成中..."):
            try:
                # APIキーの設定
                os.environ["OPENAI_API_KEY"] = api_key
                
                # アップロードされたファイルを保存
                content = uploaded_file.read().decode("utf-8")
                
                # ChatOpenAIモデルを初期化
                llm = ChatOpenAI(model=model, temperature=0.0)
                
                # PPTXAgentを初期化
                agent = PPTXAgent(llm=llm)
                
                # エージェントを実行して最終的な出力を取得
                final_output = agent.run(user_request=content)
                
                # Python コードブロックが含まれている場合の処理
                if "```python" in final_output:
                    final_output = final_output.split("```python\n")[-1].split("```")[0]
                
                # 出力をファイルに保存
                with open("workspace/output/create_pptx.py", "w") as f:
                    f.write(final_output)
                
                # Pythonコードを実行
                with st.expander("生成されたPythonコード"):
                    st.code(final_output, language="python")
                
                # コードを実行
                st.info("生成されたPythonコードを実行中...")
                exec(final_output)
                
                # 成功メッセージ
                st.success("プレゼンテーションの生成が完了しました！")
                
                # 生成されたファイルを探す
                output_files = [f for f in os.listdir("workspace/output") if f.endswith(".pptx")]
                
                if output_files:
                    # ダウンロードボタンを提供
                    for file in output_files:
                        with open(f"workspace/output/{file}", "rb") as f:
                            st.download_button(
                                label=f"{file}をダウンロード",
                                data=f,
                                file_name=file,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                else:
                    st.warning("生成されたPPTXファイルが見つかりません。コードを確認してください。")
                    
            except Exception as e:
                st.error(f"エラーが発生しました: {str(e)}")
                st.error("詳細なエラー情報を確認するには、コンソールログを確認してください。")
                raise e

    # 使用方法のガイド
    with st.expander("使用方法"):
        st.write("""
        ### 使用方法
        1. テキストファイル (.txt, .docx, .md) をアップロードします。このファイルにはプレゼンテーションに含めたい内容を記述します。
        2. テンプレートファイル (.pptx) をアップロードします。これがプレゼンテーションのベースとなります。
        3. 必要に応じて、画像や動画ファイルをアップロードします。
        4. OpenAI APIキーを入力します。
        5. 使用するモデルを選択します。
        6. 「プレゼンテーションを生成」ボタンをクリックします。
        7. 生成されたプレゼンテーションをダウンロードします。
        
        ### 画像・動画の使用について
        テキストファイル内で画像や動画を使用したい場合は、以下のように指定してください：
        
        ```
        [画像: 説明テキスト]
        [動画: 説明テキスト]
        ```
        
        AIはこれらの指示に基づいて、アップロードされた画像や動画ファイルを使用してプレゼンテーションを生成します。
        """)

if __name__ == "__main__":
    main() 