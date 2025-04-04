"""
このモジュールはパワーポイント資料を自動生成するStreamlitアプリケーションです。
ユーザーがアップロードしたテキストファイルからプレゼンテーション内容を読み取り、
AIエージェントを使用して、PowerPointスライドを生成するPythonコードを出力します。
画像や動画の挿入にも対応しています。
OpenAIとGemini両方のAPIに対応しています。

Usage:
    streamlit run app.py
"""

import os
import streamlit as st
import sys
import subprocess
import shutil
import importlib
import time
import logging

# ロギングの設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 必要なライブラリが存在するか確認し、なければインストールする
def check_and_install_dependencies():
    """
    必要な依存関係をチェックし、不足している場合はインストールする
    エラーハンドリングを強化し、インストール失敗時の代替手段も提供
    """
    required_packages = [
        "langchain-core==0.3.0",
        "langchain-openai==0.2.0",
        "langgraph==0.2.22",
        "python-pptx==1.0.2",
        "ipython"
    ]
    
    # Gemini関連のパッケージ（オプション）
    gemini_packages = [
        "google-generativeai>=0.3.0",
        "langchain-google-genai==0.1.5"
    ]
    
    # 基本パッケージのインストール
    install_packages(required_packages, critical=True)
    
    # 特別な方法でGeminiパッケージをインストール
    gemini_available = install_gemini_packages()
    
    return gemini_available

def install_gemini_packages():
    """
    Gemini API関連パッケージを特別な方法でインストールする
    
    Returns:
        bool: Geminiパッケージのインストールに成功したかどうか
    """
    try:
        # まずgoogle-generativeaiをインストール
        st.info("Google AI SDKをインストールしています...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--no-cache-dir", "google-generativeai>=0.3.0"],
            stderr=subprocess.PIPE
        )
        
        # 依存関係のチェック
        st.info("LangChain Google Genai拡張をインストールしています...")
        
        # 複数の方法を試す
        install_success = False
        
        methods = [
            # 方法1: 最新バージョンで直接インストール
            [sys.executable, "-m", "pip", "install", "--no-cache-dir", "langchain-google-genai"],
            # 方法2: 特定バージョンを指定してインストール
            [sys.executable, "-m", "pip", "install", "--no-cache-dir", "--use-pep517", "langchain-google-genai==0.1.5"],
            # 方法3: 依存関係を明示的に指定
            [sys.executable, "-m", "pip", "install", "--no-cache-dir", "langchain-google-genai==0.1.5", "typing-inspect>=0.8.0", "typing-extensions>=4.5.0"],
            # 方法4: 最小限の依存関係でインストール
            [sys.executable, "-m", "pip", "install", "--no-cache-dir", "--no-deps", "langchain-google-genai", "google-ai-generativelanguage>=0.6.0"]
        ]
        
        for i, cmd in enumerate(methods):
            try:
                st.info(f"インストール方法 {i+1}/{len(methods)} を試行中...")
                subprocess.check_call(cmd, stderr=subprocess.PIPE)
                install_success = True
                st.success(f"インストール方法 {i+1} で成功しました！")
                break
            except subprocess.CalledProcessError as e:
                st.warning(f"インストール方法 {i+1} が失敗しました: {e}")
                continue
        
        # インポートテスト
        try:
            import langchain_google_genai
            st.success("Google Gemini APIのサポートが有効化されました！")
            return True
        except ImportError as e:
            if install_success:
                st.warning(f"パッケージはインストールされましたが、インポートできませんでした: {e}")
                st.info("アプリケーションを再起動して再試行してください。")
            else:
                st.error("すべてのインストール方法が失敗しました。")
            return False
            
    except Exception as e:
        st.warning(f"Google Gemini APIサポートのインストールに失敗しました: {e}")
        st.info("OpenAI APIのみ使用可能です。手動インストールを試すには次のコマンドを実行してください:")
        st.code("pip install langchain-google-genai --no-cache-dir")
        return False

def install_packages(packages, critical=False):
    """
    パッケージのリストをインストールする

    Args:
        packages (list): インストールするパッケージのリスト
        critical (bool): 重要なパッケージかどうか。Trueの場合、インストール失敗時にエラーを表示

    Returns:
        bool: すべてのパッケージのインストールに成功したかどうか
    """
    all_success = True
    
    for package in packages:
        try:
            package_name = package.split("==")[0].split(">=")[0]
            try:
                # すでにインストールされているか確認
                importlib.import_module(package_name.replace("-", "_"))
                logger.info(f"{package_name} は既にインストールされています")
            except ImportError:
                st.info(f"{package_name} をインストールしています...")
                
                # インストール試行回数
                max_attempts = 3
                
                for attempt in range(max_attempts):
                    try:
                        # インストールコマンドを実行
                        subprocess.check_call(
                            [sys.executable, "-m", "pip", "install", "--no-cache-dir", package],
                            stderr=subprocess.STDOUT
                        )
                        st.success(f"{package_name} のインストールが完了しました！")
                        break
                    except subprocess.CalledProcessError as e:
                        logger.warning(f"{package_name} のインストールに失敗しました（試行 {attempt+1}/{max_attempts}）: {e}")
                        
                        if attempt == max_attempts - 1:
                            # すべての試行が失敗
                            if critical:
                                st.error(f"重要なパッケージ {package_name} のインストールに失敗しました。")
                                st.error("手動でのインストールをお試しください: `pip install {package}`")
                                all_success = False
                            else:
                                st.warning(f"オプションパッケージ {package_name} のインストールに失敗しました。一部の機能が制限される場合があります。")
                                all_success = False
                        else:
                            # 短い遅延後に再試行
                            time.sleep(2)
        except Exception as e:
            logger.error(f"パッケージ {package} の処理中にエラーが発生しました: {e}")
            if critical:
                st.error(f"パッケージ {package} の処理中にエラーが発生しました: {e}")
                all_success = False
            else:
                st.warning(f"オプションパッケージ {package} の処理中にエラーが発生しました。一部の機能が制限される場合があります。")
                all_success = False
    
    return all_success

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
    gemini_available = check_and_install_dependencies()
    
    # モジュールのインポート
    try:
        from langchain_openai import ChatOpenAI
        if gemini_available:
            try:
                from langchain_google_genai import ChatGoogleGenerativeAI
            except ImportError:
                gemini_available = False
        from pptx_agent import PPTXAgent
    except ImportError as e:
        st.error(f"必要なライブラリをインポートできませんでした: {e}")
        st.error("アプリケーションを再起動してください。")
        return
    
    # ディレクトリの確認
    ensure_directories()
    
    # タイトルと説明
    st.title("AIプレゼンテーション自動生成アプリ")
    st.write("プレゼンテーションの内容を記述したテキストファイルをアップロードすると、AIがパワーポイントの資料を自動生成します。")
    
    # ファイルアップロード
    uploaded_file = st.file_uploader("テキストファイルをアップロード (.txt, .docx, .md)", type=["txt", "docx", "md"])
    
    # APIプロバイダー選択
    api_provider_options = ["OpenAI"]
    if gemini_available:
        api_provider_options.append("Google Gemini")
    else:
        st.warning("Google Gemini APIの依存関係をインストールできませんでした。OpenAI APIのみ使用可能です。")
    
    api_provider = st.radio("使用するAPIプロバイダー", api_provider_options, horizontal=True)
    
    # APIキーの入力と関連設定（選択したプロバイダーによって表示を切り替え）
    if api_provider == "OpenAI":
        api_key = st.text_input("OpenAI APIキー", type="password")
        model = st.selectbox("使用するモデル", options=["gpt-4o", "gpt-3.5-turbo"], index=0)
    else:  # Gemini
        api_key = st.text_input("Google Gemini APIキー", type="password")
        model = st.selectbox("使用するモデル", options=["gemini-pro", "gemini-1.5-pro", "gemini-1.5-flash"], index=0)
    
    # バックアップオプションの追加
    use_fallback = st.checkbox("APIクォータ超過時にバックアップモデルを使用する", value=True, 
                             help="APIクォータが超過した場合、より軽量なモデルにフォールバックします")
    
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
        # テンプレートファイルのチェックを変更（ない場合も続行）
        if not os.path.exists("workspace/input/template.pptx"):
            st.warning("テンプレートファイルが見つかりません。新規プレゼンテーションとして生成します。")
            
        with st.spinner("AIがプレゼンテーションを生成中..."):
            try:
                # APIキーの設定
                if api_provider == "OpenAI":
                    os.environ["OPENAI_API_KEY"] = api_key
                else:  # Gemini
                    os.environ["GOOGLE_API_KEY"] = api_key
                
                # アップロードされたファイルを保存
                content = uploaded_file.read().decode("utf-8")
                
                # LLMモデルを初期化
                try:
                    if api_provider == "OpenAI":
                        llm = ChatOpenAI(model=model, temperature=0.0)
                        fallback_model = "gpt-3.5-turbo" if model != "gpt-3.5-turbo" else None
                    else:  # Gemini
                        llm = ChatGoogleGenerativeAI(model=model, temperature=0.0)
                        fallback_model = "gemini-1.5-flash" if model != "gemini-1.5-flash" else None
                    
                    # PPTXAgentを初期化
                    agent = PPTXAgent(llm=llm, use_fallback=use_fallback, api_provider=api_provider, fallback_model=fallback_model)
                    
                    # エージェントを実行して最終的な出力を取得
                    final_output = agent.run(user_request=content)
                except Exception as api_error:
                    # APIクォータ超過エラー・レート制限エラーの処理
                    error_msg = str(api_error).lower()
                    quota_error = any(err in error_msg for err in ["insufficient_quota", "rate_limit", "quota exceeded"])
                    
                    if quota_error:
                        st.warning(f"{api_provider} APIの制限に達しました。バックアップモデルを使用します。")
                        if use_fallback:
                            if api_provider == "OpenAI" and model != "gpt-3.5-turbo":
                                llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.0)
                                agent = PPTXAgent(llm=llm, use_fallback=use_fallback, api_provider=api_provider)
                                final_output = agent.run(user_request=content)
                            elif api_provider == "Google Gemini" and model != "gemini-1.5-flash":
                                llm = ChatGoogleGenerativeAI(model="gemini-1.5-flash", temperature=0.0)
                                agent = PPTXAgent(llm=llm, use_fallback=use_fallback, api_provider=api_provider)
                                final_output = agent.run(user_request=content)
                            else:
                                st.error(f"{api_provider} APIの制限に達しました。APIキーの制限を確認するか、別のAPIキーを使用してください。")
                                st.error("この問題を解決するには：")
                                st.error(f"1. {api_provider}ダッシュボードでAPIの使用状況と制限を確認する")
                                st.error("2. 有料プランにアップグレードする")
                                st.error("3. 別のAPIキーを使用する")
                                return
                        else:
                            st.error(f"{api_provider} APIの制限に達しました。APIキーの制限を確認するか、別のAPIキーを使用してください。")
                            st.error("この問題を解決するには：")
                            st.error(f"1. {api_provider}ダッシュボードでAPIの使用状況と制限を確認する")
                            st.error("2. 有料プランにアップグレードする")
                            st.error("3. 別のAPIキーを使用する")
                            return
                    else:
                        raise api_error
                
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
                
                # より詳細なエラーメッセージの表示
                error_msg = str(e).lower()
                if any(err in error_msg for err in ["insufficient_quota", "quota exceeded", "rate_limit"]):
                    st.error(f"{api_provider} APIの制限に達しました。APIキーの制限を確認するか、別のAPIキーを使用してください。")
                    st.error("この問題を解決するには：")
                    st.error(f"1. {api_provider}ダッシュボードでAPIの使用状況と制限を確認する")
                    st.error("2. 有料プランにアップグレードする")
                    st.error("3. 別のAPIキーを使用する")
                else:
                    st.error("詳細なエラー情報を確認するには、コンソールログを確認してください。")
                
                raise e

    # 使用方法のガイド
    with st.expander("使用方法"):
        st.write("""
        ### 使用方法
        1. テキストファイル (.txt, .docx, .md) をアップロードします。このファイルにはプレゼンテーションに含めたい内容を記述します。
        2. APIプロバイダー（OpenAIまたはGoogle Gemini）を選択します。
        3. 選択したプロバイダーのAPIキーを入力します。
        4. 使用するモデルを選択します。
        5. テンプレートファイル (.pptx) をアップロードします。これがプレゼンテーションのベースとなります。
        6. 必要に応じて、画像や動画ファイルをアップロードします。
        7. 「プレゼンテーションを生成」ボタンをクリックします。
        8. 生成されたプレゼンテーションをダウンロードします。
        
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