import streamlit as st
import pandas as pd
import json
import asyncio
import concurrent.futures
from datetime import datetime, date
import tempfile
import io
import base64
from pathlib import Path
import time
import os
from dotenv import load_dotenv

from core.invoice_checker import InvoiceChecker
from core.rule_manager import RuleManager
from core.file_processor import FileProcessor
from core.rule_suggester import RuleSuggester
from core.utils import load_custom_css
from core.folder_processor import FolderProcessor
from core.excel_manager import ExcelManager
from core.llm_client import LLMClient
from core.task_engine import TaskEngine
from core.ui_components import CommonUIComponents, ProgressManager

# セッション状態の初期化
def initialize_session_state():
    """セッション状態を初期化"""
    # 既存機能
    if "rules" not in st.session_state:
        st.session_state.rules = RuleManager()
    if "checker" not in st.session_state:
        st.session_state.checker = InvoiceChecker()
    if "processor" not in st.session_state:
        st.session_state.processor = FileProcessor()
    if "suggester" not in st.session_state:
        st.session_state.suggester = RuleSuggester()
    
    # 新機能
    if "folder_processor" not in st.session_state:
        st.session_state.folder_processor = FolderProcessor()
    if "excel_manager" not in st.session_state:
        st.session_state.excel_manager = ExcelManager()
    if "task_engine" not in st.session_state:
        st.session_state.task_engine = TaskEngine()
    
    # アプリケーションモード
    if "app_mode" not in st.session_state:
        st.session_state.app_mode = "請求書チェック"

def main():
    # 環境変数をロード
    load_dotenv()
    
    st.set_page_config(
        page_title="経理業務ツール",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    initialize_session_state()
    
    # カスタムCSSを読み込み
    load_custom_css()
    
    # カスタムヘッダー
    st.markdown("""
    <div class="main-header">
        <h1>統合経理業務ツール</h1>
        <p>請求書チェック・経理業務処理</p>
    </div>
    """, unsafe_allow_html=True)
    
    # サイドバーでモード選択
    with st.sidebar:
        st.markdown("### ツール選択")
        st.markdown("---")
        
        app_mode = st.selectbox(
            "使用するツールを選択",
            ["請求書チェック", "経理業務処理"],
            index=0 if st.session_state.app_mode == "請求書チェック" else 1,
            help="実行したい業務を選択してください"
        )
        st.session_state.app_mode = app_mode
        
        st.markdown("### 設定")
        st.markdown("---")
        
        # プロバイダー情報を表示
        provider = os.getenv("OPENAI_PROVIDER", "openai")
        st.info(f"現在のプロバイダー: {provider.upper()}")
        
        # APIキー設定（共通UIコンポーネントを使用）
        provider_info = {
            "provider": provider,
            "model": "gpt-4.1",
            "configured": False
        }
        
        # 現在の設定状況を確認
        def is_task_engine_configured():
            """タスクエンジンの設定状況を安全にチェック"""
            try:
                # task_engine の存在確認
                if not hasattr(st.session_state, 'task_engine') or st.session_state.task_engine is None:
                    return False
                
                # llm_client の存在確認
                if not hasattr(st.session_state.task_engine, 'llm_client') or st.session_state.task_engine.llm_client is None:
                    return False
                
                # client の存在確認
                if not hasattr(st.session_state.task_engine.llm_client, 'client') or st.session_state.task_engine.llm_client.client is None:
                    return False
                
                return True
            except AttributeError as e:
                # 構造的な問題がある場合はログに記録
                print(f"Warning: TaskEngine structure issue: {e}")
                return False
            except Exception as e:
                # その他の予期しないエラー
                print(f"Unexpected error checking TaskEngine configuration: {e}")
                return False
        
        provider_info["configured"] = (
            st.session_state.checker.is_configured() and 
            is_task_engine_configured()
        )
        
        api_key_input = CommonUIComponents.show_api_key_configuration(provider_info)
        
        # プロバイダー詳細情報を表示
        CommonUIComponents.show_provider_details(provider_info)
        
        # APIキーが入力された場合の設定処理
        if api_key_input and not provider_info["configured"]:
            if st.button("APIキーを設定", use_container_width=True):
                try:
                    # 既存機能にAPIキー設定
                    checker_success = st.session_state.checker.set_api_key(api_key_input)
                    suggester_success = st.session_state.suggester.set_api_key(api_key_input)
                    # 新機能にAPIキー設定
                    st.session_state.task_engine.set_api_key(api_key_input)
                    
                    if checker_success and suggester_success:
                        st.success("✅ APIキーが正常に設定されました")
                        st.rerun()
                    else:
                        st.error("❌ APIキー設定に失敗しました")
                except Exception as e:
                    st.error(f"❌ APIキー設定エラー: {str(e)}")
        
        # 処理設定
        st.markdown("### 処理設定")
        st.markdown("---")
        
        max_workers = st.slider(
            "並列処理数", 
            min_value=1, 
            max_value=12, 
            value=3,
            help="同時に処理するファイル数を設定します"
        )
        st.session_state.max_workers = max_workers
        
        # システム情報
        st.markdown("### 情報")
        st.markdown("---")
        
        system_config = {
            "モード": app_mode,
            "モデル": "GPT-4.1",
            "並列処理": f"{max_workers} ファイル",
            "APIキー": "設定済み" if provider_info["configured"] else "未設定",
            "プロバイダー": provider_info["provider"].upper()
        }
        
        CommonUIComponents.show_configuration_summary(system_config, "システム設定")
    
    # モードに応じたUIを表示
    if app_mode == "請求書チェック":
        show_invoice_checker_ui()
    else:
        show_accounting_processor_ui()

def show_invoice_checker_ui():
    """既存の請求書チェックUI"""
    # 既存のタブ構成を使用
    tab1, tab2, tab3, tab4 = st.tabs([
        "ルール設定", 
        "ファイルアップロード", 
        "チェック実行", 
        "結果表示"
    ])
    
    with tab1:
        show_rule_management()
    
    with tab2:
        show_file_upload()
    
    with tab3:
        show_check_execution()
    
    with tab4:
        show_results()

def show_accounting_processor_ui():
    """新しい経理業務処理UI"""
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "タスク設定",
        "証跡アップロード", 
        "調書アップロード",
        "処理実行",
        "結果確認"
    ])
    
    with tab1:
        show_task_configuration()
    
    with tab2:
        show_evidence_upload()
    
    with tab3:
        show_excel_upload()
    
    with tab4:
        show_processing_execution()
    
    with tab5:
        show_processing_results()

def show_task_configuration():
    """タスク設定UI"""
    st.markdown('<div class="card-header">タスク設定</div>', unsafe_allow_html=True)
    
    with st.container(border=True):
        st.markdown("### 利用可能なタスク")
        
        available_tasks = st.session_state.task_engine.get_available_tasks()
        
        if available_tasks:
            # タスク一覧を表形式で表示
            task_data = []
            for task in available_tasks:
                task_data.append({
                    "ID": task['id'],
                    "タスク名": task['name'],
                    "説明": task.get('description', ''),
                    "カスタム": "✓" if task.get('is_custom') else ""
                })
            
            tasks_df = pd.DataFrame(task_data)
            
            # 選択可能な表を表示
            st.markdown("タスク一覧（クリックして選択）:")
            
            # データグリッドで表示
            selected = st.dataframe(
                tasks_df,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row"
            )
            
            # 選択されたタスクを取得
            if selected and selected.selection and selected.selection.rows:
                selected_idx = selected.selection.rows[0]
                selected_task_id = available_tasks[selected_idx]['id']
                st.session_state.selected_task_id = selected_task_id
            elif 'selected_task_id' not in st.session_state and available_tasks:
                # デフォルトで最初のタスクを選択
                st.session_state.selected_task_id = available_tasks[0]['id']
            
            # 選択されたタスクの詳細を表示
            if 'selected_task_id' in st.session_state:
                selected_task_id = st.session_state.selected_task_id
                task_config = st.session_state.task_engine.get_task_config(selected_task_id)
                if task_config:
                    with st.expander("タスク詳細", expanded=True):
                        st.markdown(f"**説明**: {task_config.get('description', 'なし')}")
                        
                        if "output_config" in task_config:
                            output_config = task_config["output_config"]
                            st.markdown(f"**出力シート**: {output_config.get('target_sheet', '不明')}")
                            
                            # 表形式で表示
                            st.markdown(f"**開始行**: {output_config.get('start_row', '不明')}")
                            st.markdown("**出力カラム定義**:")
                            
                            # 表形式でカラム定義を表示
                            column_data = []
                            for col_letter in sorted(output_config.get("column_definitions", {}).keys()):
                                col_def = output_config["column_definitions"][col_letter]
                                column_data.append({
                                    "列": col_letter,
                                    "ヘッダー": col_def.get('header', ''),
                                    "キー": col_def.get('key', ''),
                                    "説明": col_def.get('description', '')
                                })
                            
                            if column_data:
                                df = pd.DataFrame(column_data)
                                st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("利用可能なタスクがありません")
    
    # カスタムタスク作成
    with st.container(border=True):
        st.markdown("### カスタムタスク作成")
        
        with st.expander("新しいタスクを作成", expanded=False):
            custom_name = st.text_input("タスク名", placeholder="例: 支払明細照合")
            custom_description = st.text_area("説明", placeholder="このタスクの内容を説明してください...")
            
            st.markdown("**出力設定**")
            
            col1, col2 = st.columns(2)
            with col1:
                target_sheet = st.text_input("対象シート名", value="結果")
                start_row = st.number_input("開始行", min_value=1, value=3)
            
            # カラム設定を全幅で表示
            st.markdown("**カラム設定**")
            st.markdown("列定義を以下の表で編集してください：")
            
            # 初期データの設定
            if "custom_column_df" not in st.session_state:
                default_data = [
                    {"列": "A", "キー": "row_number", "ヘッダー": "No.", "説明": "行番号"},
                    {"列": "B", "キー": "data_id", "ヘッダー": "データID", "説明": "データ識別子"},
                    {"列": "C", "キー": "result", "ヘッダー": "結果", "説明": "処理結果"}
                ]
                st.session_state.custom_column_df = pd.DataFrame(default_data)
            
            # 編集可能なデータフレーム
            edited_df = st.data_editor(
                st.session_state.custom_column_df,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "列": st.column_config.TextColumn(
                        "列",
                        help="Excel列名（A, B, C...）",
                        default="",
                        max_chars=1,
                        validate=r"^[A-Z]$"
                    ),
                    "キー": st.column_config.TextColumn(
                        "キー",
                        help="データのキー名（英数字とアンダースコアのみ）",
                        default="",
                        validate=r"^[a-zA-Z0-9_]+$"
                    ),
                    "ヘッダー": st.column_config.TextColumn(
                        "ヘッダー",
                        help="Excelに表示されるヘッダー名",
                        default=""
                    ),
                    "説明": st.column_config.TextColumn(
                        "説明",
                        help="この列の説明（オプション）",
                        default=""
                    )
                },
                hide_index=True,
                key="column_editor"
            )
            
            # 編集されたデータフレームから辞書形式に変換
            temp_definitions = {}
            for _, row in edited_df.iterrows():
                if row["列"] and row["キー"] and row["ヘッダー"]:
                    temp_definitions[row["列"]] = {
                        "key": row["キー"],
                        "header": row["ヘッダー"],
                        "description": row["説明"]
                    }
            
            st.session_state.custom_column_definitions = temp_definitions
            
            # プレビュー表示
            if temp_definitions:
                with st.expander("列定義プレビュー", expanded=False):
                    st.json(temp_definitions)
            
            if st.button("カスタムタスクを作成", type="primary"):
                if custom_name and custom_description:
                    output_config = {
                        "target_sheet": target_sheet,
                        "start_row": start_row,
                        "column_definitions": st.session_state.custom_column_definitions
                    }
                    
                    task_id = st.session_state.task_engine.create_custom_task_config(
                        custom_name, custom_description, output_config
                    )
                    
                    st.success(f"カスタムタスク '{custom_name}' が作成されました (ID: {task_id})")
                    st.rerun()
                else:
                    st.error("タスク名と説明を入力してください")

def show_evidence_upload():
    """証跡アップロードUI"""
    st.markdown('<div class="card-header">証跡フォルダのアップロード</div>', unsafe_allow_html=True)
    
    with st.container(border=True):
        st.markdown("### フォルダ構造の説明")
        st.markdown("""
        **必要な構造**: 2階層フォルダ
        - 1階層目: データ識別フォルダ（例: データ001, データ002）
        - 2階層目: 各データに関連するドキュメント（PDF、画像など）
        
        **例**:
        ```
        証跡フォルダ.zip
        ├── データ001/
        │   ├── 請求書.pdf
        │   └── 入金明細.png
        ├── データ002/
        │   ├── 請求書.pdf
        │   └── 入金明細.jpg
        └── ...
        ```
        """)
    
    uploaded_folder = st.file_uploader(
        "証跡フォルダ（ZIP形式）をアップロードしてください",
        type=['zip'],
        help="2階層構造のフォルダをZIP圧縮してアップロードしてください"
    )
    
    if uploaded_folder:
        # フォルダ構造を検証
        is_valid, error_msg = st.session_state.folder_processor.validate_folder_structure(uploaded_folder)
        
        if is_valid:
            st.success(f"フォルダ '{uploaded_folder.name}' が正常にアップロードされました")
            
            if st.button("証跡フォルダを処理", type="primary"):
                with st.spinner("証跡データを処理中..."):
                    processed_evidence = st.session_state.folder_processor.process_evidence_folder(uploaded_folder)
                    st.session_state.processed_evidence = processed_evidence
                    
                    if processed_evidence["success"]:
                        st.success("証跡データの処理が完了しました")
                        
                        # サマリー表示
                        summary = st.session_state.folder_processor.get_data_summary(processed_evidence)
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("データ数", summary.get("total_data_entries", 0))
                        with col2:
                            st.metric("ドキュメント数", summary.get("total_documents", 0))
                        with col3:
                            st.metric("ドキュメント種類", len(summary.get("document_types", {})))
                        
                        # ドキュメント種類別統計
                        if "document_types" in summary:
                            st.markdown("**ドキュメント種類別統計**")
                            for doc_type, count in summary["document_types"].items():
                                st.markdown(f"- {doc_type}: {count}件")
                        
                        # デバッグ情報（開発時のみ表示）
                        # with st.expander("詳細情報（デバッグ用）"):
                        #     metadata = processed_evidence.get("metadata", {})
                        #     if "zip_structure" in metadata:
                        #         st.markdown("**ZIPファイル内構造:**")
                        #         for file_path in metadata["zip_structure"]:
                        #             st.code(file_path)
                            
                        #     if "detected_data_folders" in metadata:
                        #         st.markdown("**検出されたデータフォルダ:**")
                        #         for folder in metadata["detected_data_folders"]:
                        #             st.markdown(f"- {folder}")
                    else:
                        st.error(f"証跡データ処理エラー: {processed_evidence.get('error', '不明なエラー')}")
        else:
            st.error(f"フォルダ構造エラー: {error_msg}")

def show_excel_upload():
    """調書アップロードUI"""
    st.markdown('<div class="card-header">調書（Excel）のアップロード</div>', unsafe_allow_html=True)
    
    uploaded_excel = st.file_uploader(
        "調書Excelファイルをアップロードしてください",
        type=['xlsx', 'xls'],
        help="結果を記載するExcelファイルをアップロードしてください"
    )
    
    if uploaded_excel:
        # Excelファイルを検証
        is_valid, error_msg = st.session_state.excel_manager.validate_excel_file(uploaded_excel)
        
        if is_valid:
            st.success(f"調書 '{uploaded_excel.name}' が正常にアップロードされました")
            
            if st.button("調書を読み込み", type="primary", key="load_excel_btn"):
                with st.spinner("調書を読み込み中..."):
                    load_result = st.session_state.excel_manager.load_workbook(uploaded_excel)
                    
                    if load_result["success"]:
                        # 読み込み結果をsession_stateに保存
                        st.session_state.excel_load_result = load_result
                        st.success("調書の読み込みが完了しました")
                    else:
                        st.error(f"調書読み込みエラー: {load_result.get('error', '不明なエラー')}")
            
            # 読み込み済みの場合、ワークブック情報とプレビューを表示
            if hasattr(st.session_state, 'excel_load_result') and st.session_state.excel_load_result:
                load_result = st.session_state.excel_load_result
                
                # ワークブック情報を表示
                st.markdown("### ワークブック情報")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"**ファイル名**: {load_result['file_name']}")
                    st.info(f"**シート数**: {len(load_result['sheet_names'])}")
                
                with col2:
                    st.markdown("**シート一覧（表形式）**:")
                    sheet_table = []
                    for sheet_name in load_result['sheet_names']:
                        sheet_info = load_result['sheet_info'][sheet_name]
                        status = "✅ データあり" if sheet_info['has_data'] else "📄 空のシート"
                        sheet_table.append({"シート名": sheet_name, "状態": status})
                    st.dataframe(pd.DataFrame(sheet_table), hide_index=True)
                
                # シートプレビュー
                st.markdown("### シートプレビュー")
                selected_sheet = st.selectbox(
                    "プレビューするシートを選択",
                    load_result['sheet_names'],
                    key="sheet_selector"
                )
                
                if selected_sheet:
                    preview_data = st.session_state.excel_manager.get_sheet_preview(selected_sheet)
                    if preview_data:
                        df_preview = pd.DataFrame(preview_data)
                        st.dataframe(df_preview, hide_index=True)
                    else:
                        st.info("このシートにはデータがありません")
        else:
            st.error(f"Excelファイルエラー: {error_msg}")

def show_processing_execution():
    """処理実行UI"""
    st.markdown('<div class="card-header">処理実行</div>', unsafe_allow_html=True)
    
    # 前提条件チェック
    status_checks = []
    
    # タスク選択チェック
    if "selected_task_id" not in st.session_state:
        status_checks.append(("タスク選択", False, "タスクを選択してください"))
    else:
        status_checks.append(("タスク選択", True, f"選択済み: {st.session_state.selected_task_id}"))
    
    # 証跡データチェック
    if "processed_evidence" not in st.session_state:
        status_checks.append(("証跡データ", False, "証跡フォルダをアップロードして処理してください"))
    else:
        evidence_data = st.session_state.processed_evidence
        if evidence_data.get("success"):
            data_count = len(evidence_data.get("data", {}))
            status_checks.append(("証跡データ", True, f"{data_count}件のデータを読み込み済み"))
        else:
            status_checks.append(("証跡データ", False, "証跡データの処理に失敗しています"))
    
    # 調書チェック
    if not hasattr(st.session_state, 'excel_load_result') or not st.session_state.excel_load_result:
        status_checks.append(("調書", False, "調書Excelファイルをアップロードして読み込んでください"))
    else:
        load_result = st.session_state.excel_load_result
        status_checks.append(("調書", True, f"調書（{load_result['file_name']}）"))
    
    # APIキーチェック
    provider = os.getenv("OPENAI_PROVIDER", "openai")
    env_api_key = os.getenv("AZURE_OPENAI_API_KEY" if provider == "azure" else "OPENAI_API_KEY")
    has_api_key = env_api_key or (hasattr(st.session_state, "openai_api_key") and st.session_state.openai_api_key and st.session_state.openai_api_key != "[環境変数から設定済み]")
    
    if not has_api_key:
        status_checks.append(("API設定", False, f"{provider.upper()} APIキーを設定してください"))
    else:
        status_checks.append(("API設定", True, f"{provider.upper()} APIキー設定済み"))
    
    # ステータス表示
    col1, col2, col3, col4 = st.columns(4)
    for i, (name, status, message) in enumerate(status_checks):
        with [col1, col2, col3, col4][i]:
            if status:
                st.success(f"**{name}**\n\n{message}")
            else:
                st.error(f"**{name}**\n\n{message}")
    
    # 全ての前提条件が満たされていない場合は終了
    if not all(status for _, status, _ in status_checks):
        return
    
    # タスク情報の表示
    selected_task_id = st.session_state.selected_task_id
    task_config = st.session_state.task_engine.get_task_config(selected_task_id)
    
    if task_config:
        with st.container(border=True):
            st.markdown("### 実行するタスク")
            st.info(f"""**タスク名**: {task_config['name']}
            
**説明**: {task_config.get('description', '')}""")
    
    # 補足事項入力
    with st.container(border=True):
        st.markdown("### 補足事項")
        
        additional_info = st.text_area(
            "タスクの補足事項があれば入力してください（任意）",
            placeholder="例: 特定の条件で追加チェックが必要な場合など、タスク実行に関する補足情報を記載してください。",
            height=100,
            help="タスク定義に含まれていない追加の指示がある場合に入力してください"
        )
        
        if additional_info:
            st.info(f"**補足事項**: {additional_info}")
    
    # 実行
    with st.container(border=True):
        st.markdown("### 実行")
        
        if st.button("処理実行", type="primary"):
                # workbookが利用可能かチェック
                if not hasattr(st.session_state.excel_manager, 'workbook') or not st.session_state.excel_manager.workbook:
                    st.error("調書が読み込まれていません。調書をアップロードして読み込んでください。")
                else:
                    # 進捗表示用のコンテナ
                    progress_container = st.container()
                    with progress_container:
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                    # 並列処理数を取得
                    max_workers = st.session_state.get("max_workers", 3)
                    
                    # データ件数を取得
                    data_count = len(st.session_state.processed_evidence.get("data", {}))
                    status_text.text(f"処理開始: {data_count}件のデータを{max_workers}並列で処理します")
                    
                    # 進捗更新用のコールバック関数
                    def update_progress(completed, total, data_id):
                        progress = completed / total
                        progress_bar.progress(progress)
                        status_text.text(f"処理中... {completed}/{total} 完了 (最新: {data_id})")
                    
                    # タスク定義の説明を使用し、補足事項があれば追加
                    task_description = task_config.get('description', '')
                    if additional_info:
                        combined_instruction = f"{task_description}\n\n補足事項: {additional_info}"
                    else:
                        combined_instruction = task_description
                    
                    execution_result = st.session_state.task_engine.execute_accounting_task(
                        st.session_state.selected_task_id,
                        st.session_state.processed_evidence,
                        st.session_state.excel_manager,
                        combined_instruction,
                        max_workers=max_workers,
                        progress_callback=update_progress
                    )
                    
                    # 処理完了後の表示更新
                    progress_bar.progress(1.0)
                    status_text.empty()
                    
                    st.session_state.execution_result = execution_result
                    
                    if execution_result["success"]:
                        st.success("経理業務処理が完了しました")
                        
                        # サマリー表示
                        summary = execution_result["summary"]
                        
                        col_a, col_b, col_c, col_d = st.columns(4)
                        with col_a:
                            st.metric("総データ数", summary.get("total_data_count", 0))
                        with col_b:
                            st.metric("成功処理数", summary["processed_data_count"])
                        with col_c:
                            st.metric("失敗数", summary.get("failed_data_count", 0))
                        with col_d:
                            st.metric("使用トークン", summary["tokens_used"])
                            
                        # エラー情報表示
                        if summary.get("failed_data_count", 0) > 0:
                            st.warning(f"⚠️ {summary['failed_data_count']}件のデータで処理エラーが発生しました")
                            if st.expander("エラー詳細を表示"):
                                errors = summary.get("processing_details", {}).get("processing_errors", [])
                                for error in errors:
                                    st.error(error)
                        
                        st.info(f"**結果書き込み先**: {summary['target_sheet']}シート {summary['excel_range']}")
                    else:
                        st.error(f"処理エラー: {execution_result.get('error', '不明なエラー')}")

def show_processing_results():
    """処理結果確認UI"""
    st.markdown('<div class="card-header">処理結果</div>', unsafe_allow_html=True)
    
    if "execution_result" not in st.session_state:
        st.info("まだ処理が実行されていません。先に経理業務処理を実行してください。")
        return
    
    execution_result = st.session_state.execution_result
    
    if not execution_result.get("success"):
        st.error("処理が失敗しています。処理実行画面で再度実行してください。")
        return
    
    # 実行サマリー
    summary = execution_result["summary"]
    
    with st.container(border=True):
        st.markdown("### 実行サマリー")
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"**タスク名**: {summary['task_name']}")
        
        with col2:
            st.info(f"**処理データ数**: {summary['processed_data_count']} 件")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"**書き込み行数**: {summary['written_rows']} 行")
        with col2:
            st.info(f"**書き込み対象シート**: {summary['target_sheet']}")
        with col3:
            st.info(f"**書き込み範囲**: {summary['excel_range']}")


    # 処理されたExcelファイルのダウンロード
    with st.container(border=True):
        st.markdown("### 結果ダウンロード")
        
        if hasattr(st.session_state.excel_manager, 'workbook') and st.session_state.excel_manager.workbook:
            try:
                excel_data = st.session_state.excel_manager.save_workbook()
                
                st.download_button(
                    label="処理済み調書をダウンロード",
                    data=excel_data,
                    file_name=f"処理済み調書_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    help="処理結果が記載されたExcelファイルをダウンロードします"
                )
            except Exception as e:
                st.error(f"ダウンロード準備エラー: {str(e)}")
        else:
            st.warning("調書が読み込まれていないため、ダウンロードできません。")
    
    # 詳細結果表示
    with st.container(border=True):
        st.markdown("### 実行結果")
        
        llm_result = execution_result.get("llm_result", {})
        
        # デバッグ情報表示
        # with st.expander("デバッグ情報", expanded=False):
        #     st.write("llm_result keys:", list(llm_result.keys()))
        #     if "raw_responses" in llm_result:
        #         st.write("raw_responses:", llm_result["raw_responses"])
        #     else:
        #         st.write("raw_responses not found in llm_result")
        
        if llm_result.get("success") and "data" in llm_result:
            llm_data = llm_result["data"]
            
            if "results" in llm_data:
                df_results = pd.DataFrame(llm_data["results"])
                st.dataframe(df_results, hide_index=True, use_container_width=True)
            
            # RAWレスポンス表示（デバッグ用）
            # with st.expander("RAWレスポンス（デバッグ用）", expanded=False):
            #     raw_responses = llm_result.get("raw_responses", {})
                
            #     if raw_responses:
            #         st.markdown(f"**{len(raw_responses)}件のLLMレスポンス**")
                    
            #         # タブでデータごとに表示
            #         if len(raw_responses) > 1:
            #             tab_names = list(raw_responses.keys())
            #             tabs = st.tabs(tab_names)
                        
            #             for i, (data_id, response) in enumerate(raw_responses.items()):
            #                 with tabs[i]:
            #                     st.text_area(
            #                         f"{data_id}のレスポンス",
            #                         value=response,
            #                         height=400,
            #                         disabled=True
            #                     )
            #         else:
            #             # 1件のみの場合はそのまま表示
            #             for data_id, response in raw_responses.items():
            #                 st.text_area(
            #                     f"{data_id}のレスポンス",
            #                     value=response,
            #                     height=400,
            #                     disabled=True
            #                 )
            #     else:
            #         # 従来の単一レスポンス表示（後方互換性）
            #         single_response = llm_result.get("raw_response", "レスポンスなし")
            #         st.text_area(
            #             "レスポンス",
            #             value=single_response,
            #             height=400,
            #             disabled=True
            #         )

# 既存の請求書チェック機能（元のapp.pyから移植）
def show_rule_management():
    """既存のルール管理機能"""
    st.markdown('<div class="card-header">チェックルール設定</div>', unsafe_allow_html=True)
    
    # 既存ルール表示（共通）
    with st.container(border=True):
        st.markdown(f'<div class="card-header">既存のルール</div>', unsafe_allow_html=True)
        show_existing_rules()

    st.divider()
    st.markdown('<div class="card-header">新しいルールを追加</div>', unsafe_allow_html=True)
    # タブを追加：手動追加とルール提案
    tab1, tab2 = st.tabs(["手動ルール追加", "ルール提案"])
    
    with tab1:
        with st.container():
            show_manual_rule_creation()
    
    with tab2:
        with st.container():
            show_rule_suggestion()

def show_manual_rule_creation():
    """手動でのルール作成UI"""      
    st.info('手動でルールを追加できます')
    with st.container(border=True):
        rule_name = st.text_input(
            "ルール名",
            placeholder="例: 請求書日付チェック",
            help="ルールを識別するための名前を入力してください",
            key="manual_rule_name"
        )
        
        rule_category = st.selectbox(
            "カテゴリ",
            ["日付チェック", "金額チェック", "承認チェック", "書式チェック", "その他"],
            help="ルールのカテゴリを選択してください",
            key="manual_rule_category"
        )
        
        rule_prompt = st.text_area(
            "チェックルールの内容",
            placeholder="ルールの内容を入力してください...\n例: 請求書の日付が期末日以前かどうかを確認してください。\n- 請求書日付を特定してください\n- 期末日は通常3月31日ですが、文書に記載がある場合はそれに従ってください",
            height=300,
            help="具体的なチェック内容を記述してください。",
            key="manual_rule_prompt"
        )
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        if st.button("ルールを追加", type="primary", use_container_width=True, key="add_manual_rule"):
            if rule_name and rule_prompt:
                st.session_state.rules.add_rule(
                    rule_name, rule_category, rule_prompt
                )
                st.success(f"ルール '{rule_name}' が追加されました")
                st.rerun()
            else:
                st.error("ルール名とチェックルールの内容を入力してください")

def show_rule_suggestion():
    """ルール提案機能のUI"""
    st.info('アップロードしたドキュメントからルールを提案します')
    with st.container(border=True):
        col1, col2 = st.columns([1, 1], gap="large")
        
        with col1:            
            uploaded_doc = st.file_uploader(
                "マニュアルや規定書をアップロードしてください",
                type=['pdf', 'docx', 'doc', 'txt', 'md', 'xlsx', 'xls'],
                help="PDF、Word、Excel、テキストファイルに対応しています",
                key="rule_suggestion_file"
            )
            
            if uploaded_doc:
                st.success(f"ファイル '{uploaded_doc.name}' がアップロードされました")
                
                if st.button("ルールを提案", type="primary", use_container_width=True):
                    with st.spinner("ドキュメントを分析中..."):
                        try:
                            # ドキュメントからテキスト抽出
                            document_content = st.session_state.suggester.process_uploaded_document(uploaded_doc)
                            
                            # 既存ルールを取得
                            existing_rules = st.session_state.rules.get_all_rules()
                            
                            # ルール提案を生成
                            suggested_rules = st.session_state.suggester.suggest_rules_from_document(
                                document_content, existing_rules
                            )
                            
                            if suggested_rules:
                                st.session_state.suggested_rules = suggested_rules
                                st.success(f"{len(suggested_rules)}件のルールが提案されました")
                                st.rerun()
                            else:
                                st.warning("提案できるルールが見つかりませんでした")
                                
                        except Exception as e:
                            st.error(f"ルール提案の生成に失敗しました: {str(e)}")
        
        with col2:
            st.markdown("**💡 ルール提案の使い方**")
            st.markdown("""
            1. ルール提案のインプットとして、マニュアルや規定書をアップロードしてください
            2. 提案されるルールを確認し、必要に応じて内容を修正してください
            3. 「このルールを追加」ボタンを押すと、ルールとして登録されます
            """)
        
        # 提案されたルールの表示
        st.divider()
        if "suggested_rules" in st.session_state and st.session_state.suggested_rules:
            st.info(f"**{len(st.session_state.suggested_rules)}件のルールが提案されました**")
            for i, suggested_rule in enumerate(st.session_state.suggested_rules):
                with st.expander(f"提案 {i+1}: {suggested_rule.get('name', '無題')}", expanded=False):
                    
                    col_details, col_actions = st.columns([2, 1])
                    
                    with col_details:
                        st.markdown(f"**カテゴリ**: {suggested_rule.get('category', '未分類')}")
                        st.markdown(f"**信頼度**: {suggested_rule.get('confidence', 0):.0%}")
                        if suggested_rule.get('reason'):
                            st.markdown(f"**提案理由**: {suggested_rule['reason']}")
                        
                        st.markdown("**チェック内容**:")
                        st.code(suggested_rule.get('prompt', ''), language="text")
                    
                    with col_actions:
                        st.markdown("**アクション**")
                        
                        # ルール編集
                        if st.button("内容を編集", key=f"edit_suggested_{i}", use_container_width=True):
                            st.session_state[f"editing_suggested_{i}"] = True
                            st.rerun()
                        
                        # ルール追加
                        if st.button("このルールを追加", key=f"add_suggested_{i}", type="primary", use_container_width=True):
                            try:
                                rule_id = st.session_state.rules.add_rule(
                                    suggested_rule['name'],
                                    suggested_rule['category'],
                                    suggested_rule['prompt']
                                )
                                st.success(f"ルール '{suggested_rule['name']}' を追加しました")
                                # 提案リストから削除
                                st.session_state.suggested_rules.pop(i)
                                st.rerun()
                            except Exception as e:
                                st.error(f"ルール追加エラー: {str(e)}")
                    
                    # 編集モード
                    if st.session_state.get(f"editing_suggested_{i}", False):
                        st.markdown("---")
                        st.markdown("**ルールを編集**")
                        
                        edited_name = st.text_input(
                            "ルール名", 
                            value=suggested_rule.get('name', ''),
                            key=f"edit_name_{i}"
                        )
                        
                        edited_category = st.selectbox(
                            "カテゴリ",
                            ["日付チェック", "金額チェック", "承認チェック", "書式チェック", "その他"],
                            index=["日付チェック", "金額チェック", "承認チェック", "書式チェック", "その他"].index(
                                suggested_rule.get('category', 'その他')
                            ),
                            key=f"edit_category_{i}"
                        )
                        
                        edited_prompt = st.text_area(
                            "チェック内容",
                            value=suggested_rule.get('prompt', ''),
                            height=200,
                            key=f"edit_prompt_{i}"
                        )
                        
                        col_save, col_cancel = st.columns(2)
                        with col_save:
                            if st.button("変更を保存", key=f"save_edit_{i}", type="primary"):
                                # 編集内容でルールを追加
                                st.session_state.rules.add_rule(edited_name, edited_category, edited_prompt)
                                st.success(f"編集したルール '{edited_name}' を追加しました")
                                # 編集モードを終了し、提案リストから削除
                                del st.session_state[f"editing_suggested_{i}"]
                                st.session_state.suggested_rules.pop(i)
                                st.rerun()
                        
                        with col_cancel:
                            if st.button("編集をキャンセル", key=f"cancel_edit_{i}"):
                                del st.session_state[f"editing_suggested_{i}"]
                                st.rerun()
            
            # 全ての提案をクリア
            if st.button("提案をクリア", type="secondary"):
                del st.session_state.suggested_rules
                st.rerun()

def show_existing_rules():
    """既存ルールの表示"""
    rules = st.session_state.rules.get_all_rules()
    
    if rules:
        # ルール数の表示
        st.info(f"登録済みルール数: {len(rules)} 件")
        st.markdown("")
        
        for rule_id, rule in rules.items():
            with st.expander(f"{rule['name']} ({rule['category']})", expanded=False):
                st.markdown(f"**チェック内容**")
                st.code(rule['prompt'], language="text")
                
                st.markdown(f"**作成日**: {rule['created_at']}")
                
                col_edit, col_delete = st.columns([1, 1])
                with col_delete:
                    if st.button("削除", key=f"delete_{rule_id}", type="secondary", use_container_width=True):
                        st.session_state.rules.delete_rule(rule_id)
                        st.success(f"ルール '{rule['name']}' が削除されました")
                        st.rerun()
    else:
        st.info("まだルールが設定されていません。上記のフォームから新しいルールを追加してください。")

def show_file_upload():
    """既存のファイルアップロード機能"""
    st.markdown('<div class="card-header">請求書ファイルのアップロード</div>', unsafe_allow_html=True)
        
    uploaded_files = st.file_uploader(
        "ファイルを選択またはドラッグ&ドロップしてください",
        type=['pdf', 'xlsx', 'xls', 'docx', 'doc'],
        accept_multiple_files=True,
        help="複数のファイルを同時に選択できます。最大サイズ: 50MB/ファイル"
    )
        
    if uploaded_files:
        st.success(f"{len(uploaded_files)} 個のファイルがアップロードされました")
        
        # ファイル一覧表示
        file_data = []
        total_size = 0
        
        for file in uploaded_files:
            size_kb = file.size / 1024
            total_size += size_kb
            
            # ファイルサイズの表示形式
            if size_kb < 1024:
                size_str = f"{size_kb:.1f} KB"
            else:
                size_str = f"{size_kb/1024:.1f} MB"
            
            file_data.append({
                "ファイル名": file.name,
                "サイズ": size_str,
                "タイプ": file.type.split('/')[-1] if file.type else "不明"
            })
        
        # サマリー情報
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("ファイル数", len(uploaded_files))
        with col2:
            if total_size < 1024:
                st.metric("合計サイズ", f"{total_size:.1f} KB")
            else:
                st.metric("合計サイズ", f"{total_size/1024:.1f} MB")
        with col3:
            avg_size = total_size / len(uploaded_files)
            if avg_size < 1024:
                st.metric("平均サイズ", f"{avg_size:.1f} KB")
            else:
                st.metric("平均サイズ", f"{avg_size/1024:.1f} MB")
        
        st.markdown("---")
        
        # ファイル詳細リスト
        df = pd.DataFrame(file_data)
        st.dataframe(
            df, 
            use_container_width=True,
            hide_index=True,
            column_config={
                "ファイル名": st.column_config.TextColumn(
                    "ファイル名",
                    width="large"
                ),
                "サイズ": st.column_config.TextColumn(
                    "サイズ",
                    width="small"
                ),
                "タイプ": st.column_config.TextColumn(
                    "タイプ",
                    width="small"
                )
            }
        )
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # セッション状態に保存
        st.session_state.uploaded_files = uploaded_files
        
        # 処理ボタン
        if st.button("ファイルを処理", type="primary", use_container_width=True):
            with st.spinner("ファイルを処理中..."):
                processed_data = st.session_state.processor.process_files(uploaded_files)
                st.session_state.processed_data = processed_data
                st.success(f"{len(processed_data)} 件のファイルが処理されました")

def show_check_execution():
    """既存のチェック実行機能"""
    st.markdown('<div class="card-header">請求書チェック実行</div>', unsafe_allow_html=True)
    
    # 前提条件チェック
    status_checks = []
    
    if "processed_data" not in st.session_state:
        status_checks.append(("ファイル処理", False, "ファイルをアップロードして処理してください"))
    else:
        status_checks.append(("ファイル処理", True, f"{len(st.session_state.processed_data)} 件のファイルが処理済み"))
    
    if not st.session_state.rules.get_all_rules():
        status_checks.append(("チェックルール", False, "チェックルールを設定してください"))
    else:
        status_checks.append(("チェックルール", True, f"{len(st.session_state.rules.get_all_rules())} 件のルールが設定済み"))
    
    provider = os.getenv("OPENAI_PROVIDER", "openai")
    env_api_key = os.getenv("AZURE_OPENAI_API_KEY" if provider == "azure" else "OPENAI_API_KEY")
    has_api_key = env_api_key or (hasattr(st.session_state, "openai_api_key") and st.session_state.openai_api_key and st.session_state.openai_api_key != "[環境変数から設定済み]")
    
    if not has_api_key:
        status_checks.append(("API設定", False, f"{provider.upper()} APIキーを設定してください"))
    else:
        status_checks.append(("API設定", True, f"GPT-4.1 APIキーが設定済み ({provider.upper()})"))
        
    col1, col2, col3 = st.columns(3)
    for i, (name, status, message) in enumerate(status_checks):
        with [col1, col2, col3][i]:
            if status:
                st.success(f"**{name}**\n\n{message}")
            else:
                st.error(f"**{name}**\n\n{message}")
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # すべての前提条件が満たされていない場合は終了
    if not all(status for _, status, _ in status_checks):
        return
    
    # チェック設定セクション
    with st.container(border=True):
        st.markdown(f'<div class="card-header">チェック設定</div>', unsafe_allow_html=True)

        col1, col2 = st.columns([1, 1], gap="large")
        
        with col1:
            st.markdown("**適用するルール**")
            available_rules = st.session_state.rules.get_all_rules()
            selected_rules = st.multiselect(
                "チェックに使用するルールを選択",
                options=list(available_rules.keys()),
                format_func=lambda x: f"{available_rules[x]['name']} ({available_rules[x]['category']})",
                default=list(available_rules.keys()),
                help="複数のルールを同時に適用できます"
            )
            
            if selected_rules:
                st.info(f"選択済み: {len(selected_rules)} 件のルール")
        
        with col2:
            st.markdown("**処理対象ファイル**")
            processed_files = list(st.session_state.processed_data.keys())
            selected_files = st.multiselect(
                "チェックするファイルを選択",
                options=processed_files,
                default=processed_files,
                help="チェックしたいファイルを選択してください"
            )
            
            if selected_files:
                st.info(f"選択済み: {len(selected_files)} 件のファイル")
    
    # 実行予定の概要
    if selected_rules and selected_files:
        with st.container(border=True):
            st.markdown(f'<div class="card-header">チェック実行</div>', unsafe_allow_html=True)
        
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("対象ファイル", len(selected_files))
            with col2:
                st.metric("適用ルール", len(selected_rules))
            with col3:
                total_checks = len(selected_files) * len(selected_rules)
                st.metric("総チェック数", total_checks)
            
            # 推定処理時間
            estimated_time = round((total_checks/st.session_state.max_workers)+1) * 3   # 1チェックあたり約3秒と仮定 並列処理を考慮
            st.markdown(f"**推定処理時間**: 約 {estimated_time//60} 分 {estimated_time%60} 秒")
            
            st.markdown("</div>", unsafe_allow_html=True)
            
            # 実行ボタン
            if st.button("チェック開始", type="primary", use_container_width=True):
                run_invoice_check(selected_rules, selected_files)
    else:
        st.warning("ルールとファイルの両方を選択してください")

def run_invoice_check(selected_rules, selected_files):
    """請求書チェックを実行"""
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_files = len(selected_files)
    results = {}
    
    # 並列処理で実行
    with concurrent.futures.ThreadPoolExecutor(max_workers=st.session_state.max_workers) as executor:
        futures = {}
        
        for i, file_name in enumerate(selected_files):
            file_data = st.session_state.processed_data[file_name]
            future = executor.submit(
                st.session_state.checker.check_invoice,
                file_data,
                selected_rules
            )
            futures[future] = file_name
        
        # 結果を収集
        completed = 0
        status_text.text("チェック中...")
        for future in concurrent.futures.as_completed(futures):
            file_name = futures[future]
            try:
                result = future.result()
                results[file_name] = result
                completed += 1
                
                progress = completed / total_files
                progress_bar.progress(progress)
                status_text.text(f"処理中... {completed}/{total_files} 完了")
                
            except Exception as e:
                st.error(f"ファイル {file_name} の処理中にエラーが発生しました: {str(e)}")
                results[file_name] = {"error": str(e)}
    
    # 結果をセッション状態に保存
    st.session_state.check_results = results
    status_text.empty()
    st.session_state.check_timestamp = datetime.now()
    
    progress_bar.progress(1.0)
    st.success(f"{len(results)}件のファイルのチェックが完了しました")

def show_results():
    """既存の結果表示機能"""
    with st.container(border=True):
        st.markdown(f'<div class="card-header">チェック結果</div>', unsafe_allow_html=True)
    
        if "check_results" not in st.session_state:
            st.info("まだチェックが実行されていません。先に請求書ファイルをアップロードし、チェック実行を行ってください。")
            return
        
        results = st.session_state.check_results
        timestamp = st.session_state.check_timestamp
        
        # 実行情報
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"**実行時刻**\n\n{timestamp.strftime('%Y年%m月%d日 %H:%M:%S')}")
        with col2:
            st.info(f"**処理件数**\n\n{len(results)} ファイル")
                    
        # メトリクス計算（ファイル単位）
        total_files = len(results)
        needs_review_files = 0  # 要確認（エラーまたは警告があるファイル）
        ok_files = 0  # すべてのチェックが正常なファイル
        
        # デバッグ用：各ファイルの分類詳細
        file_classifications = []
        
        for file_name, result in results.items():
            checks = result.get("checks", [])
            has_error_check = any(check.get("severity") == "error" for check in checks)
            has_warning_check = any(check.get("severity") == "warning" for check in checks)
            
            if has_error_check or has_warning_check:
                # エラーまたは警告があるファイル → 要確認
                needs_review_files += 1
                error_count_in_file = sum(1 for check in checks if check.get("severity") == "error")
                warning_count_in_file = sum(1 for check in checks if check.get("severity") == "warning")
                file_classifications.append(f"{file_name}: 要確認 (エラー:{error_count_in_file}, 警告:{warning_count_in_file})")
            else:
                # すべてのチェックが正常なファイル
                ok_files += 1
                file_classifications.append(f"{file_name}: 正常 ({len(checks)}チェック全て正常)")
        
        # チェック単位での詳細計算
        total_checks = 0
        total_check_errors = 0
        total_check_warnings = 0
        total_check_success = 0
        
        for result in results.values():
            if "error" not in result:
                checks = result.get("checks", [])
                total_checks += len(checks)
                total_check_errors += sum(1 for check in checks if check.get("severity") == "error")
                total_check_warnings += sum(1 for check in checks if check.get("severity") == "warning")
                total_check_success += sum(1 for check in checks if check.get("severity") == "info")
        
        # ファイル単位のサマリー（2列構成）
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                label="総ファイル数",
                value=total_files,
                help="処理された総ファイル数"
            )
        
        with col2:
            st.metric(
                label="要確認",
                value=needs_review_files,
                help="エラーまたは警告が含まれるファイル数"
            )
        
        with col3:
            st.metric(
                label="正常",
                value=ok_files,
                help="すべてのチェックが正常なファイル数"
            )
        
    # 詳細結果表示
    with st.container(border=True):
        st.markdown(f'<div class="card-header">詳細結果</div>', unsafe_allow_html=True)
    
        # フィルターとダウンロード
        col1, col2 = st.columns([1, 1], gap="large")
        
        with col1:
            filter_option = st.selectbox(
                "表示フィルター",
                ["すべて", "要確認", "正常"],
                help="結果の表示をフィルタリングできます"
            )
        
        with col2:
            st.write(" ")
            # 結果のダウンロード
            excel_data = create_excel_report(results)
            st.download_button(
                label="結果をダウンロード",
                data=excel_data,
                file_name=f"invoice_check_results_{timestamp.strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                help="チェック結果をExcel形式でダウンロードします"
            )
        
        # 結果一覧
        filtered_results = {k: v for k, v in results.items() if should_show_result(v, filter_option)}
        
        if filtered_results:
            st.markdown(f"**表示中**: {len(filtered_results)} / {len(results)} ファイル")
            
            for file_name, result in filtered_results.items():
                show_file_result(file_name, result)
        else:
            st.info(f"フィルター条件 '{filter_option}' に該当する結果がありません。")

def should_show_result(result, filter_option):
    """フィルター条件に基づいて結果を表示するかどうかを判定"""
    if filter_option == "すべて":
        return True
    elif filter_option == "要確認":
        checks = result.get("checks", [])
        return any(check.get("severity") in ["error", "warning"] for check in checks)
    elif filter_option == "正常":
        checks = result.get("checks", [])
        return not any(check.get("severity") in ["error", "warning"] for check in checks)
    return True

def show_file_result(file_name, result):
    """個別ファイルの結果を表示"""
    st.markdown('<hr style="margin-top: 4px; margin-bottom: 4px; border: none; border-top: 1.5px solid #e0e0e0;">', unsafe_allow_html=True)
    result_tab1, result_tab2 = st.columns([0.8, 1.2], gap="large")
    
    with result_tab1:
        if "error" in result:
            # 処理エラーの場合
            st.markdown(f"""
            <div style="background-color: #f8d7da; padding: 12px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid #dc3545;">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="background-color: #dc3545; color: white; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: bold; display: flex; align-items: center; gap: 5px;">
                        ❌ 処理エラー
                    </span>
                    <strong style="color: #721c24; font-size: 14px;">{file_name}</strong>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            with st.expander("エラー詳細を表示", expanded=False):
                st.error(f"**処理エラー**\n\n{result['error']}")
        else:
            checks = result.get("checks", [])
            has_warning = any(check.get("severity") == "warning" for check in checks)
            has_error = any(check.get("severity") == "error" for check in checks)
            
            if has_error or has_warning:
                # 要確認の場合
                status = "⚠️ 要確認"
                badge_color = "#856404"
                bg_color = "#fff3cd"
                border_color = "#ffc107"
                expanded = True
            else:
                # 正常の場合
                status = "✅ 正常"
                badge_color = "#155724"
                bg_color = "#d4edda"
                border_color = "#28a745"
                expanded = False
            
            st.markdown(f"""
            <div style="background-color: {bg_color}; padding: 14px; border-radius: 8px; margin-bottom: 10px; border-left: 4px solid {border_color};">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <span style="background-color: {badge_color}; color: white; padding: 5px 14px 5px 14px; border-radius: 20px; font-size: 18px; font-weight: bold; display: flex; align-items: center; gap: 5px;">
                        {status}
                    </span>
                    <strong style="color: {badge_color}; font-size: 18px;">{file_name}</strong>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
    with result_tab2:
        if "error" not in result:
            checks = result.get("checks", [])
            if checks:
                # チェック結果のサマリー計算
                check_error_count = sum(1 for check in checks if check.get("severity") == "error")
                check_warning_count = sum(1 for check in checks if check.get("severity") == "warning")
                check_success_count = sum(1 for check in checks if check.get("severity") == "info")
                check_needs_review = check_error_count + check_warning_count
                
                # バッジ風のサマリー表示
                badges = []
                if check_success_count > 0:
                    badges.append(f'<span style="background-color: #155724; color: white; padding: 5px 15px; border-radius: 15px; font-size: 14px; font-weight: bold;">正常 {check_success_count}件</span>')
                if check_needs_review > 0:
                    badges.append(f'<span style="background-color: #856404; color: white; padding: 5px 15px; border-radius: 15px; font-size: 14px; font-weight: bold;">要確認 {check_needs_review}件</span>')
                if not badges:
                    badges.append('<span style="background-color: #6c757d; color: white; padding: 5px 15px; border-radius: 15px; font-size: 14px; font-weight: bold;">該当なし</span>')
                st.markdown(f"""
                <div style="display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap;">
                    {''.join(badges)}
                </div>
                """, unsafe_allow_html=True)
                
                with st.expander("詳細結果を表示"):
                    for i, check in enumerate(checks, 1):
                        # 個別チェック結果
                        severity = check.get("severity", "info")
                        rule_name = check.get("rule_name", "不明")
                        message = check.get("message", "")
                        details = check.get("details")

                        status = "✅ 正常" if severity == "info" else "⚠️ 要確認"

                        st.markdown(f"**{i}. {rule_name}** ：{status}")
                                                        
                        # カスタムスタイルでメッセージを中央表示
                        if severity == "error":
                            st.markdown(
                                f"""
                                <div style="
                                    background-color: #ffebee; 
                                    border: 1px solid #f44336; 
                                    border-radius: 0.5rem; 
                                    padding: 1rem; 
                                    margin: 0.5rem 0;
                                    display: flex;
                                    align-items: center;
                                    justify-content: left;
                                    text-align: left;
                                    color: #c62828;
                                    font-weight: 500;
                                ">
                                    {message}
                                </div>
                                """, 
                                unsafe_allow_html=True
                            )
                        elif severity == "warning":
                            st.markdown(
                                f"""
                                <div style="
                                    background-color: #fff3e0; 
                                    border: 1px solid #ff9800; 
                                    border-radius: 0.5rem; 
                                    padding: 1rem; 
                                    margin: 0.5rem 0;
                                    display: flex;
                                    align-items: center;
                                    justify-content: left;
                                    text-align: left;
                                    color: #f57c00;
                                    font-weight: 500;
                                ">
                                    {message}
                                </div>
                                """, 
                                unsafe_allow_html=True
                            )
                        else:
                            st.markdown(
                                f"""
                                <div style="
                                    background-color: #e8f5e8; 
                                    border: 1px solid #4caf50; 
                                    border-radius: 0.5rem; 
                                    padding: 1rem; 
                                    margin: 0.5rem 0;
                                    display: flex;
                                    align-items: center;
                                    justify-content: left;
                                    text-align: left;
                                    color: #2e7d32;
                                    font-weight: 500;
                                ">
                                    {message}
                                </div>
                                """, 
                                unsafe_allow_html=True
                            )
                        if details:
                            st.text(details)
                        
                        if i < len(checks):
                            st.markdown("---")

def create_excel_report(results):
    """結果をExcelファイルに出力"""
    import xlsxwriter
    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # サマリーシート
    summary_sheet = workbook.add_worksheet('サマリー')
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'})
    
    summary_sheet.write_row(0, 0, ['項目', '件数'], header_format)
    summary_sheet.write(1, 0, '総ファイル数')
    summary_sheet.write(1, 1, len(results))
    
    error_count = sum(1 for r in results.values() if "error" in r)
    summary_sheet.write(2, 0, 'エラー件数')
    summary_sheet.write(2, 1, error_count)
    
    # 詳細シート
    detail_sheet = workbook.add_worksheet('詳細結果')
    detail_sheet.write_row(0, 0, ['ファイル名', 'ステータス', 'ルール名', '重要度', 'メッセージ'], header_format)
    
    row = 1
    for file_name, result in results.items():
        if "error" in result:
            detail_sheet.write_row(row, 0, [file_name, 'エラー', '', 'error', result['error']])
            row += 1
        else:
            checks = result.get("checks", [])
            if not checks:
                detail_sheet.write_row(row, 0, [file_name, '正常', '', 'info', 'チェック完了'])
                row += 1
            else:
                for check in checks:
                    detail_sheet.write_row(row, 0, [
                        file_name,
                        '警告' if check.get("severity") == "warning" else '正常',
                        check.get("rule_name", ""),
                        check.get("severity", ""),
                        check.get("message", "")
                    ])
                    row += 1
    
    workbook.close()
    output.seek(0)
    return output.getvalue()

if __name__ == "__main__":
    main()