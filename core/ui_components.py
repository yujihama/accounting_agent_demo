import streamlit as st
import os
from typing import Dict, Any, Optional, Callable
import pandas as pd


class CommonUIComponents:
    """
    共通UI コンポーネント
    APIキー設定、進捗表示、結果表示などを統一化
    """
    
    @staticmethod
    def show_api_key_configuration(provider_info: Dict[str, Any]) -> Optional[str]:
        """
        APIキー設定UIを表示
        
        Args:
            provider_info: プロバイダー情報
            
        Returns:
            入力されたAPIキー（または None）
        """
        provider = provider_info.get("provider", "openai")
        
        # プロバイダー情報を表示
        st.info(f"現在のプロバイダー: {provider.upper()}")
        
        # 環境変数から取得を試行
        env_api_key = None
        if provider == "azure":
            env_api_key = os.getenv("AZURE_OPENAI_API_KEY")
        else:
            env_api_key = os.getenv("OPENAI_API_KEY")
        
        # APIキー入力フィールド
        api_key_input = st.text_input(
            f"{provider.upper()} APIキー",
            type="password",
            value=env_api_key if env_api_key else "",
            help=f"{provider.upper()} APIキーを入力してください"
        )
        
        # 設定状態を表示
        if provider_info.get("configured", False):
            st.success("✅ APIキーが設定されています")
        elif api_key_input:
            st.warning("⚠️ APIキーを設定してください（保存ボタンを押下）")
        else:
            st.error("❌ APIキーが必要です")
        
        return api_key_input if api_key_input else None
    
    @staticmethod
    def show_provider_details(provider_info: Dict[str, Any]):
        """プロバイダー詳細情報を表示"""
        with st.expander("プロバイダー詳細設定", expanded=False):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**プロバイダー**")
                st.code(provider_info.get("provider", "不明"))
                
                st.write("**モデル**")
                st.code(provider_info.get("model", "不明"))
            
            with col2:
                st.write("**設定状況**")
                if provider_info.get("configured", False):
                    st.success("設定済み")
                else:
                    st.error("未設定")
    
    @staticmethod
    def show_progress_bar(current: int, total: int, operation: str = "処理"):
        """
        進捗バーを表示
        
        Args:
            current: 現在の処理数
            total: 全体の処理数
            operation: 操作名
        """
        if total > 0:
            progress = current / total
            st.progress(progress, text=f"{operation}中... {current}/{total} ({progress*100:.1f}%)")
        else:
            st.info(f"{operation}準備中...")
    
    @staticmethod
    def show_processing_status(status: str, details: str = ""):
        """
        処理ステータスを表示
        
        Args:
            status: ステータス（processing, success, error, warning）
            details: 詳細メッセージ
        """
        if status == "processing":
            st.info(f"🔄 処理中... {details}")
        elif status == "success":
            st.success(f"✅ 完了: {details}")
        elif status == "error":
            st.error(f"❌ エラー: {details}")
        elif status == "warning":
            st.warning(f"⚠️ 警告: {details}")
        else:
            st.info(details)
    
    @staticmethod
    def show_validation_errors(validation_result: Dict[str, Any]):
        """
        入力検証エラーを表示
        
        Args:
            validation_result: 検証結果
        """
        if not validation_result.get("valid", True):
            st.error("❌ 入力エラーがあります:")
            for error in validation_result.get("errors", []):
                st.error(f"• {error}")
        
        warnings = validation_result.get("warnings", [])
        if warnings:
            st.warning("⚠️ 注意事項:")
            for warning in warnings:
                st.warning(f"• {warning}")
    
    @staticmethod
    def show_check_results(results: Dict[str, Any], title: str = "チェック結果"):
        """
        チェック結果を表示（請求書チェック用）
        
        Args:
            results: チェック結果
            title: 表示タイトル
        """
        st.markdown(f'<div class="card-header">{title}</div>', unsafe_allow_html=True)
        
        if "error" in results:
            st.error(f"エラー: {results['error']}")
            return
        
        # 基本情報
        st.write(f"**ファイル名**: {results.get('file_name', '不明')}")
        st.write(f"**チェック実行時刻**: {results.get('checked_at', '不明')}")
        
        # チェック結果詳細
        checks = results.get("checks", [])
        if not checks:
            st.info("チェック項目がありません")
            return
        
        for i, check in enumerate(checks):
            severity = check.get("severity", "info")
            message = check.get("message", "メッセージなし")
            details = check.get("details", "")
            rule_name = check.get("rule_name", "")
            
            # ルール名を表示
            if rule_name:
                st.markdown(f"**{rule_name}**")
            
            # 重要度に応じて表示を変える
            if severity == "error":
                st.error(f"❌ {message}")
            elif severity == "warning":
                st.warning(f"⚠️ {message}")
            else:
                st.success(f"✅ {message}")
            
            # 詳細があれば表示
            if details:
                with st.expander("詳細"):
                    st.markdown(details)
            
            if i < len(checks) - 1:
                st.markdown("---")
    
    @staticmethod
    def show_processing_results(results: Dict[str, Any], title: str = "処理結果"):
        """
        処理結果を表示（経理業務処理用）
        
        Args:
            results: 処理結果
            title: 表示タイトル
        """
        st.markdown(f'<div class="card-header">{title}</div>', unsafe_allow_html=True)
        
        if not results.get("success", False):
            st.error(f"処理エラー: {results.get('error', '不明なエラー')}")
            return
        
        # サマリー情報
        summary = results.get("summary", {})
        if summary:
            st.markdown("### 処理サマリー")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("総データ数", summary.get("total_data_count", 0))
            with col2:
                st.metric("処理成功", summary.get("processed_data_count", 0))
            with col3:
                st.metric("処理失敗", summary.get("failed_data_count", 0))
            
            # 詳細情報
            with st.expander("詳細情報", expanded=False):
                st.write(f"**タスク名**: {summary.get('task_name', '不明')}")
                st.write(f"**書き込み行数**: {summary.get('written_rows', 0)}")
                st.write(f"**出力シート**: {summary.get('target_sheet', '不明')}")
                st.write(f"**出力範囲**: {summary.get('excel_range', '不明')}")
                st.write(f"**使用トークン数**: {summary.get('tokens_used', 0)}")
        
        # エラー詳細
        processing_details = summary.get("processing_details", {})
        processing_errors = processing_details.get("processing_errors", [])
        if processing_errors:
            st.markdown("### 処理エラー詳細")
            for error in processing_errors:
                st.error(f"• {error}")
    
    @staticmethod
    def show_data_table(data: list, title: str = "データ一覧", 
                       column_config: Optional[Dict[str, Any]] = None):
        """
        データテーブルを表示
        
        Args:
            data: 表示するデータ
            title: テーブルタイトル
            column_config: カラム設定
        """
        if not data:
            st.info("表示するデータがありません")
            return
        
        st.markdown(f"### {title}")
        
        # データフレームに変換
        if isinstance(data, list) and len(data) > 0:
            df = pd.DataFrame(data)
            
            # 列設定があれば適用
            kwargs = {"use_container_width": True, "hide_index": True}
            if column_config:
                kwargs["column_config"] = column_config
            
            st.dataframe(df, **kwargs)
        else:
            st.info("データが空です")
    
    @staticmethod
    def create_download_button(data: bytes, filename: str, 
                             button_text: str = "ダウンロード",
                             mime_type: str = "application/octet-stream"):
        """
        ダウンロードボタンを作成
        
        Args:
            data: ダウンロードするデータ
            filename: ファイル名
            button_text: ボタンテキスト
            mime_type: MIMEタイプ
        """
        st.download_button(
            label=button_text,
            data=data,
            file_name=filename,
            mime=mime_type,
            use_container_width=True
        )
    
    @staticmethod
    def show_file_upload_help(accepted_types: list, max_size_mb: int = 200):
        """
        ファイルアップロード用のヘルプを表示
        
        Args:
            accepted_types: 受け入れ可能なファイルタイプ
            max_size_mb: 最大ファイルサイズ（MB）
        """
        st.markdown("### アップロード可能なファイル")
        
        col1, col2 = st.columns(2)
        with col1:
            st.write("**対応ファイル形式**")
            for file_type in accepted_types:
                st.write(f"• {file_type}")
        
        with col2:
            st.write("**制限事項**")
            st.write(f"• 最大ファイルサイズ: {max_size_mb}MB")
            st.write("• 同時アップロード: 複数ファイル対応")
    
    @staticmethod
    def show_configuration_summary(config: Dict[str, Any], title: str = "設定サマリー"):
        """
        設定情報のサマリーを表示
        
        Args:
            config: 設定情報
            title: タイトル
        """
        with st.expander(title, expanded=False):
            for key, value in config.items():
                if isinstance(value, dict):
                    st.write(f"**{key}**:")
                    for sub_key, sub_value in value.items():
                        st.write(f"  • {sub_key}: {sub_value}")
                else:
                    st.write(f"**{key}**: {value}")


class ProgressManager:
    """
    進捗管理クラス
    """
    
    def __init__(self, total_items: int, operation_name: str = "処理"):
        self.total_items = total_items
        self.current_item = 0
        self.operation_name = operation_name
        self.progress_bar = None
        self.status_text = None
    
    def initialize_ui(self):
        """UI要素を初期化"""
        self.progress_bar = st.progress(0)
        self.status_text = st.empty()
        self._update_display()
    
    def update(self, increment: int = 1, item_name: str = ""):
        """進捗を更新"""
        self.current_item += increment
        if self.current_item > self.total_items:
            self.current_item = self.total_items
        
        self._update_display(item_name)
    
    def set_progress(self, current: int, item_name: str = ""):
        """進捗を直接設定"""
        self.current_item = current
        if self.current_item > self.total_items:
            self.current_item = self.total_items
        
        self._update_display(item_name)
    
    def _update_display(self, item_name: str = ""):
        """表示を更新"""
        if self.total_items > 0:
            progress = self.current_item / self.total_items
            
            if self.progress_bar:
                self.progress_bar.progress(progress)
            
            if self.status_text:
                percentage = progress * 100
                status_msg = f"{self.operation_name}: {self.current_item}/{self.total_items} ({percentage:.1f}%)"
                if item_name:
                    status_msg += f" - {item_name}"
                self.status_text.text(status_msg)
    
    def complete(self, final_message: str = "完了"):
        """処理完了"""
        if self.progress_bar:
            self.progress_bar.progress(1.0)
        if self.status_text:
            self.status_text.success(f"✅ {final_message}")
    
    def error(self, error_message: str):
        """エラー表示"""
        if self.status_text:
            self.status_text.error(f"❌ エラー: {error_message}")
    
    def cleanup(self):
        """UI要素をクリーンアップ"""
        if self.progress_bar:
            self.progress_bar.empty()
        if self.status_text:
            self.status_text.empty()