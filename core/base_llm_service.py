import os
from abc import ABC, abstractmethod
from typing import Dict, Any, Optional
from dotenv import load_dotenv
try:
    from langchain_openai import ChatOpenAI, AzureChatOpenAI
except ImportError:
    try:
        from langchain.chat_models import ChatOpenAI
        AzureChatOpenAI = None
    except ImportError:
        ChatOpenAI = None
        AzureChatOpenAI = None

try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False


class BaseLLMService(ABC):
    """
    LLMサービスの基盤クラス
    OpenAI/Azure OpenAIの設定と初期化を統一化
    """
    
    def __init__(self):
        # 環境変数をロード
        load_dotenv()
        
        self.llm = None
        self.client = None  # OpenAI直接クライアント用
        self.api_key = None
        self.provider = os.getenv("OPENAI_PROVIDER", "openai")
        
        # プロバイダーに応じた設定を準備
        if self.provider == "azure":
            self.azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
            self.azure_deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME", "gpt-4.1")
            self.azure_api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview")
            self.model = self.azure_deployment
        else:
            self.model = "gpt-4.1"  # GPT-4.1相当
    
    def set_api_key(self, api_key: str = None) -> bool:
        """
        APIキーを設定
        
        Args:
            api_key: APIキー（Noneの場合は環境変数から取得）
            
        Returns:
            設定成功可否
        """
        try:
            # 引数で渡されたAPIキーか、環境変数から取得
            if api_key:
                self.api_key = api_key
            else:
                if self.provider == "azure":
                    self.api_key = os.getenv("AZURE_OPENAI_API_KEY")
                else:
                    self.api_key = os.getenv("OPENAI_API_KEY")
            
            if not self.api_key:
                raise ValueError("APIキーが設定されていません。環境変数または引数で指定してください。")
            
            # LangChain用LLMを初期化
            self._initialize_langchain_llm()
            
            # OpenAI直接クライアントを初期化（必要に応じて）
            if OPENAI_AVAILABLE:
                self._initialize_openai_client()
            
            return True
            
        except Exception as e:
            print(f"APIキー設定エラー: {str(e)}")
            return False
    
    def _initialize_langchain_llm(self):
        """LangChain用LLMを初期化"""
        if not ChatOpenAI:
            raise ImportError("LangChainライブラリが正しくインストールされていません")
            
        if self.provider == "azure" and AzureChatOpenAI:
            # Azure OpenAIを使用
            self.llm = AzureChatOpenAI(
                azure_endpoint=self.azure_endpoint,
                openai_api_key=self.api_key,
                azure_deployment=self.azure_deployment,
                openai_api_version=self.azure_api_version,
                model_name=self.azure_deployment
            )
        else:
            # 通常のOpenAIを使用
            self.llm = ChatOpenAI(
                model_name=self.model,
                openai_api_key=self.api_key,
            )
    
    def _initialize_openai_client(self):
        """OpenAI直接クライアントを初期化"""
        if self.provider == "azure":
            self.client = openai.AzureOpenAI(
                api_key=self.api_key,
                azure_endpoint=self.azure_endpoint,
                api_version=self.azure_api_version
            )
        else:
            self.client = openai.OpenAI(api_key=self.api_key)
    
    def is_configured(self) -> bool:
        """APIキーが設定されているかを確認"""
        return self.api_key is not None and (self.llm is not None or self.client is not None)
    
    def get_provider_info(self) -> Dict[str, Any]:
        """プロバイダー情報を取得"""
        return {
            "provider": self.provider,
            "model": self.model,
            "configured": self.is_configured()
        }
    
    @abstractmethod
    def validate_service_specific_config(self) -> Dict[str, Any]:
        """
        サービス固有の設定を検証
        
        Returns:
            検証結果
        """
        pass
    
    def validate_configuration(self) -> Dict[str, Any]:
        """
        設定の妥当性を総合的にチェック
        
        Returns:
            検証結果
        """
        result = {
            "valid": True,
            "errors": [],
            "warnings": []
        }
        
        # 基本設定のチェック
        if not self.is_configured():
            result["valid"] = False
            result["errors"].append("APIキーが設定されていません")
        
        # プロバイダー固有のチェック
        if self.provider == "azure":
            if not self.azure_endpoint:
                result["valid"] = False
                result["errors"].append("Azure エンドポイントが設定されていません")
            
            if not self.azure_deployment:
                result["warnings"].append("Azure デプロイメント名がデフォルト値を使用しています")
        
        # サービス固有の検証
        try:
            service_validation = self.validate_service_specific_config()
            if not service_validation.get("valid", True):
                result["valid"] = False
                result["errors"].extend(service_validation.get("errors", []))
            result["warnings"].extend(service_validation.get("warnings", []))
        except Exception as e:
            result["valid"] = False
            result["errors"].append(f"サービス固有の設定検証エラー: {str(e)}")
        
        return result


class BaseDataValidator:
    """
    データ検証の共通ロジック
    """
    
    @staticmethod
    def validate_file_data(file_data: Dict[str, Any]) -> Dict[str, Any]:
        """ファイルデータの妥当性を検証"""
        result = {"valid": True, "errors": []}
        
        if not file_data:
            result["valid"] = False
            result["errors"].append("ファイルデータが空です")
            return result
        
        if "content" not in file_data:
            result["valid"] = False
            result["errors"].append("ファイルの内容が取得できません")
        
        if not file_data.get("content", "").strip():
            result["valid"] = False
            result["errors"].append("ファイルの内容が空です")
        
        return result
    
    @staticmethod
    def validate_rule_ids(rule_ids: list) -> Dict[str, Any]:
        """ルールIDリストの妥当性を検証"""
        result = {"valid": True, "errors": []}
        
        if not rule_ids:
            result["valid"] = False
            result["errors"].append("チェックルールが選択されていません")
        
        if not isinstance(rule_ids, list):
            result["valid"] = False
            result["errors"].append("ルールIDは配列である必要があります")
        
        return result
    
    @staticmethod
    def validate_evidence_data(evidence_data: Dict[str, Any]) -> Dict[str, Any]:
        """証跡データの妥当性を検証"""
        result = {"valid": True, "errors": []}
        
        if not evidence_data:
            result["valid"] = False
            result["errors"].append("証跡データが空です")
            return result
        
        if not evidence_data.get("success"):
            result["valid"] = False
            result["errors"].append("証跡データが正常に処理されていません")
        
        if not evidence_data.get("data"):
            result["valid"] = False
            result["errors"].append("証跡データが空です")
        
        return result
    
    @staticmethod
    def validate_output_config(output_config: Dict[str, Any]) -> Dict[str, Any]:
        """出力設定の妥当性を検証"""
        result = {"valid": True, "errors": []}
        
        required_keys = ["target_sheet", "start_row", "column_definitions"]
        for key in required_keys:
            if key not in output_config:
                result["valid"] = False
                result["errors"].append(f"出力設定に{key}が指定されていません")
        
        return result


class BaseProcessor(ABC):
    """
    処理基盤の共通クラス
    エラーハンドリング、ログ、進捗管理を統一化
    """
    
    def __init__(self, service_name: str):
        self.service_name = service_name
        self.validator = BaseDataValidator()
    
    def log_info(self, message: str):
        """情報ログを出力"""
        print(f"[{self.service_name}] INFO: {message}")
    
    def log_warning(self, message: str):
        """警告ログを出力"""
        print(f"[{self.service_name}] WARNING: {message}")
    
    def log_error(self, message: str):
        """エラーログを出力"""
        print(f"[{self.service_name}] ERROR: {message}")
    
    def handle_exception(self, operation: str, exception: Exception) -> Dict[str, Any]:
        """例外を統一的に処理"""
        error_message = f"{operation}中にエラーが発生しました: {str(exception)}"
        self.log_error(error_message)
        
        return {
            "success": False,
            "error": error_message,
            "operation": operation,
            "exception_type": type(exception).__name__
        }
    
    def create_progress_callback(self, operation_name: str, total_items: int):
        """進捗コールバック関数を作成"""
        def progress_callback(completed: int, total: int, current_item: str = ""):
            percentage = (completed / total) * 100 if total > 0 else 0
            item_info = f" ({current_item})" if current_item else ""
            self.log_info(f"{operation_name}: {completed}/{total} ({percentage:.1f}%){item_info}")
        
        return progress_callback
    
    @abstractmethod
    def process(self, *args, **kwargs) -> Dict[str, Any]:
        """
        実際の処理を実装
        
        Returns:
            処理結果
        """
        pass