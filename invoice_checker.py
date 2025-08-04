import asyncio
from datetime import datetime
import json
from typing import Dict, List, Any, Optional
from pydantic import BaseModel, Field, validator
from enum import Enum
try:
    from langchain_openai import ChatOpenAI
    from langchain.schema import HumanMessage, SystemMessage
    from langchain.output_parsers import PydanticOutputParser
except ImportError:
    try:
        from langchain.chat_models import ChatOpenAI
        from langchain.schema import HumanMessage, SystemMessage
        from langchain.output_parsers import PydanticOutputParser
    except ImportError:
        print("LangChainライブラリが正しくインストールされていません")
import concurrent.futures

class SeverityLevel(str, Enum):
    """重要度レベルの定義"""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"

class CheckResult(BaseModel):
    """個別チェック結果のPydanticモデル"""
    passed: bool = Field(description="チェックが成功したかどうか")
    severity: SeverityLevel = Field(description="重要度 (info/warning/error)")
    message: str = Field(min_length=1, description="チェック結果の説明")
    details: str = Field(default="", description="markdown形式で詳細な判断根拠を記載")
    
    @validator('message')
    def message_must_not_be_empty(cls, v):
        """メッセージが空でないことを確認"""
        if not v or v.strip() == "":
            raise ValueError('メッセージは空にできません')
        return v.strip()
    
    @validator('severity')
    def severity_consistency_check(cls, v, values):
        """重要度とpassedの整合性をチェック"""
        passed = values.get('passed', True)
        if v == SeverityLevel.ERROR and passed:
            raise ValueError('errorの場合、passedはfalseである必要があります')
        return v

class InvoiceCheckResponse(BaseModel):
    """GPT-4.1からのレスポンス全体のPydanticモデル"""
    check_result: CheckResult = Field(description="チェック結果")

class InvoiceChecker:
    """請求書チェック機能を提供するクラス"""
    
    def __init__(self):
        self.llm = None
        self.api_key = None
        # Structured Output用のパーサーを初期化
        self.output_parser = PydanticOutputParser(pydantic_object=InvoiceCheckResponse)
        
    def set_api_key(self, api_key: str):
        """OpenAI APIキーを設定"""
        self.api_key = api_key
        # 重要: GPT-4.1を使用（絶対に変更しないこと）
        self.llm = ChatOpenAI(
            model_name="gpt-4.1",
            openai_api_key=api_key,
        )
    
    def check_invoice(self, file_data: Dict[str, Any], rule_ids: List[str]) -> Dict[str, Any]:
        """
        単一の請求書をチェック
        
        Args:
            file_data: ファイルの内容とメタデータ
            rule_ids: 適用するルールのIDリスト
            
        Returns:
            チェック結果
        """
        if not self.llm:
            return {"error": "OpenAI APIキーが設定されていません"}
        
        try:
            # ルール管理からルールを取得
            from rule_manager import RuleManager
            rule_manager = RuleManager()
            
            checks = []
            
            for rule_id in rule_ids:
                rule = rule_manager.get_rule(rule_id)
                if rule:
                    check_result = self._apply_rule(file_data, rule)
                    checks.append(check_result)
            
            return {
                "file_name": file_data.get("file_name", "不明"),
                "checks": checks,
                "checked_at": datetime.now().isoformat()
            }
            
        except Exception as e:
            return {"error": f"チェック処理中にエラーが発生しました: {str(e)}"}
    
    def _apply_rule(self, file_data: Dict[str, Any], rule: Dict[str, Any]) -> Dict[str, Any]:
        """
        個別ルールを適用
        
        Args:
            file_data: ファイルデータ
            rule: 適用するルール
            
        Returns:
            ルール適用結果
        """
        try:
            # システムプロンプトの構築
            system_prompt = self._build_system_prompt(rule)
            
            # ユーザープロンプトの構築
            user_prompt = self._build_user_prompt(file_data, rule)
            
            # GPT-4.1に問い合わせ
            messages = [
                SystemMessage(content=system_prompt),
                HumanMessage(content=user_prompt)
            ]
            
            response = self.llm(messages)
            
            # Structured Outputを使用してレスポンスを解析
            result = self._parse_structured_response(response.content, rule)
            
            return result
            
        except Exception as e:
            return {
                "rule_name": rule.get("name", "不明"),
                "severity": SeverityLevel.ERROR.value,
                "message": f"ルール適用エラー: {str(e)}",
                "details": None
            }
    
    def _build_system_prompt(self, rule: Dict[str, Any]) -> str:
        """システムプロンプトを構築（Structured Output対応）"""
        format_instructions = self.output_parser.get_format_instructions()
        
        return f"""
あなたは経理部門として、提出された請求書の内容をチェックしてください。

チェックルール: {rule['name']}
カテゴリ: {rule['category']}

重要なガイドライン:
- 軽微な問題は "warning"
- 重大な問題は "error" 
- 問題なしは "info" で passed: true
- 請求書の内容を理解してチェックしてください
- 請求書は原本から事前に抽出されたテキストです。そのため不自然なスペースや改行が含まれていたり、文字が欠落したりする場合があります。俯瞰的に見て問題なければ、表記や体裁についての指摘はしないでください。

{format_instructions}
"""
    
    def _build_user_prompt(self, file_data: Dict[str, Any], rule: Dict[str, Any]) -> str:
        """ユーザープロンプトを構築"""
        content = file_data.get("content", "")
        metadata = file_data.get("metadata", {})
        
        prompt = f"""
以下の請求書を指示された内容でチェックしてください:

ファイル名: {file_data.get('file_name', '不明')}
ファイルタイプ: {metadata.get('file_type', '不明')}

請求書から抽出された内容:
{content}

チェック内容:
{rule['prompt']}

上記の内容について、指定されたルールに基づいてチェックしてください。
"""
        return prompt
    
    def _parse_structured_response(self, response_content: str, rule: Dict[str, Any]) -> Dict[str, Any]:
        """Structured Outputを使用してGPTレスポンスを型安全に解析"""
        try:
            # LangChainのPydanticOutputParserを使用して構造化された出力を解析
            parsed_response: InvoiceCheckResponse = self.output_parser.parse(response_content)
            
            # Pydanticモデルから辞書形式に変換
            check_result = parsed_response.check_result
            
            return {
                "rule_name": rule.get("name", "不明"),
                "severity": check_result.severity.value,  # Enumの値を取得
                "message": check_result.message,
                "details": check_result.details,
                "passed": check_result.passed
            }
            
        except Exception as e:
            # 構造化出力の解析に失敗した場合のフォールバック
            return {
                "rule_name": rule.get("name", "不明"),
                "severity": SeverityLevel.ERROR.value,
                "message": f"構造化出力解析エラー: {str(e)}",
                "details": response_content,
                "passed": False
            }
    
    async def check_invoices_async(self, files_data: List[Dict[str, Any]], rule_ids: List[str]) -> List[Dict[str, Any]]:
        """
        複数の請求書を非同期でチェック
        
        Args:
            files_data: ファイルデータのリスト
            rule_ids: 適用するルールのIDリスト
            
        Returns:
            チェック結果のリスト
        """
        tasks = []
        for file_data in files_data:
            task = asyncio.create_task(self._check_invoice_async(file_data, rule_ids))
            tasks.append(task)
        
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
        # 例外を処理
        processed_results = []
        for i, result in enumerate(results):
            if isinstance(result, Exception):
                processed_results.append({
                    "file_name": files_data[i].get("file_name", f"ファイル_{i}"),
                    "error": str(result)
                })
            else:
                processed_results.append(result)
        
        return processed_results
    
    async def _check_invoice_async(self, file_data: Dict[str, Any], rule_ids: List[str]) -> Dict[str, Any]:
        """非同期版の単一請求書チェック"""
        loop = asyncio.get_event_loop()
        with concurrent.futures.ThreadPoolExecutor() as executor:
            result = await loop.run_in_executor(
                executor, 
                self.check_invoice, 
                file_data, 
                rule_ids
            )
        return result