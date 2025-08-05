"""
LLMレスポンス用のPydanticモデル定義
"""
from typing import Dict, List, Any, Optional
from pydantic import BaseModel, Field, validator
from datetime import datetime


class AccountingResult(BaseModel):
    """経理業務処理の個別結果"""
    data_id: str = Field(description="データID")
    task_type: str = Field(description="タスクタイプ")
    status: str = Field(description="処理ステータス", default="完了")
    result_data: Dict[str, Any] = Field(description="処理結果データ")
    calculations: Optional[List[str]] = Field(default=None, description="計算過程")
    notes: Optional[str] = Field(default=None, description="備考")
    
    @validator('status')
    def validate_status(cls, v):
        valid_statuses = ["完了", "エラー", "要確認", "処理中"]
        if v not in valid_statuses:
            raise ValueError(f"ステータスは {valid_statuses} のいずれかである必要があります")
        return v


class AccountingSummary(BaseModel):
    """経理業務処理のサマリー"""
    total_processed: int = Field(description="処理済みデータ数", ge=0)
    total_amount: Optional[float] = Field(default=None, description="合計金額")
    matched_count: Optional[int] = Field(default=None, description="一致件数", ge=0)
    unmatched_count: Optional[int] = Field(default=None, description="不一致件数", ge=0)
    require_review_count: Optional[int] = Field(default=None, description="要確認件数", ge=0)


class LLMAccountingResponse(BaseModel):
    """LLMからの経理業務処理レスポンス"""
    results: List[AccountingResult] = Field(description="処理結果のリスト")
    summary: AccountingSummary = Field(description="処理サマリー")
    processing_notes: Optional[str] = Field(default=None, description="処理全体に関する備考")


class RuleSuggestion(BaseModel):
    """ルール提案の個別項目"""
    name: str = Field(description="ルール名", min_length=3)
    category: str = Field(description="カテゴリ")
    prompt: str = Field(description="ルール内容", min_length=10)
    severity: Optional[str] = Field(default="warning", description="重要度")
    examples: Optional[List[str]] = Field(default=None, description="チェック例")
    
    @validator('category')
    def validate_category(cls, v):
        valid_categories = ["日付チェック", "金額チェック", "承認チェック", "書式チェック", "その他"]
        if v not in valid_categories:
            raise ValueError(f"カテゴリは {valid_categories} のいずれかである必要があります")
        return v
    
    @validator('severity')
    def validate_severity(cls, v):
        valid_severities = ["info", "warning", "error"]
        if v and v not in valid_severities:
            raise ValueError(f"重要度は {valid_severities} のいずれかである必要があります")
        return v


class RuleSuggestionsResponse(BaseModel):
    """ルール提案のレスポンス"""
    suggestions: List[RuleSuggestion] = Field(description="提案されたルールのリスト")
    analysis_summary: Optional[str] = Field(default=None, description="分析サマリー")


class EnhancedRuleResponse(BaseModel):
    """ルール改善のレスポンス"""
    enhanced_prompt: str = Field(description="改善されたルール内容", min_length=10)
    improvements: Optional[List[str]] = Field(default=None, description="改善点のリスト")