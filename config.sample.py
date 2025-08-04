# 請求書カットオフチェックツール - 設定サンプル

"""
設定ファイルのサンプルです。
実際の使用時は、このファイルをconfig.pyとしてコピーし、
必要に応じて設定を変更してください。
"""

# OpenAI API設定
OPENAI_CONFIG = {
    # APIキー（環境変数または直接設定）
    "api_key": None,  # 実際の使用時は環境変数 OPENAI_API_KEY を使用推奨
    
    # 使用するモデル（固定：変更禁止）
    "model": "gpt-4.1",
    
    # レスポンスの温度（0.0-1.0）
    "temperature": 0.1,
    
    # 最大トークン数
    "max_tokens": 4000,
}

# ファイル処理設定
FILE_CONFIG = {
    # 最大ファイルサイズ（バイト）
    "max_file_size": 50 * 1024 * 1024,  # 50MB
    
    # サポートされるファイル形式
    "supported_extensions": [".pdf", ".xlsx", ".xls", ".docx", ".doc"],
    
    # 一度に処理可能な最大ファイル数
    "max_files_per_batch": 100,
}

# 並列処理設定
PROCESSING_CONFIG = {
    # デフォルトの並列処理数
    "default_workers": 3,
    
    # 最大並列処理数
    "max_workers": 10,
    
    # タイムアウト（秒）
    "timeout": 300,  # 5分
}

# アプリケーション設定
APP_CONFIG = {
    # Streamlitのページ設定
    "page_title": "請求書カットオフチェックツール",
    "page_icon": "📊",
    "layout": "wide",
    
    # デバッグモード
    "debug": False,
    
    # ログレベル
    "log_level": "INFO",
}

# ルール設定
RULE_CONFIG = {
    # ルールファイルのパス
    "rules_file": "rules.json",
    
    # デフォルトルールを自動作成するか
    "auto_create_default_rules": True,
    
    # ルールの検証を行うか
    "validate_rules": True,
}

# 結果出力設定
OUTPUT_CONFIG = {
    # 結果ファイルの保存ディレクトリ
    "output_directory": "results",
    
    # 結果ファイルの命名規則
    "filename_format": "invoice_check_results_{timestamp}.xlsx",
    
    # 結果ファイルの自動保存
    "auto_save": False,
    
    # 詳細ログの出力
    "detailed_logging": True,
}

# セキュリティ設定
SECURITY_CONFIG = {
    # ファイルの一時保存を行うか
    "allow_temp_files": True,
    
    # 処理後のファイル削除
    "auto_cleanup": True,
    
    # セッションタイムアウト（分）
    "session_timeout": 60,
}

# エラーハンドリング設定
ERROR_CONFIG = {
    # エラー時のリトライ回数
    "max_retries": 3,
    
    # リトライ間隔（秒）
    "retry_delay": 1,
    
    # エラーログの詳細度
    "error_detail_level": "HIGH",
}

# カスタムプロンプトテンプレート
PROMPT_TEMPLATES = {
    "system_prompt_base": """
あなたは経理の専門家として、請求書の内容をチェックしてください。

チェックルール: {rule_name}
カテゴリ: {category}
ルール内容: {prompt}

以下の形式で結果を返してください（JSON形式）:
{{
    "passed": true/false,
    "severity": "info/warning/error", 
    "message": "チェック結果の説明",
    "details": "詳細な説明（オプション）"
}}

重要なガイドライン:
- 軽微な問題は "warning"
- 重大な問題は "error" 
- 問題なしは "info" で passed: true
- 日本の会計基準と税法を考慮してください
- 明確で具体的な説明を提供してください
""",
    
    "user_prompt_base": """
以下の請求書をチェックしてください:

ファイル名: {file_name}
ファイルタイプ: {file_type}

請求書の内容:
{content}

特別なチェック指示:
{custom_instructions}

上記の内容について、指定されたルールに基づいてチェックしてください。
"""
}

# 会計基準設定
ACCOUNTING_STANDARDS = {
    # 期末日（デフォルト）
    "fiscal_year_end": "03-31",
    
    # 消費税率
    "tax_rates": [0.08, 0.10],
    
    # 支払期限の上限（日）
    "max_payment_terms": 30,
    
    # 必須項目リスト
    "required_fields": [
        "請求者名",
        "請求日", 
        "支払期限",
        "請求金額",
        "請求内容",
        "振込先"
    ],
}

# 国際化設定
I18N_CONFIG = {
    # デフォルト言語
    "default_language": "ja",
    
    # サポート言語
    "supported_languages": ["ja", "en"],
    
    # 日付形式
    "date_formats": {
        "ja": "%Y年%m月%d日",
        "en": "%Y-%m-%d"
    },
    
    # 数値形式
    "number_formats": {
        "ja": "{:,}円",
        "en": "¥{:,}"
    }
}