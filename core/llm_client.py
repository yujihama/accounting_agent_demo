import openai
from typing import Dict, List, Any, Optional, Tuple, Callable
import json
from datetime import datetime
import concurrent.futures
from pydantic import ValidationError
from .models import LLMAccountingResponse, AccountingResult, AccountingSummary
import os
from dotenv import load_dotenv

class LLMClient:
    """
    LLM統合クライアント
    GPT-4.1を使用した処理を統合管理
    """
    
    def __init__(self):
        # 環境変数をロード
        load_dotenv()
        
        self.client = None
        self.api_key = None
        self.provider = os.getenv("OPENAI_PROVIDER", "openai")  # デフォルトはopenai
        
        # プロバイダーに応じた設定
        if self.provider == "azure":
            self.model = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME", "gpt-4-turbo-preview")
            self.azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
            self.azure_api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview")
        else:
            self.model = "gpt-4-turbo-preview"  # GPT-4.1相当
    
    def set_api_key(self, api_key: str = None):
        """APIキーを設定"""
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
        
        # プロバイダーに応じてクライアントを初期化
        if self.provider == "azure":
            self.client = openai.AzureOpenAI(
                api_key=self.api_key,
                azure_endpoint=self.azure_endpoint,
                api_version=self.azure_api_version
            )
        else:
            self.client = openai.OpenAI(api_key=self.api_key)
    
    def process_accounting_task(self, instruction: str, evidence_data: Dict[str, Any], 
                              output_format: Dict[str, Any], task_config: Dict[str, Any] = None,
                              max_workers: int = 3,
                              progress_callback: Optional[Callable] = None) -> Dict[str, Any]:
        """
        経理業務タスクを処理（データごとに並列処理）
        
        Args:
            instruction: ユーザーからの指示
            evidence_data: 証跡データ
            output_format: 出力フォーマット定義
            max_workers: 最大並列処理数
            
        Returns:
            処理結果
        """
        try:
            if not self.client:
                return {"success": False, "error": "APIキーが設定されていません"}
            
            if not evidence_data.get("success") or not evidence_data.get("data"):
                return {"success": False, "error": "有効な証跡データがありません"}
            
            data_items = list(evidence_data["data"].items())
            total_data_count = len(data_items)
            
            print(f"並列処理開始: {total_data_count}件のデータを{max_workers}並列で処理")
            
            all_results = []
            total_tokens = 0
            processing_errors = []
            raw_responses = {}  # 各データのRAWレスポンスを保存
            
            # 並列処理で実行
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                # 各データの処理タスクを作成
                future_to_data = {}
                
                for data_id, data_entry in data_items:
                    single_data_evidence = {
                        "success": True,
                        "data": {data_id: data_entry},
                        "metadata": evidence_data.get("metadata", {})
                    }
                    
                    future = executor.submit(
                        self._process_single_data,
                        instruction, 
                        single_data_evidence, 
                        output_format, 
                        data_id,
                        task_config
                    )
                    future_to_data[future] = data_id
                
                # 結果を完了順に収集
                completed_count = 0
                for future in concurrent.futures.as_completed(future_to_data):
                    data_id = future_to_data[future]
                    completed_count += 1
                    
                    try:
                        single_result = future.result()
                        
                        if single_result["success"]:
                            # 結果を集約
                            if "results" in single_result["data"]:
                                all_results.extend(single_result["data"]["results"])
                            total_tokens += single_result.get("tokens_used", 0)
                            
                            # RAWレスポンスを保存
                            if "raw_response" in single_result:
                                raw_responses[data_id] = single_result["raw_response"]
                        else:
                            processing_errors.append(f"データ {data_id}: {single_result.get('error', '不明なエラー')}")
                        
                        print(f"処理完了: {completed_count}/{total_data_count} ({data_id})")
                        
                        # 進捗コールバックを呼び出し
                        if progress_callback:
                            progress_callback(completed_count, total_data_count, data_id)
                        
                    except Exception as e:
                        processing_errors.append(f"データ {data_id}: 並列処理エラー - {str(e)}")
                        print(f"エラー発生: {completed_count}/{total_data_count} ({data_id}) - {str(e)}")
            
            # 集約結果を作成
            aggregated_data = {
                "results": all_results,
                "summary": {
                    "total_processed": len(all_results),
                    "total_amount": sum(r.get("amount", 0) for r in all_results if isinstance(r.get("amount"), (int, float))),
                    "match_count": len([r for r in all_results if r.get("match_status") == "一致"]),
                    "mismatch_count": len([r for r in all_results if r.get("match_status") == "不一致"]),
                    "processing_errors": processing_errors
                }
            }
            
            print(f"並列処理完了: 成功{len(all_results)}件, エラー{len(processing_errors)}件, 総トークン{total_tokens}")
            print(f"RAWレスポンス数: {len(raw_responses)}件")
            if raw_responses:
                print(f"RAWレスポンスのキー: {list(raw_responses.keys())}")
            
            return {
                "success": True,
                "data": aggregated_data,
                "raw_responses": raw_responses,  # 各データのRAWレスポンスを追加
                "processed_at": datetime.now().isoformat(),
                "tokens_used": total_tokens,
                "processing_errors": processing_errors
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": f"LLM並列処理エラー: {str(e)}"
            }
    
    def _process_single_data(self, instruction: str, single_data_evidence: Dict[str, Any], 
                           output_format: Dict[str, Any], data_id: str, task_config: Dict[str, Any] = None) -> Dict[str, Any]:
        """
        単一データを処理
        
        Args:
            instruction: ユーザーからの指示
            single_data_evidence: 単一データの証跡データ
            output_format: 出力フォーマット定義
            data_id: データID
            
        Returns:
            処理結果
        """
        try:
            # プロンプトを構築
            prompt = self._build_accounting_prompt(instruction, single_data_evidence, output_format, task_config)
            
            # レスポンススキーマを追加
            response_schema = LLMAccountingResponse.model_json_schema()
            
            # GPT-4.1で処理（構造化出力を要求）
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": f"""あなたは経理業務の専門家です。指示に従って証跡データを分析し、正確な結果を返してください。

必ず以下のJSONスキーマに厳密に従った形式で回答してください：
{json.dumps(response_schema, ensure_ascii=False, indent=2)}"""
                    },
                    {
                        "role": "user", 
                        "content": prompt
                    }
                ],
                temperature=0,
                response_format={"type": "json_object"}
            )
            
            # レスポンスを解析
            result_text = response.choices[0].message.content
            
            # Pydanticモデルで検証とパース
            try:
                # JSONパース
                json_data = json.loads(result_text)
                
                # Pydanticモデルで検証
                validated_response = LLMAccountingResponse(**json_data)
                
                # データIDが設定されていない場合は設定
                for result in validated_response.results:
                    if not result.data_id:
                        result.data_id = data_id
                
                return {
                    "success": True,
                    "data": validated_response.model_dump(),
                    "raw_response": result_text,
                    "processed_at": datetime.now().isoformat(),
                    "tokens_used": response.usage.total_tokens
                }
                
            except (json.JSONDecodeError, ValidationError) as e:
                # 型検証に失敗した場合はフォールバック
                fallback_data = self._create_fallback_response(result_text, data_id, str(e))
                return {
                    "success": False,
                    "data": fallback_data,
                    "raw_response": result_text,
                    "error": f"レスポンス検証エラー: {str(e)}",
                    "processed_at": datetime.now().isoformat(),
                    "tokens_used": response.usage.total_tokens
                }
            
        except Exception as e:
            return {
                "success": False,
                "error": f"データ {data_id} のLLM処理エラー: {str(e)}"
            }
    
    def _build_accounting_prompt(self, instruction: str, evidence_data: Dict[str, Any], 
                               output_format: Dict[str, Any], task_config: Dict[str, Any] = None) -> str:
        """経理業務用のプロンプトを構築"""
        
        # 証跡データを文字列化
        evidence_text = self._format_evidence_data(evidence_data)
        
        # 出力フォーマットを文字列化
        format_text = self._format_output_specification(output_format)
        
        # タスク設定のプロンプトテンプレートを使用
        if task_config and "processing_prompt_template" in task_config:
            # テンプレートに値を代入
            prompt = task_config["processing_prompt_template"].format(
                instruction=instruction,
                evidence_data=evidence_text,
                output_format=format_text
            )
        else:
            # デフォルトのプロンプト
            prompt = f"""
# 経理業務処理指示

## ユーザー指示
{instruction}

## 証跡データ
{evidence_text}

## 出力フォーマット要求
{format_text}

## 処理要求
1. 証跡データを詳細に分析してください
2. ユーザー指示に基づいて適切な処理を行ってください
3. 結果は必ずJSON形式で、指定されたフォーマットに従って返してください
4. 金額の照合や計算が必要な場合は、詳細な計算過程も含めてください
5. 不明な点や問題がある場合は、備考欄に記載してください

## 重要な注意事項
- 金額は数値として正確に抽出してください
- 日付形式は統一してください（YYYY-MM-DD）
- 照合結果は「一致」「不一致」「要確認」のいずれかで判定してください
- 計算結果に端数がある場合の処理方法を明記してください

結果をJSON形式で返してください：
"""
        
        return prompt
    
    def _format_evidence_data(self, evidence_data: Dict[str, Any]) -> str:
        """証跡データを読みやすい形式に整形"""
        formatted_text = []
        
        if evidence_data.get("success") and "data" in evidence_data:
            for data_id, data_entry in evidence_data["data"].items():
                formatted_text.append(f"\n### データID: {data_id}")
                formatted_text.append(f"ドキュメント数: {data_entry.get('document_count', 0)}")
                
                if "documents" in data_entry:
                    for doc_name, doc_info in data_entry["documents"].items():
                        formatted_text.append(f"\n#### ドキュメント: {doc_name}")
                        formatted_text.append(f"タイプ: {doc_info.get('type', '不明')}")
                        formatted_text.append(f"拡張子: {doc_info.get('extension', '不明')}")
                        
                        # 実際のドキュメント内容がある場合は追加
                        if "content" in doc_info and doc_info["content"]:
                            content = doc_info["content"]
                            # 長すぎる場合は適切に切り詰める（PDFは5000文字、その他は全文）
                            if doc_info.get('extension') == '.pdf' and len(content) > 5000:
                                formatted_text.append(f"内容（最初の5000文字）:\n{content[:5000]}...\n[以下省略]")
                            else:
                                formatted_text.append(f"内容:\n{content}")
                        else:
                            formatted_text.append("内容: [コンテンツなし]")
        
        return "\n".join(formatted_text) if formatted_text else "証跡データが見つかりません"
    
    def _format_output_specification(self, output_format: Dict[str, Any]) -> str:
        """出力フォーマット仕様を文字列化"""
        format_lines = []
        
        # 列定義を出力
        format_lines.append("必要な出力カラム（列定義）:")
        for col_letter in sorted(output_format["column_definitions"].keys()):
            col_def = output_format["column_definitions"][col_letter]
            format_lines.append(f"- {col_letter}列: {col_def.get('key', '')} ({col_def.get('header', '')})")
            if "description" in col_def:
                format_lines.append(f"  説明: {col_def['description']}")
        
        format_lines.append("\nJSON形式例:")
        
        # 出力キーを取得（row_numberは除外）
        output_keys = [col_def.get('key', '') for col_def in output_format["column_definitions"].values() 
                      if col_def.get('key') != 'row_number']
        
        # JSON例を生成
        example_result = {"data_id": "データ識別子"}
        for key in output_keys:
            if key and key != "data_id":
                example_result[key] = f"{key}の値"
        
        format_lines.append(json.dumps({
            "results": [example_result],
            "summary": {
                "total_processed": "処理件数",
                "total_amount": "合計金額（該当する場合）",
                "match_count": "一致件数（照合の場合）",
                "mismatch_count": "不一致件数（照合の場合）"
            }
        }, ensure_ascii=False, indent=2))
        
        return "\n".join(format_lines)
    
    def _create_fallback_response(self, response_text: str, data_id: str, error_message: str) -> Dict[str, Any]:
        """検証エラー時のフォールバックレスポンスを作成"""
        # 従来の文字列パース処理を試みる
        try:
            # JSONブロックを探す
            start_markers = ["```json", "```", "{"]
            json_start = -1
            
            for marker in start_markers:
                pos = response_text.find(marker)
                if pos != -1:
                    if marker == "{":
                        json_start = pos
                    else:
                        json_start = pos + len(marker)
                    break
            
            if json_start == -1:
                json_start = 0
            
            # JSON終了位置を探す
            remaining_text = response_text[json_start:]
            last_brace = remaining_text.rfind("}")
            json_end = json_start + last_brace + 1 if last_brace != -1 else len(response_text)
            
            # JSON部分を抽出
            json_text = response_text[json_start:json_end].strip()
            
            # JSONパース
            result = json.loads(json_text)
            
            # データIDを設定
            if "results" in result:
                for item in result["results"]:
                    if "data_id" not in item:
                        item["data_id"] = data_id
            
            return result
            
        except Exception:
            # 完全にパースに失敗した場合のデフォルトレスポンス
            return {
                "results": [{
                    "data_id": data_id,
                    "task_type": "unknown",
                    "status": "エラー",
                    "result_data": {},
                    "notes": f"レスポンス解析エラー: {error_message}"
                }],
                "summary": {
                    "total_processed": 0,
                    "require_review_count": 1
                },
                "processing_notes": f"LLMレスポンスの解析に失敗しました: {error_message}"
            }
    
    def validate_api_key(self, api_key: str = None) -> Tuple[bool, Optional[str]]:
        """APIキーの妥当性を検証"""
        try:
            # 検証用のAPIキーを取得
            test_api_key = api_key or self.api_key
            if not test_api_key:
                if self.provider == "azure":
                    test_api_key = os.getenv("AZURE_OPENAI_API_KEY")
                else:
                    test_api_key = os.getenv("OPENAI_API_KEY")
            
            if not test_api_key:
                return False, "APIキーが設定されていません"
            
            # プロバイダーに応じてクライアントを作成
            if self.provider == "azure":
                temp_client = openai.AzureOpenAI(
                    api_key=test_api_key,
                    azure_endpoint=self.azure_endpoint,
                    api_version=self.azure_api_version
                )
            else:
                temp_client = openai.OpenAI(api_key=test_api_key)
            
            # 簡単なテストリクエスト
            response = temp_client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": "Hello"}],
                max_tokens=10
            )
            
            return True, None
            
        except Exception as e:
            return False, f"APIキー検証エラー: {str(e)}"