import json
from typing import Dict, List, Any, Optional, Callable
from datetime import datetime
from .llm_client import LLMClient
from .excel_manager import ExcelManager
from .folder_processor import FolderProcessor
from .base_llm_service import BaseLLMService, BaseDataValidator, BaseProcessor

class TaskEngine(BaseProcessor):
    """
    指示実行エンジン（共通基盤を使用）
    ユーザーの指示に基づいて証跡データを処理し、調書に結果を記載
    """
    
    def __init__(self):
        # 共通基盤を初期化
        BaseProcessor.__init__(self, "TaskEngine")
        
        self.llm_client = LLMClient()
        self.excel_manager = ExcelManager()
        self.folder_processor = FolderProcessor()
        self.task_configs = {}
        self._load_task_configs()
    
    def _load_task_configs(self):
        """タスク設定を読み込み"""
        try:
            with open("config/task_configs.json", "r", encoding="utf-8") as f:
                self.task_configs = json.load(f)
                self.log_info(f"タスク設定を読み込みました: {len(self.task_configs)}件")
        except Exception as e:
            self.log_error(f"タスク設定読み込みエラー: {str(e)}")
            self.task_configs = {}
    
    def set_api_key(self, api_key: str):
        """APIキーを設定"""
        self.llm_client.set_api_key(api_key)
        self.log_info("APIキーを設定しました")
    
    def execute_accounting_task(self, task_config_id: str, evidence_data: Dict[str, Any], 
                              excel_manager, instruction: str, 
                              custom_output_config: Optional[Dict[str, Any]] = None,
                              max_workers: int = 3,
                              progress_callback: Optional[Callable] = None) -> Dict[str, Any]:
        """
        経理業務タスクを実行
        
        Args:
            task_config_id: タスク設定ID
            evidence_data: 証跡データ
            excel_manager: ExcelManagerインスタンス
            instruction: ユーザー指示
            custom_output_config: カスタム出力設定
            
        Returns:
            実行結果
        """
        try:
            self.log_info(f"経理業務タスク実行開始: {task_config_id}")
            
            # タスク設定を取得
            if task_config_id not in self.task_configs:
                error_msg = f"タスク設定が見つかりません: {task_config_id}"
                self.log_error(error_msg)
                return {"success": False, "error": error_msg}
            
            task_config = self.task_configs[task_config_id]
            output_config = custom_output_config or task_config.get("output_config", {})
            
            # 前提条件チェック（共通バリデーターを使用）
            evidence_validation = self.validator.validate_evidence_data(evidence_data)
            if not evidence_validation["valid"]:
                self.log_error(f"証跡データ検証エラー: {evidence_validation['errors']}")
                return {"success": False, "error": f"証跡データエラー: {', '.join(evidence_validation['errors'])}"}
            
            output_validation = self.validator.validate_output_config(output_config)
            if not output_validation["valid"]:
                self.log_error(f"出力設定検証エラー: {output_validation['errors']}")
                return {"success": False, "error": f"出力設定エラー: {', '.join(output_validation['errors'])}"}
            
            # ExcelManagerチェック
            if not excel_manager or not hasattr(excel_manager, 'workbook') or not excel_manager.workbook:
                error_msg = "Excelワークブックが読み込まれていません"
                self.log_error(error_msg)
                return {"success": False, "error": error_msg}
            
            # Excel Manager を一時的に設定
            original_excel_manager = self.excel_manager
            self.excel_manager = excel_manager
            
            # データ件数を取得
            data_count = len(evidence_data.get("data", {}))
            self.log_info(f"処理開始: {data_count}件のデータを{max_workers}並列で処理します")
            
            # 進捗コールバックを作成（共通基盤を使用）
            if not progress_callback:
                progress_callback = self.create_progress_callback("経理業務処理", data_count)
            
            # LLMで処理実行（データごとに並列処理）
            # タスク設定のプロンプトテンプレートを活用
            task_prompt_template = task_config.get("processing_prompt_template", "")
            
            llm_result = self.llm_client.process_accounting_task(
                instruction=instruction,
                evidence_data=evidence_data,
                output_format=output_config,
                task_config=task_config,  # タスク設定全体を渡す
                max_workers=max_workers,
                progress_callback=progress_callback
            )
            
            if not llm_result["success"]:
                self.log_error(f"LLM処理エラー: {llm_result.get('error', '不明')}")
                return llm_result
            
            # 処理エラーがある場合は警告表示
            if llm_result.get("processing_errors"):
                error_count = len(llm_result['processing_errors'])
                self.log_warning(f"{error_count}件のデータで処理エラーが発生しました")
                for error in llm_result["processing_errors"]:
                    self.log_warning(f"  - {error}")
            
            # 結果をExcelに書き込み
            write_result = self._write_results_to_excel(llm_result["data"], output_config)
            
            if not write_result["success"]:
                return write_result
            
            # 実行サマリーを作成
            summary = self._create_execution_summary(llm_result, write_result, task_config, data_count)
            
            result = {
                "success": True,
                "summary": summary,
                "llm_result": llm_result,
                "write_result": write_result,
                "executed_at": datetime.now().isoformat()
            }
            
            # Excel Managerを元に戻す
            self.excel_manager = original_excel_manager
            
            self.log_info(f"経理業務タスク実行完了: {summary.get('processed_data_count', 0)}件処理")
            return result
            
        except Exception as e:
            # エラー時もExcel Managerを元に戻す
            if 'original_excel_manager' in locals():
                self.excel_manager = original_excel_manager
            return self.handle_exception("経理業務タスク実行", e)
    
    def process(self, *args, **kwargs) -> Dict[str, Any]:
        """
        BaseProcessorのabstractメソッド実装
        execute_accounting_taskのラッパー
        """
        if len(args) >= 4:
            return self.execute_accounting_task(args[0], args[1], args[2], args[3], **kwargs)
        else:
            return {"success": False, "error": "引数が不足しています"}
    

    
    def _write_results_to_excel(self, llm_data: Dict[str, Any], output_config: Dict[str, Any]) -> Dict[str, Any]:
        """LLM処理結果をExcelに書き込み"""
        try:
            # LLMデータから結果リストを取得
            results = llm_data.get("results", [])
            
            if not results:
                return {"success": False, "error": "書き込み可能な結果データがありません"}
            
            # Excelに書き込み
            write_result = self.excel_manager.write_results(
                target_sheet=output_config["target_sheet"],
                start_row=output_config["start_row"],
                column_definitions=output_config["column_definitions"],
                data=results
            )
            
            return write_result
            
        except Exception as e:
            return {"success": False, "error": f"Excel書き込みエラー: {str(e)}"}
    
    def _create_execution_summary(self, llm_result: Dict[str, Any], write_result: Dict[str, Any], 
                                task_config: Dict[str, Any], total_data_count: int = 0) -> Dict[str, Any]:
        """実行サマリーを作成"""
        
        llm_data = llm_result.get("data", {})
        summary_data = llm_data.get("summary", {})
        processing_errors = llm_result.get("processing_errors", [])
        
        return {
            "task_name": task_config.get("name", "不明なタスク"),
            "total_data_count": total_data_count,
            "processed_data_count": summary_data.get("total_processed", 0),
            "failed_data_count": len(processing_errors),
            "written_rows": write_result.get("written_rows", 0),
            "excel_range": write_result.get("range", "不明"),
            "target_sheet": write_result.get("target_sheet", "不明"),
            "tokens_used": llm_result.get("tokens_used", 0),
            "processing_details": {
                "total_amount": summary_data.get("total_amount", 0),
                "match_count": summary_data.get("match_count", 0),
                "mismatch_count": summary_data.get("mismatch_count", 0),
                "processing_errors": processing_errors
            }
        }
    
    def get_available_tasks(self) -> List[Dict[str, str]]:
        """利用可能なタスク一覧を取得"""
        tasks = []
        for task_id, config in self.task_configs.items():
            tasks.append({
                "id": task_id,
                "name": config.get("name", "無題"),
                "description": config.get("description", "説明なし")
            })
        return tasks
    
    def get_task_config(self, task_id: str) -> Optional[Dict[str, Any]]:
        """指定タスクの設定を取得"""
        return self.task_configs.get(task_id)
    
    def create_custom_task_config(self, name: str, description: str, 
                                output_config: Dict[str, Any]) -> str:
        """カスタムタスク設定を作成"""
        task_id = f"custom_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        self.task_configs[task_id] = {
            "name": name,
            "description": description,
            "input_structure": {
                "evidence_folders": {"required": True, "structure": "2-tier"},
                "excel_workbook": {"required": True},
                "instruction": {"required": True}
            },
            "output_config": output_config,
            "created_at": datetime.now().isoformat(),
            "is_custom": True
        }
        
        # 設定ファイルに保存
        self._save_task_configs()
        
        return task_id
    
    def _save_task_configs(self):
        """タスク設定をファイルに保存"""
        try:
            with open("config/task_configs.json", "w", encoding="utf-8") as f:
                json.dump(self.task_configs, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"タスク設定保存エラー: {str(e)}")
    
    # def preview_processing_result(self, task_config_id: str, evidence_data: Dict[str, Any], 
    #                             instruction: str, sample_count: int = 3) -> Dict[str, Any]:
    #     """処理結果のプレビューを生成（少数サンプル）"""
    #     try:
    #         if task_config_id not in self.task_configs:
    #             return {"success": False, "error": f"タスク設定が見つかりません: {task_config_id}"}
            
    #         task_config = self.task_configs[task_config_id]
    #         output_config = task_config.get("output_config", {})
            
    #         # サンプルデータを作成（最初のN件のみ）
    #         sample_evidence = self._create_sample_evidence(evidence_data, sample_count)
            
    #         # LLMで処理（プレビュー用）
    #         llm_result = self.llm_client.process_accounting_task(
    #             instruction=f"{instruction}\n\n注意：これはプレビューです。最初の{sample_count}件のデータのみを処理してください。",
    #             evidence_data=sample_evidence,
    #             output_format=output_config
    #         )
            
    #         return llm_result
            
    #     except Exception as e:
    #         return {"success": False, "error": f"プレビュー生成エラー: {str(e)}"}
    
    # def _create_sample_evidence(self, evidence_data: Dict[str, Any], sample_count: int) -> Dict[str, Any]:
    #     """プレビュー用のサンプル証跡データを作成"""
    #     if not evidence_data.get("success") or not evidence_data.get("data"):
    #         return evidence_data
        
    #     original_data = evidence_data["data"]
    #     sample_data = {}
        
    #     # 最初のN件を取得
    #     for i, (data_id, data_entry) in enumerate(original_data.items()):
    #         if i >= sample_count:
    #             break
    #         sample_data[data_id] = data_entry
        
    #     return {
    #         "success": True,
    #         "data": sample_data,
    #         "metadata": {
    #             **evidence_data.get("metadata", {}),
    #             "is_sample": True,
    #             "sample_count": len(sample_data),
    #             "original_count": len(original_data)
    #         }
    #     }