import json
import uuid
from datetime import datetime
from typing import Dict, Any, Optional, List
from pathlib import Path

class RuleManager:
    """チェックルールの管理を行うクラス"""
    
    def __init__(self, rules_file: str = "rules.json"):
        self.rules_file = Path(rules_file)
        self.rules = self._load_rules()
        self._initialize_default_rules()
    
    def _load_rules(self) -> Dict[str, Any]:
        """ルールファイルを読み込み"""
        if self.rules_file.exists():
            try:
                with open(self.rules_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, Exception):
                return {}
        return {}
    
    def _save_rules(self):
        """ルールをファイルに保存"""
        with open(self.rules_file, 'w', encoding='utf-8') as f:
            json.dump(self.rules, f, ensure_ascii=False, indent=2)
    
    def _initialize_default_rules(self):
        """デフォルトルールを初期化"""
        if not self.rules:
            default_rules = [
                {
                    "name": "請求書日付チェック",
                    "category": "日付チェック",
                    "prompt": """請求書の日付が会計期間内（期末日以前）かどうかを確認してください。
- 請求書日付を特定してください
- 期末日は通常3月31日ですが、文書に記載がある場合はそれに従ってください
- 日付が期末日より後の場合は warning としてください
- 日付が不明瞭な場合は warning としてください"""
                },
                {
                    "name": "金額整合性チェック",
                    "category": "金額チェック", 
                    "prompt": """請求書の金額計算が正しいかどうかを確認してください。
- 税抜金額と消費税額の合計が税込金額と一致するか
- 消費税率（8%または10%）が正しく適用されているか
- 計算に誤りがある場合は error としてください
- 金額が不明瞭な場合は warning としてください"""
                },
                {
                    "name": "必須項目チェック",
                    "category": "書式チェック",
                    "prompt": """請求書に以下の必須項目が記載されているかを確認してください：
- 請求者名/会社名
- 請求日
- 支払期限
- 請求金額
- 請求内容/摘要
- 振込先情報
不足している項目がある場合は warning としてください。"""
                },
                {
                    "name": "承認印チェック",
                    "category": "承認チェック",
                    "prompt": """請求書に適切な承認印や署名があるかを確認してください。
- 会社印、代表印、担当者印などの有無
- 電子印鑑の場合も含む
- 承認に関する記載の確認
承認印が不明瞭または不足している場合は warning としてください。"""
                },
                {
                    "name": "支払条件チェック", 
                    "category": "その他",
                    "prompt": """請求書の支払条件が適切かを確認してください。
- 支払期限が妥当な期間内（通常30日以内）か
- 支払方法が明記されているか
- 振込手数料の負担者が明記されているか
不適切な条件がある場合は warning としてください。"""
                }
            ]
            
            for rule_data in default_rules:
                self.add_rule(
                    rule_data["name"],
                    rule_data["category"], 
                    rule_data["prompt"]
                )
    
    def add_rule(self, name: str, category: str, prompt: str) -> str:
        """新しいルールを追加"""
        rule_id = str(uuid.uuid4())
        
        self.rules[rule_id] = {
            "name": name,
            "category": category,
            "prompt": prompt,
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat()
        }
        
        self._save_rules()
        return rule_id
    
    def get_rule(self, rule_id: str) -> Optional[Dict[str, Any]]:
        """ルールを取得"""
        return self.rules.get(rule_id)
    
    def get_all_rules(self) -> Dict[str, Any]:
        """すべてのルールを取得"""
        return self.rules.copy()
    
    def update_rule(self, rule_id: str, name: Optional[str] = None, 
                   category: Optional[str] = None, prompt: Optional[str] = None) -> bool:
        """ルールを更新"""
        if rule_id not in self.rules:
            return False
        
        if name is not None:
            self.rules[rule_id]["name"] = name
        if category is not None:
            self.rules[rule_id]["category"] = category
        if prompt is not None:
            self.rules[rule_id]["prompt"] = prompt
        
        self.rules[rule_id]["updated_at"] = datetime.now().isoformat()
        self._save_rules()
        return True
    
    def delete_rule(self, rule_id: str) -> bool:
        """ルールを削除"""
        if rule_id in self.rules:
            del self.rules[rule_id]
            self._save_rules()
            return True
        return False
    
    def get_rules_by_category(self, category: str) -> Dict[str, Any]:
        """カテゴリ別にルールを取得"""
        return {
            rule_id: rule 
            for rule_id, rule in self.rules.items() 
            if rule["category"] == category
        }
    
    def search_rules(self, keyword: str) -> Dict[str, Any]:
        """キーワードでルールを検索"""
        keyword_lower = keyword.lower()
        return {
            rule_id: rule
            for rule_id, rule in self.rules.items()
            if (keyword_lower in rule["name"].lower() or 
                keyword_lower in rule["prompt"].lower())
        }
    
    def export_rules(self, file_path: str) -> bool:
        """ルールをエクスポート"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.rules, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False
    
    def import_rules(self, file_path: str, overwrite: bool = False) -> bool:
        """ルールをインポート"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                imported_rules = json.load(f)
            
            if overwrite:
                self.rules = imported_rules
            else:
                # 既存のルールIDと重複しないようにする
                for rule_id, rule_data in imported_rules.items():
                    if rule_id not in self.rules:
                        self.rules[rule_id] = rule_data
                    else:
                        # 新しいIDを生成してインポート
                        new_id = str(uuid.uuid4())
                        self.rules[new_id] = rule_data
            
            self._save_rules()
            return True
        except Exception:
            return False
    
    def get_categories(self) -> List[str]:
        """利用可能なカテゴリのリストを取得"""
        categories = set()
        for rule in self.rules.values():
            categories.add(rule["category"])
        return sorted(list(categories))
    
    def validate_rule(self, rule_data: Dict[str, Any]) -> List[str]:
        """ルールデータの妥当性を検証"""
        errors = []
        
        required_fields = ["name", "category", "prompt"]
        for field in required_fields:
            if not rule_data.get(field):
                errors.append(f"必須フィールド '{field}' が不足しています")
        
        if rule_data.get("name") and len(rule_data["name"]) < 3:
            errors.append("ルール名は3文字以上である必要があります")
        
        if rule_data.get("prompt") and len(rule_data["prompt"]) < 10:
            errors.append("プロンプトは10文字以上である必要があります")
        
        return errors