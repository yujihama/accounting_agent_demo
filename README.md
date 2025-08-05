# 経理業務統合ツール (請求書チェック & 業務処理自動化)

経理業務における請求書のカットオフチェックと多様な経理処理を自動化する統合AIツールです。GPT-4.1を活用して、大量の請求書を効率的にチェックし、会計基準や経費精算ルールに基づいた検証を行うとともに、証跡処理や調書作成などの経理業務全般をサポートします。

## 機能

### 既存機能（請求書チェック）

- **自動請求書チェック**: GPT-4.1を使用した高精度な請求書内容の検証（Structured Output対応）
- **カスタムルール設定**: 会計基準や経費精算ルールを自由に設定可能
- **AIルール提案機能**: マニュアルや規定書をアップロードして自動的にチェックルールを提案
  - PDF、Word、Excel、テキストファイルから内容を解析
  - 既存ルールとの重複を回避した新しいルール提案
  - 提案されたルールの編集・カスタマイズ機能
- **並列処理**: 設定可能な並列処理数（1-12）で大量のファイルを効率的に処理
- **多様なファイル形式対応**: PDF、Excel (.xlsx, .xls)、Word (.docx) に対応
- **詳細な結果レポート**: Excel形式での結果出力とダウンロード
- **リアルタイム処理状況表示**: プログレスバーと処理完了数の表示
- **高度なフィルタリング**: 要確認・正常・全件での結果表示切り替え

### 🆕 拡張機能（経理業務処理）

- **証跡フォルダ処理**: 2階層構造のZIPフォルダをアップロード・処理
  - 自動的なドキュメント分類
  - メタデータ抽出
  - バリデーション
- **調書（Excel）操作**: 処理結果をExcelの指定箇所に記載
  - 動的シート選択
  - セル範囲指定
  - スタイル設定
  - データ形式変換
- **タスクエンジン**: 業務処理の統合管理
  - 設定駆動処理
  - プレビュー実行
  - バッチ処理
  - 結果検証

## 技術スタック

- **Streamlit**: Webベースユーザーインターフェース
- **LangChain**: GPT-4.1との連携とPydantic Output Parser
- **OpenAI GPT-4.1**: 請求書内容の分析と検証（モデル固定）[[memory:5049262]]
- **Pydantic v2**: 型安全なデータ検証とStructured Output
- **Python 3.8+**: バックエンド処理とファイル操作
- **pandas**: Excelファイル処理とデータ操作
- **PyPDF2 / PyMuPDF**: PDF文書テキスト抽出
- **python-docx**: Word文書処理
- **concurrent.futures**: 並列処理とスレッドプール管理

## セットアップ

### 方法1: PowerShellスクリプトを使用（推奨）

```powershell
# 仮想環境の作成、依存関係のインストール、アプリ起動を自動実行
.\run.ps1
```

### 方法2: 手動セットアップ

#### 1. 仮想環境の作成と有効化

```powershell
# 仮想環境の作成
python -m venv venv

# 仮想環境の有効化
.\venv\Scripts\Activate.ps1
```

#### 2. 必要なライブラリのインストール

```powershell
pip install -r requirements.txt
```

#### 3. APIキーの設定

OpenAIまたはAzure OpenAIのAPIキーを設定する方法は2つあります：

##### 方法1: 環境変数（.envファイル）を使用（推奨）

プロジェクトのルートディレクトリに`.env`ファイルを作成し、以下の内容を設定します：

```bash
# OpenAI/Azure OpenAI設定
# プロバイダー選択 (openai または azure)
OPENAI_PROVIDER=openai

# OpenAI API設定
OPENAI_API_KEY=your_openai_api_key_here

# Azure OpenAI設定 (OPENAI_PROVIDER=azure の場合に使用)
AZURE_OPENAI_API_KEY=your_azure_api_key_here
AZURE_OPENAI_ENDPOINT=https://your-resource-name.openai.azure.com/
AZURE_OPENAI_DEPLOYMENT_NAME=your-deployment-name
AZURE_OPENAI_API_VERSION=2024-02-15-preview
```

##### 方法2: アプリケーション内で設定

環境変数を設定していない場合は、アプリケーション起動後、サイドバーでAPIキーを入力できます。

#### 4. アプリケーションの起動

```powershell
# 既存の請求書チェック専用版
streamlit run app.py

# 拡張版（統合ツール）
streamlit run app_extended.py
```

### 依存関係

実装されている主要ライブラリ：
- streamlit >= 1.28.0
- langchain >= 0.0.340 
- langchain-openai >= 0.0.2
- openai >= 1.3.0
- pandas >= 2.0.0
- python-docx >= 0.8.11
- PyPDF2 >= 3.0.0
- openpyxl >= 3.1.0
- xlsxwriter >= 3.1.0
- pydantic >= 2.0.0
- python-dotenv >= 1.0.0

## 使用方法

### 請求書チェック機能

#### 1. APIキー設定

アプリケーション起動後の設定：

- **環境変数を設定済みの場合**: 自動的にAPIキーが読み込まれます
- **環境変数を設定していない場合**: サイドバーでAPIキーを入力してください
- **プロバイダーの切り替え**: `.env`ファイルの`OPENAI_PROVIDER`を`openai`または`azure`に設定

#### 2. ルール設定

「ルール設定」タブで請求書チェックのルールを設定します：

##### 手動ルール追加
- **ルール名**: チェックルールの識別名
- **カテゴリ**: 日付チェック、金額チェック、承認チェック、書式チェック、その他
- **チェックルールの内容**: GPT-4.1に送信する具体的な指示

##### AIルール提案機能
マニュアルや規定書から自動的にルールを提案する新機能：

1. **ドキュメントアップロード**: 会計マニュアル、請求書チェック規定、内部統制文書などをアップロード
2. **AI分析**: GPT-4.1がドキュメント内容を分析し、請求書チェックに関連するルールを抽出
3. **重複回避**: 既存ルールとの重複を自動的に回避
4. **ルール提案**: 信頼度スコア付きでルールを提案
5. **編集・追加**: 提案されたルールを確認・編集してから追加可能

**対応ファイル形式**: PDF、Word (.docx)、Excel (.xlsx)、テキスト (.txt, .md)

#### 3. ファイルアップロード

「ファイルアップロード」タブで請求書ファイルをアップロードします：

- **対応形式**: PDF、Excel (.xlsx, .xls)、Word (.docx)
- **複数ファイル**: 一括アップロード可能
- **ファイルサイズ**: 最大50MB/ファイル
- **自動処理**: アップロード後、「ファイルを処理」ボタンでテキスト抽出

#### 4. チェック実行

「チェック実行」タブでチェックを実行します：

##### 前提条件チェック
- ファイル処理: アップロードしたファイルが処理済みか
- チェックルール: 1つ以上のルールが設定されているか  
- API設定: OpenAI APIキーが設定されているか

##### 実行設定
- **適用するルール**: 実行したいチェックルールを複数選択可能
- **処理対象ファイル**: チェックするファイルを選択
- **並列処理数**: サイドバーで設定（1-12の範囲）
- **推定処理時間**: 総チェック数と並列数から自動計算

#### 5. 結果確認

「結果表示」タブで結果を確認できます：

##### サマリー表示
- **総ファイル数**: 処理されたファイルの総数
- **要確認**: エラーまたは警告があるファイル数
- **正常**: すべてのチェックが正常なファイル数

##### 詳細結果
- **フィルター機能**: すべて・要確認・正常での表示切り替え
- **ファイル別結果**: 各ファイルのチェック結果を詳細表示
- **重要度表示**: info（正常）・warning（警告）・error（エラー）
- **Excel出力**: 結果をタイムスタンプ付きExcelファイルでダウンロード

### 🆕 経理業務処理機能

#### 利用フロー

1. **ツール選択**
   - 請求書チェック（既存機能）
   - 経理業務処理（新機能）

2. **経理業務処理の場合**
   ```
   タスク設定 → 証跡アップロード → 調書アップロード → 
   処理実行 → 結果確認・ダウンロード
   ```

3. **各ステップ詳細**
   - **タスク設定**: 既存テンプレートまたはカスタム作成
   - **証跡アップロード**: ZIP形式の2階層フォルダ
   - **調書アップロード**: 結果記載用Excelファイル
   - **処理実行**: AI指示＋プレビュー実行＋本実行
   - **結果確認**: サマリー表示＋調書ダウンロード

## ファイル構成

```
extract_invoice/
├── app.py                    # 既存アプリ（請求書チェック専用）
├── app_extended.py           # 拡張版アプリ（統合）
├── core/                     # 共通コア機能
│   ├── folder_processor.py   # フォルダ処理
│   ├── excel_manager.py      # Excel操作
│   ├── llm_client.py         # LLM統合
│   └── task_engine.py        # タスク実行エンジン
├── config/
│   ├── rules.json            # ルール設定ファイル（自動生成）
│   └── task_configs.json     # タスク設定
├── requirements.txt          # 必要なライブラリ一覧
├── run.ps1                   # PowerShell起動スクリプト
├── venv/                     # Python仮想環境
└── README.md                 # このファイル
```

## 主要クラス

### 既存クラス

#### InvoiceChecker
- **GPT-4.1**: 請求書内容の分析とチェック（モデル固定）
- **LangChain**: PydanticOutputParserによる型安全なレスポンス処理
- **Structured Output**: CheckResultモデルによる構造化された出力
- **並列処理**: concurrent.futuresによる効率的なマルチファイル処理
- **エラーハンドリング**: 構造化出力解析失敗時のフォールバック機能

#### RuleManager
- **CRUD操作**: ルールの作成・読み取り・更新・削除
- **UUID管理**: 一意なルールID生成と管理
- **JSON保存**: rules.jsonファイルへの永続化
- **デフォルトルール**: 5つの基本チェックルールを初期提供
- **検索機能**: カテゴリ別・キーワード検索対応

#### FileProcessor
- **多形式対応**: PDF（PyPDF2）、Excel（pandas/openpyxl）、Word（python-docx）
- **テキスト抽出**: 各ファイル形式からの構造化テキスト抽出
- **メタデータ取得**: ファイル情報、ページ数、シート情報等
- **バリデーション**: ファイルサイズ・形式・名前の妥当性チェック
- **エラーハンドリング**: 各ファイル処理での例外処理

### 🆕 拡張クラス

#### FolderProcessor
- **ZIP処理**: 2階層構造のフォルダ解析
- **ファイル分類**: 自動的なドキュメントタイプ識別
- **メタデータ管理**: フォルダ構造の保持

#### ExcelManager
- **動的操作**: シート・セル範囲の柔軟な指定
- **スタイル管理**: 書式設定の保持と適用
- **データ変換**: 構造化データのExcel形式への変換

#### TaskEngine
- **設定駆動**: JSONベースのタスク定義
- **実行管理**: プレビュー・本実行の制御
- **結果追跡**: 処理履歴と結果の管理

## 技術詳細

### Structured Output（構造化出力）

このツールはLangChainのPydantic Output Parserを使用して型安全な出力処理を実現しています：

```python
class SeverityLevel(str, Enum):
    INFO = "info"        # 正常
    WARNING = "warning"  # 警告
    ERROR = "error"      # エラー

class CheckResult(BaseModel):
    passed: bool = Field(description="チェックが成功したかどうか")
    severity: SeverityLevel = Field(description="重要度 (info/warning/error)")
    message: str = Field(min_length=1, description="チェック結果の説明")
    details: str = Field(default="", description="markdown形式で詳細な判断根拠を記載")
    
    @validator('severity')
    def severity_consistency_check(cls, v, values):
        """重要度とpassedの整合性をチェック"""
        passed = values.get('passed', True)
        if v == SeverityLevel.ERROR and passed:
            raise ValueError('errorの場合、passedはfalseである必要があります')
        return v
```

#### 特徴
- **型安全性**: Pydantic v2によるRuntime型チェック
- **データ検証**: 自動的なバリデーション（メッセージの空文字チェック等）
- **エラーハンドリング**: 構造化されたエラー情報
- **整合性チェック**: 重要度とpassed状態の論理的整合性確認
- **フォールバック**: 構造化出力解析失敗時の安全な処理

### 並列処理アーキテクチャ

```python
# ThreadPoolExecutorによる並列処理
with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
    futures = {}
    for file_data in files:
        future = executor.submit(check_invoice, file_data, rules)
        futures[future] = file_name
    
    # 完了順に結果を収集
    for future in concurrent.futures.as_completed(futures):
        result = future.result()
        # リアルタイムプログレス更新
```

## 設定例

### タスク設定（config/task_configs.json）

```json
{
  "accounting_task_1": {
    "name": "請求書・入金明細照合",
    "description": "請求書と入金明細を照合し、結果を調書に記載",
    "output_config": {
      "target_sheet": "結果",
      "start_cell": "B3",
      "columns": [
        {"key": "data_id", "header": "データID"},
        {"key": "invoice_amount", "header": "請求金額"},
        {"key": "payment_amount", "header": "入金金額"},
        {"key": "match_status", "header": "照合結果"},
        {"key": "variance", "header": "差額"},
        {"key": "remarks", "header": "備考"}
      ]
    }
  }
}
```

## カスタマイズ

### 新しいルールの追加

1. 「ルール設定」タブでルールを追加
2. カテゴリ、ルール名、チェック内容を設定  
3. システムが自動的にStructured Outputのフォーマット指示を追加
4. rules.jsonに永続化されUUID管理

### 新しいファイル形式の対応

`file_processor.py`の`FileProcessor`クラスを拡張：

```python
def _process_new_format(self, file_content: bytes, file_name: str) -> str:
    """新しいファイル形式の処理を実装"""
    # 実装コード
    return extracted_text
```

### 新しい業務タイプの追加

タスク設定の追加のみで新業務に対応：
- カスタムプロンプトテンプレート
- 出力フォーマットのカスタマイズ
- 処理ロジックの設定

### GPT-4.1モデル設定

`invoice_checker.py`および`core/llm_client.py`でモデルが固定されています [[memory:5049262]]：

```python
self.llm = ChatOpenAI(
    model_name="gpt-4.1",  # GPT-4.1相当、絶対に変更禁止
    openai_api_key=api_key,
)
```
