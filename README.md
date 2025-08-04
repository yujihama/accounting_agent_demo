# 請求書カットオフチェックツール

経理業務における請求書のカットオフチェックを自動化するAIツールです。GPT-4.1を活用して、大量の請求書を効率的にチェックし、会計基準や経費精算ルールに基づいた検証を行います。

## 機能

- **自動請求書チェック**: GPT-4.1を使用した高精度な請求書内容の検証（Structured Output対応）
- **カスタムルール設定**: 会計基準や経費精算ルールを自由に設定可能
- **🆕 AIルール提案機能**: マニュアルや規定書をアップロードして自動的にチェックルールを提案
  - PDF、Word、Excel、テキストファイルから内容を解析
  - 既存ルールとの重複を回避した新しいルール提案
  - 提案されたルールの編集・カスタマイズ機能
- **並列処理**: 設定可能な並列処理数（1-12）で大量のファイルを効率的に処理
- **多様なファイル形式対応**: PDF、Excel (.xlsx, .xls)、Word (.docx) に対応
- **詳細な結果レポート**: Excel形式での結果出力とダウンロード
- **リアルタイム処理状況表示**: プログレスバーと処理完了数の表示
- **高度なフィルタリング**: 要確認・正常・全件での結果表示切り替え

## 技術スタック

- **Streamlit**: Webベースユーザーインターフェース
- **LangChain**: GPT-4.1との連携とPydantic Output Parser
- **OpenAI GPT-4.1**: 請求書内容の分析と検証（モデル固定）
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

#### 3. OpenAI APIキーの準備

OpenAI のAPIキーを取得してください。アプリケーション内のサイドバーで設定できます。

#### 4. アプリケーションの起動

```powershell
streamlit run app.py
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

## 使用方法

### 1. APIキー設定

アプリケーション起動後、サイドバーでOpenAI APIキーを設定してください。GPT-4.1を使用するため、適切なAPIキーが必要です。

### 2. ルール設定

「ルール設定」タブで請求書チェックのルールを設定します：

#### 手動ルール追加
- **ルール名**: チェックルールの識別名
- **カテゴリ**: 日付チェック、金額チェック、承認チェック、書式チェック、その他
- **チェックルールの内容**: GPT-4.1に送信する具体的な指示

#### 🆕 AIルール提案機能
マニュアルや規定書から自動的にルールを提案する新機能：

1. **ドキュメントアップロード**: 会計マニュアル、請求書チェック規定、内部統制文書などをアップロード
2. **AI分析**: GPT-4がドキュメント内容を分析し、請求書チェックに関連するルールを抽出
3. **重複回避**: 既存ルールとの重複を自動的に回避
4. **ルール提案**: 信頼度スコア付きでルールを提案
5. **編集・追加**: 提案されたルールを確認・編集してから追加可能

**対応ファイル形式**: PDF、Word (.docx)、Excel (.xlsx)、テキスト (.txt, .md)

#### デフォルトルール

以下のルールが初期設定されています：

- **請求書日付チェック**: 請求書の日付が会計期間内（期末日以前）かどうかを確認
- **金額整合性チェック**: 税込金額と税抜金額の計算、消費税率の確認
- **必須項目チェック**: 請求者名、請求日、支払期限、金額、内容、振込先の記載確認
- **承認印チェック**: 承認印や署名の有無を確認
- **支払条件チェック**: 支払条件や期限の適切性を確認

### 3. ファイルアップロード

「ファイルアップロード」タブで請求書ファイルをアップロードします：

- **対応形式**: PDF、Excel (.xlsx, .xls)、Word (.docx)
- **複数ファイル**: 一括アップロード可能
- **ファイルサイズ**: 最大50MB/ファイル
- **自動処理**: アップロード後、「ファイルを処理」ボタンでテキスト抽出

### 4. チェック実行

「チェック実行」タブでチェックを実行します：

#### 前提条件チェック
- ファイル処理: アップロードしたファイルが処理済みか
- チェックルール: 1つ以上のルールが設定されているか  
- API設定: OpenAI APIキーが設定されているか

#### 実行設定
- **適用するルール**: 実行したいチェックルールを複数選択可能
- **処理対象ファイル**: チェックするファイルを選択
- **並列処理数**: サイドバーで設定（1-12の範囲）
- **推定処理時間**: 総チェック数と並列数から自動計算

### 5. 結果確認

「結果表示」タブで結果を確認できます：

#### サマリー表示
- **総ファイル数**: 処理されたファイルの総数
- **要確認**: エラーまたは警告があるファイル数
- **正常**: すべてのチェックが正常なファイル数

#### 詳細結果
- **フィルター機能**: すべて・要確認・正常での表示切り替え
- **ファイル別結果**: 各ファイルのチェック結果を詳細表示
- **重要度表示**: info（正常）・warning（警告）・error（エラー）
- **Excel出力**: 結果をタイムスタンプ付きExcelファイルでダウンロード

## ファイル構成

```
extract_invoice/
├── app.py                            # メインStreamlitアプリケーション
├── invoice_checker.py                # 請求書チェック機能（GPT-4.1連携）
├── rule_manager.py                   # ルール管理機能
├── file_processor.py                 # ファイル処理機能（PDF/Excel/Word対応）
├── requirements.txt                  # 必要なライブラリ一覧
├── run.ps1                          # PowerShell起動スクリプト
├── rules.json                       # ルール設定ファイル（自動生成）
├── config.sample.py                 # 設定サンプルファイル
├── sample_usage.py                  # 使用例サンプル
├── structured_output_demo.py        # Structured Output動作確認
├── test_basic.py                    # 基本テスト
├── venv/                           # Python仮想環境
└── README.md                       # このファイル
```

## 主要クラス

### InvoiceChecker
- **GPT-4.1**: 請求書内容の分析とチェック（モデル固定）
- **LangChain**: PydanticOutputParserによる型安全なレスポンス処理
- **Structured Output**: CheckResultモデルによる構造化された出力
- **並列処理**: concurrent.futuresによる効率的なマルチファイル処理
- **エラーハンドリング**: 構造化出力解析失敗時のフォールバック機能

### RuleManager
- **CRUD操作**: ルールの作成・読み取り・更新・削除
- **UUID管理**: 一意なルールID生成と管理
- **JSON保存**: rules.jsonファイルへの永続化
- **デフォルトルール**: 5つの基本チェックルールを初期提供
- **検索機能**: カテゴリ別・キーワード検索対応

### FileProcessor
- **多形式対応**: PDF（PyPDF2）、Excel（pandas/openpyxl）、Word（python-docx）
- **テキスト抽出**: 各ファイル形式からの構造化テキスト抽出
- **メタデータ取得**: ファイル情報、ページ数、シート情報等
- **バリデーション**: ファイルサイズ・形式・名前の妥当性チェック
- **エラーハンドリング**: 各ファイル処理での例外処理

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

### カスタマイズ

#### 新しいルールの追加

1. 「ルール設定」タブでルールを追加
2. カテゴリ、ルール名、チェック内容を設定  
3. システムが自動的にStructured Outputのフォーマット指示を追加
4. rules.jsonに永続化されUUID管理

#### 新しいファイル形式の対応

`file_processor.py`の`FileProcessor`クラスを拡張：

```python
def _process_new_format(self, file_content: bytes, file_name: str) -> str:
    """新しいファイル形式の処理を実装"""
    # 実装コード
    return extracted_text
```

#### GPT-4.1モデル設定

`invoice_checker.py`でモデルが固定されています [[memory:5049262]]：

```python
self.llm = ChatOpenAI(
    model_name="gpt-4.1",  # 絶対に変更禁止
    openai_api_key=api_key,
)
```
