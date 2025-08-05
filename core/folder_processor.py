import io
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
from datetime import datetime
import json

# PDF処理ライブラリ
try:
    import PyPDF2
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

class FolderProcessor:
    """
    フォルダ構造を持つ証跡データの処理機能
    2階層構造（データ/ドキュメント）に対応
    """
    
    def __init__(self):
        self.supported_document_types = ['.pdf', '.txt', '.png', '.jpg', '.jpeg', '.xlsx', '.xls', '.docx']
    
    def process_evidence_folder(self, uploaded_file) -> Dict[str, Any]:
        """
        アップロードされたZIPファイル（フォルダ）を処理
        
        Args:
            uploaded_file: StreamlitのZIPファイルアップロード
            
        Returns:
            階層化された証跡データ
        """
        try:
            # ZIPファイルを展開
            zip_content = self._extract_zip_file(uploaded_file)
            
            # デバッグ情報：zipファイル構造を記録
            zip_structure = list(zip_content.keys())
            
            # 2階層構造を解析
            structured_data = self._analyze_folder_structure(zip_content)
            
            # デバッグ: 構造化データの内容を表示
            print(f"解析された構造:")
            for data_id, files in structured_data.items():
                print(f"  {data_id}: {len(files)}ファイル - {files}")
            
            # 各ドキュメントを処理
            processed_data = self._process_documents(structured_data, zip_content)
            
            return {
                "success": True,
                "data": processed_data,
                "metadata": {
                    "total_data_folders": len(processed_data),
                    "total_documents": sum(len(data["documents"]) for data in processed_data.values()),
                    "processed_at": datetime.now().isoformat(),
                    "zip_structure": zip_structure,  # デバッグ情報
                    "detected_data_folders": list(structured_data.keys())  # デバッグ情報
                }
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": f"フォルダ処理エラー: {str(e)}",
                "data": {}
            }
    
    def _extract_zip_file(self, uploaded_file) -> Dict[str, bytes]:
        """ZIPファイルを展開"""
        zip_content = {}
        
        with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as zip_file:
            print(f"ZIPファイル内のファイル一覧:")
            for file_info in zip_file.filelist:
                print(f"  {file_info.filename} (ディレクトリ: {file_info.is_dir()})")
                
                if not file_info.is_dir():
                    file_path = Path(file_info.filename)
                    file_extension = file_path.suffix.lower()
                    
                    if file_extension in self.supported_document_types:
                        zip_content[file_info.filename] = zip_file.read(file_info)
                        print(f"    ✓ 読み込み: {file_info.filename} ({file_extension})")
                    else:
                        print(f"    ✗ スキップ: {file_info.filename} ({file_extension} はサポート外)")
        
        print(f"読み込まれたファイル数: {len(zip_content)}")
        return zip_content
    
    def _analyze_folder_structure(self, zip_content: Dict[str, bytes]) -> Dict[str, List[str]]:
        """
        2階層フォルダ構造を解析
        
        Returns:
            {データID: [ドキュメントパスのリスト]}
        """
        structured_data = {}
        
        for file_path in zip_content.keys():
            parts = Path(file_path).parts
            
            # zipファイル内の構造を解析
            # 通常: root_folder/data_folder/file.ext の形式
            if len(parts) >= 3:
                # 2階層目をデータIDとして使用（通常の2階層構造）
                data_id = parts[1]
            elif len(parts) == 2:
                # 1階層目をデータIDとして使用（フラット構造）
                data_id = parts[0]
            else:
                # ルートにあるファイルはスキップ
                continue
                
            if data_id not in structured_data:
                structured_data[data_id] = []
            
            structured_data[data_id].append(file_path)
        
        return structured_data
    
    def _process_documents(self, structured_data: Dict[str, List[str]], zip_content: Dict[str, bytes]) -> Dict[str, Dict[str, Any]]:
        """各データフォルダ内のドキュメントを処理"""
        processed_data = {}
        
        print(f"\nドキュメント処理開始:")
        for data_id, document_paths in structured_data.items():
            documents = {}
            print(f"\n処理中: {data_id} ({len(document_paths)}ファイル)")
            
            for doc_path in document_paths:
                try:
                    doc_name = Path(doc_path).name
                    doc_extension = Path(doc_path).suffix.lower()
                    
                    # ファイル内容を取得
                    file_content = zip_content.get(doc_path, b"")
                    print(f"  処理: {doc_name} ({doc_extension}, {len(file_content)} bytes)")
                    
                    # テキスト内容を抽出
                    content_text = ""
                    if doc_extension == ".txt":
                        try:
                            content_text = file_content.decode('utf-8')
                        except:
                            try:
                                content_text = file_content.decode('shift-jis')
                            except:
                                content_text = "テキストデコードエラー"
                    elif doc_extension == ".pdf":
                        content_text = self._extract_pdf_text(file_content)
                    elif doc_extension in [".xlsx", ".xls"]:
                        content_text = self._extract_excel_text(file_content, doc_name)
                    elif doc_extension == ".docx":
                        content_text = "[.docxファイル - 内容抽出は未実装]"
                    
                    documents[doc_name] = {
                        "path": doc_path,
                        "type": self._classify_document_type(doc_name),
                        "extension": doc_extension,
                        "content": content_text,  # 実際のコンテンツを追加
                        "size": len(file_content),
                        "processed_at": datetime.now().isoformat(),
                        "status": "処理完了"
                    }
                    print(f"    ✓ 完了: 内容サイズ = {len(content_text)} 文字")
                    
                except Exception as e:
                    documents[doc_name] = {
                        "error": f"ドキュメント処理エラー: {str(e)}"
                    }
            
            processed_data[data_id] = {
                "data_id": data_id,
                "documents": documents,
                "document_count": len(documents)
            }
        
        return processed_data
    
    def _classify_document_type(self, filename: str) -> str:
        """ファイル名からドキュメントタイプを推定"""
        filename_lower = filename.lower()
        
        if "請求" in filename_lower or "invoice" in filename_lower:
            return "請求書"
        elif "入金" in filename_lower or "payment" in filename_lower or "明細" in filename_lower:
            return "入金明細"
        elif "領収" in filename_lower or "receipt" in filename_lower:
            return "領収書"
        elif "契約" in filename_lower or "contract" in filename_lower:
            return "契約書"
        else:
            return "その他"
    
    def get_data_summary(self, processed_data: Dict[str, Any]) -> Dict[str, Any]:
        """処理されたデータのサマリーを取得"""
        if not processed_data.get("success"):
            return {"error": "データ処理が完了していません"}
        
        data = processed_data["data"]
        
        # ドキュメントタイプ別統計
        doc_type_stats = {}
        total_docs = 0
        
        for data_entry in data.values():
            for doc_name, doc_info in data_entry["documents"].items():
                doc_type = doc_info.get("type", "不明")
                doc_type_stats[doc_type] = doc_type_stats.get(doc_type, 0) + 1
                total_docs += 1
        
        return {
            "total_data_entries": len(data),
            "total_documents": total_docs,
            "document_types": doc_type_stats,
            "data_ids": list(data.keys())
        }
    
    def validate_folder_structure(self, uploaded_file) -> Tuple[bool, Optional[str]]:
        """
        フォルダ構造の妥当性を検証
        
        Returns:
            (is_valid, error_message)
        """
        try:
            # ファイルサイズチェック
            max_size = 100 * 1024 * 1024  # 100MB
            if uploaded_file.size > max_size:
                return False, f"ファイルサイズが大きすぎます（最大100MB）: {uploaded_file.size / 1024 / 1024:.1f}MB"
            
            # ZIPファイルかチェック
            if not uploaded_file.name.lower().endswith('.zip'):
                return False, "ZIPファイルをアップロードしてください"
            
            # ZIPファイルの内容をチェック
            with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as zip_file:
                file_list = zip_file.filelist
                
                # 最低1つのファイルが含まれているかチェック
                valid_files = [f for f in file_list if not f.is_dir() and 
                             Path(f.filename).suffix.lower() in self.supported_document_types]
                
                if not valid_files:
                    return False, "サポートされているドキュメントが見つかりません"
                
                # 2階層構造のチェック
                folder_structure = {}
                for file_info in valid_files:
                    parts = Path(file_info.filename).parts
                    # zipファイル内の構造に応じて適切な階層を選択
                    if len(parts) >= 3:
                        # root_folder/data_folder/file.ext の形式
                        data_folder = parts[1]
                    elif len(parts) == 2:
                        # data_folder/file.ext の形式
                        data_folder = parts[0]
                    else:
                        continue
                        
                    folder_structure[data_folder] = folder_structure.get(data_folder, 0) + 1
                
                if not folder_structure:
                    return False, "適切な2階層フォルダ構造が見つかりません"
            
            # ファイルポインタをリセット
            uploaded_file.seek(0)
            
            return True, None
            
        except Exception as e:
            return False, f"フォルダ構造検証エラー: {str(e)}"
    
    def _extract_pdf_text(self, file_content: bytes) -> str:
        """PDFファイルからテキストを抽出"""
        
        # PyMuPDFを優先的に使用（より高精度）
        if FITZ_AVAILABLE:
            try:
                pdf_document = fitz.open(stream=file_content, filetype="pdf")
                text_content = []
                
                for page_num in range(pdf_document.page_count):
                    page = pdf_document.load_page(page_num)
                    page_text = page.get_text()
                    if page_text.strip():
                        text_content.append(f"--- ページ {page_num + 1} ---")
                        text_content.append(page_text)
                
                pdf_document.close()
                
                if text_content:
                    return "\n".join(text_content)
                else:
                    return "PDFからテキストを抽出できませんでした（空のPDF）"
                    
            except Exception as e:
                # PyMuPDFでエラーが発生した場合、PyPDF2にフォールバック
                if PDF_AVAILABLE:
                    return self._extract_pdf_with_pypdf2(file_content)
                else:
                    return f"PyMuPDF処理エラー: {str(e)}"
        
        # PyPDF2を使用
        elif PDF_AVAILABLE:
            return self._extract_pdf_with_pypdf2(file_content)
        
        else:
            return "PDF処理ライブラリがインストールされていません"
    
    def _extract_pdf_with_pypdf2(self, file_content: bytes) -> str:
        """PyPDF2を使用してPDFからテキストを抽出"""
        try:
            pdf_file = io.BytesIO(file_content)
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            
            text_content = []
            for page_num, page in enumerate(pdf_reader.pages):
                page_text = page.extract_text()
                if page_text.strip():
                    text_content.append(f"--- ページ {page_num + 1} ---")
                    text_content.append(page_text)
            
            if text_content:
                return "\n".join(text_content)
            else:
                return "PDFからテキストを抽出できませんでした（空のPDF）"
                
        except Exception as e:
            return f"PDF処理エラー（PyPDF2）: {str(e)}"
    
    def _extract_excel_text(self, file_content: bytes, file_name: str) -> str:
        """Excelファイルからテキストを抽出"""
        try:
            excel_file = io.BytesIO(file_content)
            
            # すべてのシートを読み込み
            excel_data = pd.read_excel(excel_file, sheet_name=None)
            
            text_content = []
            for sheet_name, df in excel_data.items():
                text_content.append(f"--- シート: {sheet_name} ---")
                
                # データフレームを文字列に変換
                df_string = df.to_string(index=False, na_rep='')
                text_content.append(df_string)
                text_content.append("")
            
            if text_content:
                return "\n".join(text_content)
            else:
                return "Excelファイルが空です"
                
        except Exception as e:
            return f"Excel処理エラー: {str(e)}"