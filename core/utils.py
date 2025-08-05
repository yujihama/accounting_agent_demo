import streamlit as st

# カスタムCSS
def load_custom_css():
    st.markdown("""
    <style>
    /* メインコンテナのスタイル */
    .main {
        padding-top: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    
    /* ヘッダーのスタイル */
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
        font-size: 3rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    
    .main-header h1 {
        font-size: 5rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .main-header p {
        font-size: 2.5rem;
        opacity: 0.9;
        margin: 0;
    }
    
    /* タブのスタイル */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f8f9fa;
        padding: 8px;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        width: 280px;
        white-space: pre-wrap;
        background-color: white;
        border-radius: 8px;
        color: #495057;
        font-weight: 500;
        border: 1px solid #e9ecef;
        transition: all 0.3s ease;
    }
                
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    /* カードスタイル */
    .card {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        border: 1px solid #e9ecef;
        margin-bottom: 1.5rem;
    }
    
    .card-header {
        font-size: 1.25rem;
        font-weight: 600;
        color: #495057;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #f8f9fa;
    }
    
    /* サイドバーのスタイル */
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);
    }
    
    /* ボタンのスタイル */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        font-weight: 500;
        transition: all 0.3s ease;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
         .stButton > button:hover {
         transform: translateY(-2px);
         box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
     }
     
    /* メトリクス */
    [data-testid="metric-container"] {
        background: white;
        border: 1px solid #e9ecef;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }
    
    /* 成功メッセージ */
    .stSuccess {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* 警告メッセージ */
    .stWarning {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* エラーメッセージ */
    .stError {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* プログレスバー */
    .stProgress .stProgress-bar {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 4px;
    }
    
    /* データフレーム */
    .dataframe {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    /* エクスパンダー - より幅広いセレクター対応 */
    [data-testid="stExpander"] {
        border: none !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        border-radius: 12px;
        overflow: hidden;
        margin-bottom: 1rem;
    }
    
    [data-testid="stExpander"] > div > div:first-child,
    .streamlit-expanderHeader,
    [data-testid="stExpander"] summary {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 50%, #f8f9fa 100%) !important;
        border: 1px solid #dee2e6 !important;
        border-radius: 12px 12px 0 0 !important;
        padding: 1rem 1.5rem !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
        font-weight: 500 !important;
        position: relative !important;
        cursor: pointer !important;
    }
    
    [data-testid="stExpander"] > div > div:first-child:hover,
    .streamlit-expanderHeader:hover,
    [data-testid="stExpander"] summary:hover {
        background: linear-gradient(135deg, #e9ecef 0%, #dee2e6 50%, #e9ecef 100%) !important;
        border-color: #667eea !important;
        box-shadow: 0 4px 8px rgba(102, 126, 234, 0.15) !important;
        transform: translateY(-1px) !important;
    }
    
    [data-testid="stExpander"] > div > div:first-child::before,
    .streamlit-expanderHeader::before,
    [data-testid="stExpander"] summary::before {
        content: '' !important;
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        width: 4px !important;
        height: 100% !important;
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%) !important;
        transition: all 0.3s ease !important;
    }
    
    [data-testid="stExpander"] > div > div:first-child:hover::before,
    .streamlit-expanderHeader:hover::before,
    [data-testid="stExpander"] summary:hover::before {
        width: 6px !important;
    }
    
    /* エクスパンダーのテキスト（ボタンを除く） */
    [data-testid="stExpander"] [data-testid="stMarkdownContainer"] p:not(.stButton):not([class*="stButton"]),
    [data-testid="stExpander"] summary,
    .streamlit-expanderHeader [data-testid="stMarkdownContainer"]:not(.stButton):not([class*="stButton"]) {
        color: #495057;
        font-weight: 600 !important;
        letter-spacing: 0.3px !important;
        margin: 0 !important;
        font-size: 1.1rem !important;
        line-height: 1.4 !important;
    }
    
    /* エクスパンダー内のボタンの文字色を白に保持 */
    [data-testid="stExpander"] .stButton > button,
    [data-testid="stExpander"] button,
    [data-testid="stExpander"] .stButton button,
    [data-testid="stExpander"] .stButton > button *,
    [data-testid="stExpander"] button[kind="secondary"],
    [data-testid="stExpander"] button[data-testid*="baseButton"] {
        color: white !important;
    }
    
    # [data-testid="stExpander"]:hover [data-testid="stMarkdownContainer"] p,
    # [data-testid="stExpander"]:hover summary,
    # .streamlit-expanderHeader:hover [data-testid="stMarkdownContainer"] {
    #     color: #667eea !important;
    }
    
    /* エクスパンダーの矢印アイコン */
    [data-testid="stExpander"] svg,
    .streamlit-expanderHeader svg {
        transition: all 0.3s ease !important;
        filter: drop-shadow(0 1px 2px rgba(0, 0, 0, 0.1)) !important;
    }
    
    [data-testid="stExpander"]:hover svg,
    .streamlit-expanderHeader:hover svg {
        color: #667eea !important;
        transform: scale(1.1) !important;
    }
    
     /* エクスパンダーのコンテンツ部分 */
     [data-testid="stExpander"] > div > div:nth-child(2),
     [data-testid="stExpander"] > div:nth-child(2),
     .streamlit-expanderContent {
         background-color: #ffffff !important;
         border: 1px solid #e9ecef !important;
         border-top: none !important;
         border-radius: 0 0 12px 12px !important;
         padding: 2rem 1.5rem !important;
         box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
         position: relative !important;
         margin-top: 0 !important;
     }
     
     /* エクスパンダー内のすべての直接的な子要素にもパディングを適用 */
     [data-testid="stExpander"] [data-testid="stVerticalBlock"],
     [data-testid="stExpander"] .element-container {
         padding-top: 0.5rem !important;
         padding-bottom: 0.5rem !important;
     }
    
    [data-testid="stExpander"] > div > div:nth-child(2)::before,
    .streamlit-expanderContent::before {
        content: '' !important;
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        width: 4px !important;
        height: 100% !important;
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%) !important;
    }
    
    /* アクティブ状態（展開中）のエクスパンダー */
    [data-testid="stExpander"][open] > div > div:first-child,
    [data-testid="stExpander"][aria-expanded="true"] > div > div:first-child,
    [data-testid="stExpander"][open] summary,
    [data-testid="stExpander"][aria-expanded="true"] .streamlit-expanderHeader {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        border-color: #667eea !important;
        color: white !important;
        box-shadow: 0 4px 8px rgba(102, 126, 234, 0.25) !important;
    }
    
    [data-testid="stExpander"][open] [data-testid="stMarkdownContainer"] p,
    [data-testid="stExpander"][open] summary,
    [data-testid="stExpander"][aria-expanded="true"] [data-testid="stMarkdownContainer"],
    [data-testid="stExpander"][aria-expanded="true"] .streamlit-expanderHeader [data-testid="stMarkdownContainer"] {
        color: white !important;
    }
    
    [data-testid="stExpander"][open] svg,
    [data-testid="stExpander"][aria-expanded="true"] svg,
    [data-testid="stExpander"][aria-expanded="true"] .streamlit-expanderHeader svg {
        color: white !important;
    }
    
    [data-testid="stExpander"][open] > div > div:first-child::before,
    [data-testid="stExpander"][open] summary::before,
    [data-testid="stExpander"][aria-expanded="true"] .streamlit-expanderHeader::before {
        background: rgba(255, 255, 255, 0.3) !important;
        width: 6px !important;
    }
    
    /* セレクトボックス */
    .stSelectbox > div > div {
        border-radius: 8px;
        border: 1px solid #e9ecef;
    }
    
    /* テキスト入力 */
    .stTextInput > div > div > input {
        border-radius: 8px;
        border: 1px solid #e9ecef;
    }
    
    /* テキストエリア */
    .stTextArea > div > div > textarea {
        border-radius: 8px;
        border: 1px solid #e9ecef;
    }
    
    /* ファイルアップローダー */
    .stFileUploader {
        border: 2px dashed #e9ecef;
        border-radius: 8px;
        padding: 2rem;
        text-align: center;
        background-color: #f8f9fa;
        transition: all 0.3s ease;
    }
    
    .stFileUploader:hover {
        border-color: #667eea;
        background-color: #f0f3ff;
    }
    
    /* スライダー */
    .stSlider > div > div > div > div {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    /* ページタイトル非表示（カスタムヘッダーを使用） */
    .main > div:first-child > div:first-child > div:first-child {
        padding-top: 0;
    }
    
    h1 {
        display: none;
    }
    </style>
    """, unsafe_allow_html=True)
