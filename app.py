import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
import fitz  # PyMuPDF
from PIL import Image
import io
import pytesseract
import tempfile

# Tesseractのパス設定
def setup_tesseract():
    """Tesseractのパスを自動検出して設定"""
    # Streamlit Cloud環境では、tesseractは既にPATHに含まれている
    try:
        import shutil
        tesseract_path = shutil.which('tesseract')
        if tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            return True
        
        # 環境変数PATHから直接検索
        import subprocess
        result = subprocess.run(['tesseract', '--version'], 
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            return True
    except:
        pass
    
    # ローカル環境（Windows）でのフォールバック
    if os.name == 'nt':
        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            r'C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'.format(os.getenv('USERNAME', '')),
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                return True
    
    return False

# Tesseractのセットアップ
tesseract_available = setup_tesseract()

# デバッグ情報
if tesseract_available:
    st.sidebar.success("✅ Tesseract OCR が利用可能です")
else:
    st.sidebar.error("❌ Tesseract OCR が見つかりません")

# ページ設定
st.set_page_config(
    page_title="PDF検査データ抽出アプリ",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# タイトルと説明
st.title("📊 PDF検査データ抽出アプリ")
st.markdown("""
このアプリケーションは、検査報告書PDFから27項目のデータを自動抽出し、
Excelファイルとして出力します。
""")

# サイドバーにアプリの説明
st.sidebar.title("📋 使用方法")
st.sidebar.markdown("""
1. **PDFファイルをアップロード**
   - 検査報告書のPDFファイルを選択してください
   
2. **データ抽出**
   - 「データ抽出開始」ボタンをクリック
   - 処理完了までお待ちください
   
3. **Excelファイルダウンロード**
   - 抽出完了後、Excelファイルをダウンロードできます
""")

# 抽出対象項目のリスト
TARGET_ITEMS_27 = [
    'ShipNo.', 'Kind of Material', 'Place', 'Inspection Date', 'Inspection Time',
    'Weather', 'Dry bulb Temp', 'Wet bulb Temp', 'Relative Humidity', 'Dew Point',
    'Surface Temp', 'Judgement', 'Surface Cleanliness', 'Surface Profile',
    'Water Soluble Salt', 'Dust', 'Oil / Grease', 'contamination of abrasive',
    'Manufacturer', 'Product name', 'ID number', 'Batch No Base', 'Batch No Hard',
    'Lower', 'Upper', 'Measured D.F.T', 'Curing'
]

def process_value(item_name, text):
    """
    項目名に応じて値を処理する関数
    """
    if not text or not text.strip():
        return ""
    
    text = text.strip()
    
    # Inspection Dateの処理
    if item_name == 'Inspection Date':
        # 日付形式に変換（例: "2025-03-07" -> "2025/03/07"）
        date_match = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', text)
        if date_match:
            year, month, day = date_match.groups()
            return f"{year}/{month.zfill(2)}/{day.zfill(2)}"
        return text
    
    # Weatherの処理（スラッシュ前後を別セルに分離）
    elif item_name == 'Weather':
        # 天気の日本語変換
        weather_map = {
            'fine': '晴',
            'cloud': '曇', 
            'rain': '雨'
        }
        
        # スラッシュで分割
        parts = text.split('/')
        processed_parts = []
        
        for part in parts:
            part = part.strip()
            # ハイフン（未記入）をスキップ
            if part in ['-', '-/-', '', '- -']:
                continue
            
            # 天気の日本語変換
            for eng, jpn in weather_map.items():
                if eng in part.lower():
                    processed_parts.append(jpn)
                    break
            else:
                # 変換できない場合はそのまま（ただし空でない場合のみ）
                if part and not re.match(r'^[\s\-]+$', part):
                    processed_parts.append(part)
        
        return '/'.join(processed_parts) if processed_parts else ""
    
    # Surface Profileの処理（ハイフンと単位を保持）
    elif item_name == 'Surface Profile':
        # ハイフンを含む範囲と単位を保持（例: "30-75 um" -> "30-75 µm"）
        range_match = re.search(r'(\d+-\d+)\s*um', text)
        if range_match:
            return range_match.group(1) + ' µm'  # uをµに変換
        return text
    
    # Dustの処理（ISO認識問題を修正）
    elif item_name == 'Dust':
        # ISO認識問題を修正（例: "1S08502-3" -> "ISO8502-3"）
        dust_match = re.search(r'([A-Za-z0-9]+-[0-9]+)', text)
        if dust_match:
            dust_value = dust_match.group(1)
            # 1S0をISOに修正
            if dust_value.startswith('1S0'):
                dust_value = 'ISO' + dust_value[3:]
            return dust_value
        return text
    
    # ID numberの処理（ハイフンを保持）
    elif item_name == 'ID number':
        # ハイフンを含む値を保持（例: "M2-10000A"）
        id_match = re.search(r'([A-Za-z0-9]+-[A-Za-z0-9]+)', text)
        if id_match:
            return id_match.group(1)
        return text
    
    # Inspection Timeの処理（時間形式を保持）
    elif item_name == 'Inspection Time':
        # 時間形式を保持（例: "07:45 - -" -> "07:45"）
        time_match = re.search(r'(\d{1,2}:\d{2})', text)
        if time_match:
            return time_match.group(1)
        return text
    
    # その他の項目の処理
    else:
        # ハイフン（未記入）を除去（ただし数値のマイナスは保持）
        text = re.sub(r'\s*-\s*-\s*', '', text)  # 連続ハイフンを除去
        text = re.sub(r'\s*-\s*$', '', text)    # 末尾のハイフンを除去
        
        # 数値のみの項目（温度、湿度、膜厚など）の処理
        numeric_items = ['Dry bulb Temp', 'Wet bulb Temp', 'Relative Humidity', 'Dew Point', 
                        'Surface Temp', 'Water Soluble Salt', 'Lower', 'Upper', 'Measured D.F.T']
        
        if item_name in numeric_items:
            # 数値と単位を分離（単位は表示しない、マイナス値を保持）
            number_pattern = r'-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?'
            number_match = re.search(number_pattern, text)
            
            if number_match:
                return number_match.group()
            return text
        else:
            # テキスト項目の場合はそのまま返す（ハイフン除去後）
            return text

def extract_data_from_pdf(pdf_file):
    """
    PDFからデータを抽出する関数
    """
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    all_pages_data = []
    
    try:
        # 一時ファイルに保存
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(pdf_file.read())
            tmp_file_path = tmp_file.name
        
        with pdfplumber.open(tmp_file_path) as pdf:
            total_pages = len(pdf.pages)
            status_text.text(f"総ページ数: {total_pages}")
            
            for i, page in enumerate(pdf.pages):
                progress = (i + 1) / total_pages
                progress_bar.progress(progress)
                status_text.text(f"{i+1}/{total_pages}ページ目を処理中...")
                
                page_data = {}
                
                # PyMuPDFでページを画像として取得
                doc = fitz.open(tmp_file_path)
                pdf_page = doc[page.page_number - 1]
                
                # 高解像度でページを画像に変換
                mat = fitz.Matrix(3.0, 3.0)
                pix = pdf_page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                
                # PIL Imageオブジェクトに変換
                image = Image.open(io.BytesIO(img_data))
                
                # OCRでテキストを抽出
                try:
                    page_text = pytesseract.image_to_string(image, lang='eng')
                    
                    if page_text and page_text.strip():
                        lines = page_text.split('\n')
                        
                        for line in lines:
                            cleaned_line = re.sub(r'\s+', ' ', line).strip()
                            
                            if cleaned_line:
                                for item in TARGET_ITEMS_27:
                                    # 基本的な検索
                                    if cleaned_line.startswith(item):
                                        value_text = cleaned_line[len(item):].strip()
                                        value_text = re.sub(r'^[:：\s]+', '', value_text)
                                        
                                        if value_text:
                                            processed_value = process_value(item, value_text)
                                            
                                            # Weatherの特別処理：常にWeather_1とWeather_2に分離
                                            if item == 'Weather':
                                                if '/' in processed_value:
                                                    # スラッシュがある場合：前後を分離
                                                    weather_parts = processed_value.split('/')
                                                    page_data[f"{item}_1"] = weather_parts[0] if len(weather_parts) > 0 else ""
                                                    page_data[f"{item}_2"] = weather_parts[1] if len(weather_parts) > 1 else ""
                                                else:
                                                    # スラッシュがない場合：単独の天気をWeather_1に、Weather_2は空
                                                    page_data[f"{item}_1"] = processed_value
                                                    page_data[f"{item}_2"] = ""
                                            else:
                                                page_data[item] = processed_value
                                        else:
                                            if item == 'Weather':
                                                page_data[f"{item}_1"] = ""
                                                page_data[f"{item}_2"] = ""
                                            else:
                                                page_data[item] = ""
                                        break
                                    
                                    # ShipNo.の特別な処理（Ship No.としても検索）
                                    elif item == 'ShipNo.' and re.search(r'Ship\s*No\.?\s*', cleaned_line, re.IGNORECASE):
                                        match = re.search(r'Ship\s*No\.?\s*[:\s]*([^\n\r]+)', cleaned_line, re.IGNORECASE)
                                        if match:
                                            value_text = match.group(1).strip()
                                            processed_value = process_value(item, value_text)
                                            page_data[item] = processed_value
                                        break
                                    
                                    # Dry bulb Tempの特別な処理（Dry bulb Temp.としても検索）
                                    elif item == 'Dry bulb Temp' and re.search(r'Dry\s+bulb\s+Temp\.?\s*', cleaned_line, re.IGNORECASE):
                                        match = re.search(r'Dry\s+bulb\s+Temp\.?\s*[:\s]*([^\n\r]+)', cleaned_line, re.IGNORECASE)
                                        if match:
                                            value_text = match.group(1).strip()
                                            processed_value = process_value(item, value_text)
                                            page_data[item] = processed_value
                                        break
                except Exception as ocr_error:
                    st.warning(f"OCR処理でエラーが発生しました: {str(ocr_error)}")
                
                finally:
                    doc.close()
                
                if page_data:
                    all_pages_data.append(page_data)
        
        # 一時ファイルを削除
        os.unlink(tmp_file_path)
        
        progress_bar.progress(1.0)
        status_text.text("データ抽出完了！")
        
        return all_pages_data
        
    except Exception as e:
        status_text.text(f"エラーが発生しました: {str(e)}")
        return None

# メイン処理
def main():
    # Tesseractの可用性チェック
    if not tesseract_available:
        st.error("""
        ❌ **Tesseract OCRがインストールされていません**
        
        以下の手順でインストールしてください：
        
        **Windows:**
        ```bash
        winget install tesseract-ocr.tesseract
        ```
        
        **または Chocolatey:**
        ```bash
        choco install tesseract
        ```
        
        インストール後、アプリケーションを再起動してください。
        """)
        return
    
    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "PDFファイルをアップロードしてください",
        type=['pdf'],
        help="検査報告書のPDFファイルを選択してください"
    )
    
    if uploaded_file is not None:
        st.success(f"ファイル '{uploaded_file.name}' がアップロードされました")
        
        # ファイル情報を表示
        file_size = len(uploaded_file.getvalue())
        st.info(f"ファイルサイズ: {file_size:,} bytes")
        
        # データ抽出ボタン
        if st.button("🚀 データ抽出開始", type="primary"):
            with st.spinner("データを抽出中..."):
                result_data = extract_data_from_pdf(uploaded_file)
            
            if result_data:
                # データフレーム作成
                df = pd.DataFrame(result_data)
                
                # 結果表示
                st.success(f"✅ 抽出完了！ {len(result_data)}ページのデータを抽出しました")
                
                # データプレビュー
                st.subheader("📋 抽出データプレビュー")
                st.dataframe(df, use_container_width=True)
                
                # 統計情報
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("総ページ数", len(result_data))
                with col2:
                    st.metric("抽出項目数", len(df.columns))
                with col3:
                    st.metric("データ総数", len(df) * len(df.columns))
                
                # Excelファイル生成
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='検査データ', index=False)
                
                excel_buffer.seek(0)
                
                # ダウンロードボタン
                st.download_button(
                    label="📥 Excelファイルをダウンロード",
                    data=excel_buffer.getvalue(),
                    file_name=f"検査データ_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # 列名一覧
                with st.expander("📝 抽出された項目一覧"):
                    for i, col in enumerate(df.columns, 1):
                        st.write(f"{i:2d}. {col}")
            
            else:
                st.error("❌ データの抽出に失敗しました。PDFファイルを確認してください。")

# フッター
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray;'>
    <small>PDF検査データ抽出アプリ - 27項目の自動抽出機能</small>
</div>
""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
