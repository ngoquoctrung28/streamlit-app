import streamlit as st
import pandas as pd
import os
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import zipfile
import re
import json

# CSS tùy chỉnh cho giao diện chuyên nghiệp
st.markdown("""
    <style>
    /* Tùy chỉnh màu sắc chung */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    /* Tiêu đề và header */
    h1, h2, h3 {
        color: #1f1f1f;
        font-weight: 600;
        margin-bottom: 1.5rem !important;
    }
    
    h3 {
        margin-top: 2rem !important;
        margin-bottom: 1rem !important;
    }
    
    h4 {
        margin-top: 1.5rem !important;
        margin-bottom: 0.75rem !important;
    }
    
    /* Nền và text */
    .stApp {
        background-color: #ffffff;
    }
    
    .stMarkdown {
        color: #1f1f1f;
    }
    
    /* Khoảng cách giữa các sections */
    .stMarkdown {
        margin-bottom: 1.5rem;
    }
    
    /* Form elements - tăng khoảng cách */
    .stTextInput > div > div > input,
    .stSelectbox > div > div > select,
    .stNumberInput > div > div > input {
        margin-bottom: 1.5rem;
    }
    
    /* Khoảng cách cho các widget */
    .element-container {
        margin-bottom: 1.5rem !important;
    }
    
    /* Text input spacing */
    div[data-testid="stTextInput"] {
        margin-bottom: 1.5rem !important;
    }
    
    /* Selectbox spacing */
    div[data-testid="stSelectbox"] {
        margin-bottom: 1.5rem !important;
    }
    
    /* Multiselect spacing */
    div[data-testid="stMultiSelect"] {
        margin-bottom: 1.5rem !important;
    }
    
    /* Number input spacing */
    div[data-testid="stNumberInput"] {
        margin-bottom: 1.5rem !important;
    }
    
    /* Button spacing */
    .stButton {
        margin-top: 0.5rem;
        margin-bottom: 1.5rem;
    }
    
    .stButton > button {
        background-color: #0d6efd;
        color: white;
        border-radius: 4px;
        border: none;
        font-weight: 500;
        padding: 0.5rem 1.5rem;
        margin-top: 0.5rem;
    }
    
    .stButton > button:hover {
        background-color: #0b5ed7;
    }
    
    /* Columns spacing */
    [data-testid="column"] {
        padding-left: 1rem;
        padding-right: 1rem;
    }
    
    [data-testid="column"]:first-child {
        padding-left: 0;
    }
    
    [data-testid="column"]:last-child {
        padding-right: 0;
    }
    
    /* Sidebar */
    .css-1d391kg {
        background-color: #f8f9fa;
    }
    
    /* Tab */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        margin-bottom: 2rem;
    }
    
    .stTabs [data-baseweb="tab"] {
        color: #1f1f1f;
        font-weight: 500;
    }
    
    /* Info, success, warning boxes */
    .stInfo {
        background-color: #e7f3ff;
        border-left: 4px solid #0d6efd;
        padding: 1rem;
        margin-bottom: 1.5rem;
    }
    
    .stSuccess {
        background-color: #d1e7dd;
        border-left: 4px solid #198754;
        padding: 1rem;
        margin-bottom: 1.5rem;
    }
    
    .stWarning {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 1rem;
        margin-bottom: 1.5rem;
    }
    
    /* Subheader spacing */
    .stSubheader {
        margin-top: 2rem !important;
        margin-bottom: 1.5rem !important;
    }
    
    /* Metric spacing */
    [data-testid="stMetricValue"] {
        margin-bottom: 0.5rem;
    }
    
    /* Dataframe spacing */
    [data-testid="stDataFrame"] {
        margin-top: 1rem;
        margin-bottom: 1.5rem;
    }
    
    /* Expander spacing */
    [data-testid="stExpander"] {
        margin-top: 1rem;
        margin-bottom: 1.5rem;
    }
    
    /* Download button spacing */
    [data-testid="stDownloadButton"] {
        margin-top: 1rem;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# Cấu hình trang
st.set_page_config(
    page_title="Tổng hợp & Tra cứu Excel",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Tiêu đề ứng dụng
st.title("Ứng dụng Tổng hợp & Tra cứu Excel")
st.markdown("---")

# Khởi tạo session state
if 'dataframes' not in st.session_state:
    st.session_state.dataframes = {}
if 'combined_df' not in st.session_state:
    st.session_state.combined_df = None
if 'search_results' not in st.session_state:
    st.session_state.search_results = None
if 'student_table' not in st.session_state:
    st.session_state.student_table = []
if 'files_loaded' not in st.session_state:
    # Tự động load các file đã lưu khi khởi động (chỉ load 1 lần)
    st.session_state.files_loaded = False

UPLOADED_FILES_DIR = "uploaded_files"
BACKUP_DIR = "backup_data"
STUDENT_TABLE_FILE = os.path.join(BACKUP_DIR, "student_table.json")
COMBINED_DF_DIR = os.path.join(BACKUP_DIR, "combined_data")
EXPORTED_DATA_DIR = os.path.join(BACKUP_DIR, "exported_data")

def ensure_upload_dir():
    """Tạo thư mục lưu file nếu chưa tồn tại"""
    if not os.path.exists(UPLOADED_FILES_DIR):
        os.makedirs(UPLOADED_FILES_DIR)

def ensure_backup_dirs():
    """Tạo các thư mục backup nếu chưa tồn tại"""
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)
    if not os.path.exists(COMBINED_DF_DIR):
        os.makedirs(COMBINED_DF_DIR)
    if not os.path.exists(EXPORTED_DATA_DIR):
        os.makedirs(EXPORTED_DATA_DIR)

def save_file_to_disk(file_bytes, filename):
    """Lưu file vào disk"""
    try:
        ensure_upload_dir()
        file_path = os.path.join(UPLOADED_FILES_DIR, filename)
        with open(file_path, 'wb') as f:
            f.write(file_bytes)
        return file_path
    except Exception as e:
        st.error(f"Lỗi khi lưu file {filename}: {str(e)}")
        return None

def load_saved_files():
    """Tải lại các file đã lưu từ disk"""
    saved_files = {}
    ensure_upload_dir()
    
    try:
        if os.path.exists(UPLOADED_FILES_DIR):
            for filename in os.listdir(UPLOADED_FILES_DIR):
                if filename.endswith(('.xlsx', '.xls')):
                    file_path = os.path.join(UPLOADED_FILES_DIR, filename)
                    try:
                        sheets = load_excel_file(file_path)
                        if sheets:
                            saved_files[filename] = sheets
                    except Exception as e:
                        continue  # Bỏ qua file lỗi
    except Exception as e:
        pass
    
    return saved_files

def delete_saved_file(filename):
    """Xóa file đã lưu"""
    try:
        file_path = os.path.join(UPLOADED_FILES_DIR, filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            return True
    except Exception as e:
        return False
    return False

def save_student_table():
    """Lưu bảng thông tin học sinh vào file JSON"""
    try:
        ensure_backup_dirs()
        with open(STUDENT_TABLE_FILE, 'w', encoding='utf-8') as f:
            json.dump(st.session_state.student_table, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        return False

def load_student_table():
    """Tải lại bảng thông tin học sinh từ file JSON"""
    try:
        if os.path.exists(STUDENT_TABLE_FILE):
            with open(STUDENT_TABLE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data
    except Exception as e:
        pass
    return []

def save_combined_df(df):
    """Lưu dữ liệu tổng hợp với timestamp (tránh ghi đè)"""
    try:
        ensure_backup_dirs()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"combined_data_{timestamp}.xlsx"
        file_path = os.path.join(COMBINED_DF_DIR, filename)
        
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Dữ liệu tổng hợp')
        
        # Lưu file mới nhất (để load lại nhanh)
        latest_file = os.path.join(COMBINED_DF_DIR, "latest_combined_data.xlsx")
        with pd.ExcelWriter(latest_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Dữ liệu tổng hợp')
        
        return True
    except Exception as e:
        return False

def load_latest_combined_df():
    """Tải lại dữ liệu tổng hợp mới nhất"""
    try:
        latest_file = os.path.join(COMBINED_DF_DIR, "latest_combined_data.xlsx")
        if os.path.exists(latest_file):
            df = pd.read_excel(latest_file, sheet_name='Dữ liệu tổng hợp')
            return df
    except Exception as e:
        pass
    return None

def save_exported_data(df, export_type='excel'):
    """Lưu dữ liệu đã xuất với timestamp"""
    try:
        ensure_backup_dirs()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        if export_type == 'excel':
            filename = f"exported_data_{timestamp}.xlsx"
            file_path = os.path.join(EXPORTED_DATA_DIR, filename)
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Dữ liệu xuất')
        else:
            filename = f"exported_data_{timestamp}.csv"
            file_path = os.path.join(EXPORTED_DATA_DIR, filename)
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
        
        return True
    except Exception as e:
        return False

def load_excel_file(file):
    """Đọc file Excel và trả về dictionary các sheet"""
    try:
        if isinstance(file, str):
            # Nếu là đường dẫn file
            excel_file = pd.ExcelFile(file)
        else:
            # Nếu là file object từ upload
            excel_file = pd.ExcelFile(file)
        
        sheets_data = {}
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file, sheet_name=sheet_name)
            sheets_data[sheet_name] = df
        return sheets_data
    except Exception as e:
        if isinstance(file, str):
            st.error(f"Lỗi khi đọc file {file}: {str(e)}")
        else:
            st.error(f"Lỗi khi đọc file {file.name}: {str(e)}")
        return None

# Tự động load các file đã lưu khi khởi động (sau khi các hàm đã được định nghĩa)
if not st.session_state.files_loaded:
    # Load file Excel đã upload
    saved_files = load_saved_files()
    if saved_files:
        for file_name, sheets in saved_files.items():
            if file_name not in st.session_state.dataframes:
                st.session_state.dataframes[file_name] = sheets
    
    # Load bảng thông tin học sinh đã lưu
    saved_table = load_student_table()
    if saved_table:
        st.session_state.student_table = saved_table
    
    # Load dữ liệu tổng hợp mới nhất
    saved_combined = load_latest_combined_df()
    if saved_combined is not None:
        st.session_state.combined_df = saved_combined
    
    st.session_state.files_loaded = True

def combine_dataframes(dataframes_dict, existing_df=None):
    """Tổng hợp tất cả các dataframe từ nhiều file, có thể append vào dữ liệu hiện có"""
    combined_data = []
    
    # Nếu có dữ liệu hiện có, thêm vào danh sách
    if existing_df is not None and not existing_df.empty:
        combined_data.append(existing_df)
    
    # Thêm dữ liệu mới từ các file
    for file_name, sheets in dataframes_dict.items():
        for sheet_name, df in sheets.items():
            df_copy = df.copy()
            df_copy['Nguồn_File'] = file_name
            df_copy['Sheet'] = sheet_name
            combined_data.append(df_copy)
    
    if combined_data:
        # Loại bỏ trùng lặp dựa trên tất cả các cột (trừ Nguồn_File và Sheet có thể khác nhau)
        result = pd.concat(combined_data, ignore_index=True)
        # Có thể thêm logic loại bỏ trùng lặp nếu cần
        return result
    return None

def search_dataframe(df, search_columns, search_value, match_type='contains'):
    """Tra cứu dữ liệu trong dataframe"""
    if df is None or df.empty:
        return None
    
    try:
        results = pd.DataFrame()
        for col in search_columns:
            if col in df.columns:
                if match_type == 'contains':
                    mask = df[col].astype(str).str.contains(str(search_value), case=False, na=False)
                elif match_type == 'exact':
                    mask = df[col].astype(str).str.lower() == str(search_value).lower()
                elif match_type == 'starts_with':
                    mask = df[col].astype(str).str.lower().str.startswith(str(search_value).lower())
                elif match_type == 'ends_with':
                    mask = df[col].astype(str).str.lower().str.endswith(str(search_value).lower())
                else:
                    mask = pd.Series([False] * len(df))
                
                results = pd.concat([results, df[mask]], ignore_index=True)
        
        # Loại bỏ trùng lặp
        if not results.empty:
            results = results.drop_duplicates()
        
        return results
    except Exception as e:
        st.error(f"Lỗi khi tra cứu: {str(e)}")
        return None

def export_to_excel(df, filename='bao_cao'):
    """Xuất dataframe ra file Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Báo cáo')
    output.seek(0)
    return output.getvalue()

def export_to_csv(df):
    """Xuất dataframe ra file CSV"""
    return df.to_csv(index=False).encode('utf-8-sig')

def normalize_file_name(file_name):
    """Chuẩn hóa tên file: bỏ khoảng trắng, chuyển thành lowercase"""
    return str(file_name).replace(' ', '').replace('\t', '').lower().strip()

def find_student_info_by_file(df, name, file_name):
    """Tìm thông tin học sinh dựa trên tên và tên file Excel"""
    if df is None or df.empty:
        return None
    
    try:
        # Lọc theo tên file (cột Nguồn_File)
        if 'Nguồn_File' in df.columns:
            # Chuẩn hóa tên file input (bỏ khoảng trắng, lowercase)
            normalized_input = normalize_file_name(file_name)
            
            # Tìm file khớp (chuẩn hóa cả input và tên file trong database)
            def matches_file(nguon_file):
                normalized_file = normalize_file_name(nguon_file)
                return normalized_input in normalized_file or normalized_file in normalized_input
            
            file_mask = df['Nguồn_File'].astype(str).apply(matches_file)
            filtered_df = df[file_mask]
        else:
            filtered_df = df
        
        if filtered_df.empty:
            return None
        
        # Tìm các cột có thể chứa tên
        name_columns = [col for col in filtered_df.columns if any(keyword in str(col).lower() for keyword in ['tên', 'name', 'họ', 'ho ten', 'hoten'])]
        
        if not name_columns or not name:
            # Nếu không có tên, không trả về dữ liệu (không khả dụng)
            return None
        
        # Chuẩn hóa tên để so sánh
        name_normalized = str(name).strip().lower()
        name_parts = name_normalized.split()
        
        # Tìm kiếm dựa trên tên - ưu tiên khớp chính xác hơn
        best_match = None
        best_score = 0
        
        for idx, row in filtered_df.iterrows():
            score = 0
            for col in name_columns:
                if col in filtered_df.columns:
                    row_name = str(row[col]).strip().lower() if pd.notna(row[col]) else ''
                    
                    # Khớp chính xác hoàn toàn (điểm cao nhất)
                    if row_name == name_normalized:
                        score = 100
                        break
                    # Khớp chính xác từng từ (điểm cao)
                    elif name_normalized in row_name or row_name in name_normalized:
                        score = max(score, 80)
                    # Khớp chứa tất cả các từ trong tên
                    elif all(part in row_name for part in name_parts if len(part) > 2):
                        score = max(score, 60)
                    # Khớp một phần (điểm thấp)
                    elif any(part in row_name for part in name_parts if len(part) > 2):
                        score = max(score, 30)
            
            if score > best_score:
                best_score = score
                best_match = row
        
        # Nếu tìm thấy khớp tốt (score >= 30), trả về kết quả
        if best_match is not None and best_score >= 30:
            return best_match
        
        # Nếu không tìm thấy theo tên, trả về None (không trả về dòng đầu tiên nữa)
        return None
    except Exception as e:
        return None

def check_duplicate_student(student_table, name, khoa, ngay_sinh=''):
    """Kiểm tra trùng dữ liệu học sinh trong bảng"""
    if not student_table:
        return None, None
    
    name_normalized = str(name).strip().lower() if name else ''
    khoa_normalized = str(khoa).strip().lower() if khoa else ''
    ngay_sinh_normalized = str(ngay_sinh).strip().lower() if ngay_sinh else ''
    
    duplicates = []
    for idx, row in enumerate(student_table):
        row_name = str(row.get('Họ và tên', '')).strip().lower()
        row_khoa = str(row.get('Khoá', '')).strip().lower()
        row_ngay_sinh = str(row.get('Ngày sinh', '')).strip().lower()
        
        # Kiểm tra trùng tên và khóa
        if row_name == name_normalized and row_khoa == khoa_normalized:
            if row_ngay_sinh == ngay_sinh_normalized:
                # Trùng hoàn toàn
                duplicates.append({
                    'index': idx + 1,
                    'name': row.get('Họ và tên', ''),
                    'khoa': row.get('Khoá', ''),
                    'ngay_sinh': row.get('Ngày sinh', ''),
                    'type': 'exact'  # Trùng hoàn toàn
                })
            else:
                # Trùng tên + khóa nhưng khác ngày sinh
                duplicates.append({
                    'index': idx + 1,
                    'name': row.get('Họ và tên', ''),
                    'khoa': row.get('Khoá', ''),
                    'ngay_sinh': row.get('Ngày sinh', ''),
                    'type': 'different_dob'  # Trùng nhưng khác ngày sinh
                })
    
    if duplicates:
        # Tìm trùng hoàn toàn trước
        exact_duplicate = [d for d in duplicates if d['type'] == 'exact']
        if exact_duplicate:
            return 'exact', exact_duplicate[0]
        # Nếu không có trùng hoàn toàn, trả về trùng khác ngày sinh
        return 'different_dob', duplicates[0]
    
    return None, None

def is_date_format(value):
    """Kiểm tra xem giá trị có phải là định dạng ngày tháng (dd/mm/yyyy, dd-mm-yyyy, etc.)"""
    if pd.isna(value):
        return False
    value_str = str(value).strip()
    # Các pattern ngày tháng phổ biến
    date_patterns = [
        r'^\d{1,2}[/-]\d{1,2}[/-]\d{4}$',  # dd/mm/yyyy, dd-mm-yyyy
        r'^\d{4}[/-]\d{1,2}[/-]\d{1,2}$',  # yyyy/mm/dd, yyyy-mm-dd
        r'^\d{1,2}[/-]\d{1,2}[/-]\d{2}$',  # dd/mm/yy, dd-mm-yy
    ]
    for pattern in date_patterns:
        if re.match(pattern, value_str):
            return True
    return False

def is_all_digits(value):
    """Kiểm tra xem giá trị có phải là toàn số (dài 9-12 chữ số - CCCD)"""
    if pd.isna(value):
        return False
    value_str = str(value).strip().replace('.', '').replace(',', '').replace(' ', '')
    # CCCD thường có 9-12 chữ số
    if value_str.isdigit() and 9 <= len(value_str) <= 12:
        return True
    return False

def extract_khoa_from_filename(file_name):
    """Trích xuất khóa từ tên file (ví dụ: Bk16 từ bao cao 1- Bk16.xlsx)"""
    if not file_name:
        return ''
    
    file_str = str(file_name).strip()
    
    # Tìm pattern như Bk16, BK16, bk16 (chữ cái + số)
    pattern = r'([A-Za-z]+\d+)'
    matches = re.findall(pattern, file_str)
    
    if matches:
        # Lấy match cuối cùng (thường là khóa)
        khoa = matches[-1]
        # Viết hoa chữ cái đầu
        if len(khoa) > 1:
            khoa = khoa[0].upper() + khoa[1:].lower()
        return khoa
    
    # Nếu không tìm thấy, thử tìm số (ví dụ: 16, 2024)
    pattern_number = r'(\d{2,4})'
    matches_number = re.findall(pattern_number, file_str)
    if matches_number:
        return matches_number[-1]
    
    return ''

def capitalize_words(text):
    """Viết hoa chữ cái đầu của mỗi từ"""
    if not text or pd.isna(text):
        return ''
    
    text_str = str(text).strip()
    if not text_str or text_str.lower() == 'nan':
        return ''
    
    # Tách thành các từ và viết hoa chữ cái đầu
    words = text_str.split()
    capitalized_words = []
    for word in words:
        if len(word) > 0:
            capitalized_words.append(word[0].upper() + word[1:].lower() if len(word) > 1 else word.upper())
        else:
            capitalized_words.append(word)
    
    return ' '.join(capitalized_words)

def map_column_names(df):
    """Ánh xạ tên cột trong dataframe với tên cột chuẩn dựa trên tên và nội dung"""
    mapping = {}
    column_mapping = {
        'ngày sinh': ['ngày sinh', 'ngay sinh', 'date of birth', 'dob', 'sinh ngày'],
        'cccd': ['cccd', 'cmnd', 'chứng minh', 'chung minh', 'số cmnd', 'so cmnd', 'cmnd/cccd', 'so cmnd/cccd', 'số cccd', 'so cccd', 'căn cước'],
        'địa chỉ': ['địa chỉ', 'dia chi', 'address', 'địa điểm', 'dia diem', 'nơi ở', 'noi o'],
        'thầy': ['thầy', 'thay', 'giáo viên', 'giao vien', 'teacher', 'gv', 'người hướng dẫn', 'nguoi huong dan', 'giảng viên', 'giang vien', 'cô', 'co', 'thầy/cô', 'thay/co']
    }
    
    # Lấy sample dữ liệu để phân tích (lấy 100 dòng đầu hoặc tất cả nếu ít hơn)
    sample_size = min(100, len(df))
    sample_df = df.head(sample_size) if sample_size > 0 else df
    
    # Tìm cột ngày sinh - ưu tiên: tên cột có từ khóa > giá trị có định dạng ngày
    if 'ngày sinh' not in mapping:
        matched_cols_by_name = []
        for col in df.columns:
            if col in ['Nguồn_File', 'Sheet']:
                continue
            col_str = str(col).lower().strip()
            if any(keyword.lower().strip() in col_str for keyword in column_mapping['ngày sinh']):
                matched_cols_by_name.append(col)
        
        if matched_cols_by_name:
            mapping['ngày sinh'] = matched_cols_by_name[0]
        else:
            # Tìm theo nội dung - cột có nhiều giá trị định dạng ngày nhất
            best_col = None
            best_count = 0
            for col in df.columns:
                if col in ['Nguồn_File', 'Sheet']:
                    continue
                date_count = sample_df[col].apply(is_date_format).sum()
                if date_count > best_count and date_count > sample_size * 0.3:  # Ít nhất 30% là ngày
                    best_count = date_count
                    best_col = col
            if best_col:
                mapping['ngày sinh'] = best_col
    
    # Tìm cột CCCD - ưu tiên: tên cột có từ khóa > giá trị toàn số (9-12 chữ số)
    if 'cccd' not in mapping:
        matched_cols_by_name = []
        for col in df.columns:
            if col in ['Nguồn_File', 'Sheet']:
                continue
            col_str = str(col).lower().strip()
            if any(keyword.lower().strip() in col_str for keyword in column_mapping['cccd']):
                matched_cols_by_name.append(col)
        
        if matched_cols_by_name:
            mapping['cccd'] = matched_cols_by_name[0]
        else:
            # Tìm theo nội dung - cột có nhiều giá trị toàn số nhất
            best_col = None
            best_count = 0
            for col in df.columns:
                if col in ['Nguồn_File', 'Sheet', mapping.get('ngày sinh')]:
                    continue
                digit_count = sample_df[col].apply(is_all_digits).sum()
                if digit_count > best_count and digit_count > sample_size * 0.5:  # Ít nhất 50% là số
                    best_count = digit_count
                    best_col = col
            if best_col:
                mapping['cccd'] = best_col
    
    # Tìm cột địa chỉ - ưu tiên: cột ngay sau cột CCCD > tên cột có từ "địa chỉ" > cột có văn bản dài
    if 'địa chỉ' not in mapping:
        # Ưu tiên 1: Tìm cột ngay sau cột CCCD (nếu đã tìm được CCCD)
        if 'cccd' in mapping:
            cccd_col = mapping['cccd']
            # Lấy danh sách cột (loại trừ Nguồn_File, Sheet)
            valid_cols = [col for col in df.columns if col not in ['Nguồn_File', 'Sheet']]
            if cccd_col in valid_cols:
                cccd_idx = valid_cols.index(cccd_col)
                # Tìm cột ngay sau CCCD
                if cccd_idx + 1 < len(valid_cols):
                    next_col = valid_cols[cccd_idx + 1]
                    # Kiểm tra xem cột sau có phải là "thầy" không
                    next_col_str = str(next_col).lower().strip()
                    is_thay_col = any(keyword.lower().strip() in next_col_str for keyword in column_mapping['thầy'])
                    if not is_thay_col and next_col != cccd_col:
                        # Luôn lấy cột ngay sau CCCD (không kiểm tra độ dài)
                        mapping['địa chỉ'] = next_col
        
        # Nếu chưa tìm được, thử tìm theo tên cột
        if 'địa chỉ' not in mapping:
            matched_cols_by_name = []
            for col in df.columns:
                if col in ['Nguồn_File', 'Sheet', mapping.get('cccd')]:
                    continue
                col_str = str(col).lower().strip()
                # Kiểm tra xem cột có phải là "thầy" không (tránh nhầm lẫn)
                is_thay_col = any(keyword.lower().strip() in col_str for keyword in column_mapping['thầy'])
                if not is_thay_col and any(keyword.lower().strip() in col_str for keyword in column_mapping['địa chỉ']):
                    matched_cols_by_name.append(col)
            
            if matched_cols_by_name:
                mapping['địa chỉ'] = matched_cols_by_name[0]
        
        # Nếu vẫn chưa tìm được, tìm theo nội dung
        if 'địa chỉ' not in mapping:
            # Tìm theo nội dung - ưu tiên cột có địa danh (chữ và số dài)
            best_col = None
            best_score = 0
            for col in df.columns:
                if col in ['Nguồn_File', 'Sheet', mapping.get('ngày sinh'), mapping.get('cccd'), mapping.get('thầy')]:
                    continue
                col_str = str(col).lower().strip()
                # Kiểm tra xem cột có phải là "thầy" không
                is_thay_col = any(keyword.lower().strip() in col_str for keyword in column_mapping['thầy'])
                if not is_thay_col and df[col].dtype == 'object':
                    avg_length = sample_df[col].astype(str).str.len().mean()
                    if avg_length > 20:  # Trung bình > 20 ký tự
                        # Kiểm tra tỷ lệ chữ và số (địa danh thường có cả chữ và số)
                        sample_values = sample_df[col].astype(str).dropna().head(50)
                        if len(sample_values) > 0:
                            has_letters = sample_values.str.contains(r'[a-zA-ZÀ-ỹ]', na=False, regex=True).sum()
                            has_digits = sample_values.str.contains(r'\d', na=False, regex=True).sum()
                            # Điểm cao hơn nếu có cả chữ và số (địa danh)
                            score = avg_length
                            if has_letters > len(sample_values) * 0.3 and has_digits > len(sample_values) * 0.2:
                                score = avg_length * 1.5  # Tăng điểm nếu có cả chữ và số
                            else:
                                score = avg_length
                            
                            if score > best_score:
                                best_score = score
                                best_col = col
            if best_col:
                mapping['địa chỉ'] = best_col
    
    # Tìm cột thầy - chỉ dựa vào tên cột, không dựa vào nội dung
    if 'thầy' not in mapping:
        # Tìm cột có tên khớp với từ khóa "thầy"
        best_thay_col = None
        best_score = 0
        
        for col in df.columns:
            if col in ['Nguồn_File', 'Sheet', mapping.get('ngày sinh'), mapping.get('cccd'), mapping.get('địa chỉ')]:
                continue
            col_str = str(col).lower().strip()
            
            # Tính điểm dựa trên từ khóa khớp
            for keyword in column_mapping['thầy']:
                keyword_lower = keyword.lower().strip()
                if keyword_lower in col_str:
                    # Điểm cao hơn nếu khớp chính xác hơn
                    if col_str == keyword_lower:
                        score = 100
                    elif col_str.startswith(keyword_lower) or col_str.endswith(keyword_lower):
                        score = 80
                    else:
                        score = 50
                    
                    if score > best_score:
                        best_score = score
                        best_thay_col = col
                    break
        
        if best_thay_col:
            mapping['thầy'] = best_thay_col
    
    return mapping

# Sidebar - Upload files
with st.sidebar:
    st.header("Quản lý File")
    
    uploaded_files = st.file_uploader(
        "Chọn các file Excel",
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        if st.button("Tải lên & Xử lý", type="primary"):
            with st.spinner("Đang xử lý files..."):
                for file in uploaded_files:
                    if file.name not in st.session_state.dataframes:
                        # Đọc dữ liệu từ file
                        file_bytes = file.read()
                        file.seek(0)  # Reset file pointer
                        
                        sheets = load_excel_file(file)
                        if sheets:
                            # Lưu file vào disk
                            file_path = save_file_to_disk(file_bytes, file.name)
                            if file_path:
                                st.session_state.dataframes[file.name] = sheets
                                st.success(f"Đã tải và lưu: {file.name}")
                            else:
                                # Vẫn lưu vào session state nếu không lưu được vào disk
                                st.session_state.dataframes[file.name] = sheets
                                st.success(f"Đã tải: {file.name}")
    
    # Hiển thị danh sách file đã tải
    if st.session_state.dataframes:
        st.markdown("---")
        st.subheader("Files đã tải:")
        
        for file_name in list(st.session_state.dataframes.keys()):
            col_file, col_del = st.columns([4, 1])
            with col_file:
                st.write(f"• {file_name}")
            with col_del:
                if st.button("🗑️", key=f"del_{file_name}", help=f"Xóa {file_name}"):
                    # Xóa khỏi session state
                    del st.session_state.dataframes[file_name]
                    # Xóa khỏi disk
                    delete_saved_file(file_name)
                    # Reset combined_df nếu đang dùng
                    if st.session_state.combined_df is not None:
                        # Kiểm tra xem file bị xóa có trong combined_df không
                        if 'Nguồn_File' in st.session_state.combined_df.columns:
                            st.session_state.combined_df = st.session_state.combined_df[
                                st.session_state.combined_df['Nguồn_File'] != file_name
                            ]
                            if st.session_state.combined_df.empty or len(st.session_state.combined_df) == 0:
                                st.session_state.combined_df = None
                            else:
                                # Tạo lại combined_df nếu còn file khác
                                remaining_files = {k: v for k, v in st.session_state.dataframes.items() if k != file_name}
                                if remaining_files:
                                    st.session_state.combined_df = combine_dataframes(remaining_files)
                                else:
                                    st.session_state.combined_df = None
                    st.rerun()
        
        st.markdown("---")
        if st.button("Xóa tất cả", type="secondary"):
            # Xóa tất cả file khỏi disk
            for file_name in list(st.session_state.dataframes.keys()):
                delete_saved_file(file_name)
            
            # Xóa khỏi session state
            st.session_state.dataframes = {}
            st.session_state.combined_df = None
            st.session_state.search_results = None
            st.rerun()

# Main content
tab0, tab1, tab2, tab3, tab4 = st.tabs(["Bảng Thông tin", "Tổng hợp", "Tra cứu", "Thống kê", "Xuất báo cáo"])

# Tab 0: Bảng Thông tin
with tab0:
    st.header("Bảng Thông tin Học sinh/Sinh viên")
    
    if st.session_state.combined_df is not None:
        st.info("Nhập tên và tên file Excel để tự động điền thông tin. Đảm bảo đã tổng hợp dữ liệu ở tab 'Tổng hợp' trước.")
        
        # Tạo cột mapping
        col_mapping = map_column_names(st.session_state.combined_df)
        
        # Khởi tạo bảng nếu chưa có
        if len(st.session_state.student_table) == 0:
            st.session_state.student_table = [{
                'STT': 1,
                'Họ và tên': '',
                'Khoá': '',
                'Ngày sinh': '',
                'CCCD': '',
                'Địa chỉ': '',
                'Thầy': ''
            }]
        
        # Form để thêm/sửa dòng
        st.markdown("### Nhập thông tin")
        st.markdown("")  # Khoảng cách
        
        # Sử dụng st.form để có thể bấm Enter để submit
        with st.form("form_tim_them", clear_on_submit=False):
            col_form1, col_form2 = st.columns(2)
            
            with col_form1:
                new_name = st.text_input("Họ và tên:", key="new_name_form", placeholder="Ví dụ: Nguyễn Văn A")
                # Hiển thị lỗi trùng dữ liệu (nếu có)
                if 'duplicate_error' in st.session_state:
                    st.markdown(f"<div style='color: #dc3545; font-size: 0.9em; margin-top: -1rem; margin-bottom: 1rem;'>{st.session_state.duplicate_error}</div>", unsafe_allow_html=True)
                    del st.session_state.duplicate_error
            with col_form2:
                new_file = st.text_input("Khóa:", key="new_file_form", placeholder="Ví dụ: Bk16, B01K14", help="Nhập khóa (có thể có khoảng trắng hoặc viết hoa/thường). Bấm Enter để tự động tìm và thêm.")
            
            # Nút submit (sẽ được trigger khi bấm Enter)
            submitted = st.form_submit_button("Tìm và Thêm (hoặc bấm Enter)", type="primary", use_container_width=True)
        
        # Xử lý khi form được submit (bấm Enter hoặc bấm nút)
        if submitted:
            if new_file:
                # Tìm thông tin theo tên file (và tên nếu có)
                    found_info = find_student_info_by_file(st.session_state.combined_df, new_name, new_file)
                    
                    if found_info is not None:
                        # Lấy thông tin từ kết quả tìm được
                        ngay_sinh = ''
                        cccd = ''
                        dia_chi = ''
                        thay = ''
                        
                        # Lấy thông tin từ mapping cho ngày sinh
                        if 'ngày sinh' in col_mapping:
                            col_name = col_mapping['ngày sinh']
                            if col_name in found_info.index:
                                value = found_info[col_name]
                                if pd.notna(value):
                                    ngay_sinh = str(value).strip()
                                    if ngay_sinh.lower() == 'nan':
                                        ngay_sinh = ''
                        
                        # Lấy CCCD từ cột thứ 4 (sau khi loại trừ Nguồn_File, Sheet)
                        # Lấy danh sách cột hợp lệ (loại trừ Nguồn_File, Sheet)
                        valid_cols = [col for col in found_info.index if col not in ['Nguồn_File', 'Sheet']]
                        if len(valid_cols) >= 4:
                            col_cccd = valid_cols[3]  # Cột thứ 4 (index 3)
                            if col_cccd in found_info.index:
                                value = found_info[col_cccd]
                                if pd.notna(value):
                                    cccd = str(value).strip()
                                    if cccd.lower() != 'nan':
                                        cccd = cccd
                        
                        # Lấy Địa chỉ từ cột thứ 5 (sau khi loại trừ Nguồn_File, Sheet)
                        if len(valid_cols) >= 5:
                            col_dia_chi = valid_cols[4]  # Cột thứ 5 (index 4)
                            if col_dia_chi in found_info.index:
                                value = found_info[col_dia_chi]
                                if pd.notna(value):
                                    dia_chi = str(value).strip()
                                    if dia_chi.lower() != 'nan':
                                        dia_chi = dia_chi
                        
                        # Không tự động điền cột Thầy - để người dùng tự nhập
                        thay = ''
                        
                        # Lấy tên từ dữ liệu tìm được (nếu có)
                        display_name = new_name if new_name else ''
                        if not display_name and 'ngày sinh' in col_mapping:
                            # Thử lấy tên từ cột tên nếu có
                            name_cols = [col for col in found_info.index if any(kw in str(col).lower() for kw in ['tên', 'name', 'họ'])]
                            if name_cols:
                                display_name = str(found_info[name_cols[0]]) if pd.notna(found_info[name_cols[0]]) else ''
                        
                        # Lấy khóa trực tiếp từ input và viết hoa chữ cái đầu
                        khoa = new_file.strip() if new_file else ''
                        if khoa:
                            # Viết hoa chữ cái đầu, giữ nguyên phần còn lại
                            khoa = khoa[0].upper() + khoa[1:] if len(khoa) > 1 else khoa.upper()
                        
                        # Viết hoa chữ cái đầu cho các giá trị (trừ Thầy - không điền tự động)
                        display_name = capitalize_words(display_name if display_name else new_name)
                        dia_chi = capitalize_words(dia_chi)
                        
                        # Kiểm tra trùng dữ liệu
                        dup_type, dup_info = check_duplicate_student(
                            st.session_state.student_table, 
                            display_name, 
                            khoa, 
                            ngay_sinh
                        )
                        
                        if dup_type == 'exact':
                            # Trùng hoàn toàn
                            error_msg = f"⚠️ Trùng dữ liệu! Học viên '{display_name}' (Khóa: {khoa}, Ngày sinh: {ngay_sinh}) đã tồn tại ở dòng STT {dup_info['index']}."
                            st.session_state.duplicate_error = error_msg
                            st.error(error_msg)
                            st.rerun()
                        elif dup_type == 'different_dob':
                            # Trùng tên + khóa nhưng khác ngày sinh
                            error_msg = f"⚠️ Trùng dữ liệu! Học viên '{display_name}' (Khóa: {khoa}) đã tồn tại ở dòng STT {dup_info['index']} với ngày sinh: {dup_info['ngay_sinh']}. Ngày sinh hiện tại: {ngay_sinh if ngay_sinh else '(trống)'}."
                            st.session_state.duplicate_error = error_msg
                            st.error(error_msg)
                            st.rerun()
                        else:
                            # Không trùng, thêm dòng mới
                            new_stt = len(st.session_state.student_table) + 1
                            st.session_state.student_table.append({
                                'STT': new_stt,
                                'Họ và tên': display_name,
                                'Khoá': khoa,
                                'Ngày sinh': ngay_sinh,
                                'CCCD': cccd,
                                'Địa chỉ': dia_chi,
                                'Thầy': ''  # Không tự động điền - để người dùng tự nhập
                            })
                            # Tự động sao lưu
                            save_student_table()
                            st.success(f"Đã tìm thấy và thêm thông tin từ file {new_file}!")
                            st.rerun()
                    else:
                        # Không tìm thấy - chỉ báo, KHÔNG add vào bảng
                        st.warning(f"Không tìm thấy thông tin học viên '{new_name if new_name else '(không có tên)'}' với khóa {new_file}!")
                        st.info("Vui lòng kiểm tra lại tên học viên và khóa.")
            else:
                st.warning("Vui lòng nhập khóa!")
        
        # Các nút khác
        st.markdown("")  # Khoảng cách
        col_btn2, col_btn3 = st.columns(2)
        with col_btn2:
            if st.button("Thêm Dòng Trống"):
                new_stt = len(st.session_state.student_table) + 1
                st.session_state.student_table.append({
                    'STT': new_stt,
                    'Họ và tên': '',
                    'Khoá': '',
                    'Ngày sinh': '',
                    'CCCD': '',
                    'Địa chỉ': '',
                    'Thầy': ''
                })
                # Tự động sao lưu
                save_student_table()
                st.rerun()
        
        with col_btn3:
            if st.button("Xóa Tất cả"):
                st.session_state.student_table = []
                # Tự động sao lưu
                save_student_table()
                st.rerun()
        
        st.markdown("---")
        
        # Hiển thị bảng
        if st.session_state.student_table:
            st.markdown("### Bảng Thông tin")
            
            # Cập nhật STT
            for i, row in enumerate(st.session_state.student_table):
                row['STT'] = i + 1
            
            # Chuyển đổi sang DataFrame để hiển thị
            df_display = pd.DataFrame(st.session_state.student_table)
            
            # Loại bỏ cột not_found khỏi hiển thị (nếu có)
            display_cols = [col for col in df_display.columns if col != 'not_found']
            df_to_show = df_display[display_cols].copy()
            
            # Cấu hình cột cho data_editor - chỉ cho phép chỉnh sửa cột "Thầy"
            column_config = {}
            for col in df_to_show.columns:
                if col == 'Thầy':
                    column_config[col] = st.column_config.TextColumn(
                        col,
                        help="Có thể chỉnh sửa thông tin",
                        default=""
                    )
                elif col == 'STT':
                    column_config[col] = st.column_config.NumberColumn(
                        col,
                        disabled=True
                    )
                else:
                    column_config[col] = st.column_config.TextColumn(
                        col,
                        disabled=True
                    )
            
            # Hiển thị với data_editor để cho phép chỉnh sửa cột Thầy
            edited_df = st.data_editor(
                df_to_show,
                column_config=column_config,
                use_container_width=True,
                height=400,
                key="student_table_editor"
            )
            
            # Cập nhật lại session_state nếu có thay đổi
            if not edited_df.equals(df_to_show):
                # Cập nhật dữ liệu từ edited_df về session_state
                for idx in range(min(len(edited_df), len(st.session_state.student_table))):
                    if 'Thầy' in edited_df.columns:
                        new_value = edited_df.iloc[idx]['Thầy']
                        if pd.notna(new_value):
                            st.session_state.student_table[idx]['Thầy'] = str(new_value).strip()
                        else:
                            st.session_state.student_table[idx]['Thầy'] = ''
                # Tự động sao lưu
                save_student_table()
                st.rerun()
            
            # Chức năng cập nhật lại thông tin
            st.markdown("#### Cập nhật lại thông tin")
            st.markdown("")  # Khoảng cách
            col_update1, col_update2 = st.columns([3, 1])
            with col_update1:
                st.info("Chức năng này sẽ tự động tìm lại thông tin từ Excel và cập nhật lại các trường: Ngày sinh, CCCD, Địa chỉ (dựa trên tên và khóa).")
            with col_update2:
                if st.button("Cập nhật lại tất cả", type="primary"):
                    updated_count = 0
                    not_found_count = 0
                    
                    with st.spinner("Đang cập nhật thông tin..."):
                        for idx, row in enumerate(st.session_state.student_table):
                            student_name = row.get('Họ và tên', '').strip()
                            student_khoa = row.get('Khoá', '').strip()
                            
                            if not student_name or not student_khoa:
                                continue
                            
                            # Tìm file Excel dựa trên khóa
                            # Khóa thường chứa tên file (ví dụ: "Bk16" trong "bao cao 1- Bk16.xlsx")
                            found_info = None
                            
                            # Thử tìm trong tất cả các file (chuẩn hóa để bỏ khoảng trắng, không phân biệt hoa thường)
                            normalized_khoa = normalize_file_name(student_khoa)
                            for file_name in st.session_state.dataframes.keys():
                                normalized_file = normalize_file_name(file_name)
                                if normalized_khoa in normalized_file or normalized_file in normalized_khoa:
                                    # Tìm thông tin học viên trong file này
                                    found_info = find_student_info_by_file(
                                        st.session_state.combined_df, 
                                        student_name, 
                                        file_name
                                    )
                                    if found_info is not None:
                                        break
                            
                            # Nếu không tìm thấy theo khóa, thử tìm trong tất cả file
                            if found_info is None:
                                for file_name in st.session_state.dataframes.keys():
                                    found_info = find_student_info_by_file(
                                        st.session_state.combined_df, 
                                        student_name, 
                                        file_name
                                    )
                                    if found_info is not None:
                                        break
                            
                            if found_info is not None:
                                # Cập nhật thông tin từ Excel
                                # Lấy thông tin từ mapping cho ngày sinh
                                if 'ngày sinh' in col_mapping:
                                    col_name = col_mapping['ngày sinh']
                                    if col_name in found_info.index:
                                        value = found_info[col_name]
                                        if pd.notna(value):
                                            new_ngay_sinh = str(value).strip()
                                            if new_ngay_sinh.lower() != 'nan':
                                                st.session_state.student_table[idx]['Ngày sinh'] = new_ngay_sinh
                                
                                # Lấy CCCD từ cột thứ 4 (sau khi loại trừ Nguồn_File, Sheet)
                                valid_cols = [col for col in found_info.index if col not in ['Nguồn_File', 'Sheet']]
                                if len(valid_cols) >= 4:
                                    col_cccd = valid_cols[3]  # Cột thứ 4 (index 3)
                                    if col_cccd in found_info.index:
                                        value = found_info[col_cccd]
                                        if pd.notna(value):
                                            new_cccd = str(value).strip()
                                            if new_cccd.lower() != 'nan':
                                                st.session_state.student_table[idx]['CCCD'] = new_cccd
                                
                                # Lấy Địa chỉ từ cột thứ 5 (sau khi loại trừ Nguồn_File, Sheet)
                                if len(valid_cols) >= 5:
                                    col_dia_chi = valid_cols[4]  # Cột thứ 5 (index 4)
                                    if col_dia_chi in found_info.index:
                                        value = found_info[col_dia_chi]
                                        if pd.notna(value):
                                            new_dia_chi = str(value).strip()
                                            if new_dia_chi.lower() != 'nan':
                                                new_dia_chi = capitalize_words(new_dia_chi)
                                                st.session_state.student_table[idx]['Địa chỉ'] = new_dia_chi
                                
                                updated_count += 1
                            else:
                                not_found_count += 1
                    
                    # Tự động sao lưu
                    save_student_table()
                    
                    if updated_count > 0:
                        st.success(f"Đã cập nhật {updated_count} học viên!")
                    if not_found_count > 0:
                        st.warning(f"Không tìm thấy thông tin cho {not_found_count} học viên.")
                    st.rerun()
            
            st.markdown("---")
            
            # Chức năng xóa từng dòng
            st.markdown("#### Xóa dòng")
            if len(st.session_state.student_table) > 0:
                delete_col1, delete_col2 = st.columns([3, 1])
                with delete_col1:
                    delete_index = st.number_input(
                        "Nhập STT của dòng cần xóa:",
                        min_value=1,
                        max_value=len(st.session_state.student_table),
                        value=1,
                        step=1,
                        key="delete_index"
                    )
                with delete_col2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("Xóa dòng này", type="secondary"):
                        if 1 <= delete_index <= len(st.session_state.student_table):
                            st.session_state.student_table.pop(delete_index - 1)
                            # Tự động sao lưu
                            save_student_table()
                            st.success(f"Đã xóa dòng {delete_index}!")
                            st.rerun()
            
            # Tùy chọn xuất
            st.markdown("---")
            st.markdown("### Xuất báo cáo")
            export_df = pd.DataFrame(st.session_state.student_table)
            
            col_exp1, col_exp2 = st.columns(2)
            
            with col_exp1:
                file_data_excel = export_to_excel(export_df, f"bang_thong_tin_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
                # Tự động sao lưu dữ liệu xuất
                save_exported_data(export_df, 'excel')
                st.download_button(
                    label="Tải file Excel (.xlsx)",
                    data=file_data_excel,
                    file_name=f"bang_thong_tin_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col_exp2:
                file_data_csv = export_to_csv(export_df)
                # Tự động sao lưu dữ liệu xuất
                save_exported_data(export_df, 'csv')
                st.download_button(
                    label="Tải file CSV (.csv)",
                    data=file_data_csv,
                    file_name=f"bang_thong_tin_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
        else:
            st.info("Bảng trống. Hãy thêm dòng mới để bắt đầu!")
    elif st.session_state.dataframes:
        st.warning("Vui lòng tổng hợp dữ liệu ở tab 'Tổng hợp' trước khi sử dụng chức năng này!")
    else:
        st.info("Vui lòng tải lên các file Excel ở sidebar và tổng hợp dữ liệu trước!")

# Tab 1: Tổng hợp
with tab1:
    st.header("Tổng hợp Dữ liệu")
    
    if st.session_state.dataframes:
        if st.button("Tổng hợp Tất cả", type="primary"):
            with st.spinner("Đang tổng hợp dữ liệu..."):
                # Tổng hợp thêm vào dữ liệu hiện có (nếu có)
                st.session_state.combined_df = combine_dataframes(
                    st.session_state.dataframes, 
                    existing_df=st.session_state.combined_df
                )
                if st.session_state.combined_df is not None:
                    # Tự động sao lưu dữ liệu tổng hợp (tránh ghi đè)
                    save_combined_df(st.session_state.combined_df)
                    st.success(f"Đã tổng hợp {len(st.session_state.combined_df)} dòng dữ liệu và tự động sao lưu!")
        
        if st.session_state.combined_df is not None:
            st.subheader("Dữ liệu tổng hợp")
            st.markdown("")  # Khoảng cách
            
            # Hiển thị thông tin cơ bản
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Tổng số dòng", len(st.session_state.combined_df))
            with col2:
                st.metric("Tổng số cột", len(st.session_state.combined_df.columns))
            with col3:
                st.metric("Số file nguồn", st.session_state.combined_df['Nguồn_File'].nunique())
            
            st.markdown("")  # Khoảng cách
            
            # Lọc và hiển thị dữ liệu
            st.markdown("### Lọc dữ liệu")
            st.markdown("")  # Khoảng cách
            
            filter_col1, filter_col2 = st.columns(2)
            
            with filter_col1:
                file_filter = st.multiselect(
                    "Chọn file nguồn",
                    options=st.session_state.combined_df['Nguồn_File'].unique(),
                    default=st.session_state.combined_df['Nguồn_File'].unique()
                )
            
            with filter_col2:
                sheet_filter = st.multiselect(
                    "Chọn sheet",
                    options=st.session_state.combined_df['Sheet'].unique(),
                    default=st.session_state.combined_df['Sheet'].unique()
                )
            
            filtered_df = st.session_state.combined_df[
                (st.session_state.combined_df['Nguồn_File'].isin(file_filter)) &
                (st.session_state.combined_df['Sheet'].isin(sheet_filter))
            ]
            
            st.dataframe(filtered_df, use_container_width=True, height=400)
            
            # Tùy chọn hiển thị thêm
            with st.expander("Xem chi tiết cấu trúc dữ liệu"):
                st.write("**Thông tin cột:**")
                col_info = pd.DataFrame({
                    'Cột': filtered_df.columns,
                    'Kiểu dữ liệu': [str(dtype) for dtype in filtered_df.dtypes],
                    'Giá trị null': filtered_df.isnull().sum().values,
                    'Giá trị duy nhất': [filtered_df[col].nunique() for col in filtered_df.columns]
                })
                st.dataframe(col_info, use_container_width=True)
    else:
        st.info("Vui lòng tải lên các file Excel ở sidebar để bắt đầu!")

# Tab 2: Tra cứu
with tab2:
    st.header("Tra cứu Thông tin")
    
    if st.session_state.combined_df is not None:
        search_col1, search_col2 = st.columns([2, 1])
        
        with search_col1:
            search_value = st.text_input("Nhập từ khóa tra cứu:", placeholder="Ví dụ: tên, mã số, v.v.")
        
        with search_col2:
            match_type = st.selectbox(
                "Kiểu tìm kiếm:",
                options=['contains', 'exact', 'starts_with', 'ends_with'],
                format_func=lambda x: {
                    'contains': 'Chứa',
                    'exact': 'Chính xác',
                    'starts_with': 'Bắt đầu với',
                    'ends_with': 'Kết thúc bằng'
                }[x]
            )
        
        # Chọn cột để tra cứu
        available_columns = [col for col in st.session_state.combined_df.columns 
                           if col not in ['Nguồn_File', 'Sheet']]
        search_columns = st.multiselect(
            "Chọn cột để tra cứu:",
            options=available_columns,
            default=available_columns[:3] if len(available_columns) >= 3 else available_columns
        )
        
        if st.button("Tìm kiếm", type="primary"):
            if search_value and search_columns:
                with st.spinner("Đang tra cứu..."):
                    results = search_dataframe(
                        st.session_state.combined_df,
                        search_columns,
                        search_value,
                        match_type
                    )
                    st.session_state.search_results = results
                    
                    if results is not None and not results.empty:
                        st.success(f"Tìm thấy {len(results)} kết quả!")
                    else:
                        st.warning("Không tìm thấy kết quả nào!")
            else:
                st.warning("Vui lòng nhập từ khóa và chọn ít nhất một cột!")
        
        # Hiển thị kết quả tra cứu
        if st.session_state.search_results is not None and not st.session_state.search_results.empty:
            st.markdown("### Kết quả tra cứu")
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.dataframe(st.session_state.search_results, use_container_width=True, height=400)
            
            with col2:
                st.metric("Số kết quả", len(st.session_state.search_results))
                
                # Thống kê nhanh
                st.markdown("**Theo nguồn:**")
                source_counts = st.session_state.search_results['Nguồn_File'].value_counts()
                for source, count in source_counts.items():
                    st.write(f"• {source}: {count}")
    elif st.session_state.dataframes:
        st.info("Vui lòng tổng hợp dữ liệu trước ở tab 'Tổng hợp'!")
    else:
        st.info("Vui lòng tải lên các file Excel ở sidebar để bắt đầu!")

# Tab 3: Thống kê
with tab3:
    st.header("Thống kê & Phân tích")
    
    if st.session_state.combined_df is not None:
        st.subheader("Thống kê mô tả")
        
        # Chọn cột số để thống kê
        numeric_columns = st.session_state.combined_df.select_dtypes(include=['number']).columns.tolist()
        
        if numeric_columns:
            selected_numeric_col = st.selectbox("Chọn cột số để phân tích:", numeric_columns)
            
            if selected_numeric_col:
                col1, col2 = st.columns(2)
                
                with col1:
                    # Thống kê cơ bản
                    stats = st.session_state.combined_df[selected_numeric_col].describe()
                    st.markdown("**Thống kê cơ bản:**")
                    st.dataframe(stats)
                
                with col2:
                    # Biểu đồ phân bố
                    fig_hist = px.histogram(
                        st.session_state.combined_df,
                        x=selected_numeric_col,
                        nbins=30,
                        title=f"Phân bố {selected_numeric_col}"
                    )
                    st.plotly_chart(fig_hist, use_container_width=True)
        
        st.markdown("")  # Khoảng cách
        
        # Thống kê theo nhóm
        st.markdown("### Thống kê theo nhóm")
        st.markdown("")  # Khoảng cách
        
        # Lấy danh sách cột có sẵn
        available_cols = [col for col in st.session_state.combined_df.columns 
                         if col not in ['Nguồn_File', 'Sheet']]
        
        group_col1, group_col2 = st.columns(2)
        
        with group_col1:
            group_by = st.selectbox(
                "Nhóm theo:",
                options=['Nguồn_File', 'Sheet'] + available_cols,
                key='group_by'
            )
        
        with group_col2:
            if numeric_columns:
                agg_column = st.selectbox(
                    "Cột tính toán:",
                    options=numeric_columns,
                    key='agg_column'
                )
        
        if group_by and numeric_columns:
            if st.button("Tính toán", type="primary"):
                if agg_column:
                    grouped_stats = st.session_state.combined_df.groupby(group_by)[agg_column].agg([
                        'count', 'sum', 'mean', 'median', 'std'
                    ]).round(2)
                    grouped_stats.columns = ['Số lượng', 'Tổng', 'Trung bình', 'Trung vị', 'Độ lệch chuẩn']
                    st.dataframe(grouped_stats, use_container_width=True)
                    
                    # Biểu đồ cột
                    fig_bar = px.bar(
                        grouped_stats.reset_index(),
                        x=group_by,
                        y='Tổng',
                        title=f"Tổng {agg_column} theo {group_by}"
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)
    elif st.session_state.dataframes:
        st.info("Vui lòng tổng hợp dữ liệu trước ở tab 'Tổng hợp'!")
    else:
        st.info("Vui lòng tải lên các file Excel ở sidebar để bắt đầu!")

# Tab 4: Xuất báo cáo
with tab4:
    st.header("Xuất Báo cáo")
    
    if st.session_state.combined_df is not None:
        st.subheader("Chọn dữ liệu để xuất")
        
        export_option = st.radio(
            "Chọn dữ liệu:",
            options=['Tất cả dữ liệu tổng hợp', 'Kết quả tra cứu'],
            horizontal=True
        )
        
        if export_option == 'Tất cả dữ liệu tổng hợp':
            export_df = st.session_state.combined_df
        else:
            if st.session_state.search_results is not None and not st.session_state.search_results.empty:
                export_df = st.session_state.search_results
            else:
                st.warning("Không có kết quả tra cứu để xuất!")
                export_df = None
        
        if export_df is not None and not export_df.empty:
            st.info(f"Sẽ xuất {len(export_df)} dòng dữ liệu")
            
            export_format = st.selectbox(
                "Chọn định dạng:",
                options=['Excel (.xlsx)', 'CSV (.csv)']
            )
            
            filename = st.text_input(
                "Tên file (không cần đuôi):",
                value=f"bao_cao_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("Tải xuống", type="primary", use_container_width=True):
                    if export_format == 'Excel (.xlsx)':
                        file_data = export_to_excel(export_df, filename)
                        # Tự động sao lưu dữ liệu xuất
                        save_exported_data(export_df, 'excel')
                        st.download_button(
                            label="Tải file Excel",
                            data=file_data,
                            file_name=f"{filename}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        file_data = export_to_csv(export_df)
                        # Tự động sao lưu dữ liệu xuất
                        save_exported_data(export_df, 'csv')
                        st.download_button(
                            label="Tải file CSV",
                            data=file_data,
                            file_name=f"{filename}.csv",
                            mime="text/csv"
                        )
            
            with col2:
                st.markdown("### Xem trước dữ liệu")
                st.dataframe(export_df.head(100), use_container_width=True, height=300)
    elif st.session_state.dataframes:
        st.info("Vui lòng tổng hợp dữ liệu trước ở tab 'Tổng hợp'!")
    else:
        st.info("Vui lòng tải lên các file Excel ở sidebar để bắt đầu!")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666; padding: 1rem 0;'>"
    "Ứng dụng Tổng hợp & Tra cứu Excel | Powered by Streamlit"
    "</div>",
    unsafe_allow_html=True
)
