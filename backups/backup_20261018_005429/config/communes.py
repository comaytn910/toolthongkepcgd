import json
import os
import re
import unicodedata

# Đường dẫn danh mục 92 xã
CONFIG_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(CONFIG_DIR)
DANHMUC_PATH = os.path.join(BASE_DIR, 'danhmuc_92_xa.json')

def load_canonical_communes():
    """Tải danh mục 92 xã chuẩn từ file danhmuc_92_xa.json"""
    if os.path.exists(DANHMUC_PATH):
        with open(DANHMUC_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

CANONICAL_COMMUNES = load_canonical_communes()

def normalize_text(text: str) -> str:
    """Chuẩn hóa chuỗi văn bản tiếng Việt để so khớp không phân biệt hoa thường/dấu"""
    if not text:
        return ""
    # Chuyển về dạng chuẩn Unicode NFC
    text = unicodedata.normalize('NFC', text)
    text = text.lower().strip()
    # Bỏ các tiền tố STT dạng "01.", "1.", "16. "
    text = re.sub(r'^\d+[\.\s_\-]+', '', text)
    # Chuẩn hóa khoảng trắng
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def match_commune_name(raw_name: str):
    """
    Nhận diện tên xã/phường từ tên file hoặc nội dung ô trong Excel
    Trả về: (stt, tên chuẩn, độ tin cậy) hoặc (None, None, 0)
    """
    if not raw_name:
        return None, None, 0
    
    norm_input = normalize_text(raw_name)
    
    # 1. So khớp chính xác
    for c in CANONICAL_COMMUNES:
        norm_canonical = normalize_text(c['raw_name'])
        if norm_input == norm_canonical or norm_canonical in norm_input:
            return c['stt'], c['raw_name'], 100
        
    # 2. Bỏ chữ "xã", "phường", "thị trấn" để so tên riêng
    clean_input = re.sub(r'^(xã|phường|thị trấn)\s+', '', norm_input).strip()
    for c in CANONICAL_COMMUNES:
        clean_canonical = re.sub(r'^(xã|phường|thị trấn)\s+', '', normalize_text(c['raw_name'])).strip()
        if clean_input and (clean_input == clean_canonical or clean_canonical in clean_input):
            return c['stt'], c['raw_name'], 90
            
    return None, None, 0
