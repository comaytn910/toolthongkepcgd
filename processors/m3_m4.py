"""
Bộ xử lý dữ liệu Biểu 3.M4: Hiện trạng mù chữ Mức 1 (15-25 tuổi)
"""
from typing import Dict, List, Any, Optional
from processors.m3_mu_chu import process_3_mu_chu_single_file, aggregate_3_mu_chu

def process_3_m4_single_file(file_path: str, commune_name: str) -> Optional[Dict[str, Any]]:
    return process_3_mu_chu_single_file(file_path, commune_name, '3.M4')

def aggregate_3_m4(communes_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    return aggregate_3_mu_chu(communes_data, '3.M4')
