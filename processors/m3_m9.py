"""
Bộ xử lý dữ liệu Biểu 3.M9: Hiện trạng mù chữ Mức 2 (15-60 tuổi)
"""
from typing import Dict, List, Any, Optional
from processors.m3_mu_chu import process_3_mu_chu_single_file, aggregate_3_mu_chu

def process_3_m9_single_file(file_path: str, commune_name: str) -> Optional[Dict[str, Any]]:
    return process_3_mu_chu_single_file(file_path, commune_name, '3.M9')

def aggregate_3_m9(communes_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    return aggregate_3_mu_chu(communes_data, '3.M9')
