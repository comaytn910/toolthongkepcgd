import os
import sys
import json
import glob
from datetime import datetime

def get_data_state(data_dir: str):
    """
    Quét thư mục data/ và trả về trạng thái hiện tại của tất cả các file nguồn.
    """
    files = glob.glob(os.path.join(data_dir, "*.xlsx"))
    state = {
        'count': len(files),
        'files': {}
    }
    max_mtime = 0.0
    for f in sorted(files):
        fname = os.path.basename(f)
        try:
            stat = os.stat(f)
            mtime = stat.st_mtime
            size = stat.st_size
            if mtime > max_mtime:
                max_mtime = mtime
            state['files'][fname] = {
                'size': size,
                'mtime': mtime,
                'mtime_str': datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            }
        except Exception:
            pass
            
    state['latest_mtime'] = max_mtime
    state['latest_mtime_str'] = datetime.fromtimestamp(max_mtime).strftime("%Y-%m-%d %H:%M:%S") if max_mtime > 0 else ""
    return state

def save_sync_metadata(metadata_path: str, data_dir: str, summary_info: dict = None):
    """
    Lưu metadata sau khi tổng hợp thành công toàn tỉnh.
    """
    cur_state = get_data_state(data_dir)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    now_ts = datetime.now().timestamp()
    
    meta = {
        'last_sync_time': now_str,
        'last_sync_timestamp': now_ts,
        'source_file_count': cur_state['count'],
        'latest_source_mtime': cur_state['latest_mtime'],
        'latest_source_mtime_str': cur_state['latest_mtime_str'],
        'summary': summary_info or {},
        'files_snapshot': cur_state['files']
    }
    
    os.makedirs(os.path.dirname(metadata_path), exist_ok=True)
    with open(metadata_path, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
        
    return meta

def check_data_freshness(metadata_path: str, data_dir: str):
    """
    Kiểm tra độ mới của kết quả so với dữ liệu nguồn trong data/.
    Trả về dict:
      - status: 'FRESH' | 'STALE' | 'NO_METADATA'
      - message: thông điệp hiển thị cho người dùng
      - changed_files: danh sách file nguồn có mtime mới hơn lần tổng hợp gần nhất
    """
    if not os.path.exists(metadata_path):
        return {
            'status': 'NO_METADATA',
            'is_fresh': False,
            'message': 'Chưa có thông tin tổng hợp lưu sẵn. Vui lòng thực hiện tổng hợp.',
            'last_sync_time': None,
            'changed_files': []
        }
        
    try:
        with open(metadata_path, 'r', encoding='utf-8') as f:
            meta = json.load(f)
    except Exception as e:
        return {
            'status': 'NO_METADATA',
            'is_fresh': False,
            'message': f'Lỗi đọc metadata: {e}',
            'last_sync_time': None,
            'changed_files': []
        }
        
    last_sync_ts = meta.get('last_sync_timestamp', 0.0)
    last_sync_str = meta.get('last_sync_time', 'Chưa xác định')
    saved_files = meta.get('files_snapshot', {})
    
    cur_state = get_data_state(data_dir)
    changed_files = []
    
    # 1. Kiểm tra số lượng file
    if cur_state['count'] != meta.get('source_file_count', 0):
        changed_files.append(f"Số lượng file thay đổi: Hiện có {cur_state['count']} file, trước đó tổng hợp {meta.get('source_file_count')} file.")
        
    # 2. Kiểm tra từng file
    for fname, info in cur_state['files'].items():
        if fname not in saved_files:
            changed_files.append(f"File mới: {fname}")
        else:
            saved_info = saved_files[fname]
            # Cho phép dung sai 1 giây
            if info['mtime'] > saved_info.get('mtime', 0.0) + 1.0 or info['size'] != saved_info.get('size', 0):
                changed_files.append(f"{fname} (Cập nhật lúc: {info['mtime_str']})")
                
    # 3. Kết luận
    if changed_files:
        return {
            'status': 'STALE',
            'is_fresh': False,
            'message': f'CẢNH BÁO: Dữ liệu nguồn trong data/ đã có thay đổi mới hơn lần tổng hợp lúc {last_sync_str}! Kết quả hiện tại đang CŨ.',
            'last_sync_time': last_sync_str,
            'changed_files': changed_files,
            'summary': meta.get('summary', {})
        }
    else:
        return {
            'status': 'FRESH',
            'is_fresh': True,
            'message': f'Kết quả tổng hợp khớp 100% với dữ liệu nguồn mới nhất. (Thời điểm tổng hợp: {last_sync_str} - {cur_state["count"]} xã)',
            'last_sync_time': last_sync_str,
            'changed_files': [],
            'summary': meta.get('summary', {})
        }
