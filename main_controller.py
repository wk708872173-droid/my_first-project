import subprocess
import os
import sys
import shutil
import time
from datetime import datetime

# --- 统一配置区 ---
TOOL_DIR = r'C:\Users\10174306\Tool'
INPUT_ROOT = r'D:\Input'   # 所有的输入文件夹现在都在这里
ARCHIVE_ROOT = r'D:\Archive' # 处理完后的存档处

# 脚本与对应子文件夹的映射
JOBS = [
    {"script": "go_store_update.py", "folder": "master_store"},
    {"script": "emplloyee_update.py", "folder": "master_employee"},
    {"script": "product_mst_update.py", "folder": "product_mst"},
    {"script": "系统推荐数.py", "folder": "shw_orderqty_data"},
    {"script": "sht_order_detail_append.py", "folder": "sht_order_detail"},
    {"script": "purchase_update.py", "folder": "purchase"},
    {"script": "trans.py", "folder": "trans"},
    {"script": "loss_update.py", "folder": "loss"},
    {"script": "goma_discard_update.py", "folder": "goma_discard"},
    {"script": "キャパ数調整_データ加工.py", "folder": "capacity_adjustment_data"},
    {"script": "data_process_refactored_v2.py", "folder": None}
]

def run_job(job):
    script_name = job["script"]
    folder_name = job["folder"]
    script_path = os.path.join(TOOL_DIR, script_name)
    target_input = os.path.join(INPUT_ROOT, folder_name) if folder_name else None

    print(f"\n🚀 执行脚本: {script_name}...")
    
    try:
        env = os.environ.copy()
        if target_input:
            env["DYNAMIC_INPUT_PATH"] = target_input

        # 使用 subprocess.run 会阻塞主线程，直到子脚本完全退出
        result = subprocess.run([sys.executable, script_path], check=True, env=env)
        
        # --- 重点修改：子脚本运行成功后再进行归档 ---
        if result.returncode == 0:
            if target_input and os.path.exists(target_input):
                # 增加 2 秒延迟，确保文件句柄已完全释放
                time.sleep(2) 
                archive_processed_files(folder_name, target_input)
            return True
        else:
            return False
    except Exception as e:
        print(f"💥 脚本 {script_name} 内部发生错误: {e}")
        return False

def archive_processed_files(folder_name, src_dir):
    """只移动 Excel 原始文件，保留生成的中间 CSV（或按需清理）"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    dst_dir = os.path.join(ARCHIVE_ROOT, folder_name, timestamp)
    
    # 修改：只查找原始 Excel 文件进行归档，不移动生成的临时 CSV
    files = [f for f in os.listdir(src_dir) if f.lower().endswith(('.xlsx', '.xls', '.csv'))]
    if not files: return

    os.makedirs(dst_dir, exist_ok=True)
    for f in files:
        src_path = os.path.join(src_dir, f)
        dst_path = os.path.join(dst_dir, f)
        try:
            # 使用 copy + delete 替代 move，防止跨盘符或文件占用导致的崩溃
            shutil.copy2(src_path, dst_path)
            os.remove(src_path)
        except Exception as e:
            print(f"⚠️ 归档文件 {f} 失败（可能正在被占用）: {e}")
            
    print(f"📦 原始文件已移至存档: {dst_dir}")

def main():
    if not os.path.exists(ARCHIVE_ROOT): os.makedirs(ARCHIVE_ROOT)
        
    print("="*60)
    print("🌟 自动化链路管理系统 - 动态路径模式 🌟")
    for job in JOBS:
        if not run_job(job) and "master" in job["script"]:
            print("🛑 关键基础表更新失败，任务终止。")
            break
    print("="*60)

if __name__ == "__main__":
    main()