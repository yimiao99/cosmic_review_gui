import os
from datetime import datetime

class RuntimeLogger:
    """运行日志管理器，用于记录并保存详细的校验过程"""
    
    _instance = None
    _logs = []
    _start_time = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(RuntimeLogger, cls).__new__(cls)
            cls._logs = []
            cls._start_time = datetime.now()
        return cls._instance

    @classmethod
    def log(cls, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = f"[{timestamp}] [{level}] {message}"
        print(log_entry) # 同时输出到控制台
        cls._logs.append(log_entry)

    @classmethod
    def clear(cls):
        cls._logs = []
        cls._start_time = datetime.now()

    @classmethod
    def save_to_file(cls, filename, output_dir="."):
        """保存日志到文件，返回绝对路径"""
        try:
            # 净化文件名
            clean_name = filename.replace(" ", "_").replace("/", "").replace("\\", "")
            base_name = f"RunLog_{clean_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            target_path = os.path.abspath(os.path.join(output_dir, base_name))
            
            with open(target_path, "w", encoding="utf-8") as f:
                f.write(f"=== CosmicReviewGUI 运行日志 ===\n")
                f.write(f"项目名称: {filename}\n")
                f.write(f"开始时间: {cls._start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"保存时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"{'='*30}\n\n")
                f.write("\n".join(cls._logs))
                f.write(f"\n\n{'='*30}\n日志结束。\n")
                
            return target_path
        except Exception as e:
            print(f"保存日志失败: {e}")
            return None

    @classmethod
    def get_logs(cls):
        return "\n".join(cls._logs)
