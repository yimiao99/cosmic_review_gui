import os
import sys
from datetime import datetime


class RuntimeLogger:
    """运行日志管理器，用于记录并保存详细的校验过程"""
    _instance = None
    _logs = []
    _original_stdout = sys.stdout
    _original_stderr = sys.stderr
    _current_project = None
    _current_log_filename = None

    # 【新增】日志级别控制与文件句柄缓存
    _log_level = "INFO"  # 默认级别: NONE, ERROR, WARN, INFO, DEBUG
    _file_handle = None
    _level_map = {"NONE": 0, "ERROR": 1, "WARN": 2, "INFO": 3, "DEBUG": 4}

    class StreamToLogger:
        def __init__(self, original_stream, level="INFO"):
            self.original_stream = original_stream
            self.level = level
            self.line_buffer = ""

        def write(self, buf):
            for line in buf.splitlines(True):
                if self.original_stream:
                    try:
                        self.original_stream.write(line)
                    except:
                        pass
                self.line_buffer += line
                if self.line_buffer.endswith("\n"):
                    msg = self.line_buffer.rstrip("\n")
                    if msg.strip():
                        RuntimeLogger.log(msg, self.level)
                    self.line_buffer = ""

        def flush(self):
            if self.original_stream:
                try:
                    self.original_stream.flush()
                except:
                    pass

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(RuntimeLogger, cls).__new__(cls)
            cls._logs = []
            cls._setup_redirection()
        return cls._instance

    @classmethod
    def _setup_redirection(cls):
        if not isinstance(sys.stdout, cls.StreamToLogger):
            sys.stdout = cls.StreamToLogger(cls._original_stdout, "INFO")
            sys.stderr = cls.StreamToLogger(cls._original_stderr, "ERROR")

    @classmethod
    def set_log_level(cls, level: str):
        """允许外部（如UI设置）动态设置日志级别"""
        old_level = cls._log_level
        cls._log_level = level.upper() if level else "INFO"

        # 【调试】打印级别变更
        print(f"[LOG-LEVEL-CHANGE] 日志级别从 {old_level} 变更为 {cls._log_level}")
        print(f"[LOG-LEVEL-CHANGE] _level_map值: {cls._level_map.get(cls._log_level, 3)}")

    @classmethod
    def log(cls, message, level="INFO"):
        """记录日志，增加级别过滤"""
        # 【核心优化1】级别过滤，直接拦截不需要输出的日志
        if cls._level_map.get(level, 3) > cls._level_map.get(cls._log_level, 3):
            return

        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = f"[{timestamp}] [{level}] {message}"
        cls._logs.append(log_entry)
        cls._write_to_file(log_entry)

    @classmethod
    def get_current_log_path(cls):
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        output_dir = config.get("storage", {}).get("logs", "logs")
        if not os.path.exists(output_dir) or not cls._current_log_filename:
            return None
        return os.path.abspath(os.path.join(output_dir, cls._current_log_filename))

    @classmethod
    def _write_to_file(cls, log_entry):
        try:
            from extend.matcher_config import MatcherConfig
            config = MatcherConfig.load()
            output_dir = config.get("storage", {}).get("logs", "logs")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            if not cls._current_log_filename:
                time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                if cls._current_project:
                    safe_name = "".join([c for c in cls._current_project if c not in '<>:"/\\|?*']).strip()
                    cls._current_log_filename = f"{safe_name}_{time_str}.log"
                else:
                    cls._current_log_filename = f"Runtime_{time_str}.log"

            target_path = os.path.join(output_dir, cls._current_log_filename)

            # 【核心优化2】保持文件句柄打开，避免百万次 open/close
            if cls._file_handle is None or cls._file_handle.name != target_path:
                if cls._file_handle:
                    cls._file_handle.close()
                cls._file_handle = open(target_path, "a", encoding="utf-8")

            cls._file_handle.write(log_entry + "\n")
            # 每100条 flush 一次，平衡性能与数据安全
            if len(cls._logs) % 100 == 0:
                cls._file_handle.flush()

        except Exception as e:
            if cls._original_stdout:
                try:
                    cls._original_stdout.write(f"写入日志文件失败: {e}\n")
                except:
                    pass

    @classmethod
    def close_file(cls):
        """任务结束时调用，关闭文件句柄释放资源"""
        if cls._file_handle:
            cls._file_handle.close()
            cls._file_handle = None

    @classmethod
    def clear(cls):
        cls._logs = []

    @classmethod
    def start_session(cls):
        cls()
        cls.log("=== CosmicReviewGUI 运行开始 ===")

    @classmethod
    def save_session(cls):
        return None

    @classmethod
    def save_to_file(cls, filename, output_dir=None):
        return cls.get_current_log_path()

    @classmethod
    def get_logs(cls):
        return "\n".join(cls._logs)

    @classmethod
    def set_project(cls, project_name):
        cls._current_project = project_name
        cls._current_log_filename = None
        cls._logs = []
        if project_name:
            cls.log(f"=== 项目 [{project_name}] 日志开始 ===")