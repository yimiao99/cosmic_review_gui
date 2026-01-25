import os
import sys
from datetime import datetime


class RuntimeLogger:
    """运行日志管理器，用于记录并保存详细的校验过程"""

    _instance = None
    _logs = []
    _original_stdout = sys.stdout
    _original_stderr = sys.stderr
    _current_project = None  # 当前项目名称
    _current_log_filename = None  # 当前运行会话的日志文件名

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
            # 初始化时重定向一次即可
            cls._setup_redirection()
        return cls._instance

    @classmethod
    def _setup_redirection(cls):
        """重定向 stdout 和 stderr"""
        if not isinstance(sys.stdout, cls.StreamToLogger):
            sys.stdout = cls.StreamToLogger(cls._original_stdout, "INFO")
            sys.stderr = cls.StreamToLogger(cls._original_stderr, "ERROR")

    @classmethod
    def log(cls, message, level="INFO"):
        """记录日志，同时写入当前日期的日志文件"""
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = f"[{timestamp}] [{level}] {message}"
        cls._logs.append(log_entry)

        # 实时写入文件
        cls._write_to_file(log_entry)

    @classmethod
    def get_current_log_path(cls):
        """获取当前会话生成的日志文件路径"""
        from extend.matcher_config import MatcherConfig

        config = MatcherConfig.load()
        output_dir = config.get("storage", {}).get("logs", "logs")
        if not os.path.exists(output_dir):
            return None
        if not cls._current_log_filename:
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

            # 如果还没有生成本次运行的文件名，则生成一个
            if not cls._current_log_filename:
                time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                if cls._current_project:
                    # 清理项目名称中的非法字符
                    safe_name = "".join(
                        [c for c in cls._current_project if c not in '<>:"/\\|?*']
                    ).strip()
                    cls._current_log_filename = f"{safe_name}_{time_str}.log"
                else:
                    cls._current_log_filename = f"Runtime_{time_str}.log"

            target_path = os.path.join(output_dir, cls._current_log_filename)

            with open(target_path, "a", encoding="utf-8") as f:
                f.write(log_entry + "\n")
        except Exception as e:
            if cls._original_stdout:
                try:
                    cls._original_stdout.write(f"写入日志文件失败: {e}\n")
                except:
                    pass

    @classmethod
    def clear(cls):
        """清空内存中的日志缓存 (不影响已写入文件的部分)"""
        cls._logs = []

    @classmethod
    def start_session(cls):
        """向下兼容，现在自动调用初始化"""
        cls()
        cls.log("=== CosmicReviewGUI 运行开始 ===")

    @classmethod
    def save_session(cls):
        """向下兼容，不再需要手动保存整个会话"""
        return None

    @classmethod
    def save_to_file(cls, filename, output_dir=None):
        """向下兼容"""
        return None

    @classmethod
    def get_logs(cls):
        return "\n".join(cls._logs)

    @classmethod
    def set_project(cls, project_name):
        """设置当前项目名称，用于后续日志文件名的生成"""
        cls._current_project = project_name
        cls._current_log_filename = (
            None  # 重置文件名，以便下次写入时生成包含项目名的新文件
        )
        cls._logs = []  # 切换项目时清空内存日志
        if project_name:
            cls.log(f"=== 项目 [{project_name}] 日志开始 ===")
