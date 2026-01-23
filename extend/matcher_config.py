import json
import os
from utils.path_utils import get_resource_path


class MatcherConfig:
    # 配置文件固定放在用户目录下，确保可写且不会随程序删除而消失
    _USER_DIR = os.path.join(os.path.expanduser("~"), ".cosmic_review")
    if not os.path.exists(_USER_DIR):
        os.makedirs(_USER_DIR, exist_ok=True)

    CONFIG_FILE = os.path.join(_USER_DIR, "matcher_config.json")

    @classmethod
    def load(cls):
        config = cls.get_defaults()
        if os.path.exists(cls.CONFIG_FILE):
            try:
                with open(cls.CONFIG_FILE, "r", encoding="utf-8") as f:
                    user_config = json.load(f)
                    # 深度更新配置
                    cls._deep_update(config, user_config)
            except Exception as e:
                print(f"Error loading config: {e}")

        # 统一路径格式
        if "storage" in config:
            for key, path in config["storage"].items():
                if path:
                    config["storage"][key] = os.path.normpath(path)

        return config

    @staticmethod
    def _deep_update(d, u):
        for k, v in u.items():
            if isinstance(v, dict) and k in d and isinstance(d[k], dict):
                MatcherConfig._deep_update(d[k], v)
            else:
                d[k] = v
        return d

    @classmethod
    def save(cls, config):
        try:
            with open(cls.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving config: {e}")

    @classmethod
    def get_defaults(cls):
        # 获取用户文档目录作为默认存放根目录
        user_docs = os.path.normpath(os.path.expanduser("~/Documents/CosmicReview"))
        return {
            "hierarchy": {
                "level1_col": 1,
                "level2_col": 2,
                "level3_col": 3,
                "sheet_name": 2,  # Default to 3rd sheet (index 2)
            },
            "process": {
                "column": 6,  # Default functional process column (7th column, index 6)
                "sheet_name": 2,  # Default to 3rd sheet (index 2)
            },
            "storage": {
                "initial_review": os.path.join(user_docs, "InitialReview"),
                "re_review": os.path.join(user_docs, "ReReview"),
                "receipt": os.path.join(user_docs, "Receipt"),
                "logs": os.path.join(user_docs, "Logs"),
            },
            "theme": {
                "primary_color": "#2563eb",
                "text_color_light": "#1e293b",
                "text_color_dark": "#f3f4f6",
                "font_family": "Microsoft YaHei UI",
                "font_size": 14,
                "is_dark": False,
            },
            "automation": {"auto_open": True},
        }
