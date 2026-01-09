import json
import os
from utils.path_utils import get_resource_path

class MatcherConfig:
    # 使用 get_resource_path 获取配置文件路径
    CONFIG_FILE = get_resource_path("matcher_config.json")
    
    @classmethod
    def load(cls):
        if os.path.exists(cls.CONFIG_FILE):
            try:
                with open(cls.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading config: {e}")
        return cls.get_defaults()
    
    @classmethod
    def save(cls, config):
        try:
            with open(cls.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving config: {e}")
            
    @classmethod
    def get_defaults(cls):
        return {
            "hierarchy": {
                "level1_col": 1,
                "level2_col": 2,
                "level3_col": 3,
                "sheet_name": 2 # Default to 3rd sheet (index 2)
            },
            "process": {
                "column": 6, # Default functional process column (7th column, index 6)
                "sheet_name": 2 # Default to 3rd sheet (index 2)
            }
        }
