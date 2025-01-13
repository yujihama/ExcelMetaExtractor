
import logging
import json
from datetime import datetime

class Logger:
    def __init__(self):
        # watchdogのデバッグログを無効化
        logging.getLogger('watchdog.observers.inotify_buffer').setLevel(logging.WARNING)
        
        # 基本設定
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler('extraction.log', encoding='utf-8')
            ]
        )
        self.logger = logging.getLogger('ExcelMetadataExtractor')

    def method_start(self, method_name):
        self.logger.info(f"============== Method Start: {method_name} ==============")

    def method_end(self, method_name):
        self.logger.info(f"============== Method End: {method_name} ==============")

    def gpt_prompt(self, prompt):
        self.logger.info(f"GPT Prompt:\n{prompt}")

    def gpt_response(self, response):
        self.logger.info(f"GPT Response:\n{json.dumps(response, ensure_ascii=False, indent=2)}")

    def region_detected(self, region_type, range_str):
        self.logger.info(f"[Region Detected] Type: {region_type}, Range: {range_str}")

    def processing_region(self, region_type, range_str):
        self.logger.info(f"[Processing Region] Type: {region_type}, Range: {range_str}")

    def start_region_processing(self, region_info):
        region_type = region_info.get("type", "unknown")
        range_str = region_info.get("range", "N/A")
        self.logger.info(f">>>>> Start Processing Region - Type: {region_type}, Range: {range_str}")

    def end_region_processing(self, region_info):
        region_type = region_info.get("type", "unknown")
        range_str = region_info.get("range", "N/A")
        self.logger.info(f"<<<<< End Processing Region - Type: {region_type}, Range: {range_str}")

    def info(self, message):
        self.logger.info(message)

    def error(self, message, error=None):
        if error:
            self.logger.error(f"{message}: {str(error)}")
        else:
            self.logger.error(message)

    def debug(self, message):
        self.logger.debug(message)
        
    def debug_region(self, row, col, value, region_type=None):
        self.logger.info(f"Processing cell at row={row}, col={col}, value={value}, detected_type={region_type}")

    def debug_boundaries(self, start_row, start_col, max_row, max_col):
        self.logger.info(f"Region boundaries: ({start_row},{start_col}) to ({max_row},{max_col})")
