"""
Logger Module
アプリケーション全体のログ管理を行うモジュール

主な機能:
- メソッドの実行開始/終了のログ記録
- GPTプロンプトとレスポンスのログ
- 領域検出と処理のログ
- エラー処理とスタックトレースの記録
- デバッグ情報の出力制御
"""

import logging
import json
from datetime import datetime

class Logger:
    def __init__(self):
        # watchdogのデバッグログを無効化
        logging.getLogger('watchdog.observers.inotify_buffer').setLevel(logging.WARNING)

        # 基本設定
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s [%(levelname)s] %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler('extraction.log', encoding='utf-8')
            ]
        )
        self.logger = logging.getLogger('ExcelMetadataExtractor')

    def method_start(self, method_name):
        """メソッドの開始をログに記録"""
        self.logger.info(f"============== Method Start: {method_name} ==============")

    def method_end(self, method_name):
        """メソッドの終了をログに記録"""
        self.logger.info(f"============== Method End: {method_name} ==============")

    def gpt_prompt(self, prompt):
        """GPTプロンプトをログに記録"""
        self.logger.info(f"GPT Prompt:\n{prompt}")

    def gpt_response(self, response):
        """GPTレスポンスをログに記録"""
        self.logger.info(f"GPT Response:\n{json.dumps(response, ensure_ascii=False, indent=2)}")

    def region_detected(self, region_type, range_str):
        """領域検出をログに記録"""
        self.logger.info(f"[Region Detected] Type: {region_type}, Range: {range_str}")

    def processing_region(self, region_type, range_str):
        """領域処理をログに記録"""
        self.logger.info(f"[Processing Region] Type: {region_type}, Range: {range_str}")

    def start_region_processing(self, region_info):
        """領域処理の開始をログに記録"""
        region_type = region_info.get("type", "unknown")
        range_str = region_info.get("range", "N/A")
        self.logger.info(f">>>>> Start Processing Region - Type: {region_type}, Range: {range_str}")

    def end_region_processing(self, region_info):
        """領域処理の終了をログに記録"""
        region_type = region_info.get("type", "unknown")
        range_str = region_info.get("range", "N/A")
        self.logger.info(f"<<<<< End Processing Region - Type: {region_type}, Range: {range_str}")

    def info(self, message):
        """一般情報をログに記録"""
        self.logger.info(message)

    def error(self, message, error=None):
        """エラー情報をログに記録"""
        if error:
            self.logger.error(f"{message}: {str(error)}")
        else:
            self.logger.error(message)

    def exception(self, error):
        """例外のスタックトレースを記録"""
        self.logger.exception(error)

    def debug(self, message):
        """デバッグ情報をログに記録"""
        self.logger.debug(message)

    def debug_region(self, row, col, value, region_type=None):
        """領域のデバッグ情報をログに記録"""
        self.logger.info(f"Processing cell at row={row}, col={col}, value={value}, detected_type={region_type}")

    def debug_boundaries(self, start_row, start_col, max_row, max_col):
        """領域境界のデバッグ情報をログに記録"""
        self.logger.info(f"Region boundaries: ({start_row},{start_col}) to ({max_row},{max_col})")