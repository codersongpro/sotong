"""Persistent application configuration."""

import json
import logging
import os


CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".chungbuk_auto_config.json")


class Config:
    DEFAULTS = {
        'search_field_x': None, 'search_field_y': None,
        'result_first_x': None, 'result_first_y': None,
        'add_button_x':   None, 'add_button_y':   None,
        'empty_pixel_rgb': None,
        'search_delay': 0.5,
        'manual_confirm': True,
    }

    def __init__(self):
        self.data = dict(self.DEFAULTS)
        self._load()

    def _load(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    self.data.update(json.load(f))
        except Exception as exc:
            logging.warning("설정 파일을 읽지 못했습니다: %s", exc)

    def save(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def is_calibrated(self):
        return (self.data.get('search_field_x') is not None and
                self.data.get('search_field_y') is not None and
                self.data.get('result_first_x') is not None and
                self.data.get('result_first_y') is not None and
                self.data.get('add_button_x')   is not None and
                self.data.get('add_button_y')   is not None)

