import os
import re
import io
import subprocess
import gc
import multiprocessing
import hashlib
import tempfile
import concurrent.futures
import pandas as pd
from datetime import datetime, timedelta, time as dt_time
import logging
import sys
import shutil
import json
from io import BytesIO
from typing import List, Dict, Set, Optional, Tuple
import warnings
import functools
import contextlib
import time
import threading
import ipaddress
import socket
import uuid
import smtplib
import zipfile
import ssl
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import queue
import requests
import psutil

try:
    import keyring
except ImportError:
    keyring = None

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import pdfplumber
from docx import Document
import openpyxl
import olefile
from openpyxl.utils.exceptions import InvalidFileException, IllegalCharacterError
from PIL import Image, ImageFilter, ImageEnhance
from PIL.Image import Resampling
from zipfile import ZipFile, is_zipfile
import tarfile

TIKA_AVAILABLE = False
try:
    from tika import parser as tika_parser
    TIKA_AVAILABLE = True
except ImportError:
    pass

PY7ZR_AVAILABLE = False
try:
    import py7zr
    PY7ZR_AVAILABLE = True
except ImportError:
    pass

RARFILE_AVAILABLE = False
try:
    import rarfile
    RARFILE_AVAILABLE = True
except ImportError:
    pass

PPTX_AVAILABLE = False
try:
    from pptx import Presentation
    PPTX_AVAILABLE = True
except ImportError:
    pass

EXCHANGELIB_AVAILABLE = False
try:
    from exchangelib import DELEGATE, Account, Credentials, Configuration, Message, FileAttachment
    EXCHANGELIB_AVAILABLE = True
except ImportError:
    pass

# --- Tesserocr инициализируется позже, после создания логгера ---
TESSEROCR_AVAILABLE = False
TESSEROCR_WORKING = False

WIN_COM_AVAILABLE = False
XLWINGS_AVAILABLE = False
WIN32NET_AVAILABLE = False
WIN32SECURITY_AVAILABLE = False
if sys.platform == "win32":
    try:
        import pythoncom
        from win32com.client import DispatchEx
        WIN_COM_AVAILABLE = True
    except ImportError:
        pass
    try:
        import xlwings
        XLWINGS_AVAILABLE = True
    except ImportError:
        pass
    try:
        import win32net
        import win32netcon
        WIN32NET_AVAILABLE = True
    except ImportError:
        pass
    try:
        import win32security
        WIN32SECURITY_AVAILABLE = True
    except ImportError:
        pass
else:
    try:
        import pwd
    except ImportError:
        pwd = None

TESSERACT_AVAILABLE = False
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    pytesseract = None

XLRD_AVAILABLE = False
try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    pass

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

logger = logging.getLogger(__name__)
CONFIG_FILE = 'config.json'
CACHE_FILE = 'hash_cache.json'
TIKA_SESSION = None
TIKA_SESSION_LOCK = threading.Lock()

# --- Инициализация Tesserocr после создания логгера ---
try:
    import tesserocr
    TESSEROCR_AVAILABLE = True
    # Проверяем, доступны ли языковые данные
    test_img = Image.new('RGB', (10, 10))
    tesserocr.image_to_text(test_img, lang='rus+eng')
    TESSEROCR_WORKING = True
    logger.info("Tesserocr инициализирован успешно")
except ImportError:
    TESSEROCR_AVAILABLE = False
    logger.debug("Tesserocr не установлен. Будет использован pytesseract.")
except Exception as e:
    TESSEROCR_WORKING = False
    logger.debug(f"Tesserocr недоступен: {e}. Будет использован pytesseract.")

TEXT_BASED_EXTENSIONS = ['.csv', '.json', '.xml', '.yaml', '.yml', '.md', '.html', '.htm', '.cfg', '.conf', '.ini', '.rtf', '.sql', '.log', '.txt', '.bat', '.sh', '.env', '.properties', '.vbs', '.ps1', '.py', '.js', '.php', '.reg', '.cmd']
OFFICE_EXTENSIONS = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.xlsm', '.ppt', '.pptx']
ARCHIVE_EXTENSIONS = ['.zip', '.rar', '.7z', '.tar', '.gz', '.bz2']
ALL_SUPPORTED_EXTENSIONS = OFFICE_EXTENSIONS + TEXT_BASED_EXTENSIONS + ARCHIVE_EXTENSIONS

HANDLER_MAP = {
    '.docx': 'check_word_document', '.doc': 'check_word_document',
    '.pdf': 'check_pdf_document',
    '.xlsx': 'check_excel_document', '.xls': 'check_excel_document', '.xlsm': 'check_excel_document',
    '.pptx': 'check_presentation_document', '.ppt': 'check_presentation_document'
}
for text_ext in TEXT_BASED_EXTENSIONS:
    HANDLER_MAP[text_ext] = 'check_text_document'


class Config:
    def __init__(self, settings: Dict):
        self.settings = settings

    @property
    def ocr_enabled(self) -> bool:
        return (
            (TESSERACT_AVAILABLE or TESSEROCR_AVAILABLE) and
            self.settings.get('use_ocr', False)
        )

    @property
    def tika_enabled(self) -> bool:
        return self.settings.get('use_tika', False)

    @property
    def tika_server_enabled(self) -> bool:
        return self.settings.get('tika_server_enabled', False)

    @property
    def log_level(self) -> str:
        return self.settings.get('log_level', 'INFO')

    def get(self, key: str, default=None):
        return self.settings.get(key, default)

    @functools.cached_property
    def text_grifs_regex(self) -> re.Pattern:
        patterns = self.settings.get('text_grif_patterns', [])
        flags = 0 if self.settings.get('text_regex_case_sensitive', False) else re.IGNORECASE
        if not patterns:
            return re.compile(r'$.^')
        return re.compile('|'.join(f'(?:{p})' for p in patterns), flags)

    @functools.cached_property
    def ocr_grifs_regex(self) -> re.Pattern:
        patterns = self.settings.get('ocr_grif_patterns', [])
        flags = 0 if self.settings.get('ocr_regex_case_sensitive', False) else re.IGNORECASE
        if not patterns:
            return re.compile(r'$.^')
        return re.compile('|'.join(f'(?:{p})' for p in patterns), flags)

    @functools.cached_property
    def exclude_path_regex(self) -> Optional[re.Pattern]:
        patterns = self.settings.get('exclude_path_patterns', [])
        if not patterns:
            return None
        valid_patterns = []
        for p in patterns:
            try:
                re.compile(p, re.IGNORECASE)
                valid_patterns.append(p)
            except re.error:
                escaped = re.escape(p)
                try:
                    re.compile(escaped, re.IGNORECASE)
                    valid_patterns.append(escaped)
                    logger.debug(f"Паттерн исключения '{p}' не является корректным regex. Используется literal: {escaped}")
                except re.error:
                    logger.debug(f"Некорректный паттерн исключения '{p}' пропущен")
        if not valid_patterns:
            return None
        combined = '|'.join(f'(?:{p})' for p in valid_patterns)
        try:
            return re.compile(combined, re.IGNORECASE)
        except re.error as e:
            logger.error(f"Ошибка компиляции итогового regex исключений: {e}")
            return None

    @property
    def enabled_extensions(self) -> Tuple[str, ...]:
        text_exts = {ext for ext, enabled in self.settings.get('text_scanned_extensions', {}).items() if enabled}
        ocr_exts = {ext for ext, enabled in self.settings.get('ocr_scanned_extensions', {}).items() if enabled}
        return tuple(text_exts.union(ocr_exts))

    @property
    def tika_server_endpoint(self) -> str:
        return self.settings.get('tika_endpoint', 'http://127.0.0.1:9998')

    @property
    def tesseract_cmd(self) -> Optional[str]:
        tesseract_path_arg = self.settings.get('tesseract_path', '')
        if not TESSERACT_AVAILABLE:
            return None
        if tesseract_path_arg and os.path.exists(tesseract_path_arg):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path_arg
            return tesseract_path_arg
        tesseract_path_in_env = shutil.which('tesseract')
        if tesseract_path_in_env:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path_in_env
            return tesseract_path_in_env
        default_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(default_path):
            pytesseract.pytesseract.tesseract_cmd = default_path
            return default_path
        return None

    @property
    def unrar_cmd(self) -> Optional[str]:
        unrar_path_arg = self.settings.get('unrar_path', '')
        if unrar_path_arg and os.path.exists(unrar_path_arg):
            return unrar_path_arg
        unrar_tool = shutil.which('UnRAR.exe') or shutil.which('unrar')
        if unrar_tool:
            return unrar_tool
        for path in [
            r'C:\Program Files\WinRAR\UnRAR.exe',
            r'C:\Program Files (x86)\WinRAR\UnRAR.exe'
        ]:
            if os.path.exists(path):
                return path
        return None


def get_default_settings() -> Dict:
    default_patterns = [
        r"(?:гриф|пометка|метка)[\s-]*(?:конфиденциально|секретно|дсп)", r"(?:строго\s*)?конфиденциально(?:сть)?",
        r"(?:совершенно\s*)?секретно", r"для\s*служебного\s*пользования", r"\bдсп\b", r"коммерческая[\s-]*тайна",
        r"ком[\s.,-]*тайна", r"ноу[- ]хау", r"confidential", r"proprietary", r"restricted", r"internal\s*use\s*only",
        r"secret", r"trade[\s-]*secret", r"commercial[\s-]*secret", r"\b\d{2}\s?\d{2}\s?\d{6}\b"
    ]
    return {
        "scan_path": os.path.expanduser("~"), "report_output_dir": os.path.join(os.path.expanduser("~"), "Отчеты грифов"),
        "directory_list_file": "", "ip_range": "192.168.1.1-192.168.1.254", "network_resolve_fqdn": True,
        "network_scan_threads": 100, "use_ocr": True, "use_tika": False, "tika_server_enabled": False, "tesseract_path": "", "unrar_path": "", 
        "tika_jar_path": "", "tika_endpoint": "http://127.0.0.1:9998", "max_file_size_mb": 200,
        "pdf_text_page_limit": 10, "pdf_ocr_page_limit": 5, "ocr_resolution": 300, "max_workers": os.cpu_count() or 4,
        "use_date_filter": False, "date_filter_days": 30, "text_grif_patterns": default_patterns,
        "ocr_grif_patterns": default_patterns, "log_level": "INFO", "scan_search_mode": "both",
        "text_regex_case_sensitive": False, "ocr_regex_case_sensitive": False,
        "text_scanned_extensions": {ext: True for ext in ALL_SUPPORTED_EXTENSIONS},
        "ocr_scanned_extensions": {ext: True for ext in OFFICE_EXTENSIONS + ARCHIVE_EXTENSIONS}, "use_hash_cache": True,
        "cache_config_keys": [
            "text_grif_patterns", "ocr_grif_patterns", "text_regex_case_sensitive", "ocr_regex_case_sensitive",
            "scan_search_mode", "text_scanned_extensions", "ocr_scanned_extensions", "use_ocr", "use_tika", "tika_server_enabled", "ocr_resolution",
            "unrar_path", "exclude_path_patterns"
        ], "scheduler_jobs": [], "email_enabled": False, "email_backend": "smtp", "email_recipient": "user@example.com",
        "email_subject": "Отчет сканера документов", "email_report_format": "xlsx", "email_smtp_server": "smtp.example.com",
        "email_smtp_port": 587, "email_smtp_user": "user@example.com", "email_smtp_password": "", "email_use_tls": True,
        "exchange_autodiscover": True, "exchange_server": "mail.example.com", "exchange_email": "sender@example.com",
        "exchange_username": "DOMAIN\\user", "exchange_password": "", "exchange_auth_mode": "specific_user",
        "exchange_ews_url": "", "exchange_verify_ssl": True,
        "play_sound_on_finish": True,
        "temp_files_dir": "",
        "exclude_path_patterns": [
            r"\\\$Recycle\.Bin\\",
            "System Volume Information",
            r"/\.DocumentRevisions-V100/",
            r"/\.Trashes/",
            "/__MACOSX/",
            r"~\$",
            r"\\\\AppData\\\\Local\\\\Temp\\\\"
        ],
        "network_list_file": "",
        "use_file_timeout": False,
        "file_timeout_seconds": 120,
        "show_results_in_gui": True,
        "tika_java_heap_size": 2048,
    }


def load_config() -> Dict:
    defaults = get_default_settings()
    if not os.path.exists(CONFIG_FILE):
        save_config(defaults)
        return defaults
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            loaded_settings = json.load(f)

        if 'grif_patterns' in loaded_settings:
            loaded_settings['text_grif_patterns'] = loaded_settings.pop('grif_patterns')
            loaded_settings['ocr_grif_patterns'] = loaded_settings.get('text_grif_patterns')
        if 'regex_case_sensitive' in loaded_settings:
            loaded_settings['text_regex_case_sensitive'] = loaded_settings.pop('regex_case_sensitive')
            loaded_settings['ocr_regex_case_sensitive'] = loaded_settings.get('text_regex_case_sensitive')

        for key, value in defaults.items():
            loaded_settings.setdefault(key, value)

        if 'scanned_extensions' in loaded_settings:
            if 'text_scanned_extensions' not in loaded_settings:
                loaded_settings['text_scanned_extensions'] = loaded_settings.get('scanned_extensions')
            if 'ocr_scanned_extensions' not in loaded_settings:
                ocr_exts = {ext: enabled for ext, enabled in loaded_settings.get('scanned_extensions').items() if ext in OFFICE_EXTENSIONS + ARCHIVE_EXTENSIONS}
                loaded_settings['ocr_scanned_extensions'] = ocr_exts
            loaded_settings.pop('scanned_extensions', None)
            save_config(loaded_settings)

        return loaded_settings
    except (json.JSONDecodeError, TypeError):
        save_config(defaults)
        return defaults


def save_config(settings: Dict):
    _save_config_to_file(settings, CONFIG_FILE)


def _save_config_to_file(settings: Dict, filepath: str):
    try:
        settings_to_save = settings.copy()
        settings_to_save.pop('email_smtp_password', None)
        settings_to_save.pop('exchange_password', None)
        settings_to_save.pop('proxy_password', None)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(settings_to_save, f, ensure_ascii=False, indent=4)
    except Exception as e:
        logger.error(f"Не удалось сохранить конфигурацию в {filepath}: {e}")


class AutoFlushFileHandler(logging.FileHandler):
    def emit(self, record):
        super().emit(record)
        self.flush()


class GuiLogger(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        if self.text_widget and self.text_widget.winfo_exists():
            self.text_widget.after(0, self._append_text, msg + '\n')

    def _append_text(self, text):
        if self.text_widget.winfo_exists():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, text)
            self.text_widget.see(tk.END)
            self.text_widget.configure(state='disabled')


def initialize_basic_logging():
    global logger
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_filename = os.path.join(log_dir, f"grif_scan_{timestamp}.log")
    file_handler = AutoFlushFileHandler(log_filename, 'w', 'utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(file_handler)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(console_handler)
    logger.info(f"Логирование сессии будет вестись в файл: {log_filename}")


def setup_logging(level_str: str, gui_handler: Optional[GuiLogger] = None):
    global logger
    log_level = getattr(logging, level_str.upper(), logging.INFO)
    logger.setLevel(log_level)
    if gui_handler is not None:
        for h in logger.handlers[:]:
            if isinstance(h, GuiLogger):
                logger.removeHandler(h)
        gui_handler.setLevel(log_level)
        gui_handler.setFormatter(logging.Formatter('%(asctime)s %(message)s', datefmt='%H:%M:%S'))
        logger.addHandler(gui_handler)


def get_file_owner(filepath: str) -> str:
    try:
        if sys.platform == "win32":
            if not WIN32SECURITY_AVAILABLE: return "N/A"
            sec_desc = win32security.GetFileSecurity(filepath, win32security.OWNER_SECURITY_INFORMATION)
            owner_sid = sec_desc.GetSecurityDescriptorOwner()
            name, domain, _ = win32security.LookupAccountSid(None, owner_sid)
            return f"{domain}\\{name}"
        else:
            if pwd: return pwd.getpwuid(os.stat(filepath).st_uid).pw_name
            else: return "N/A"
    except Exception as e:
        logger.debug(f"Не удалось получить владельца для файла {filepath}: {e}")
        return "Неизвестно"


def get_file_hash(filepath: str) -> Optional[str]:
    try:
        hash_md5 = hashlib.md5()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except IOError as e:
        logger.error(f"Ошибка хеширования {filepath}: {e}")
        return None


def search_grifs_in_text(text: Optional[str], compiled_regex: re.Pattern) -> Set[str]:
    if not text:
        return set()
    processed_text = ' '.join(text.split())
    return set(match.group(0) for match in compiled_regex.finditer(processed_text))


class TikaServerManager:
    def __init__(self, logger_instance, settings: Dict, status_callback=None):
        self.logger = logger_instance
        self.process = None
        self.status_callback = status_callback
        self.settings = settings
        self.tika_endpoint = 'http://127.0.0.1:9998'

    def _log_tika_output(self, pipe):
        last_error_time = 0
        try:
            for line_bytes in iter(pipe.readline, b''):
                line = line_bytes.decode('utf-8', errors='backslashreplace').strip()
                if not line: continue
                line_u = line.upper()
                if 'ERROR' in line_u or 'SEVERE' in line_u:
                    curr_time = time.monotonic()
                    if curr_time - last_error_time > 2.0:
                        logger.debug(f"[Tika Server] {line}")
                        last_error_time = curr_time
            pipe.close()
        except:
            pass

    def is_running(self):
        return self.process and self.process.poll() is None

    def start_server_async(self):
        if self.is_running():
            if self.status_callback: self.status_callback("Запущен")
            return
        thread = threading.Thread(target=self._start_thread_target, daemon=True)
        thread.start()

    def _start_thread_target(self):
        import subprocess
        import time
        jar_path = self.settings.get('tika_jar_path', '')
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        if jar_path and not os.path.isabs(jar_path):
            jar_path = os.path.normpath(os.path.join(script_dir, jar_path))
        
        if not jar_path or not os.path.exists(jar_path):
            potential = [
                os.path.join(script_dir, 'Tools', 'Tika', 'tika-server.jar'),
                os.path.join(script_dir, 'tika-server.jar')
            ]
            for p in potential:
                if os.path.exists(p):
                    jar_path = os.path.normpath(p)
                    break
        
        if not jar_path or not os.path.exists(jar_path):
            self.logger.error("JAR Tika не найден. Укажите путь в настройках.")
            if self.status_callback: self.status_callback("Ошибка")
            return

        java_exe = shutil.which('java') or 'java'
        heap = self.settings.get("tika_java_heap_size", 2048)
        
        cmd = [java_exe, f"-Xmx{heap}m", "-cp", jar_path, "org.apache.tika.server.core.TikaServerCli", "--host", "127.0.0.1", "--port", "9998", "--noFork"]
        
        if self.status_callback: self.status_callback("Запуск...")
        
        try:
            cf = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            self.process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=cf)
            
            threading.Thread(target=self._log_tika_output, args=[self.process.stdout], daemon=True).start()
            threading.Thread(target=self._log_tika_output, args=[self.process.stderr], daemon=True).start()
            
            start_t = time.time()
            ok = False
            while time.time() - start_t < 60:
                if not self.is_running(): break
                try:
                    resp = requests.get(f"{self.tika_endpoint}/tika", timeout=2)
                    if resp.status_code == 200:
                        ok = True
                        break
                except:
                    time.sleep(2)
            
            if ok:
                self.logger.info("Tika Server запущен успешно.")
                if self.status_callback: self.status_callback("Запущен")
                global TIKA_SESSION
                with TIKA_SESSION_LOCK:
                    TIKA_SESSION = requests.Session()
                    adapter = requests.adapters.HTTPAdapter(pool_connections=10, pool_maxsize=10, max_retries=1)
                    TIKA_SESSION.mount('http://', adapter)
            else:
                self.logger.error("Tika Server не ответил на порту.")
                if self.status_callback: self.status_callback("Ошибка")
                self.stop_server()
        except Exception as e:
            self.logger.error(f"Ошибка запуска Tika: {e}")
            if self.status_callback: self.status_callback("Ошибка")

    def stop_server(self):
        if not self.is_running():
            return
        import subprocess
        pid = self.process.pid
        try:
            if sys.platform == "win32":
                subprocess.run(['taskkill', '/F', '/T', '/PID', str(pid)], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
            else:
                self.process.terminate()
        except:
            pass
        self.process = None
        if self.status_callback: self.status_callback("Остановлен")


def extract_text_via_tika(filepath: str, config: Config) -> Optional[str]:
    if not config.tika_enabled:
        return None
    global TIKA_SESSION
    try:
        endpoint = config.tika_server_endpoint.rstrip('/')
        url = f"{endpoint}/tika"
        
        headers = {
            "Accept": "text/plain",
            "X-Tika-OCRLanguage": "rus+eng",
            "Connection": "close"
        }
        
        if TIKA_SESSION is None:
            with TIKA_SESSION_LOCK:
                if TIKA_SESSION is None:
                    TIKA_SESSION = requests.Session()
                    adapter = requests.adapters.HTTPAdapter(pool_connections=10, pool_maxsize=10, max_retries=1)
                    TIKA_SESSION.mount('http://', adapter)
        
        sess = TIKA_SESSION
        tika_timeout = config.get('file_timeout_seconds', 300)
        
        with open(filepath, 'rb') as f:
            resp = sess.put(url, data=f, headers=headers, timeout=tika_timeout)
        
        if resp.status_code == 200:
            return resp.text
        elif resp.status_code == 204:
            logger.debug(f"Tika: файл {os.path.basename(filepath)} не содержит текста.")
            return None
        else:
            logger.debug(f"Tika вернул статус {resp.status_code} для {os.path.basename(filepath)}")
            return None
    except Exception as e:
        logger.debug(f"Tika не смог обработать {os.path.basename(filepath)}: {e}")
        return None


def preprocess_image_for_ocr(image: Image.Image) -> Image.Image:
    img = image.convert('L')
    width, height = img.size
    if width < 2500 or height < 2500:
        scale = 2
        if width < 1200 or height < 1200:
            scale = 3
        img = img.resize((width * scale, height * scale), Resampling.LANCZOS)
    img = ImageEnhance.Contrast(img).enhance(2.5)
    img = img.filter(ImageFilter.DETAIL)
    img = img.filter(ImageFilter.SHARPEN)
    img = img.point(lambda p: p > 140 and 255)
    return img


def ocr_image(image: Image.Image, config: Config) -> Set[str]:
    if not config.ocr_enabled or image is None:
        return set()
    try:
        processed_img = preprocess_image_for_ocr(image)

        # 1. Пробуем быстрый tesserocr
        if TESSEROCR_AVAILABLE and TESSEROCR_WORKING:
            try:
                text = tesserocr.image_to_text(processed_img, lang='rus+eng')
                if text.strip():
                    found = search_grifs_in_text(text, config.ocr_grifs_regex)
                    if found:
                        return found
            except Exception as e:
                logger.debug(f"Tesserocr error: {e}")

        # 2. Если tesserocr не дал результатов, пробуем pytesseract (только 2 режима PSM)
        if TESSERACT_AVAILABLE:
            t_path = config.get('tesseract_path')
            if not t_path:
                t_path = getattr(config, 'tesseract_cmd', None)
            if t_path and os.path.exists(t_path):
                pytesseract.pytesseract.tesseract_cmd = t_path
            elif not shutil.which('tesseract'):
                for common_path in [
                    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                    r'C:\opt\Project\Grif-scan\Tools\Tesseract-OCR\tesseract.exe'
                ]:
                    if os.path.exists(common_path):
                        pytesseract.pytesseract.tesseract_cmd = common_path
                        break

            # Только два самых эффективных режима
            configs_to_try = ['--psm 3', '--psm 6']
            for psm in configs_to_try:
                try:
                    text = pytesseract.image_to_string(processed_img, lang='rus+eng', config=psm)
                    if text.strip():
                        found = search_grifs_in_text(text, config.ocr_grifs_regex)
                        if found:
                            return found
                except Exception as e:
                    logger.debug(f"pytesseract error ({psm}): {e}")
        return set()
    except Exception as e:
        logger.warning(f"Ошибка в ocr_image: {e}")
        return set()


@contextlib.contextmanager
def _com_application(app_name: str):
    if not WIN_COM_AVAILABLE: raise ImportError("pywin32 не доступен")
    app = None
    pythoncom.CoInitialize()
    try:
        app = DispatchEx(app_name)
        if app_name != 'PowerPoint.Application':
            app.Visible = False
        app.DisplayAlerts = False
        yield app
    finally:
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        app = None
        gc.collect()
        pythoncom.CoUninitialize()


def _convert_doc_to_docx(input_path: str, output_path: str) -> bool:
    if sys.platform != "win32":
        return False
    doc = None
    try:
        with _com_application('Word.Application') as word_app:
            doc = word_app.Documents.Open(
                os.path.abspath(input_path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False
            )
            doc.SaveAs(os.path.abspath(output_path), FileFormat=16)
            doc.Close(SaveChanges=False)
            doc = None
        return True
    except Exception as e:
        logger.error(f"Ошибка конвертации .doc {os.path.basename(input_path)}: {e}")
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        return False


def _robust_remove_dir(path: str, retries: int = 3, delay: float = 0.5):
    for i in range(retries):
        try:
            shutil.rmtree(path)
            return True
        except OSError as e:
            logger.warning(f"Попытка {i + 1}/{retries} удаления {path}: {e}. Повтор через {delay} сек.")
            time.sleep(delay)
    logger.error(f"Не удалось удалить директорию {path} после {retries} попыток.")
    return False


def _robust_remove_file(path: str, retries: int = 5, delay: float = 1.0):
    for i in range(retries):
        try:
            os.remove(path)
            return True
        except PermissionError:
            time.sleep(delay)
        except FileNotFoundError:
            return True
        except Exception as e:
            logger.error(f"Не удалось удалить файл {path} при попытке {i + 1}: {e}")
            time.sleep(delay)
    logger.error(f"Не удалось удалить файл {path} после {retries} попыток.")
    return False


def _convert_xls_to_xlsx(input_path: str, output_path: str) -> bool:
    if sys.platform != "win32":
        return False
    if XLWINGS_AVAILABLE:
        app = None
        try:
            app = xlwings.App(visible=False, add_book=False)
            app.display_alerts = False
            wb = app.books.open(input_path)
            wb.save(output_path, file_format=51)
            wb.close()
            return True
        except Exception as e:
            logger.error(f"Ошибка xlwings для {os.path.basename(input_path)}: {e}")
        finally:
            if app:
                app.quit()
            app = None
            gc.collect()
        return False
    if WIN_COM_AVAILABLE:
        wb = None
        try:
            with _com_application('Excel.Application') as excel_app:
                excel_app.AutomationSecurity = 3
                wb = excel_app.Workbooks.Open(
                    os.path.abspath(input_path),
                    UpdateLinks=0,
                    ReadOnly=True
                )
                wb.SaveAs(os.path.abspath(output_path), FileFormat=51)
                wb.Close(SaveChanges=False)
                wb = None
            return True
        except Exception as e:
            logger.error(f"Ошибка COM для {os.path.basename(input_path)}: {e}")
        return False
    logger.error("Для конвертации .xls не найдено библиотек (xlwings или pywin32).")
    return False


def _convert_ppt_to_pptx(input_path: str, output_path: str) -> bool:
    if sys.platform != "win32":
        return False
    presentation = None
    try:
        with _com_application('PowerPoint.Application') as pp_app:
            pp_app.AutomationSecurity = 3
            pp_app.FeatureInstall = 0
            try:
                pp_app.FileValidation = 0
            except AttributeError:
                logger.debug("Свойство FileValidation не поддерживается этой версией PowerPoint.")
            presentation = pp_app.Presentations.Open(
                os.path.abspath(input_path),
                ReadOnly=True,
                Untitled=False,
                WithWindow=False
            )
            presentation.SaveAs(os.path.abspath(output_path), FileFormat=24)
            presentation.Close()
        return True
    except Exception as e:
        logger.error(f"Критическая ошибка COM при конвертации .ppt файла '{os.path.basename(input_path)}'", exc_info=True)
        try:
            if presentation:
                presentation.Close()
        except Exception:
            pass
        return False
    finally:
        if presentation:
            try:
                presentation.Close()
            except:
                pass
            presentation = None


def extract_images_from_xlsx(filepath: str) -> List[Image.Image]:
    images = []
    try:
        with ZipFile(filepath) as z:
            for filename in z.namelist():
                if filename.startswith('xl/media/'):
                    try:
                        with z.open(filename) as f:
                            images.append(Image.open(BytesIO(f.read())))
                    except Exception as e:
                        logger.debug(f"Ошибка извлечения изображения из XLSX: {e}")
    except Exception as e:
        logger.error(f"Ошибка извлечения изображений из XLSX {os.path.basename(filepath)}: {e}")
    return images


def extract_images_from_docx(filepath: str) -> List[Image.Image]:
    images = []
    try:
        abs_path = os.path.abspath(filepath)
        if not os.path.exists(abs_path):
            return images
        with open(abs_path, 'rb') as f:
            source_stream = BytesIO(f.read())
        if source_stream.getbuffer().nbytes == 0:
            return images
        doc = Document(source_stream)
        for rel in doc.part.rels.values():
            if "image" in getattr(rel, 'target_ref', ''):
                try:
                    images.append(Image.open(BytesIO(rel.target_part.blob)))
                except Exception as e:
                    logger.debug(f"Ошибка извлечения изображения из DOCX: {e}")
    except Exception as e:
        logger.error(f"Ошибка извлечения изображений из DOCX {os.path.basename(filepath)}: {e}")
    return images


def extract_images_from_pptx(filepath: str) -> List[Image.Image]:
    images = []
    if not PPTX_AVAILABLE:
        return images
    try:
        pres = Presentation(filepath)
        for slide in pres.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'image'):
                    try:
                        images.append(Image.open(BytesIO(shape.image.blob)))
                    except Exception as e:
                        logger.debug(f"Ошибка извлечения изображения из PPTX: {e}")
    except Exception as e:
        logger.error(f"Ошибка извлечения изображений из PPTX {os.path.basename(filepath)}: {e}")
    return images


def extract_images_from_xls(filepath: str) -> List[Image.Image]:
    images = []
    try:
        if olefile.isOleFile(filepath):
            with olefile.OleFileIO(filepath) as ole:
                for stream_name in ole.listdir(streams=True):
                    try:
                        images.append(Image.open(BytesIO(ole.openstream(stream_name).read())))
                    except Exception as e:
                        logger.debug(f"Ошибка извлечения изображения из XLS: {e}")
    except Exception as e:
        logger.error(f"Ошибка Olefile при обработке {os.path.basename(filepath)}: {e}")
    return images


def check_text_document(filepath: str, config: Config, temp_dir: Optional[str] = None, stop_event: Optional[threading.Event] = None) -> Dict[str, Set[str]]:
    results = {'text_grifs': set(), 'ocr_grifs': set()}
    if stop_event and stop_event.is_set():
        return results

    ext = os.path.splitext(filepath)[1].lower()
    search_mode = config.get('scan_search_mode', 'both')
    is_text_scan_allowed = (search_mode in ('text', 'both') and config.get('text_scanned_extensions', {}).get(ext))
    if not is_text_scan_allowed:
        return results

    encodings_to_try = ['utf-8', 'cp1251', 'utf-16']
    chunk_size = 128 * 1024
    overlap = 1024
    
    try:
        file_size = os.path.getsize(filepath)
        if file_size == 0:
            return results

        chosen_enc = None
        with open(filepath, 'rb') as f:
            sample = f.read(4096)
            for enc in encodings_to_try:
                try:
                    sample.decode(enc)
                    chosen_enc = enc
                    break
                except:
                    continue
        
        if not chosen_enc:
            chosen_enc = 'utf-8'

        with open(filepath, 'rb') as f:
            leftover_raw = b""
            while True:
                if stop_event and stop_event.is_set():
                    return results
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                try:
                    combined_chunk = leftover_raw + chunk
                    text_chunk = combined_chunk.decode(chosen_enc, errors='ignore')
                    if len(text_chunk.strip()) >= 3:
                        found = search_grifs_in_text(text_chunk, config.text_grifs_regex)
                        if found:
                            results['text_grifs'].update(found)
                except Exception:
                    pass
                if len(chunk) > overlap:
                    leftover_raw = chunk[-overlap:]
                else:
                    leftover_raw = chunk
        return results
    except Exception as e:
        logger.debug(f"Ошибка чтения {os.path.basename(filepath)}: {e}")
        return results


def check_word_document(filepath: str, config: Config, temp_dir: str, stop_event: Optional[threading.Event] = None) -> Dict[str, Set[str]]:
    results = {'text_grifs': set(), 'ocr_grifs': set()}
    if stop_event and stop_event.is_set():
        return results
    ext = os.path.splitext(filepath)[1].lower()
    search_mode = config.get('scan_search_mode', 'both')

    is_text_scan_allowed = (search_mode in ('text', 'both') and config.get('text_scanned_extensions', {}).get(ext))
    is_ocr_scan_allowed = (config.ocr_enabled and search_mode in ('ocr', 'both') and config.get('ocr_scanned_extensions', {}).get(ext))

    if ext == '.docx':
        doc = None
        try:
            abs_path = os.path.abspath(filepath)
            if not os.path.exists(abs_path):
                logger.error(f"Файл не найден: {abs_path}")
                return results
            with open(abs_path, 'rb') as f:
                source_stream = BytesIO(f.read())
            if source_stream.getbuffer().nbytes == 0:
                logger.warning(f"Файл пуст: {os.path.basename(filepath)}")
                return results

            doc = Document(source_stream)
            if is_text_scan_allowed:
                for p in doc.paragraphs:
                    if stop_event and stop_event.is_set():
                        break
                    results['text_grifs'].update(search_grifs_in_text(p.text, config.text_grifs_regex))
                for section in doc.sections:
                    if stop_event and stop_event.is_set():
                        break
                    for header in [section.header, section.footer]:
                        if header:
                            for p in header.paragraphs:
                                results['text_grifs'].update(search_grifs_in_text(p.text, config.text_grifs_regex))
                for table in doc.tables:
                    if stop_event and stop_event.is_set():
                        break
                    for row in table.rows:
                        for cell in row.cells:
                            results['text_grifs'].update(search_grifs_in_text(cell.text, config.text_grifs_regex))
            if is_ocr_scan_allowed:
                images = extract_images_from_docx(filepath)
                for img in images:
                    if stop_event and stop_event.is_set():
                        break
                    results['ocr_grifs'].update(ocr_image(img, config))
        except Exception as e:
            logger.error(f"Ошибка чтения DOCX {os.path.basename(filepath)}: {e}")
        finally:
            if doc is not None and hasattr(doc.part, 'close'):
                doc.part.close()
    elif ext == '.doc':
        if WIN_COM_AVAILABLE:
            conversion_succeeded = False
            docx_path = os.path.join(temp_dir, f"{os.path.basename(filepath)}.docx")
            try:
                if _convert_doc_to_docx(filepath, docx_path):
                    conversion_succeeded = True
            except Exception as e:
                logger.error(f"Критический сбой при конвертации '{os.path.basename(filepath)}': {e}")
            if conversion_succeeded:
                converted_results = check_word_document(docx_path, config, temp_dir, stop_event)
                results['text_grifs'].update(converted_results['text_grifs'])
                results['ocr_grifs'].update(converted_results['ocr_grifs'])
        else:
            logger.warning(f"Файл {os.path.basename(filepath)} является DOC. Для конвертации в DOCX требуются pywin32.")
    return results


def check_excel_document(filepath: str, config: Config, temp_dir: str, stop_event: Optional[threading.Event] = None) -> Dict[str, Set[str]]:
    results = {'text_grifs': set(), 'ocr_grifs': set()}
    if stop_event and stop_event.is_set():
        return results
    ext = os.path.splitext(filepath)[1].lower()
    search_mode = config.get('scan_search_mode', 'both')

    is_text_scan_allowed = (search_mode in ('text', 'both') and config.get('text_scanned_extensions', {}).get(ext))
    is_ocr_scan_allowed = (config.ocr_enabled and search_mode in ('ocr', 'both') and config.get('ocr_scanned_extensions', {}).get(ext))

    if ext in ('.xlsx', '.xlsm'):
        workbook = None
        try:
            if is_text_scan_allowed:
                workbook = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
                for sheet in workbook.worksheets:
                    if stop_event and stop_event.is_set():
                        break
                    for row in sheet.values:
                        if not row:
                            continue
                        row_text = " ".join(str(val) for val in row if val is not None)
                        if row_text:
                            results['text_grifs'].update(search_grifs_in_text(row_text, config.text_grifs_regex))
            if is_ocr_scan_allowed:
                images = extract_images_from_xlsx(filepath)
                for img in images:
                    if stop_event and stop_event.is_set():
                        break
                    results['ocr_grifs'].update(ocr_image(img, config))
        except (InvalidFileException, zipfile.BadZipFile):
            logger.warning(f"Не удалось открыть XLSX/XLSM (возможно, поврежден/зашифрован): {os.path.basename(filepath)}")
        except IllegalCharacterError:
            logger.warning(f"Пропущен XLSX с недопустимыми символами (часто из-за macOS quarantine): {os.path.basename(filepath)}")
        except Exception as e:
            logger.error(f"Ошибка чтения XLSX/XLSM {os.path.basename(filepath)}: {e}")
        finally:
            if workbook:
                workbook.close()
    elif ext == '.xls':
        if is_text_scan_allowed:
            if XLRD_AVAILABLE:
                try:
                    workbook = xlrd.open_workbook(filepath, logfile=open(os.devnull, 'w'))
                    for sheet in workbook.sheets():
                        if stop_event and stop_event.is_set():
                            break
                        for row_idx in range(sheet.nrows):
                            for col_idx in range(sheet.ncols):
                                if str(sheet.cell(row_idx, col_idx).value):
                                    results['text_grifs'].update(search_grifs_in_text(str(sheet.cell(row_idx, col_idx).value), config.text_grifs_regex))
                except Exception as e:
                    logger.error(f"Ошибка обработки .xls через xlrd: {os.path.basename(filepath)}: {e}")
            else:
                logger.warning(f"xlrd не установлен. Текстовый анализ для .xls пропущен.")
        if is_ocr_scan_allowed:
            conversion_succeeded = False
            xlsx_path = os.path.join(temp_dir, f"{os.path.basename(filepath)}.xlsx")
            if (WIN_COM_AVAILABLE or XLWINGS_AVAILABLE):
                try:
                    if _convert_xls_to_xlsx(filepath, xlsx_path):
                        conversion_succeeded = True
                except Exception as e:
                    logger.error(f"Критический сбой при конвертации '{os.path.basename(filepath)}' для OCR: {e}")
            if conversion_succeeded:
                images = extract_images_from_xlsx(xlsx_path)
            else:
                images = extract_images_from_xls(filepath)
            for img in images:
                if stop_event and stop_event.is_set():
                    break
                results['ocr_grifs'].update(ocr_image(img, config))
    return results


def check_presentation_document(filepath: str, config: Config, temp_dir: str, stop_event: Optional[threading.Event] = None) -> Dict[str, Set[str]]:
    results = {'text_grifs': set(), 'ocr_grifs': set()}
    if stop_event and stop_event.is_set():
        return results
    ext = os.path.splitext(filepath)[1].lower()
    search_mode = config.get('scan_search_mode', 'both')

    is_text_scan_allowed = (search_mode in ('text', 'both') and config.get('text_scanned_extensions', {}).get(ext))
    is_ocr_scan_allowed = (config.ocr_enabled and search_mode in ('ocr', 'both') and config.get('ocr_scanned_extensions', {}).get(ext))

    if not PPTX_AVAILABLE:
        logger.warning(f"Для обработки презентаций требуется python-pptx. Пропуск: {os.path.basename(filepath)}")
        return results

    def process_shapes(shapes):
        for shape in shapes:
            if stop_event and stop_event.is_set():
                break
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    results['text_grifs'].update(search_grifs_in_text(paragraph.text, config.text_grifs_regex))
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        results['text_grifs'].update(search_grifs_in_text(cell.text, config.text_grifs_regex))

    if ext == '.pptx':
        try:
            logger.debug(f"Попытка открытия PPTX: {os.path.basename(filepath)}")
            pres = Presentation(filepath)
            logger.debug(f"PPTX открыт. Начало обработки слайдов.")
            if is_text_scan_allowed:
                for slide in pres.slides:
                    if stop_event and stop_event.is_set():
                        break
                    if slide.has_notes_slide:
                        results['text_grifs'].update(search_grifs_in_text(slide.notes_slide.notes_text_frame.text, config.text_grifs_regex))
                    process_shapes(slide.shapes)
            if is_ocr_scan_allowed:
                images = extract_images_from_pptx(filepath)
                for img in images:
                    if stop_event and stop_event.is_set():
                        break
                    results['ocr_grifs'].update(ocr_image(img, config))
        except Exception as e:
            logger.error(f"Ошибка чтения PPTX {os.path.basename(filepath)}: {e}")
    elif ext == '.ppt':
        if WIN_COM_AVAILABLE:
            try:
                pptx_path = os.path.join(temp_dir, f"{os.path.basename(filepath)}.pptx")
                if _convert_ppt_to_pptx(filepath, pptx_path):
                    converted_results = check_presentation_document(pptx_path, config, temp_dir, stop_event)
                    results['text_grifs'].update(converted_results['text_grifs'])
                    results['ocr_grifs'].update(converted_results['ocr_grifs'])
            except Exception as e:
                logger.error(f"Ошибка при подготовке .ppt {os.path.basename(filepath)} к конвертации: {e}")
        else:
            logger.warning(f"Файл {os.path.basename(filepath)} является PPT. Для конвертации требуется pywin32.")
    return results


def check_pdf_document(filepath: str, config: Config, temp_dir: Optional[str] = None, stop_event: Optional[threading.Event] = None) -> Dict[str, Set[str]]:
    results = {'text_grifs': set(), 'ocr_grifs': set()}
    if stop_event and stop_event.is_set():
        return results

    search_mode = config.get('scan_search_mode', 'both')
    ext = os.path.splitext(filepath)[1].lower()

    is_text_scan_allowed = (search_mode in ('text', 'both') and config.get('text_scanned_extensions', {}).get(ext))
    is_ocr_scan_allowed = (config.ocr_enabled and search_mode in ('ocr', 'both') and config.get('ocr_scanned_extensions', {}).get(ext))
    if not is_text_scan_allowed and not is_ocr_scan_allowed:
        return results

    try:
        with pdfplumber.open(filepath) as pdf:
            max_text_pages = config.get('pdf_text_page_limit', 10) if is_text_scan_allowed else 0
            max_ocr_pages = config.get('pdf_ocr_page_limit', 5) if is_ocr_scan_allowed else 0
            max_pages_to_process = max(max_text_pages, max_ocr_pages)

            for i, page in enumerate(pdf.pages[:max_pages_to_process]):
                if stop_event and stop_event.is_set():
                    break

                if is_text_scan_allowed and i < max_text_pages:
                    try:
                        text = page.extract_text(x_tolerance=3) or ""
                        if text.strip():
                            results['text_grifs'].update(search_grifs_in_text(text, config.text_grifs_regex))
                    except Exception as e:
                        logger.warning(f"Ошибка извлечения текста (pdfplumber) на стр. {i+1}: {e}")

                if is_ocr_scan_allowed and i < max_ocr_pages:
                    try:
                        pil_image = page.to_image(resolution=config.get('ocr_resolution', 300)).original
                        found_ocr = ocr_image(pil_image, config)
                        if found_ocr:
                            results['ocr_grifs'].update(found_ocr)
                    except Exception as e:
                        if "access violation" in str(e) or "0xc0000005" in str(e):
                            logger.error(f"Критический сбой Tesseract OCR на стр. {i+1} PDF {os.path.basename(filepath)}: {e}")
                        else:
                            logger.error(f"Ошибка OCR на стр. {i+1} PDF {os.path.basename(filepath)}: {e}")
    except Exception as e:
        if "No /Root object" in str(e) or "trailer not found" in str(e):
            logger.warning(f"Пропущен невалидный PDF-файл: {os.path.basename(filepath)}")
        else:
            logger.error(f"Критическая ошибка обработки PDF {os.path.basename(filepath)}: {e}")
    return results


def safe_extract(zip_ref, member, path):
    """Безопасное извлечение файла из архива (защита от Zip‑slip)."""
    target = os.path.join(path, member.filename)
    abs_target = os.path.normpath(os.path.abspath(target))
    abs_root = os.path.normpath(os.path.abspath(path))
    if not abs_target.startswith(abs_root + os.sep):
        logger.warning(f"Пропущен потенциально опасный файл в архиве: {member.filename}")
        return False
    zip_ref.extract(member, path)
    return True


def check_archive_file(archive_path: str, config: Config, stop_event: Optional[threading.Event] = None) -> List[Dict]:
    results = []
    exclude_regex = config.exclude_path_regex
    if stop_event and stop_event.is_set():
        return results
    ext = os.path.splitext(archive_path)[1].lower()
    temp_dir_arg = config.get('temp_files_dir') or None
    temp_dir_archive = tempfile.mkdtemp(prefix="grifscan_archive_", dir=temp_dir_arg)
    enabled_inner_extensions = config.enabled_extensions
    if not enabled_inner_extensions:
        _robust_remove_dir(temp_dir_archive)
        return []

    try:
        archive_size = os.path.getsize(archive_path)
        max_size = config.get('max_file_size_mb', 200) * 1024 * 1024
        if archive_size > max_size:
            logger.warning(f"Архив {os.path.basename(archive_path)} превышает лимит ({round(archive_size/(1024*1024),2)} MB) — пропуск.")
            return []
    except Exception:
        pass

    archive_error_types = [zipfile.BadZipFile, tarfile.TarError]
    if PY7ZR_AVAILABLE:
        archive_error_types.append(py7zr.exceptions.Bad7zFile)
    if RARFILE_AVAILABLE:
        archive_error_types.append(rarfile.BadRarFile)
        archive_error_types.append(rarfile.RarCannotExec)
    archive_errors_to_catch = tuple(archive_error_types)

    try:
        if ext == '.zip' and is_zipfile(archive_path):
            with ZipFile(archive_path, 'r') as zf:
                if any(m.flag_bits & 0x1 for m in zf.infolist()):
                    logger.warning(f"Архив '{os.path.basename(archive_path)}' зашифрован и будет пропущен.")
                    return []
                for member in zf.infolist():
                    inner_path_normalized = member.filename.replace('\\', '/')
                    if exclude_regex and exclude_regex.search(inner_path_normalized):
                        logger.debug(f"Пропущен файл в архиве (фильтр исключений): {archive_path} -> {inner_path_normalized}")
                        continue
                    if not member.is_dir() and member.filename.lower().endswith(enabled_inner_extensions):
                        safe_extract(zf, member, temp_dir_archive)
        elif ext in ['.tar', '.gz', '.bz2'] and tarfile.is_tarfile(archive_path):
            with tarfile.open(archive_path, 'r:*') as tf:
                for member in tf.getmembers():
                    inner_path_normalized = member.name.replace('\\', '/')
                    if exclude_regex and exclude_regex.search(inner_path_normalized):
                        logger.debug(f"Пропущен файл в архиве (фильтр исключений): {archive_path} -> {inner_path_normalized}")
                        continue
                    if member.isfile() and member.name.lower().endswith(enabled_inner_extensions):
                        tf.extract(member, temp_dir_archive)
        elif ext == '.rar' and RARFILE_AVAILABLE:
            unrar_tool_path = config.unrar_cmd
            if not unrar_tool_path:
                logger.error("Утилита UnRAR не найдена. Укажите путь в настройках или установите WinRAR. Пропуск .rar файлов.")
                return []
            rarfile.UNRAR_TOOL = unrar_tool_path
            with rarfile.RarFile(archive_path, 'r') as rf:
                if rf.needs_password():
                    logger.warning(f"Архив '{os.path.basename(archive_path)}' зашифрован и будет пропущен.")
                    return []
                for member in rf.infolist():
                    inner_path_normalized = member.filename.replace('\\', '/')
                    if exclude_regex and exclude_regex.search(inner_path_normalized):
                        logger.debug(f"Пропущен файл в архиве (фильтр исключений): {archive_path} -> {inner_path_normalized}")
                        continue
                    if not member.is_dir() and member.filename.lower().endswith(enabled_inner_extensions):
                        rf.extract(member, temp_dir_archive)
        elif ext == '.7z' and PY7ZR_AVAILABLE:
            with py7zr.SevenZipFile(archive_path, 'r') as szf:
                if szf.needs_password():
                    logger.warning(f"Архив '{os.path.basename(archive_path)}' зашифрован и будет пропущен.")
                    return []
                all_files = szf.getnames()
                targets = []
                for f in all_files:
                    inner_path_normalized = f.replace('\\', '/')
                    if exclude_regex and exclude_regex.search(inner_path_normalized):
                        logger.debug(f"Пропущен файл в архиве (фильтр исключений): {archive_path} -> {inner_path_normalized}")
                        continue
                    if not f.endswith(os.path.sep) and f.lower().endswith(enabled_inner_extensions):
                        targets.append(f)
                if targets:
                    szf.extract(path=temp_dir_archive, targets=targets)

        for root, _, files in os.walk(temp_dir_archive):
            if stop_event and stop_event.is_set():
                break
            for file in files:
                if stop_event and stop_event.is_set():
                    break
                inner_file_path = os.path.join(root, file)
                handler_name = HANDLER_MAP.get(os.path.splitext(inner_file_path)[1].lower())
                if not handler_name:
                    continue

                handler = globals()[handler_name]
                temp_dir_handler = tempfile.mkdtemp(prefix="grifscan_handler_", dir=temp_dir_arg)
                try:
                    found_data = handler(inner_file_path, config, temp_dir_handler, stop_event)
                    if found_data and (found_data['text_grifs'] or found_data['ocr_grifs']):
                        relative_path = os.path.relpath(inner_file_path, temp_dir_archive)
                        result = {
                            'path': f"{archive_path} -> {relative_path}", 'archive_path': os.path.normpath(archive_path),
                            'filename': relative_path, 'text_grifs_display': f"Найдено: {len(found_data.get('text_grifs', set()))}" if found_data.get('text_grifs') else "",
                            'ocr_grifs_display': f"Найдено: {len(found_data.get('ocr_grifs', set()))}" if found_data.get('ocr_grifs') else "",
                            'full_text_grifs': found_data.get('text_grifs', set()), 'full_ocr_grifs': found_data.get('ocr_grifs', set()),
                            'size_mb': f"{round(os.path.getsize(archive_path) / (1024 * 1024), 2)} MB", 'owner': get_file_owner(archive_path)
                        }
                        results.append(result)
                finally:
                    _robust_remove_dir(temp_dir_handler)
    except archive_errors_to_catch as e:
        if isinstance(e, rarfile.RarCannotExec):
            logger.error(f"Ошибка RAR: утилита UnRAR не найдена или неисправна. Проверьте путь в Настройки -> Интеграции и OCR. Пропуск: {os.path.basename(archive_path)}")
        else:
            logger.error(f"Не удалось прочитать архив '{os.path.basename(archive_path)}' (возможно, поврежден): {e}")
    except Exception as e:
        logger.error(f"Ошибка при обработке архива {os.path.basename(archive_path)}: {e}", exc_info=True)
    finally:
        _robust_remove_dir(temp_dir_archive)
    return results


def is_network_path(filepath: str) -> bool:
    if sys.platform == "win32":
        return os.path.normpath(filepath).startswith('\\\\')
    return False


def robust_network_copy(source: str, destination: str, retries: int = 3, delay: float = 2.0):
    for attempt in range(retries):
        try:
            shutil.copy2(source, destination)
            return True
        except (OSError, ConnectionAbortedError, socket.timeout) as e:
            logger.warning(f"Попытка {attempt + 1}/{retries} копирования '{source}' не удалась: {e}. Повтор через {delay} сек.")
            if attempt + 1 == retries:
                raise
            time.sleep(delay)
    return False


def process_file(filepath: str, config: Config, hash_cache: Dict, hash_lock: threading.Lock, stop_event: threading.Event) -> Tuple[Optional[List[Dict]], Optional[str], str]:
    if stop_event and stop_event.is_set():
        return None, None, "None"

    ext = os.path.splitext(filepath)[1].lower()
    search_mode = config.get('scan_search_mode', 'both')

    is_text_scan_allowed = (
        search_mode in ('text', 'both') and
        config.get('text_scanned_extensions', {}).get(ext, False)
    )
    is_ocr_scan_allowed = (
        config.ocr_enabled and
        search_mode in ('ocr', 'both') and
        config.get('ocr_scanned_extensions', {}).get(ext, False)
    )

    if not is_text_scan_allowed and not is_ocr_scan_allowed:
        return None, None, "None"

    logger.debug(f"Обработка: {filepath}")

    try:
        max_file_size_bytes = config.get('max_file_size_mb', 200) * 1024 * 1024
        file_size = os.path.getsize(filepath)
        if file_size > max_file_size_bytes:
            logger.warning(f"Файл {filepath} слишком большой, пропуск.")
            return None, None, "Size Limit Skip"
    except OSError as e:
        logger.error(f"Не удалось получить размер файла {filepath}: {e}. Пропуск.")
        return None, None, "OSError"

    local_temp_file = None
    path_to_process = filepath
    temp_dir_arg = config.get('temp_files_dir') or None

    complex_extensions = ['.doc', '.xls', '.ppt', '.docx', '.xlsx', '.xlsm', '.pptx', '.pdf', '.zip', '.rar', '.7z']
    if is_network_path(filepath) and ext in complex_extensions:
        try:
            fd, local_temp_file = tempfile.mkstemp(suffix=ext, prefix="grifscan_net_", dir=temp_dir_arg)
            os.close(fd)
            logger.info(f"Копирование сетевого файла: {filepath}")
            robust_network_copy(filepath, local_temp_file)
            path_to_process = local_temp_file
        except Exception as e:
            logger.error(f"Не удалось скопировать сетевой файл {filepath}: {e}")
            if local_temp_file:
                _robust_remove_file(local_temp_file)
            return None, None, "Network Fail"

    try:
        try:
            st = os.stat(filepath)
            current_hash = f"M:{st.st_size}:{st.st_mtime}"
        except Exception:
            current_hash = None

        norm_path = os.path.normpath(filepath)
        with hash_lock:
            cached = hash_cache.get(norm_path) if config.get('use_hash_cache', False) else None
        if cached == current_hash:
            return None, current_hash, "Cache"

        if ext in ARCHIVE_EXTENSIONS:
            archive_results = check_archive_file(path_to_process, config, stop_event)
            if archive_results:
                logger.info(f"НАЙДЕНО в архиве {filepath}: {len(archive_results)} совпадений.")
                return archive_results, current_hash, "Archive Scanner"
            else:
                logger.info(f"Не обнаружено (архив): {filepath}")
                return None, current_hash, "Archive Scanner"

        handler_name = HANDLER_MAP.get(ext)
        if not handler_name and not config.tika_enabled:
            logger.debug(f"Файл {os.path.basename(filepath)} пропущен: нет обработчика для расширения {ext} и Tika отключена.")
            return None, current_hash, "None"

        found_data = {'text_grifs': set(), 'ocr_grifs': set()}
        scanner_name = "None"

        if handler_name:
            scanner_name = f"Встроенный поиск"
            handler = globals()[handler_name]
            temp_dir_path = tempfile.mkdtemp(prefix="grifscan_handler_", dir=temp_dir_arg)
            try:
                std_found = handler(path_to_process, config, temp_dir_path, stop_event)
                if std_found:
                    found_data['text_grifs'].update(std_found.get('text_grifs', set()))
                    found_data['ocr_grifs'].update(std_found.get('ocr_grifs', set()))
            except Exception as e:
                logger.error(f"Ошибка нативного обработчика {handler_name}: {e}")
            finally:
                _robust_remove_dir(temp_dir_path)

        if config.tika_enabled and not found_data['text_grifs'] and not found_data['ocr_grifs']:
            is_complex_format = ext in OFFICE_EXTENSIONS
            if is_complex_format or config.get('tika_for_all', False):
                tika_text = extract_text_via_tika(path_to_process, config)
                if tika_text and len(tika_text.strip()) > 0:
                    scanner_name = "Apache Tika" if not handler_name else "Встроенный поиск + Tika Fallback"
                    found_text = search_grifs_in_text(tika_text, config.text_grifs_regex)
                    if found_text:
                        found_data['text_grifs'].update(found_text)
                        logger.info(f"[{scanner_name}] Дополнительно найдено {len(found_text)} совпадений.")

        if found_data.get('text_grifs') or found_data.get('ocr_grifs'):
            method_str = "Локальная копия" if local_temp_file else "Прямой доступ"
            logger.info(f"НАЙДЕНО: {filepath} [{method_str}, {scanner_name}]. Найдено: Текст={len(found_data.get('text_grifs', set()))}, OCR={len(found_data.get('ocr_grifs', set()))}")
            mod_time_str = "N/A"
            try:
                mod_time_str = datetime.fromtimestamp(os.path.getmtime(path_to_process)).strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                pass

            result = {
                'path': norm_path, 'archive_path': '', 'filename': os.path.basename(filepath),
                'text_grifs_display': f"Найдено: {len(found_data.get('text_grifs', set()))}" if found_data.get('text_grifs') else "",
                'ocr_grifs_display': f"Найдено: {len(found_data.get('ocr_grifs', set()))}" if found_data.get('ocr_grifs') else "",
                'full_text_grifs': found_data.get('text_grifs', set()), 'full_ocr_grifs': found_data.get('ocr_grifs', set()),
                'size_mb': f"{round(file_size / (1024 * 1024), 2)} MB", 'owner': get_file_owner(filepath),
                'modified': mod_time_str, 'hash': current_hash
            }
            return [result], current_hash, scanner_name

        return None, current_hash, scanner_name

    except Exception as e:
        logger.error(f"Ошибка при обработке файла {filepath}: {e}", exc_info=True)
        return None, None, "Error"
    finally:
        if local_temp_file:
            _robust_remove_file(local_temp_file)


def scan_directory(root_path: str, extensions_to_scan: Tuple[str, ...], use_date_filter: bool, filter_days: int) -> List[str]:
    files_to_process: List[str] = []
    if not extensions_to_scan:
        return []

    cutoff_time = None
    if use_date_filter:
        cutoff_time = datetime.now() - timedelta(days=filter_days)

    for root, _, files in os.walk(root_path):
        for f in files:
            if f.lower().endswith(extensions_to_scan) and not f.startswith('~$') and not f.startswith('._'):
                filepath = os.path.normpath(os.path.join(root, f))
                if use_date_filter and cutoff_time:
                    try:
                        mod_time = datetime.fromtimestamp(os.path.getmtime(filepath))
                        if mod_time >= cutoff_time:
                            files_to_process.append(filepath)
                    except OSError as e:
                        logger.debug(f"Не удалось получить дату изменения для файла {filepath}: {e}")
                else:
                    files_to_process.append(filepath)
    return files_to_process


def read_directories_from_file(filepath: str) -> List[str]:
    directories = []
    if not os.path.exists(filepath):
        logger.warning(f"Файл со списком каталогов не найден: {filepath}")
        return []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                path = line.strip()
                if path and not path.startswith('#'):
                    directories.append(os.path.normpath(path))
        return directories
    except Exception as e:
        logger.error(f"Ошибка при чтении файла со списком каталогов {filepath}: {e}")
        return []


def read_networks_from_file(filepath: str) -> List[str]:
    networks = []
    if not os.path.exists(filepath):
        logger.warning(f"Файл со списком сетей не найден: {filepath}")
        return []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                net_range = line.strip()
                if net_range and not net_range.startswith('#'):
                    networks.append(net_range)
        return networks
    except Exception as e:
        logger.error(f"Ошибка при чтении файла со списком сетей {filepath}: {e}")
        return []


def is_port_open(host: str, port: int, timeout: float = 0.5) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except (socket.timeout, ConnectionRefusedError, OSError):
        return False
    except Exception:
        return False


def check_directory_accessibility(path: str) -> Tuple[str, bool, str]:
    try:
        if path.lower().startswith('\\\\host.docker.internal\\'):
            path = path.lower().replace('host.docker.internal', 'localhost', 1)
            logger.info(f"Обнаружен путь host.docker.internal. Преобразовано в: {path}")

        if path.startswith('\\\\'):
            if sys.platform != "win32":
                return path, False, "Сетевые пути не поддерживаются на этой ОС"

        if not os.path.exists(path):
            return path, False, "Путь не существует"

        if not os.path.isdir(path):
            return path, False, "Не является директорией"

        if not os.access(path, os.R_OK):
            return path, False, "Нет прав на чтение"

        if path.startswith('\\\\'):
            try:
                _ = os.listdir(path)
            except OSError as e:
                return path, False, f"Ресурс недоступен: {e}"

        return path, True, "Доступен"
    except Exception as e:
        return path, False, f"Ошибка: {e}"


def enumerate_shares(server_ip: str) -> List[str]:
    if not WIN32NET_AVAILABLE:
        logger.warning("Перечисление сетевых ресурсов доступно только на Windows с pywin32.")
        return []

    unc_paths = []
    try:
        shares, _, _ = win32net.NetShareEnum(server_ip, 1)
        for share in shares:
            if share['type'] == win32netcon.STYPE_DISKTREE:
                unc_path = f"\\\\{server_ip}\\{share['netname']}"
                unc_paths.append(unc_path)
    except Exception as e:
        logger.debug(f"Не удалось получить список ресурсов с \\\\{server_ip}: {e}")
    return unc_paths


def generate_reports(results: List[Dict], base_filepath: str, file_count: int, scan_paths: List[str], is_scheduled: bool = False) -> Dict[str, str]:
    if not results:
        logger.info("Документы с грифами не найдены. Отчеты не созданы.")
        return {}

    logger.info(f"Найдено документов с грифами: {len(results)}. Генерация отчетов...")
    df = pd.DataFrame(results)

    def get_outer_path(p):
        return p.split(' -> ')[0] if ' -> ' in p else p

    if 'modified' not in df.columns:
        try:
            df['modified'] = df.apply(
                lambda row: datetime.fromtimestamp(os.path.getmtime(row['archive_path'] or row['path'])).strftime('%Y-%m-%d %H:%M:%S')
                if os.path.exists(row['archive_path'] or row['path']) else 'N/A',
                axis=1
            )
        except Exception:
            df['modified'] = 'N/A'

    if 'hash' not in df.columns:
        try:
            df['hash'] = df.apply(
                lambda row: get_file_hash(row['archive_path'] or row['path']),
                axis=1
            )
        except Exception:
            df['hash'] = 'N/A'

    df['Грифы в тексте'] = df['full_text_grifs'].apply(lambda s: ", ".join(sorted(list(s))) if s else "")
    df['Грифы в штампах (OCR)'] = df['full_ocr_grifs'].apply(lambda s: ", ".join(sorted(list(s))) if s else "")

    df.rename(columns={
        'archive_path': 'Путь к архиву',
        'filename': 'Файл',
        'owner': 'Владелец',
        'size_mb': 'Размер',
        'modified': 'Дата изменения',
        'hash': 'Хэш'
    }, inplace=True)

    df.loc[df['Путь к архиву'] == '', 'Файл'] = df['path']

    report_columns = [
        'Путь к архиву',
        'Файл',
        'Грифы в тексте',
        'Грифы в штампах (OCR)',
        'Владелец',
        'Размер',
        'Дата изменения',
        'Хэш'
    ]

    df = df[report_columns]

    scan_paths_str = "\n".join([os.path.normpath(p) for p in scan_paths]) if scan_paths else "Не указано"
    summary_data = {
        "Параметр": ["Пути сканирования", "Время сканирования", "Всего проверено файлов", "Найдено документов с грифами", "Процент найденных"],
        "Значение": [scan_paths_str, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), file_count, len(results), f"{round(len(results) / file_count * 100, 2)}%" if file_count > 0 else "0.00%"]
    }
    summary_df = pd.DataFrame(summary_data)

    text_series = df['Грифы в тексте'].str.split(', ').explode()
    ocr_series = df['Грифы в штампах (OCR)'].str.split(', ').explode()
    all_grifs_series = pd.concat([text_series, ocr_series]).dropna().str.strip()
    all_grifs_series = all_grifs_series[all_grifs_series != '']
    grif_counts = all_grifs_series.value_counts().reset_index()
    grif_counts.columns = ['Гриф', 'Частота']

    def clean_unicode_surrogates(s):
        if isinstance(s, str):
            return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\uD800-\uDFFF]', '', s)
        return s

    for data_frame in [df, summary_df, grif_counts]:
        for col in data_frame.columns:
            if data_frame[col].dtype == 'object':
                data_frame[col] = data_frame[col].apply(clean_unicode_surrogates)

    report_paths = {}

    xlsx_path = f"{base_filepath}.xlsx"
    try:
        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            MAX_ROWS_PER_SHEET = 1048570
            if len(df) > MAX_ROWS_PER_SHEET:
                logger.warning(f"Детальный отчет слишком велик ({len(df)} строк). Он будет разбит на несколько листов.")
                chunks = [df[i:i + MAX_ROWS_PER_SHEET] for i in range(0, df.shape[0], MAX_ROWS_PER_SHEET)]
                for i, chunk in enumerate(chunks):
                    sheet_name = f'Детальный отчет ({i + 1})'
                    chunk.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df.to_excel(writer, sheet_name='Детальный отчет', index=False)

            summary_df.to_excel(writer, sheet_name='Сводка', index=False)
            grif_counts.to_excel(writer, sheet_name='Частота грифов', index=False)

            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                if sheet_name.startswith('Детальный отчет'):
                    sheet_df = df
                elif sheet_name == 'Сводка':
                    sheet_df = summary_df
                elif sheet_name == 'Частота грифов':
                    sheet_df = grif_counts
                else:
                    continue
                for idx, col in enumerate(sheet_df):
                    series = sheet_df[col]
                    max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 2
                    worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = max_len

        logger.info(f"Отчет XLSX успешно сохранен: {os.path.normpath(xlsx_path)}")
        report_paths['xlsx'] = xlsx_path
    except Exception as e:
        logger.error(f"Не удалось сохранить отчет XLSX: {e}", exc_info=True)
        if not is_scheduled:
            messagebox.showerror("Ошибка отчета XLSX", f"Не удалось сохранить отчет XLSX:\n{e}")

    csv_path = f"{base_filepath}.csv"
    try:
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        logger.info(f"Отчет CSV успешно сохранен: {os.path.normpath(csv_path)}")
        report_paths['csv'] = csv_path
    except Exception as e:
        logger.error(f"Не удалось сохранить отчет CSV: {e}")

    html_path = f"{base_filepath}.html"
    try:
        html_string = f"""
        <html>
        <head>
            <title>Отчет сканера документов</title>
            <meta charset="UTF-8">
            <style>
                body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; line-height: 1.6; }}
                table {{ border-collapse: collapse; margin: 25px 0; font-size: 0.9em; min-width: 400px; box-shadow: 0 0 20px rgba(0, 0, 0, 0.15); }}
                thead tr {{ background-color: #009879; color: #ffffff; text-align: left; }}
                th, td {{ padding: 12px 15px; border: 1px solid #dddddd; text-align: left; }}
                tbody tr {{ border-bottom: 1px solid #dddddd; }}
                tbody tr:nth-of-type(even) {{ background-color: #f3f3f3; }}
                tbody tr:last-of-type {{ border-bottom: 2px solid #009879; }}
                h1, h2 {{ color: #333; }}
                h2 {{ border-bottom: 2px solid #009879; padding-bottom: 5px; }}
            </style>
        </head>
        <body>
            <h1>Отчет сканера документов</h1>
            <h2>Сводка</h2>
            {summary_df.to_html(index=False, classes='table table-striped')}
            <h2>Частота грифов</h2>
            {grif_counts.to_html(index=False, classes='table table-striped')}
            <h2>Детальный отчет</h2>
            {df.to_html(index=False, classes='table table-striped')}
        </body>
        </html>
        """
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_string)
        logger.info(f"Отчет HTML успешно сохранен: {os.path.normpath(html_path)}")
        report_paths['html'] = html_path
    except Exception as e:
        logger.error(f"Не удалось сохранить отчет HTML: {e}")

    return report_paths


def _send_email_smtp(settings: Dict, summary: str, attachments: Optional[List[str]] = None):
    try:
        msg = MIMEMultipart()
        msg['From'] = settings.get('email_smtp_user')
        msg['To'] = settings.get('email_recipient')
        msg['Subject'] = settings.get('email_subject', 'Отчет сканера документов')

        msg.attach(MIMEText(summary, 'plain', 'utf-8'))

        if attachments:
            for report_path in attachments:
                if not (report_path and os.path.exists(report_path)):
                    continue
                basename = os.path.basename(report_path)
                with open(report_path, "rb") as attachment_file:
                    part = MIMEApplication(attachment_file.read(), Name=basename)
                part['Content-Disposition'] = f'attachment; filename="{basename}"'
                msg.attach(part)
                logger.info(f"Файл отчета '{report_path}' прикреплен к письму (SMTP).")

        server = None
        port = settings.get('email_smtp_port')
        host = settings.get('email_smtp_server')
        user = settings.get('email_smtp_user')
        password = settings.get('email_smtp_password')

        logger.info(f"Подключение к SMTP-серверу: {host}:{port}")

        if port == 465:
            logger.info("Используется SSL-соединение (порт 465).")
            server = smtplib.SMTP_SSL(host, port, timeout=30)
        else:
            server = smtplib.SMTP(host, port, timeout=30)
            if settings.get('email_use_tls'):
                logger.info("Инициируется TLS-шифрование (STARTTLS)...")
                server.starttls(context=ssl.create_default_context())

        logger.info(f"Аутентификация пользователя {user}...")
        server.login(user, password)
        logger.info("Отправка сообщения...")
        server.send_message(msg)
        server.quit()

        logger.info(f"Отчет успешно отправлен на адрес: {settings.get('email_recipient')} через SMTP.")

    except smtplib.SMTPAuthenticationError as e:
        logger.error("Ошибка аутентификации SMTP. Проверьте логин и пароль.", exc_info=True)
        raise e
    except smtplib.SMTPConnectError as e:
        logger.error("Ошибка подключения к SMTP серверу. Проверьте адрес сервера и порт.", exc_info=True)
        raise e
    except socket.timeout as e:
        logger.error("Тайм-аут при подключении к SMTP серверу. Проверьте доступность сервера и настройки файрвола.", exc_info=True)
        raise e
    except Exception as e:
        logger.error(f"Неизвестная ошибка при отправке письма через SMTP: {e}", exc_info=True)
        raise e


def _send_email_exchange(settings: Dict, summary: str, attachments: Optional[List[str]] = None):
    if not EXCHANGELIB_AVAILABLE:
        logger.error("Библиотека 'exchangelib' не установлена. Отправка через Exchange невозможна.")
        raise ImportError("Библиотека 'exchangelib' не установлена.")

    from exchangelib.protocol import BaseProtocol
    from requests.adapters import HTTPAdapter

    KERBEROS_AVAILABLE = False
    try:
        from requests_kerberos import HTTPKerberosAuth
        KERBEROS_AVAILABLE = True
    except ImportError:
        pass

    try:
        if settings.get('proxy_enabled', False) and settings.get('proxy_address'):
            logger.info("Обнаружены настройки прокси. Применяем их для Exchange.")
            proxy_user = settings.get('proxy_user')
            proxy_password = settings.get('proxy_password')
            proxy_address = settings.get('proxy_address')
            if proxy_user and proxy_password:
                proxy_url = f"http://{proxy_user}:{proxy_password}@{proxy_address}"
            else:
                proxy_url = f"http://{proxy_address}"
            proxies = {'https': proxy_url, 'http': proxy_url}

            import requests

            class ProxiedHTTPAdapter(requests.adapters.HTTPAdapter):
                def __init__(self, *args, **kwargs):
                    self.proxies = kwargs.pop('proxies')
                    super().__init__(*args, **kwargs)

                def send(self, request, **kwargs):
                    kwargs['proxies'] = self.proxies
                    return super().send(request, **kwargs)

            adapter_kwargs = {'proxies': proxies}
            BaseProtocol.HTTP_ADAPTER_CLS = lambda: ProxiedHTTPAdapter(**adapter_kwargs)

        auth_type = 'ntlm'
        auth_mode = settings.get('exchange_auth_mode', 'specific_user')

        if auth_mode == 'specific_user':
            logger.info("Используется аутентификация с указанием имени пользователя и пароля (NTLM).")
            creds = Credentials(username=settings.get('exchange_username'), password=settings.get('exchange_password'))
        elif auth_mode == 'current_user':
            logger.info("Используется интегрированная аутентификация Windows (NTLM).")
            creds = None
        elif auth_mode == 'kerberos':
            if not KERBEROS_AVAILABLE:
                logger.error("Выбран режим Kerberos, но библиотека 'requests-kerberos' не установлена. Выполните: pip install requests-kerberos")
                raise ImportError("Библиотека 'requests-kerberos' не установлена.")
            logger.info("Используется интегрированная аутентификация Windows (Kerberos).")
            auth_type = 'kerberos'
            creds = None

        if not settings.get('exchange_verify_ssl', True):
            logger.warning("Проверка SSL-сертификата для Exchange отключена. Используйте с осторожностью.")
            from exchangelib.protocol import NoVerifyHTTPAdapter
            if not settings.get('proxy_enabled'):
                BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
            else:
                logger.warning("Включен прокси-сервер. Отключение проверки SSL может не работать. Убедитесь, что прокси настроен корректно.")

        config = None
        ews_url = settings.get('exchange_ews_url', '')
        if ews_url:
            logger.info(f"Подключение к EWS напрямую по указанному URL: {ews_url}")
            config = Configuration(server=settings.get('exchange_server'), credentials=creds, auth_type=auth_type, service_endpoint=ews_url)
        elif settings.get('exchange_autodiscover'):
            logger.info(f"Запуск автообнаружения для EWS для почты {settings.get('exchange_email')}...")
        else:
            server_hostname = settings.get('exchange_server')
            logger.info(f"Подключение к серверу EWS вручную (без автообнаружения): {server_hostname}")
            config = Configuration(server=server_hostname, credentials=creds, auth_type=auth_type)

        account = Account(
            primary_smtp_address=settings.get('exchange_email'),
            config=config,
            autodiscover=settings.get('exchange_autodiscover') and not ews_url,
            access_type=DELEGATE
        )
        logger.info(f"Успешно подключились к EWS. Версия API: {account.version.api_version}")

        m = Message(
            account=account,
            folder=account.sent,
            subject=settings.get('email_subject'),
            body=summary,
            to_recipients=[settings.get('email_recipient')]
        )

        if attachments:
            for report_path in attachments:
                if not (report_path and os.path.exists(report_path)):
                    continue
                with open(report_path, 'rb') as f:
                    content = f.read()
                attachment = FileAttachment(name=os.path.basename(report_path), content=content)
                m.attach(attachment)
                logger.info(f"Файл отчета '{report_path}' прикреплен к письму (Exchange).")

        m.send_and_save()
        logger.info(f"Отчет успешно отправлен на адрес: {settings.get('email_recipient')} через Exchange.")

    except Exception as e:
        logger.error(f"Ошибка при отправке отчета через Exchange. Проверьте настройки подключения, логин/пароль и доступность сервера.", exc_info=True)
        raise e
    finally:
        from exchangelib.protocol import BaseProtocol
        from requests.adapters import HTTPAdapter
        BaseProtocol.HTTP_ADAPTER_CLS = HTTPAdapter


def send_email_report(settings: Dict, summary: str, attachments: Optional[List[str]] = None):
    if not settings.get('email_enabled'):
        return

    backend = settings.get('email_backend', 'smtp')
    logger.info(f"Попытка отправки отчета по электронной почте через {backend.upper()}...")

    try:
        if backend == 'smtp':
            _send_email_smtp(settings, summary, attachments)
        elif backend == 'exchange':
            _send_email_exchange(settings, summary, attachments)
        else:
            logger.error(f"Неизвестный email-backend: {backend}. Отправка отменена.")
            raise ValueError(f"Неизвестный email-backend: {backend}")
    except Exception as e:
        raise e


def format_duration(seconds_total: float) -> str:
    total_seconds = int(seconds_total)
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    duration_parts = []
    if hours > 0:
        duration_parts.append(f"{hours} ч")
    if minutes > 0:
        duration_parts.append(f"{minutes} мин")
    if total_seconds < 60 or (hours == 0 and minutes > 0):
        duration_parts.append(f"{seconds} сек")
    return " ".join(duration_parts) if duration_parts else "0 сек"


def cleanup_current_session_temp_files(temp_dir_override: Optional[str] = None):
    temp_dir = temp_dir_override or tempfile.gettempdir()
    logger.info(f"Запуск финальной очистки временных файлов в директории: {temp_dir}")
    prefixes_to_clean = ('grifscan_archive_', 'grifscan_handler_', 'grifscan_net_', 'grifscan_')
    cleaned_count = 0
    try:
        with os.scandir(temp_dir) as it:
            for entry in it:
                if entry.name.startswith(prefixes_to_clean):
                    full_path = entry.path
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            if _robust_remove_dir(full_path, retries=2, delay=0.2):
                                logger.debug(f"Финальная очистка: удалена директория {full_path}")
                                cleaned_count += 1
                        else:
                            if _robust_remove_file(full_path, retries=2, delay=0.2):
                                logger.debug(f"Финальная очистка: удален файл {full_path}")
                                cleaned_count += 1
                    except Exception as e:
                        logger.debug(f"Финальная очистка: пропуск '{full_path}': {e}")
        if cleaned_count > 0:
            logger.info(f"Финальная очистка завершена. Удалено {cleaned_count} временных ресурсов.")
        else:
            logger.info("Финальная очистка: активных временных файлов не найдено.")
    except Exception as e:
        logger.error(f"Ошибка при фоновой очистке: {e}", exc_info=True)


def cleanup_stale_temp_files(max_age_hours: int = 1, temp_dir_override: Optional[str] = None):
    temp_dir = temp_dir_override or tempfile.gettempdir()
    logger.info(f"Запуск очистки старых временных файлов в директории: {temp_dir}")
    prefixes_to_clean = ('grifscan_archive_', 'grifscan_handler_', 'grifscan_net_')
    cutoff_time = time.time() - (max_age_hours * 3600)
    cleaned_count = 0

    try:
        for filename in os.listdir(temp_dir):
            if filename.startswith(prefixes_to_clean):
                full_path = os.path.join(temp_dir, filename)
                try:
                    file_mod_time = os.path.getmtime(full_path)
                    if file_mod_time < cutoff_time:
                        if os.path.isdir(full_path):
                            logger.info(f"Удаление старой временной директории: {full_path}")
                            if _robust_remove_dir(full_path):
                                cleaned_count += 1
                        elif os.path.isfile(full_path):
                            logger.info(f"Удаление старого временного файла: {full_path}")
                            if _robust_remove_file(full_path):
                                cleaned_count += 1
                except FileNotFoundError:
                    continue
                except Exception as e:
                    logger.warning(f"Не удалось обработать/удалить временный ресурс '{full_path}': {e}")
        if cleaned_count > 0:
            logger.info(f"Очистка завершена. Удалено {cleaned_count} старых временных ресурсов.")
        else:
            logger.info("Старые временные файлы для очистки не найдены.")
    except Exception as e:
        logger.error(f"Произошла ошибка во время сканирования временной директории: {e}")


def _resolve_fqdn_with_timeout(ip: str, timeout: int = 2) -> str:
    result = {"fqdn": "Ошибка разрешения (таймаут)"}
    def target():
        try:
            resolved_fqdn = socket.getfqdn(ip)
            if resolved_fqdn != ip:
                result["fqdn"] = resolved_fqdn
            else:
                result["fqdn"] = "Не удалось разрешить"
        except Exception as e:
            logger.debug(f"Не удалось получить FQDN для {ip}: {e}")
            result["fqdn"] = "Ошибка разрешения"

    thread = threading.Thread(target=target, daemon=True)
    thread.start()
    thread.join(timeout)
    return result["fqdn"]


# --- Основной класс приложения ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🕵️ ОКИИА. Комплексный сканер документов и сети")
        self.geometry("1100x800")
        self.minsize(1000, 750)
        self.center_window()

        self.current_config_file_var = tk.StringVar(value="по умолчанию (config.json)")

        self._update_title_with_time()
        self.settings = load_config()
        self.config = Config(self.settings)

        self.tika_manager = TikaServerManager(logger, self.settings, status_callback=self._update_tika_status_callback)

        # Заменяем multiprocessing.Manager на обычные структуры с блокировками
        self.current_stop_event = threading.Event()
        self.hash_cache: Dict[str, str] = {}
        self.hash_lock = threading.Lock()

        self.current_scan_thread = None
        self.scan_lock = threading.Lock()

        self.traffic_light_colors = {"on": "#FF4136", "off": "#2ECC40", "disabled": "#AAAAAA"}

        self.scan_results = []
        self.report_path = ""
        self.network_report_dir = ""
        self.accessible_directories: List[str] = []
        self.found_network_results: List[Dict] = []

        self.current_config_hash: Optional[str] = None

        self.scheduler_thread = None
        self.scheduler_stop_event = threading.Event()

        self.sash_position = 400
        self.gui_logger_handler = GuiLogger(None)

        self.create_widgets()

        self._configure_styles()

        self.gui_logger_handler.text_widget = self.log_text
        self._apply_logging_settings()

        logger.info("Приложение инициализируется.")
        logger.info(f"Логирование активно. Уровень: {self.settings.get('log_level', 'INFO')}.")

        self.check_dependencies()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.start_scheduler_thread()

        if self.settings.get('tika_server_enabled'):
            self.after(500, self.tika_manager.start_server_async)

    def select_temp_dir(self):
        path = filedialog.askdirectory(
            initialdir=self.settings_vars['temp_files_dir'].get(),
            title="Выберите каталог для временных файлов"
        )
        if path:
            self.settings_vars['temp_files_dir'].set(os.path.normpath(path))

    def _generate_reports_in_background(self, results_copy: List[Dict], base_filepath: str, processed_count: int, paths_to_scan: List[str], is_scheduled: bool):
        try:
            logger.info("Фоновая генерация отчетов запущена...")
            report_paths = generate_reports(results_copy, base_filepath, processed_count, paths_to_scan, is_scheduled)
            if report_paths:
                self.report_path = report_paths.get('xlsx', self.settings.get('report_output_dir'))
                self.after(0, self.update_status, f"Отчеты успешно созданы в: {os.path.dirname(self.report_path)}")
                self.after(0, lambda: self.report_button.config(state='normal'))
            else:
                self.after(0, self.update_status, "Отчеты не были созданы (нет результатов).")

            email_format = self.settings.get('email_report_format', 'xlsx').lower()
            report_path_for_email = report_paths.get(email_format)
            summary = (f"Сканирование документов завершено.\n"
                       f"Проверено файлов: {processed_count}\n"
                       f"Найдено документов с грифами: {len(results_copy)}\n"
                       f"Пути сканирования: {', '.join(paths_to_scan)}")
            attachments = [report_path_for_email] if report_path_for_email else None

            if self.settings.get('email_enabled'):
                self._send_email_in_background(self.settings.copy(), summary, attachments)
            import gc
            gc.collect()
        except Exception as e:
            logger.error(f"Критическая ошибка при фоновой генерации отчетов: {e}", exc_info=True)
            self.after(0, self.update_status, "Ошибка при создании отчетов! См. лог.")

    def _update_traffic_light(self, is_scanning: bool):
        if is_scanning:
            color = self.traffic_light_colors["on"]
        else:
            color = self.traffic_light_colors["off"]
        if self.traffic_light_canvas.winfo_exists():
            self.traffic_light_canvas.itemconfig(self.traffic_light_indicator, fill=color, outline=color)

    def _cleanup_in_background(self):
        logger.info("Запуск фоновой очистки временных файлов...")
        try:
            temp_dir_setting = self.settings.get('temp_files_dir') or None
            cleanup_current_session_temp_files(temp_dir_override=temp_dir_setting)
            logger.info("Фоновая очистка временных файлов завершена.")
            logger.info("--- Система полностью готова к следующему сканированию ---")
        except Exception as e:
            logger.error(f"Ошибка при фоновой очистке: {e}", exc_info=True)

    def _configure_styles(self):
        style = ttk.Style(self)
        style.configure("Treeview.Heading", font=('Helvetica', 10, 'bold'))
        style.map('Treeview',
                  background=[('selected', '#0078d7')],
                  foreground=[('selected', 'white')])
        self.results_tree.tag_configure('oddrow', background='#F7F7F7')
        self.results_tree.tag_configure('evenrow', background='white')

    def _create_tooltip(self):
        if hasattr(self, 'tooltip_window') and self.tooltip_window.winfo_exists():
            return
        self.tooltip_window = tk.Toplevel(self)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_withdraw()
        self.tooltip_label = ttk.Label(self.tooltip_window, text="", justify=tk.LEFT,
                                       background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                                       padding="5", wraplength=800)
        self.tooltip_label.pack()
        self._tooltip_after_id = None

    def _schedule_tooltip(self, event):
        self._hide_tooltip()
        item_id = self.results_tree.identify_row(event.y)
        if item_id:
            if self._tooltip_after_id:
                self.after_cancel(self._tooltip_after_id)
            self._tooltip_after_id = self.after(500, lambda e=event, i=item_id: self._show_tooltip(e, i))

    def _show_tooltip(self, event, item_id):
        if not item_id:
            return
        result_data = next((res for res in self.scan_results if res['path'] == item_id or res.get('path') == item_id), None)
        if not result_data:
            return
        
        tooltip_parts = []
        text_grifs = result_data.get('full_text_grifs', set())
        ocr_grifs = result_data.get('full_ocr_grifs', set())
        
        if text_grifs:
            sorted_grifs = "\n• ".join(sorted(list(text_grifs)))
            tooltip_parts.append(f"Совпадений в тексте:\n• {sorted_grifs}")
        if ocr_grifs:
            sorted_grifs = "\n• ".join(sorted(list(ocr_grifs)))
            tooltip_parts.append(f"Совпадений на изображениях (OCR):\n• {sorted_grifs}")
            
        tooltip_text = "\n\n".join(tooltip_parts)
        if tooltip_text:
            self.tooltip_label.config(text=tooltip_text)
            self.tooltip_window.wm_geometry(f"+{event.x_root + 15}+{event.y_root + 10}")
            self.tooltip_window.wm_deiconify()

    def _hide_tooltip(self, event=None):
        if self._tooltip_after_id:
            self.after_cancel(self._tooltip_after_id)
            self._tooltip_after_id = None
        if hasattr(self, 'tooltip_window') and self.tooltip_window.winfo_exists():
            self.tooltip_window.wm_withdraw()

    def _update_title_with_time(self):
        now_str = datetime.now().strftime("%H:%M:%S")
        base_title = "🕵️ ОКИИА. Комплексный сканер документов и сети"
        self.title(f"{base_title} | Текущее время: {now_str}")
        self.after(1000, self._update_title_with_time)

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def _get_config_hash(self) -> str:
        keys_to_hash = self.settings.get("cache_config_keys", [])
        config_subset = {key: self.settings.get(key) for key in keys_to_hash}
        config_string = json.dumps(config_subset, sort_keys=True)
        return hashlib.md5(config_string.encode('utf-8')).hexdigest()

    def _load_hash_cache(self):
        self.current_config_hash = self._get_config_hash()
        if not self.settings.get('use_hash_cache'):
            logger.info("Использование кэша отключено в настройках. Будут проверяться все файлы")
            with self.hash_lock:
                self.hash_cache.clear()
            return
        if not os.path.exists(CACHE_FILE):
            logger.info("Файл кэша не найден. Будет создан новый.")
            with self.hash_lock:
                self.hash_cache.clear()
            return
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                disk_cache = json.load(f)
            if disk_cache.get("config_hash") == self.current_config_hash:
                with self.hash_lock:
                    self.hash_cache.clear()
                    self.hash_cache.update(disk_cache.get("scanned_files", {}))
                logger.info(f"Кэш успешно загружен. Найдено {len(self.hash_cache)} записей. Файлы с совпадающими хэшами будут пропущены")
            else:
                logger.warning("Конфигурация изменилась. Кэш будет очищен и создан заново.")
                with self.hash_lock:
                    self.hash_cache.clear()
        except (json.JSONDecodeError, TypeError, IOError) as e:
            logger.error(f"Ошибка загрузки кэша. Файл будет создан заново. Ошибка: {e}")
            with self.hash_lock:
                self.hash_cache.clear()

    def clear_hash_cache(self):
        if messagebox.askyesno("Очистка кэша", "Вы уверены, что хотите очистить кэш проверенных файлов? Следующее сканирование будет полным."):
            with self.hash_lock:
                self.hash_cache.clear()
            if os.path.exists(CACHE_FILE):
                try:
                    os.remove(CACHE_FILE)
                    logger.info("Файл кэша успешно удален.")
                    messagebox.showinfo("Успешно", "Кэш очищен.")
                except OSError as e:
                    logger.error(f"Не удалось удалить файл кэша: {e}")
                    messagebox.showerror("Ошибка", f"Не удалось удалить файл кэша:\n{e}")
            else:
                logger.info("Кэш очищен (файл не существовал).")
                messagebox.showinfo("Успешно", "Кэш очищен.")

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.scan_tab = ttk.Frame(self.notebook)
        self.settings_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.scan_tab, text="Сканирование")
        self.notebook.add(self.settings_tab, text="Настройки")

        self.create_scan_tab()
        self.create_settings_tab()

        self.after(100, lambda: self.scan_mode_var.set("single"))
        self.after(100, self._on_scan_mode_change)

    def create_scan_tab(self):
        mode_frame = ttk.LabelFrame(self.scan_tab, text="Режим сканирования", padding="10")
        mode_frame.pack(fill="x", pady=(10, 5), padx=10)

        self.scan_mode_var = tk.StringVar(value="single")

        ttk.Radiobutton(mode_frame, text="Сеть (поиск ресурсов)", variable=self.scan_mode_var, value="network", command=self._on_scan_mode_change).pack(side="left", padx=10)
        ttk.Radiobutton(mode_frame, text="По списку каталогов", variable=self.scan_mode_var, value="list", command=self._on_scan_mode_change).pack(side="left", padx=10)
        ttk.Radiobutton(mode_frame, text="Один каталог", variable=self.scan_mode_var, value="single", command=self._on_scan_mode_change).pack(side="left", padx=10)

        self.control_panel = ttk.Frame(self.scan_tab, padding="10")
        self.control_panel.pack(fill="x", padx=10)

        self._create_mode_control_frames()

        action_buttons_frame = ttk.Frame(self.scan_tab, padding="10")
        action_buttons_frame.pack(fill="x", padx=10)
        self.traffic_light_canvas = tk.Canvas(action_buttons_frame, width=28, height=28, highlightthickness=0)
        self.traffic_light_canvas.pack(side="left", padx=(0, 10), pady=2)
        self.traffic_light_indicator = self.traffic_light_canvas.create_oval(4, 4, 24, 24, fill=self.traffic_light_colors["off"], outline=self.traffic_light_colors["off"])

        self.start_button = ttk.Button(action_buttons_frame, text="🚀 Старт", command=self.start_scan)
        self.start_button.pack(side="left", padx=5)

        self.stop_button = ttk.Button(action_buttons_frame, text="🛑 Стоп", command=self.stop_scan, state="disabled")
        self.stop_button.pack(side="left", padx=5)

        self.check_shares_button = ttk.Button(action_buttons_frame, text="Проверить доступность", command=self.start_discovered_shares_check, state="disabled")
        self.check_shares_button.pack(side="left", padx=5)

        self.save_unc_button = ttk.Button(action_buttons_frame, text="💾 Сохранить пути...", command=self.save_unc_paths_to_file, state="disabled")
        self.save_unc_button.pack(side="left", padx=5)

        self.clear_button = ttk.Button(action_buttons_frame, text="🧹 Очистить", command=self._clear_results)
        self.clear_button.pack(side="left", padx=5)

        self.report_button = ttk.Button(action_buttons_frame, text="📊 Открыть отчет", command=self.open_report, state="disabled")
        self.report_button.pack(side="right", padx=5)

        self.main_pane = ttk.PanedWindow(self.scan_tab, orient=tk.VERTICAL)
        self.main_pane.pack(expand=True, fill="both", padx=10, pady=(5,10))

        results_frame = ttk.LabelFrame(self.main_pane, text="Результаты", padding="10")
        self.results_container = results_frame

        results_cols = {
            "filepath": ("Путь к файлу(архиву)", 450),
            "text_grifs_display": ("Совпадений в тексте", 250),
            "ocr_grifs_display": ("Совпадений в штампах (OCR)", 250),
            "owner": ("Владелец", 120),
            "size_mb": ("Размер", 80)
        }

        self.results_tree = ttk.Treeview(results_frame, columns=list(results_cols.keys()), show="headings")
        for name, (text, width) in results_cols.items():
            self.results_tree.heading(name, text=text)
            self.results_tree.column(name, width=width, anchor='w')

        self._make_treeview_sortable(self.results_tree, list(results_cols.keys()))

        self.results_tree.bind("<Double-1>", self.on_result_double_click)
        self.results_tree.bind("<Button-3>", self.show_result_context_menu)
        self.results_tree.tag_configure("whitelisted", foreground="grey")

        self._create_tooltip()
        self.results_tree.bind("<Motion>", self._schedule_tooltip)
        self.results_tree.bind("<Leave>", self._hide_tooltip)

        network_cols = {
            "resource": ("Сетевой ресурс", 350), "ip": ("IP-адрес", 120),
            "fqdn": ("Имя компьютера (FQDN)", 200), "status": ("Статус доступа", 150)
        }
        self.network_tree = ttk.Treeview(results_frame, columns=list(network_cols.keys()), show="headings")
        for name, (text, width) in network_cols.items():
            self.network_tree.heading(name, text=text)
            self.network_tree.column(name, width=width, anchor='w')

        self._make_treeview_sortable(self.network_tree, list(network_cols.keys()))

        self.network_tree.tag_configure("accessible", foreground="green")
        self.network_tree.tag_configure("inaccessible", foreground="red")

        self.results_tree.pack(side="left", expand=True, fill="both")
        self.network_tree.pack_forget()

        scrollbar_tree = ttk.Scrollbar(results_frame, orient="vertical")
        self.results_tree.configure(yscrollcommand=scrollbar_tree.set)
        self.network_tree.configure(yscrollcommand=scrollbar_tree.set)
        scrollbar_tree.config(command=lambda *args: (self.results_tree.yview(*args), self.network_tree.yview(*args)))
        scrollbar_tree.pack(side="right", fill="y")

        self.main_pane.add(results_frame, weight=1)

        bottom_frame = ttk.Frame(self.main_pane)
        self.progress_bar = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=5, pady=2)
        self.status_label = ttk.Label(bottom_frame, text="Готов к работе.", anchor="w")
        self.status_label.pack(fill="x", padx=5, pady=2)

        log_frame_container = ttk.LabelFrame(bottom_frame, text="Журнал событий", padding="10")
        log_frame_container.pack(expand=True, fill="both", pady=(5,0))
        self.log_text = scrolledtext.ScrolledText(log_frame_container, height=8, state="disabled", wrap="word", bg="#f0f0f0")
        self.log_text.pack(expand=True, fill="both")
        self.main_pane.add(bottom_frame, weight=1)

        self.after(50, lambda: self.main_pane.sashpos(0, self.sash_position))

    def _create_mode_control_frames(self):
        self.single_scan_controls = ttk.Frame(self.control_panel)
        self.scan_path_var = tk.StringVar(value=os.path.normpath(self.settings.get("scan_path")))
        ttk.Label(self.single_scan_controls, text="Папка для сканирования:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.path_entry_single = ttk.Entry(self.single_scan_controls, textvariable=self.scan_path_var, width=80)
        self.path_entry_single.grid(row=1, column=0, sticky="ew", padx=5)
        ttk.Button(self.single_scan_controls, text="Обзор...", command=self.select_directory).grid(row=1, column=1, padx=5)
        self.single_scan_controls.columnconfigure(0, weight=1)

        self.list_scan_controls = ttk.Frame(self.control_panel)
        list_controls_frame = ttk.Frame(self.list_scan_controls)
        list_controls_frame.pack(fill='x', expand=True)

        self.directory_list_file_var = tk.StringVar(value=os.path.normpath(self.settings.get("directory_list_file", "")))
        ttk.Label(list_controls_frame, text="Файл со списком каталогов:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.path_entry_list = ttk.Entry(list_controls_frame, textvariable=self.directory_list_file_var, width=80)
        self.path_entry_list.grid(row=1, column=0, sticky="ew", padx=5)
        ttk.Button(list_controls_frame, text="Обзор...", command=self.select_directory_list_file).grid(row=1, column=1, padx=5)

        list_buttons_frame = ttk.Frame(list_controls_frame)
        list_buttons_frame.grid(row=2, column=0, columnspan=2, pady=5, sticky="w")
        ttk.Button(list_buttons_frame, text="Загрузить список", command=self.load_directory_list).pack(side="left", padx=5)
        ttk.Button(list_buttons_frame, text="Тест UNC путей ", command=self.check_all_directories_accessibility).pack(side="left", padx=5)
        list_controls_frame.columnconfigure(0, weight=1)

        self.dir_list_tree = ttk.Treeview(self.list_scan_controls, columns=("path", "status"), show="headings", height=5)
        self.dir_list_tree.heading("path", text="Путь к каталогу")
        self.dir_list_tree.heading("status", text="Статус")
        self.dir_list_tree.column("path", width=500, anchor="w")
        self.dir_list_tree.column("status", width=200, anchor="w")
        dir_list_scrollbar_y = ttk.Scrollbar(self.list_scan_controls, orient="vertical", command=self.dir_list_tree.yview)
        self.dir_list_tree.configure(yscrollcommand=dir_list_scrollbar_y.set)
        self.dir_list_tree.pack(side='left', fill='both', expand=True, pady=5)
        dir_list_scrollbar_y.pack(side='right', fill='y')

        self.network_scan_controls = ttk.Frame(self.control_panel)

        self.network_input_mode_var = tk.StringVar(value="manual")

        net_mode_frame = ttk.Frame(self.network_scan_controls)
        net_mode_frame.pack(fill="x", pady=(0, 5))
        ttk.Radiobutton(net_mode_frame, text="Ручной ввод", variable=self.network_input_mode_var, value="manual", command=self._on_network_input_mode_change).pack(side="left", padx=10)
        ttk.Radiobutton(net_mode_frame, text="Загрузка из файла", variable=self.network_input_mode_var, value="file", command=self._on_network_input_mode_change).pack(side="left", padx=10)

        self.network_manual_frame = ttk.Frame(self.network_scan_controls)
        self.net_ip_range_var = tk.StringVar(value=self.settings.get("ip_range"))
        ttk.Label(self.network_manual_frame, text="Диапазон IP, CIDR или FQDN хоста:").pack(anchor="w", padx=5)
        ttk.Entry(self.network_manual_frame, textvariable=self.net_ip_range_var).pack(fill="x", padx=5, pady=2)

        self.network_file_frame = ttk.Frame(self.network_scan_controls)
        self.network_list_file_var = tk.StringVar(value=self.settings.get("network_list_file", ""))

        net_file_controls = ttk.Frame(self.network_file_frame)
        net_file_controls.pack(fill='x', expand=True)
        ttk.Label(net_file_controls, text="Файл со списком сетей/хостов:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(net_file_controls, textvariable=self.network_list_file_var, width=80).grid(row=1, column=0, sticky="ew", padx=5)
        ttk.Button(net_file_controls, text="Обзор...", command=self.select_network_list_file).grid(row=1, column=1, padx=5)
        ttk.Button(net_file_controls, text="Загрузить список", command=self.load_network_list).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        net_file_controls.columnconfigure(0, weight=1)

        self.net_list_tree = ttk.Treeview(self.network_file_frame, columns=("subnet",), show="headings", height=5)
        self.net_list_tree.heading("subnet", text="Подсеть/Диапазон/Хост")
        self.net_list_tree.column("subnet", width=500, anchor="w")
        net_list_scrollbar_y = ttk.Scrollbar(self.network_file_frame, orient="vertical", command=self.net_list_tree.yview)
        self.net_list_tree.configure(yscrollcommand=net_list_scrollbar_y.set)
        self.net_list_tree.pack(side='left', fill='both', expand=True, pady=5)
        net_list_scrollbar_y.pack(side='right', fill='y')

        self.after(100, self._on_network_input_mode_change)

    def _on_network_input_mode_change(self):
        mode = self.network_input_mode_var.get()
        if mode == 'manual':
            self.network_manual_frame.pack(fill='x', expand=True)
            self.network_file_frame.pack_forget()
        else:
            self.network_manual_frame.pack_forget()
            self.network_file_frame.pack(fill='both', expand=True)

    def _on_scan_mode_change(self):
        mode = self.scan_mode_var.get()
        self.single_scan_controls.pack_forget()
        self.list_scan_controls.pack_forget()
        self.network_scan_controls.pack_forget()
        self.report_button.pack_forget()
        self.check_shares_button.pack_forget()
        self.save_unc_button.pack_forget()

        if mode == 'network':
            self.start_button.config(text="🚀 Начать поиск ресурсов")
            self.check_shares_button.pack(side="left", padx=5)
            self.save_unc_button.pack(side="left", padx=5)
            self.report_button.pack(side="right", padx=5)
            self.report_button.config(state='disabled')
            self.results_tree.pack_forget()
            self.network_tree.pack(side="left", expand=True, fill="both")
            self.results_container.config(text="Результаты сканирования сети")
        else:
            self.start_button.config(text="🚀 Старт сканирования")
            self.report_button.pack(side="right", padx=5)
            self.report_button.config(state='disabled' if not (self.report_path and os.path.exists(self.report_path)) else 'normal')
            self.network_tree.pack_forget()
            self.results_tree.pack(side="left", expand=True, fill="both")
            self.results_container.config(text="Найденные документы")

        if mode == 'single':
            self.single_scan_controls.pack(fill="x", expand=True)
        elif mode == 'list':
            self.list_scan_controls.pack(fill="both", expand=True)
        elif mode == 'network':
            self.network_scan_controls.pack(fill="x", expand=True)

        self.update_status(f"Режим '{mode}' выбран. Готов к работе.")

    def create_settings_tab(self):
        buttons_frame = ttk.Frame(self.settings_tab, padding="10")
        buttons_frame.pack(fill='x', side='bottom')
        ttk.Button(buttons_frame, text="✅ Сохранить настройки", command=self.save_all_settings).pack(side="left", padx=10)
        ttk.Button(buttons_frame, text="🔄 Сбросить до первоначальных", command=self.restore_default_settings).pack(side="left", padx=10)

        settings_notebook = ttk.Notebook(self.settings_tab)
        settings_notebook.pack(expand=True, fill="both", padx=5, pady=5)

        def create_scrollable_tab(notebook, text):
            outer = ttk.Frame(notebook)
            notebook.add(outer, text=text)

            canvas = tk.Canvas(outer, highlightthickness=0)
            scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

            return scrollable_frame

        self.general_settings_tab = create_scrollable_tab(settings_notebook, "Общие")
        self.patterns_tab = create_scrollable_tab(settings_notebook, "Паттерны поиска")
        self.performance_settings_tab = create_scrollable_tab(settings_notebook, "Производительность")
        self.exclusions_tab = create_scrollable_tab(settings_notebook, "Исключения")
        self.integrations_tab = create_scrollable_tab(settings_notebook, "Интеграции и OCR")
        self.network_settings_tab = create_scrollable_tab(settings_notebook, "Сеть")
        self.email_settings_tab = create_scrollable_tab(settings_notebook, "Email")
        self.scheduler_tab = create_scrollable_tab(settings_notebook, "Планировщик")

        self._init_settings_vars()
        self._create_general_settings_widgets()
        self._create_patterns_tab_widgets()
        self._create_performance_settings_widgets()
        self._create_integrations_settings_widgets()
        self._create_network_settings_widgets()
        self._create_exclusions_tab_widgets()
        self._create_email_settings_widgets()
        self.create_scheduler_tab()

    def _init_settings_vars(self):
        smtp_pass_from_keyring = ""
        exchange_pass_from_keyring = ""
        proxy_pass_from_keyring = ""
        if keyring:
            try:
                smtp_pass_from_keyring = keyring.get_password('grif_scan_app', 'smtp_password') or ""
                exchange_pass_from_keyring = keyring.get_password('grif_scan_app', 'exchange_password') or ""
                proxy_pass_from_keyring = keyring.get_password('grif_scan_app', 'proxy_password') or ""
                if smtp_pass_from_keyring or exchange_pass_from_keyring:
                    logger.info("Пароли успешно загружены из системного хранилища.")
            except Exception as e:
                logger.error(f"Не удалось загрузить пароли из системного хранилища: {e}")

        self.settings_vars = {
            'use_ocr': tk.BooleanVar(value=self.settings.get('use_ocr')),
            'tesseract_path': tk.StringVar(value=os.path.normpath(self.settings.get('tesseract_path', ''))),
            'unrar_path': tk.StringVar(value=os.path.normpath(self.settings.get('unrar_path', ''))),
            'max_file_size_mb': tk.IntVar(value=self.settings.get('max_file_size_mb')),
            'pdf_text_page_limit': tk.IntVar(value=self.settings.get('pdf_text_page_limit')),
            'pdf_ocr_page_limit': tk.IntVar(value=self.settings.get('pdf_ocr_page_limit')),
            'ocr_resolution': tk.IntVar(value=self.settings.get('ocr_resolution')),
            'max_workers': tk.IntVar(value=self.settings.get('max_workers')),
            'report_output_dir': tk.StringVar(value=os.path.normpath(self.settings.get('report_output_dir', ''))),
            'directory_list_file': tk.StringVar(value=os.path.normpath(self.settings.get('directory_list_file', ''))),
            'network_list_file': tk.StringVar(value=os.path.normpath(self.settings.get('network_list_file', ''))),
            'ip_range': tk.StringVar(value=self.settings.get('ip_range', "192.168.1.1-192.168.1.254")),
            'network_resolve_fqdn': tk.BooleanVar(value=self.settings.get('network_resolve_fqdn', True)),
            'network_scan_threads': tk.IntVar(value=self.settings.get('network_scan_threads', 100)),
            'log_level': tk.StringVar(value=self.settings.get('log_level', 'INFO')),
            'use_hash_cache': tk.BooleanVar(value=self.settings.get('use_hash_cache')),
            'use_date_filter': tk.BooleanVar(value=self.settings.get('use_date_filter', False)),
            'date_filter_days': tk.IntVar(value=self.settings.get('date_filter_days', 30)),
            'text_regex_case_sensitive': tk.BooleanVar(value=self.settings.get('text_regex_case_sensitive', False)),
            'ocr_regex_case_sensitive': tk.BooleanVar(value=self.settings.get('ocr_regex_case_sensitive', False)),
            'scan_search_mode': tk.StringVar(value=self.settings.get('scan_search_mode', 'both')),
            'play_sound_on_finish': tk.BooleanVar(value=self.settings.get('play_sound_on_finish', True)),
            'temp_files_dir': tk.StringVar(value=self.settings.get('temp_files_dir', '')),
            'email_enabled': tk.BooleanVar(value=self.settings.get('email_enabled')),
            'email_backend': tk.StringVar(value=self.settings.get('email_backend')),
            'email_recipient': tk.StringVar(value=self.settings.get('email_recipient')),
            'email_subject': tk.StringVar(value=self.settings.get('email_subject')),
            'email_report_format': tk.StringVar(value=self.settings.get('email_report_format')),
            'email_smtp_server': tk.StringVar(value=self.settings.get('email_smtp_server')),
            'email_smtp_port': tk.IntVar(value=self.settings.get('email_smtp_port')),
            'email_smtp_user': tk.StringVar(value=self.settings.get('email_smtp_user')),
            'email_smtp_password': tk.StringVar(value=smtp_pass_from_keyring or self.settings.get('email_smtp_password', '')),
            'email_use_tls': tk.BooleanVar(value=self.settings.get('email_use_tls')),
            'exchange_autodiscover': tk.BooleanVar(value=self.settings.get('exchange_autodiscover')),
            'exchange_server': tk.StringVar(value=self.settings.get('exchange_server')),
            'exchange_email': tk.StringVar(value=self.settings.get('exchange_email')),
            'exchange_username': tk.StringVar(value=self.settings.get('exchange_username')),
            'exchange_password': tk.StringVar(value=exchange_pass_from_keyring or self.settings.get('exchange_password', '')),
            'exchange_auth_mode': tk.StringVar(value=self.settings.get('exchange_auth_mode')),
            'exchange_ews_url': tk.StringVar(value=self.settings.get('exchange_ews_url', '')),
            'exchange_verify_ssl': tk.BooleanVar(value=self.settings.get('exchange_verify_ssl', True)),
            'proxy_enabled': tk.BooleanVar(value=self.settings.get('proxy_enabled', False)),
            'proxy_address': tk.StringVar(value=self.settings.get('proxy_address', '')),
            'proxy_user': tk.StringVar(value=self.settings.get('proxy_user', '')),
            'proxy_password': tk.StringVar(value=proxy_pass_from_keyring or self.settings.get('proxy_password', '')),
            'use_file_timeout': tk.BooleanVar(value=self.settings.get('use_file_timeout', False)),
            'file_timeout_seconds': tk.IntVar(value=self.settings.get('file_timeout_seconds', 120)),
            'show_results_in_gui': tk.BooleanVar(value=self.settings.get('show_results_in_gui', True)),
            'use_tika': tk.BooleanVar(value=self.settings.get('use_tika', False)),
            'tika_jar_path': tk.StringVar(value=os.path.normpath(self.settings.get('tika_jar_path', ''))),
            'tika_endpoint': tk.StringVar(value=self.settings.get('tika_endpoint', 'http://localhost:9998')),
            'tika_server_enabled': tk.BooleanVar(value=self.settings.get('tika_server_enabled', False)),
            'tika_java_heap_size': tk.IntVar(value=self.settings.get('tika_java_heap_size', 2048)),
        }
        self.settings_vars['text_scanned_extensions_vars'] = {
            ext: tk.BooleanVar(value=is_enabled)
            for ext, is_enabled in self.settings.get('text_scanned_extensions', {}).items()
        }
        self.settings_vars['ocr_scanned_extensions_vars'] = {
            ext: tk.BooleanVar(value=is_enabled)
            for ext, is_enabled in self.settings.get('ocr_scanned_extensions', {}).items()
        }

        self.settings_vars['log_level'].trace_add('write', self._apply_logging_settings)
    
    def _create_general_settings_widgets(self):
        scrollable_frame = self.general_settings_tab
        row = 0

        profiles_frame = ttk.LabelFrame(scrollable_frame, text="Профили конфигурации", padding="10")
        profiles_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=(5, 10))
        row += 1

        buttons_subframe = ttk.Frame(profiles_frame)
        buttons_subframe.pack(fill='x')
        ttk.Button(buttons_subframe, text="Загрузить из файла...", command=self.load_configuration_from_file).pack(side="left", padx=5, pady=5)
        ttk.Button(buttons_subframe, text="Сохранить в файл...", command=self.save_configuration_as).pack(side="left", padx=5, pady=5)

        display_subframe = ttk.Frame(profiles_frame)
        display_subframe.pack(fill='x', padx=5, pady=(5,0))
        ttk.Label(display_subframe, text="Текущий профиль:").pack(side="left")
        ttk.Label(display_subframe, textvariable=self.current_config_file_var, font="-weight bold").pack(side="left", padx=5)

        search_mode_frame = ttk.LabelFrame(scrollable_frame, text="Режим поиска", padding="10")
        search_mode_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        row += 1
        search_mode_subframe = ttk.Frame(search_mode_frame)
        search_mode_subframe.pack(fill='x')
        ttk.Radiobutton(search_mode_subframe, text="Текст и OCR (вместе)", variable=self.settings_vars['scan_search_mode'], value="both").pack(side="left", padx=10)
        ttk.Radiobutton(search_mode_subframe, text="Только текст", variable=self.settings_vars['scan_search_mode'], value="text").pack(side="left", padx=10)
        ttk.Radiobutton(search_mode_subframe, text="Только OCR", variable=self.settings_vars['scan_search_mode'], value="ocr").pack(side="left", padx=10)

        report_path_frame = ttk.LabelFrame(scrollable_frame, text="Настройки отчета", padding="10")
        report_path_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        row += 1
        ttk.Label(report_path_frame, text="Каталог для сохранения отчета:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(report_path_frame, textvariable=self.settings_vars['report_output_dir'], width=60).grid(row=1, column=0, sticky="ew", padx=5)
        ttk.Button(report_path_frame, text="Обзор...", command=self.select_report_output_directory).grid(row=1, column=1, padx=5)
        report_path_frame.columnconfigure(0, weight=1)

        dir_list_file_frame = ttk.LabelFrame(scrollable_frame, text="Файл со списком каталогов по умолчанию", padding="10")
        dir_list_file_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        row += 1
        ttk.Label(dir_list_file_frame, text="Путь к файлу со списком:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(dir_list_file_frame, textvariable=self.settings_vars['directory_list_file'], width=60).grid(row=1, column=0, sticky="ew", padx=5)
        ttk.Button(dir_list_file_frame, text="Обзор...", command=self.select_directory_list_file_for_settings).grid(row=1, column=1, padx=5)
        dir_list_file_frame.columnconfigure(0, weight=1)

        log_level_frame = ttk.LabelFrame(scrollable_frame, text="Настройки логирования", padding="10")
        log_level_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        row += 1
        ttk.Label(log_level_frame, text="Уровень логирования:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.log_level_combobox = ttk.Combobox(
            log_level_frame,
            textvariable=self.settings_vars['log_level'],
            values=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
            state="readonly"
        )
        self.log_level_combobox.grid(row=0, column=1, sticky="ew", padx=5)
        log_level_frame.columnconfigure(1, weight=1)

        misc_frame = ttk.LabelFrame(scrollable_frame, text="Прочие настройки", padding="10")
        misc_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        row += 1

        ttk.Checkbutton(
            misc_frame,
            text="Проигрывать звуковой сигнал по завершении сканирования",
            variable=self.settings_vars['play_sound_on_finish']
        ).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 5))

        ttk.Label(misc_frame, text="Каталог для временных файлов (пусто = системный):").grid(row=1, column=0, columnspan=3, sticky="w", padx=5, pady=2)
        entry_frame = ttk.Frame(misc_frame)
        entry_frame.grid(row=2, column=0, columnspan=3, sticky='ew', padx=5)
        ttk.Entry(entry_frame, textvariable=self.settings_vars['temp_files_dir'], width=60).pack(side="left", fill="x", expand=True)
        ttk.Button(entry_frame, text="Обзор...", command=self.select_temp_dir).pack(side="left", padx=5)

        scrollable_frame.grid_columnconfigure(0, weight=1)

    def _create_patterns_tab_widgets(self):
        patterns_notebook = ttk.Notebook(self.patterns_tab)
        patterns_notebook.pack(expand=True, fill="both")

        text_patterns_frame = ttk.Frame(patterns_notebook, padding="10")
        ocr_patterns_frame = ttk.Frame(patterns_notebook, padding="10")

        patterns_notebook.add(text_patterns_frame, text="Поиск в тексте")
        patterns_notebook.add(ocr_patterns_frame, text="Поиск в OCR")

        text_options_frame = ttk.Frame(text_patterns_frame)
        text_options_frame.pack(fill='x', pady=(0, 5))
        ttk.Checkbutton(
            text_options_frame,
            text="Учитывать регистр символов",
            variable=self.settings_vars['text_regex_case_sensitive']
        ).pack(side='left', padx=5)

        text_patterns_lf = ttk.LabelFrame(text_patterns_frame, text="Паттерны (регулярные выражения, по одному на строку)", padding="10")
        text_patterns_lf.pack(expand=True, fill="both")

        self.text_patterns_text = scrolledtext.ScrolledText(text_patterns_lf, height=10, wrap="word")
        self.text_patterns_text.pack(expand=True, fill="both")
        self.text_patterns_text.insert(tk.END, "\n".join(self.settings.get('text_grif_patterns', [])))
        self._setup_text_field_bindings(self.text_patterns_text)

        self._create_file_type_selector(text_patterns_frame, "Типы файлов для обработки", 'text_scanned_extensions_vars')

        ocr_options_frame = ttk.Frame(ocr_patterns_frame)
        ocr_options_frame.pack(fill='x', pady=(0, 5))
        ttk.Checkbutton(
            ocr_options_frame,
            text="Учитывать регистр символов",
            variable=self.settings_vars['ocr_regex_case_sensitive']
        ).pack(side='left', padx=5)

        ocr_patterns_lf = ttk.LabelFrame(ocr_patterns_frame, text="Паттерны (регулярные выражения, по одному на строку)", padding="10")
        ocr_patterns_lf.pack(expand=True, fill="both")

        self.ocr_patterns_text = scrolledtext.ScrolledText(ocr_patterns_lf, height=10, wrap="word")
        self.ocr_patterns_text.pack(expand=True, fill="both")
        self.ocr_patterns_text.insert(tk.END, "\n".join(self.settings.get('ocr_grif_patterns', [])))
        self._setup_text_field_bindings(self.ocr_patterns_text)

        self._create_file_type_selector(ocr_patterns_frame, "Типы файлов для обработки", 'ocr_scanned_extensions_vars')

    def _create_file_type_selector(self, parent_frame, title, vars_dict_key):
        container = ttk.LabelFrame(parent_frame, text=title, padding="10")
        container.pack(fill="x", expand=True, pady=(10, 5))

        horizontal_container = ttk.Frame(container)
        horizontal_container.pack(fill='both', expand=True)

        office_frame = ttk.LabelFrame(horizontal_container, text="Office и PDF", padding="10")
        office_frame.pack(side="left", fill="y", expand=True, padx=5, pady=5, anchor='n')

        text_frame = ttk.LabelFrame(horizontal_container, text="Текстовые файлы", padding="10")
        text_frame.pack(side="left", fill="y", expand=True, padx=5, pady=5, anchor='n')

        archive_frame = ttk.LabelFrame(horizontal_container, text="Архивы", padding="10")
        archive_frame.pack(side="left", fill="y", expand=True, padx=5, pady=5, anchor='n')

        vars_dict = self.settings_vars[vars_dict_key]

        all_frames = {
            'office': (office_frame, sorted(OFFICE_EXTENSIONS)),
            'text': (text_frame, sorted(TEXT_BASED_EXTENSIONS)),
            'archive': (archive_frame, sorted(ARCHIVE_EXTENSIONS))
        }

        cols = 4
        for frame_key, (frame, extensions) in all_frames.items():
            row, col = 0, 0
            has_exts_in_frame = False
            for ext in extensions:
                if ext in vars_dict:
                    has_exts_in_frame = True
                    var = vars_dict[ext]
                    ttk.Checkbutton(frame, text=ext, variable=var).grid(row=row, column=col, sticky='w', padx=5, pady=2)
                    col += 1
                    if col >= cols:
                        col = 0
                        row += 1
            if not has_exts_in_frame:
                frame.pack_forget()

    def _create_performance_settings_widgets(self):
        scrollable_frame = self.performance_settings_tab

        perf_frame = ttk.LabelFrame(scrollable_frame, text="Производительность и лимиты", padding="10")
        perf_frame.grid(row=0, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        ttk.Label(perf_frame, text="Макс. размер файла (МБ):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(perf_frame, from_=1, to=1024, textvariable=self.settings_vars['max_file_size_mb']).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(perf_frame, text="Кол-во процессов (файлы):").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        ttk.Spinbox(perf_frame, from_=1, to=(os.cpu_count() or 1)*2, textvariable=self.settings_vars['max_workers']).grid(row=0, column=3, sticky="w", padx=5)
        ttk.Label(perf_frame, text="Лимит страниц PDF (текст):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(perf_frame, from_=1, to=1000, textvariable=self.settings_vars['pdf_text_page_limit']).grid(row=1, column=1, sticky="w", padx=5)
        ttk.Label(perf_frame, text="Лимит страниц PDF (OCR):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(perf_frame, from_=1, to=100, textvariable=self.settings_vars['pdf_ocr_page_limit']).grid(row=2, column=1, sticky="w", padx=5)

        cache_frame = ttk.LabelFrame(scrollable_frame, text="Кэширование файлов", padding="10")
        cache_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        ttk.Checkbutton(cache_frame, text="Использовать кэш для пропуска чистых и неизмененных файлов", variable=self.settings_vars['use_hash_cache']).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Button(cache_frame, text="Очистить кэш", command=self.clear_hash_cache).grid(row=0, column=1, sticky="w", padx=10, pady=2)

        date_filter_frame = ttk.LabelFrame(scrollable_frame, text="Фильтрация по дате изменения", padding="10")
        date_filter_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        self.date_filter_checkbutton = ttk.Checkbutton(
            date_filter_frame,
            text="Сканировать только файлы, измененные за последние:",
            variable=self.settings_vars['use_date_filter'],
            command=self._toggle_date_filter_widgets
        )
        self.date_filter_checkbutton.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        self.date_filter_spinbox = ttk.Spinbox(
            date_filter_frame,
            from_=1,
            to=3650,
            textvariable=self.settings_vars['date_filter_days'],
            width=8
        )
        self.date_filter_spinbox.grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(date_filter_frame, text="дней").grid(row=0, column=2, sticky="w", padx=5)

        self.after(100, self._toggle_date_filter_widgets)

        timeout_frame = ttk.LabelFrame(scrollable_frame, text="Тайм-аут обработки файла", padding="10")
        timeout_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        self.timeout_checkbutton = ttk.Checkbutton(
            timeout_frame,
            text="Отменять обработку файла, если она длится дольше:",
            variable=self.settings_vars['use_file_timeout'],
            command=self._toggle_timeout_widgets
        )
        self.timeout_checkbutton.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        self.timeout_spinbox = ttk.Spinbox(
            timeout_frame,
            from_=10,
            to=600,
            increment=10,
            textvariable=self.settings_vars['file_timeout_seconds'],
            width=8
        )
        self.timeout_spinbox.grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(timeout_frame, text="секунд").grid(row=0, column=2, sticky="w", padx=5)

        self.after(100, self._toggle_timeout_widgets)

        gui_perf_frame = ttk.LabelFrame(scrollable_frame, text="Настройки интерфейса", padding="10")
        gui_perf_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        ttk.Checkbutton(gui_perf_frame, text="Отображать найденные результаты в таблице интерфейса", variable=self.settings_vars['show_results_in_gui']).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Label(gui_perf_frame, text="⚠️ Отключение ускоряет сканирование при очень больших объемах данных", foreground="grey").grid(row=1, column=0, sticky="w", padx=25)

    def _toggle_date_filter_widgets(self):
        if self.settings_vars['use_date_filter'].get():
            self.date_filter_spinbox.config(state='normal')
        else:
            self.date_filter_spinbox.config(state='disabled')

    def _toggle_timeout_widgets(self):
        if self.settings_vars['use_file_timeout'].get():
            self.timeout_spinbox.config(state='normal')
        else:
            self.timeout_spinbox.config(state='disabled')

    def _create_integrations_settings_widgets(self):
        scrollable_frame = self.integrations_tab
        ocr_frame = ttk.LabelFrame(scrollable_frame, text="Настройки OCR", padding="10")
        ocr_frame.grid(row=0, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        ttk.Checkbutton(ocr_frame, text="Использовать OCR (требует Tesseract)", variable=self.settings_vars['use_ocr']).grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
        ttk.Label(ocr_frame, text="Путь к tesseract.exe:").grid(row=1, column=0, sticky="w", padx=5)
        ttk.Entry(ocr_frame, textvariable=self.settings_vars['tesseract_path'], width=60).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(ocr_frame, text="Обзор...", command=self.select_tesseract_path).grid(row=1, column=2, padx=5)
        ttk.Label(ocr_frame, text="Разрешение OCR (DPI):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(ocr_frame, from_=150, to=600, increment=50, textvariable=self.settings_vars['ocr_resolution']).grid(row=2, column=1, sticky="w", padx=5)
        ocr_frame.columnconfigure(1, weight=1)

        unrar_frame = ttk.LabelFrame(scrollable_frame, text="Настройки UnRAR (для .rar архивов)", padding="10")
        unrar_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        ttk.Label(unrar_frame, text="Путь к UnRAR.exe:").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Entry(unrar_frame, textvariable=self.settings_vars['unrar_path'], width=60).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(unrar_frame, text="Обзор...", command=self.select_unrar_path).grid(row=0, column=2, padx=5)
        unrar_frame.columnconfigure(1, weight=1)

        tika_frame = ttk.LabelFrame(scrollable_frame, text="Настройки Apache Tika", padding="10")
        tika_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        ttk.Checkbutton(tika_frame, text="Использовать Apache Tika (быстрее для некоторых форматов)", variable=self.settings_vars['use_tika'], command=self._toggle_tika_widgets).grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
        self.tika_server_enabled_check = ttk.Checkbutton(tika_frame, text="Запускать локальный сервер Tika автоматически при старте", variable=self.settings_vars['tika_server_enabled'])
        self.tika_server_enabled_check.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        ttk.Label(tika_frame, text="Путь к tika-server.jar:").grid(row=2, column=0, sticky="w", padx=5)
        ttk.Entry(tika_frame, textvariable=self.settings_vars['tika_jar_path'], width=60).grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(tika_frame, text="Обзор...", command=self.select_tika_jar_path).grid(row=2, column=2, padx=5)
        ttk.Label(tika_frame, text="Tika Server Endpoint:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(tika_frame, textvariable=self.settings_vars['tika_endpoint'], width=60).grid(row=3, column=1, sticky="ew", padx=5)
        ttk.Label(tika_frame, text="Память для Tika Server (МБ):").grid(row=4, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(tika_frame, from_=512, to=8192, increment=512, textvariable=self.settings_vars['tika_java_heap_size'], width=10).grid(row=4, column=1, sticky="w", padx=5)
        
        self.tika_status_frame = ttk.Frame(tika_frame)
        self.tika_status_frame.grid(row=5, column=0, columnspan=3, sticky="w", padx=5, pady=10)

        self.tika_status_light = tk.Canvas(self.tika_status_frame, width=12, height=12, highlightthickness=0)
        self.tika_status_light.pack(side="left", padx=2)
        self.tika_light_circle = self.tika_status_light.create_oval(2, 2, 10, 10, fill="grey", outline="grey")

        self.tika_status_label = ttk.Label(self.tika_status_frame, text="Tika: Инициализация...", font=('Helvetica', 9))
        self.tika_status_label.pack(side="left", padx=5)

        self.restart_tika_btn = ttk.Button(self.tika_status_frame, text="Запустить / Перезапустить сервер Tika", command=self._restart_tika_server_manually)
        self.restart_tika_btn.pack(side="left", padx=10)
        
        tika_frame.columnconfigure(1, weight=1)

        self.after(100, self._toggle_tika_widgets)

    def _toggle_tika_widgets(self):
        is_tika_on = self.settings_vars['use_tika'].get()
        state = 'normal' if is_tika_on else 'disabled'
        
        if hasattr(self, 'tika_server_enabled_check'):
            self.tika_server_enabled_check.config(state=state)
        if hasattr(self, 'restart_tika_btn'):
            self.restart_tika_btn.config(state=state)
            
        if not is_tika_on:
            if hasattr(self, 'tika_manager'):
                logger.info("Apache Tika отключен в настройках. Остановка сервера...")
                self.tika_manager.stop_server()

    def _create_exclusions_tab_widgets(self):
        main_frame = self.exclusions_tab
        info_label = ttk.Label(
            main_frame,
            text="Здесь можно указать части пути или имена файлов, которые следует пропускать при сканировании.\nКаждый шаблон указывается с новой строки. Сравнение происходит без учета регистра.",
            wraplength=700,
            justify=tk.LEFT
        )
        info_label.pack(fill='x', pady=(0, 10))

        patterns_lf = ttk.LabelFrame(main_frame, text="Шаблоны для исключения из сканирования", padding="10")
        patterns_lf.pack(expand=True, fill="both")

        self.exclude_patterns_text = scrolledtext.ScrolledText(patterns_lf, height=15, wrap="word")
        self.exclude_patterns_text.pack(expand=True, fill="both")
        initial_patterns = self.settings.get('exclude_path_patterns', [])
        self.exclude_patterns_text.insert(tk.END, "\n".join(initial_patterns))
        self._setup_text_field_bindings(self.exclude_patterns_text)

    def _create_network_settings_widgets(self):
        scrollable_frame = self.network_settings_tab
        network_settings_frame = ttk.LabelFrame(scrollable_frame, text="Настройки сканирования сети", padding="10")
        network_settings_frame.grid(row=0, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        ttk.Label(network_settings_frame, text="Диапазон IP-адресов по умолчанию (для ручного ввода):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(network_settings_frame, textvariable=self.settings_vars['ip_range'], width=60).grid(row=0, column=1, sticky="ew", padx=5)

        ttk.Checkbutton(network_settings_frame, text="Определять FQDN (полное имя компьютера) по IP-адресу", variable=self.settings_vars['network_resolve_fqdn']).grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=5)

        ttk.Label(network_settings_frame, text="Количество потоков для сканирования сети:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Spinbox(network_settings_frame, from_=10, to=500, increment=10, textvariable=self.settings_vars['network_scan_threads'], width=10).grid(row=2, column=1, sticky="w", padx=5)

        net_list_file_frame = ttk.LabelFrame(scrollable_frame, text="Файл со списком сетей по умолчанию", padding="10")
        net_list_file_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        ttk.Label(net_list_file_frame, text="Путь к файлу:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(net_list_file_frame, textvariable=self.settings_vars['network_list_file'], width=60).grid(row=1, column=0, sticky="ew", padx=5)
        ttk.Button(net_list_file_frame, text="Обзор...", command=lambda: self.select_network_list_file(for_settings=True)).grid(row=1, column=1, padx=5)
        net_list_file_frame.columnconfigure(0, weight=1)

        network_settings_frame.columnconfigure(1, weight=1)

    def _toggle_exchange_creds_widgets(self):
        if self.settings_vars['exchange_auth_mode'].get() in ('current_user', 'kerberos'):
            state = 'disabled'
        else:
            state = 'normal'
        if hasattr(self, 'exchange_username_label'):
            self.exchange_username_label.config(state=state)
            self.exchange_username_entry.config(state=state)
            self.exchange_password_label.config(state=state)
            self.exchange_password_entry.config(state=state)

    def _create_email_settings_widgets(self):
        frame = self.email_settings_tab
        main_frame = ttk.LabelFrame(frame, text="Настройки отправки отчетов по Email", padding="10")
        main_frame.pack(fill='x', expand=True, pady=5, padx=5)

        general_email_frame = ttk.Frame(main_frame)
        general_email_frame.pack(fill='x', expand=True)

        ttk.Checkbutton(general_email_frame, text="Включить отправку отчетов по почте", variable=self.settings_vars['email_enabled']).grid(row=0, column=0, columnspan=3, sticky='w', pady=5)
        ttk.Label(general_email_frame, text="Email получателя:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(general_email_frame, textvariable=self.settings_vars['email_recipient'], width=50).grid(row=1, column=1, columnspan=2, sticky='ew', padx=5, pady=2)
        ttk.Label(general_email_frame, text="Тема письма:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(general_email_frame, textvariable=self.settings_vars['email_subject'], width=50).grid(row=2, column=1, columnspan=2, sticky='ew', padx=5, pady=2)
        ttk.Label(general_email_frame, text="Формат отчета:").grid(row=3, column=0, sticky='w', padx=5, pady=2)
        report_format_combo = ttk.Combobox(
            general_email_frame, textvariable=self.settings_vars['email_report_format'],
            values=["xlsx", "html", "csv"], state="readonly", width=10
        )
        report_format_combo.grid(row=3, column=1, sticky='w', padx=5, pady=2)
        general_email_frame.columnconfigure(1, weight=1)

        backend_frame = ttk.Frame(main_frame)
        backend_frame.pack(fill='x', pady=(10, 5))
        ttk.Label(backend_frame, text="Способ отправки:").pack(side='left', padx=5)
        ttk.Radiobutton(backend_frame, text="SMTP", variable=self.settings_vars['email_backend'], value="smtp", command=self._toggle_email_backend_widgets).pack(side='left', padx=5)
        ttk.Radiobutton(backend_frame, text="Exchange (EWS)", variable=self.settings_vars['email_backend'], value="exchange", command=self._toggle_email_backend_widgets).pack(side='left', padx=5)

        self.smtp_settings_frame = ttk.LabelFrame(main_frame, text="Настройки SMTP", padding="10")
        self.exchange_settings_frame = ttk.LabelFrame(main_frame, text="Настройки Exchange", padding="10")

        sf = self.smtp_settings_frame
        ttk.Label(sf, text="SMTP Сервер:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(sf, textvariable=self.settings_vars['email_smtp_server'], width=40).grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        port_frame = ttk.Frame(sf)
        port_frame.grid(row=1, column=1, sticky='ew', padx=5, pady=2)
        ttk.Label(sf, text="SMTP Порт:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        ttk.Spinbox(port_frame, from_=0, to=65535, textvariable=self.settings_vars['email_smtp_port'], width=10).pack(side='left')
        ttk.Label(port_frame, text="(587 для TLS, 465 для SSL)").pack(side='left', padx=10)
        ttk.Checkbutton(sf, text="Использовать TLS (STARTTLS)", variable=self.settings_vars['email_use_tls']).grid(row=2, column=1, sticky='w', pady=5, padx=5)
        ttk.Label(sf, text="Имя пользователя (логин):").grid(row=3, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(sf, textvariable=self.settings_vars['email_smtp_user'], width=40).grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        ttk.Label(sf, text="Пароль:").grid(row=4, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(sf, textvariable=self.settings_vars['email_smtp_password'], width=40, show='*').grid(row=4, column=1, sticky='ew', padx=5, pady=2)
        sf.columnconfigure(1, weight=1)

        ef = self.exchange_settings_frame

        auth_frame = ttk.Frame(ef)
        auth_frame.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ttk.Label(auth_frame, text="Аутентификация:").pack(side='left', padx=(5, 10))
        ttk.Radiobutton(
            auth_frame, text="Текущий пользователь Windows (без пароля)",
            variable=self.settings_vars['exchange_auth_mode'], value="current_user",
            command=self._toggle_exchange_creds_widgets
        ).pack(side='left', padx=5)
        ttk.Radiobutton(
            auth_frame, text="Указать пользователя и пароль",
            variable=self.settings_vars['exchange_auth_mode'], value="specific_user",
            command=self._toggle_exchange_creds_widgets
        ).pack(side='left', padx=5)
        ttk.Radiobutton(
            auth_frame, text="Kerberos",
            variable=self.settings_vars['exchange_auth_mode'], value="kerberos",
            command=self._toggle_exchange_creds_widgets
        ).pack(side='left', padx=5)

        self.exchange_autodiscover_check = ttk.Checkbutton(ef, text="Автообнаружение сервера (рекомендуется)", variable=self.settings_vars['exchange_autodiscover'], command=self._toggle_exchange_server_entry)
        self.exchange_autodiscover_check.grid(row=1, column=0, columnspan=2, sticky='w', pady=5)

        self.exchange_server_label = ttk.Label(ef, text="Сервер Exchange (если автообнаружение выкл.):")
        self.exchange_server_label.grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.exchange_server_entry = ttk.Entry(ef, textvariable=self.settings_vars['exchange_server'], width=40)
        self.exchange_server_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=2)

        self.exchange_url_label = ttk.Label(ef, text="Полный EWS URL (если ничего не помогает):")
        self.exchange_url_label.grid(row=3, column=0, sticky='w', padx=5, pady=2)
        self.exchange_url_entry = ttk.Entry(ef, textvariable=self.settings_vars['exchange_ews_url'], width=40)
        self.exchange_url_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=2)

        ttk.Label(ef, text="Email отправителя:").grid(row=4, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(ef, textvariable=self.settings_vars['exchange_email'], width=40).grid(row=4, column=1, sticky='ew', padx=5, pady=2)

        self.exchange_username_label = ttk.Label(ef, text="Имя пользователя (DOMAIN\\user):")
        self.exchange_username_label.grid(row=5, column=0, sticky='w', padx=5, pady=2)
        self.exchange_username_entry = ttk.Entry(ef, textvariable=self.settings_vars['exchange_username'], width=40)
        self.exchange_username_entry.grid(row=5, column=1, sticky='ew', padx=5, pady=2)

        self.exchange_password_label = ttk.Label(ef, text="Пароль:")
        self.exchange_password_label.grid(row=6, column=0, sticky='w', padx=5, pady=2)
        self.exchange_password_entry = ttk.Entry(ef, textvariable=self.settings_vars['exchange_password'], width=40, show='*')
        self.exchange_password_entry.grid(row=6, column=1, sticky='ew', padx=5, pady=2)

        ssl_frame = ttk.Frame(ef)
        ssl_frame.grid(row=7, column=0, columnspan=2, sticky='w', pady=(10,0))
        ttk.Checkbutton(ssl_frame, text="Отключить проверку SSL-сертификата", variable=self.settings_vars['exchange_verify_ssl']).pack(side='left')
        ttk.Label(ssl_frame, text="⚠️ (Небезопасно! Только для доверенных сетей)").pack(side='left', padx=5)

        ef.columnconfigure(1, weight=1)

        self.proxy_settings_frame = ttk.LabelFrame(main_frame, text="Настройки прокси-сервера (для Exchange)", padding="10")
        self.proxy_settings_frame.pack(fill='x', expand=True, pady=10)

        psf = self.proxy_settings_frame
        ttk.Checkbutton(psf, text="Использовать прокси-сервер для подключения к Exchange", variable=self.settings_vars['proxy_enabled']).grid(row=0, column=0, columnspan=2, sticky='w', pady=5)

        ttk.Label(psf, text="Адрес прокси (хост:порт):").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(psf, textvariable=self.settings_vars['proxy_address'], width=40).grid(row=1, column=1, sticky='ew', padx=5, pady=2)

        ttk.Label(psf, text="Имя пользователя прокси (если нужно):").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(psf, textvariable=self.settings_vars['proxy_user'], width=40).grid(row=2, column=1, sticky='ew', padx=5, pady=2)

        ttk.Label(psf, text="Пароль прокси (если нужно):").grid(row=3, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(psf, textvariable=self.settings_vars['proxy_password'], width=40, show='*').grid(row=3, column=1, sticky='ew', padx=5, pady=2)

        psf.columnconfigure(1, weight=1)

        test_button_frame = ttk.Frame(main_frame)
        test_button_frame.pack(fill='x', pady=(15, 5))

        self.test_email_button = ttk.Button(
            test_button_frame,
            text="📧 Отправить тестовое письмо",
            command=self.send_test_email
        )
        self.test_email_button.pack(side='left', padx=5)

        self.test_email_status_label = ttk.Label(test_button_frame, text="")
        self.test_email_status_label.pack(side='left', padx=10)

        self.after(100, self._toggle_email_backend_widgets)
        self.after(100, self._toggle_exchange_creds_widgets)

    def send_test_email(self):
        test_settings = {key: var.get() for key, var in self.settings_vars.items() if not key.endswith('_vars')}
        test_settings['email_enabled'] = True

        recipient = test_settings.get('email_recipient')
        if not recipient:
            messagebox.showerror("Ошибка", "Не указан Email получателя.", parent=self.email_settings_tab)
            return

        self.test_email_button.config(state='disabled')
        self.test_email_status_label.config(text="Отправка...", foreground="blue")

        thread = threading.Thread(
            target=self._send_test_email_thread,
            args=(test_settings, recipient),
            daemon=True
        )
        thread.start()

    def _send_test_email_thread(self, settings: Dict, recipient: str):
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            summary = (
                f"Это тестовое письмо от сканера документов.\n\n"
                f"Письмо отправлено: {timestamp}\n"
                f"Способ отправки: {settings.get('email_backend', 'N/A').upper()}\n\n"
                f"Если вы получили это письмо, значит, настройки почты корректны."
            )
            send_email_report(settings, summary, attachments=None)
            self.after(0, self._update_test_email_status, True, f"Успешно отправлено на {recipient}!")
        except Exception as e:
            logger.error(f"Ошибка при отправке тестового письма: {e}", exc_info=True)
            self.after(0, self._update_test_email_status, False, f"Ошибка! См. лог grif_scan.log")

    def _update_test_email_status(self, success: bool, message: str):
        if success:
            self.test_email_status_label.config(text=message, foreground="green")
        else:
            self.test_email_status_label.config(text=message, foreground="red")
        self.test_email_button.config(state='normal')

    def _toggle_email_backend_widgets(self):
        backend = self.settings_vars['email_backend'].get()
        if backend == 'smtp':
            self.smtp_settings_frame.pack(fill='x', expand=True, pady=5)
            self.exchange_settings_frame.pack_forget()
        elif backend == 'exchange':
            self.smtp_settings_frame.pack_forget()
            self.exchange_settings_frame.pack(fill='x', expand=True, pady=5)
            if not EXCHANGELIB_AVAILABLE:
                messagebox.showwarning(
                    "Зависимость отсутствует",
                    "Для использования Exchange необходима библиотека 'exchangelib'.\n\n"
                    "Установите ее командой: pip install exchangelib",
                    parent=self.email_settings_tab
                )
        self._toggle_exchange_server_entry()

    def _toggle_exchange_server_entry(self):
        if self.settings_vars['email_backend'].get() == 'exchange':
            is_autodiscover = self.settings_vars['exchange_autodiscover'].get()
            manual_state = 'disabled' if is_autodiscover else 'normal'
            self.exchange_server_label.config(state=manual_state)
            self.exchange_server_entry.config(state=manual_state)
            self.exchange_url_label.config(state='normal')
            self.exchange_url_entry.config(state='normal')

    def create_scheduler_tab(self):
        main_frame = ttk.Frame(self.scheduler_tab, padding="10")
        main_frame.pack(fill="both", expand=True)

        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x', pady=5)
        ttk.Button(buttons_frame, text="➕ Добавить задачу", command=self.add_or_edit_schedule_job).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="✏️ Редактировать", command=lambda: self.add_or_edit_schedule_job(edit_mode=True)).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="❌ Удалить", command=self.remove_schedule_job).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="🔄 Обновить статус", command=self.populate_scheduler_tree).pack(side='right', padx=5)

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill='both', expand=True, pady=5)

        self.scheduler_tree = ttk.Treeview(
            tree_frame,
            columns=("enabled", "type", "target", "schedule", "next_run"),
            show="headings"
        )
        self.scheduler_tree.heading("enabled", text="Вкл.", anchor='center')
        self.scheduler_tree.heading("type", text="Тип сканирования")
        self.scheduler_tree.heading("target", text="Цель")
        self.scheduler_tree.heading("schedule", text="Расписание")
        self.scheduler_tree.heading("next_run", text="Следующий запуск")

        self.scheduler_tree.column("enabled", width=40, anchor='center', stretch=False)
        self.scheduler_tree.column("type", width=150)
        self.scheduler_tree.column("target", width=250)
        self.scheduler_tree.column("schedule", width=200)
        self.scheduler_tree.column("next_run", width=150)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.scheduler_tree.yview)
        self.scheduler_tree.configure(yscrollcommand=scrollbar.set)

        self.scheduler_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.populate_scheduler_tree()

    def get_next_run_time(self, job: Dict) -> Optional[datetime]:
        now = datetime.now()
        last_run = datetime.fromisoformat(job['last_run']) if job.get('last_run') else None

        if job['schedule_mode'] == 'interval':
            interval = timedelta(minutes=job.get('interval_minutes', 60))
            if not last_run:
                return now
            scheduled = last_run + interval
            return now if scheduled <= now else scheduled
        elif job['schedule_mode'] == 'daily':
            try:
                target_time = dt_time.fromisoformat(job.get('daily_time', '02:00'))
                target_today = now.replace(hour=target_time.hour, minute=target_time.minute, second=0, microsecond=0)
                if last_run and last_run.date() == now.date() and last_run.time() >= target_time:
                    return target_today + timedelta(days=1)
                return now if now >= target_today else target_today
            except (ValueError, TypeError):
                return None
        return None

    def populate_scheduler_tree(self):
        for item in self.scheduler_tree.get_children():
            self.scheduler_tree.delete(item)

        scan_type_map_display = {"single": "Один каталог", "list": "По списку", "network": "Сеть"}

        for job in self.settings.get('scheduler_jobs', []):
            job_id = job.get('id')
            enabled_str = "✔️" if job.get('enabled') else "❌"
            scan_type_str = scan_type_map_display.get(job.get('scan_type'), "Неизвестно")
            target_str = job.get('target', 'Не указано')

            schedule_str = ""
            if job.get('schedule_mode') == 'interval':
                schedule_str = f"Каждые {job.get('interval_minutes', '?')} мин."
            elif job.get('schedule_mode') == 'daily':
                schedule_str = f"Ежедневно в {job.get('daily_time', '??:??')}"

            next_run_dt = self.get_next_run_time(job) if job.get('enabled') else None
            next_run_str = next_run_dt.strftime('%Y-%m-%d %H:%M:%S') if next_run_dt else "Отключено"

            self.scheduler_tree.insert("", "end", iid=job_id, values=(
                enabled_str,
                scan_type_str,
                target_str,
                schedule_str,
                next_run_str
            ))

    def add_or_edit_schedule_job(self, edit_mode=False):
        job_data = None
        if edit_mode:
            selected_items = self.scheduler_tree.selection()
            if not selected_items:
                messagebox.showwarning("Нет выбора", "Выберите задачу для редактирования.", parent=self.scheduler_tab)
                return
            job_id = selected_items[0]
            job_data = next((job for job in self.settings['scheduler_jobs'] if job['id'] == job_id), None)
            if not job_data:
                messagebox.showerror("Ошибка", "Не удалось найти данные для выбранной задачи.", parent=self.scheduler_tab)
                return
        SchedulerJobEditor(self, self.settings, self.save_all_settings, job_data)

    def remove_schedule_job(self):
        selected_items = self.scheduler_tree.selection()
        if not selected_items:
            messagebox.showwarning("Нет выбора", "Выберите задачу для удаления.", parent=self.scheduler_tab)
            return
        job_id = selected_items[0]
        if messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить выбранную задачу?", parent=self.scheduler_tab):
            self.settings['scheduler_jobs'] = [job for job in self.settings['scheduler_jobs'] if job.get('id') != job_id]
            self.save_all_settings()
            logger.info(f"Задача планировщика {job_id} удалена.")

    def _update_tika_status_callback(self, status: str):
        def update():
            try:
                if not self.winfo_exists() or not hasattr(self, 'tika_status_frame') or not self.tika_status_frame.winfo_exists():
                    return
                
                color = "grey"
                text = f"Tika: {status}"
                
                is_tika_on = self.settings_vars['use_tika'].get()
                button_state = "normal" if is_tika_on else "disabled"
                
                if status == "Запущен":
                    color = "#2ECC40"
                    self.restart_tika_btn.config(state=button_state)
                elif status == "Ошибка":
                    color = "#FF4136"
                    self.restart_tika_btn.config(state=button_state)
                elif status == "Запуск...":
                    color = "#FFDC00"
                    self.restart_tika_btn.config(state="disabled")
                elif status == "Остановлен":
                    color = "grey"
                    self.restart_tika_btn.config(state=button_state)
                    
                self.tika_status_light.itemconfig(self.tika_light_circle, fill=color, outline=color)
                self.tika_status_label.config(text=text)
            except Exception:
                pass
            
        self.after(0, update)

    def _restart_tika_server_manually(self):
        logger.info("Запрошен ручной перезапуск Tika Server.")
        self.tika_manager.stop_server()
        self.after(500, self.tika_manager.start_server_async)

    def select_directory(self):
        path = filedialog.askdirectory(initialdir=self.scan_path_var.get(), title="Выберите папку для сканирования")
        if path: self.scan_path_var.set(os.path.normpath(path))

    def select_unrar_path(self):
        path = filedialog.askopenfilename(title="Выберите UnRAR.exe", filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if path: self.settings_vars['unrar_path'].set(os.path.normpath(path))

    def select_tesseract_path(self):
        path = filedialog.askopenfilename(title="Выберите tesseract.exe", filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if path: self.settings_vars['tesseract_path'].set(os.path.normpath(path))

    def select_tika_jar_path(self):
        path = filedialog.askopenfilename(title="Выберите tika-server.jar", filetypes=[("JAR files", "*.jar"), ("All files", "*.*")])
        if path: self.settings_vars['tika_jar_path'].set(os.path.normpath(path))

    def select_report_output_directory(self):
        path = filedialog.askdirectory(initialdir=self.settings_vars['report_output_dir'].get(), title="Выберите каталог для сохранения отчета")
        if path: self.settings_vars['report_output_dir'].set(os.path.normpath(path))

    def select_directory_list_file(self):
        path = filedialog.askopenfilename(title="Выберите файл со списком каталогов", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            normalized_path = os.path.normpath(path)
            self.directory_list_file_var.set(normalized_path)
            self.settings_vars['directory_list_file'].set(normalized_path)
            self.load_directory_list()

    def select_directory_list_file_for_settings(self):
        path = filedialog.askopenfilename(title="Выберите файл со списком каталогов", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            self.settings_vars['directory_list_file'].set(os.path.normpath(path))

    def select_network_list_file(self, for_settings=False):
        path = filedialog.askopenfilename(title="Выберите файл со списком сетей", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            normalized_path = os.path.normpath(path)
            if for_settings:
                self.settings_vars['network_list_file'].set(normalized_path)
            else:
                self.network_list_file_var.set(normalized_path)
                self.load_network_list()

    def load_directory_list(self):
        filepath = self.directory_list_file_var.get()
        self.dir_list_tree.delete(*self.dir_list_tree.get_children())
        self.accessible_directories = []

        if not filepath or not os.path.exists(filepath):
            msg = "Путь к файлу со списком каталогов не указан или файл не найден."
            logger.warning(msg)
            self.update_status(msg)
            return

        directories = read_directories_from_file(filepath)
        if not directories:
            logger.info("Файл со списком каталогов пуст или не содержит корректных путей.")
            self.update_status("Файл пуст или не содержит корректных путей.")
            return

        for directory_path in directories:
            self.dir_list_tree.insert("", "end", values=(directory_path, "Ожидание проверки..."))
        self.update_status(f"Загружено {len(directories)} каталогов из файла.")
        logger.info(f"Загружено {len(directories)} каталогов из файла: {filepath}")
        self.settings['directory_list_file'] = filepath

    def load_network_list(self):
        filepath = self.network_list_file_var.get()
        self.net_list_tree.delete(*self.net_list_tree.get_children())

        if not filepath or not os.path.exists(filepath):
            msg = "Путь к файлу со списком сетей не указан или файл не найден."
            logger.warning(msg)
            self.update_status(msg)
            return

        networks = read_networks_from_file(filepath)
        if not networks:
            logger.info("Файл со списком сетей пуст.")
            self.update_status("Файл со списком сетей пуст.")
            return

        for net_range in networks:
            self.net_list_tree.insert("", "end", values=(net_range,))
        self.update_status(f"Загружено {len(networks)} сетей/хостов из файла.")
        logger.info(f"Загружено {len(networks)} сетей/хостов из файла: {filepath}")

    def check_all_directories_accessibility(self):
        if not self.scan_lock.acquire(blocking=False):
            messagebox.showwarning("Внимание", "Другая операция сканирования или проверки уже запущена.", parent=self)
            return

        try:
            filepath = self.directory_list_file_var.get()
            if not filepath or not os.path.exists(filepath):
                messagebox.showwarning("Нет файла", "Сначала выберите файл со списком каталогов.")
                self.scan_lock.release()
                return

            directories = read_directories_from_file(filepath)
            if not directories:
                messagebox.showinfo("Нет каталогов", "Файл со списком каталогов пуст или не содержит путей.")
                self.scan_lock.release()
                return

            self.update_status("Начало проверки доступности каталогов...")
            self.set_controls_state(is_scanning=True)
            self.progress_bar.config(value=0)

            self.dir_list_tree.delete(*self.dir_list_tree.get_children())
            self.accessible_directories.clear()

            items_to_check = []
            for path in directories:
                iid = self.dir_list_tree.insert("", "end", values=(path, "Ожидание..."))
                items_to_check.append((path, iid))

            self.dir_list_tree.tag_configure("accessible", foreground="green")
            self.dir_list_tree.tag_configure("inaccessible", foreground="red")
            self.current_stop_event.clear()
            self.current_scan_thread = threading.Thread(
                target=self._run_accessibility_check_thread,
                args=(items_to_check,),
                daemon=True
            )
            self.current_scan_thread.start()
        except Exception:
            self.scan_lock.release()
            raise

    def _run_accessibility_check_thread(self, items_to_check: List[Tuple[str, str]]):
        try:
            total = len(items_to_check)
            self.after(0, lambda: self.progress_bar.config(value=0, maximum=total))

            for i, (path, iid) in enumerate(items_to_check):
                if self.current_stop_event.is_set():
                    self.after(0, self.update_status, "Проверка доступности отменена.")
                    break

                self.after(0, self.update_status, f"Проверка: {os.path.basename(path)}")

                _, is_accessible, status_msg = check_directory_accessibility(path)

                log_msg = f"Проверка доступности: '{path}'. Статус: {status_msg}"
                if is_accessible:
                    logger.info(log_msg)
                else:
                    logger.warning(log_msg)

                self.after(0, self._update_single_dir_status, iid, path, is_accessible, status_msg)
                self.after(0, lambda val=i + 1: self.progress_bar.config(value=val))

            accessible_count = len(self.accessible_directories)
            final_msg = f"Проверка завершена. Доступно каталогов: {accessible_count} из {total}."
            if self.current_stop_event.is_set():
                final_msg = "Проверка доступности отменена."

            self.after(0, self.update_status, final_msg)
            logger.info(final_msg)
            if not self.accessible_directories and not self.current_stop_event.is_set():
                self.after(0, lambda: messagebox.showwarning("Нет доступных", "Ни один каталог из списка не доступен для сканирования."))
        except Exception as e:
            logger.error(f"Ошибка в потоке проверки доступности: {e}", exc_info=True)
        finally:
            self.after(0, self.set_controls_state, False)
            self.scan_lock.release()

    def _update_single_dir_status(self, iid: str, path: str, is_accessible: bool, status_msg: str):
        if not self.dir_list_tree.exists(iid):
            return
        tag = "accessible" if is_accessible else "inaccessible"
        self.dir_list_tree.item(iid, values=(path, status_msg), tags=(tag,))
        if is_accessible:
            self.accessible_directories.append(path)

    def save_all_settings(self, silent=False):
        keys_to_exclude = ['text_scanned_extensions_vars', 'ocr_scanned_extensions_vars']
        new_settings = {key: var.get() for key, var in self.settings_vars.items() if key not in keys_to_exclude}

        text_patterns = self.text_patterns_text.get("1.0", tk.END).strip().split("\n")
        new_settings['text_grif_patterns'] = [p for p in text_patterns if p]
        ocr_patterns = self.ocr_patterns_text.get("1.0", tk.END).strip().split("\n")
        new_settings['ocr_grif_patterns'] = [p for p in ocr_patterns if p]

        exclude_patterns = self.exclude_patterns_text.get("1.0", tk.END).strip().split("\n")
        new_settings['exclude_path_patterns'] = [p for p in exclude_patterns if p.strip()]

        new_settings['text_scanned_extensions'] = {
            ext: var.get() for ext, var in self.settings_vars['text_scanned_extensions_vars'].items()
        }
        new_settings['ocr_scanned_extensions'] = {
            ext: var.get() for ext, var in self.settings_vars['ocr_scanned_extensions_vars'].items()
        }

        self.settings.update(new_settings)

        if keyring:
            try:
                smtp_password = self.settings.get('email_smtp_password', '')
                exchange_password = self.settings.get('exchange_password', '')
                proxy_password = self.settings.get('proxy_password', '')

                keyring.set_password('grif_scan_app', 'smtp_password', smtp_password)
                keyring.set_password('grif_scan_app', 'exchange_password', exchange_password)
                keyring.set_password('grif_scan_app', 'proxy_password', proxy_password)

                logger.info("Пароли SMTP, Exchange, Proxy сохранены в системное хранилище.")

                self.settings.pop('email_smtp_password', None)
                self.settings.pop('exchange_password', None)
                self.settings.pop('proxy_password', None)

            except Exception as e:
                logger.error(f"Не удалось сохранить пароли в системное хранилище: {e}")
                if not silent:
                    messagebox.showerror("Ошибка безопасности",
                                         f"Не удалось сохранить пароли в системное хранилище.\n"
                                         f"Они не будут записаны в config.json.\n\nОшибка: {e}")
        else:
            logger.warning("Библиотека keyring не найдена. Пароли НЕ будут сохранены в config.json (безопасно).")
            self.settings.pop('email_smtp_password', None)
            self.settings.pop('exchange_password', None)
            self.settings.pop('proxy_password', None)

        _save_config_to_file(self.settings, CONFIG_FILE)

        self.config = Config(self.settings)

        logger.info("Конфигурация сохранена в config.json")

        if not silent and self.winfo_exists():
            messagebox.showinfo("Сохранено", "Все настройки успешно сохранены.")

        self._apply_logging_settings()
        self.populate_scheduler_tree()

    def _gather_all_ui_settings(self):
        logger.debug("Сбор всех настроек из UI...")
        keys_to_exclude = ['text_scanned_extensions_vars', 'ocr_scanned_extensions_vars']
        for key, var in self.settings_vars.items():
            if key not in keys_to_exclude:
                self.settings[key] = var.get()

        text_patterns = self.text_patterns_text.get("1.0", tk.END).strip().split("\n")
        self.settings['text_grif_patterns'] = [p for p in text_patterns if p]

        ocr_patterns = self.ocr_patterns_text.get("1.0", tk.END).strip().split("\n")
        self.settings['ocr_grif_patterns'] = [p for p in ocr_patterns if p]

        exclude_patterns = self.exclude_patterns_text.get("1.0", tk.END).strip().split("\n")
        self.settings['exclude_path_patterns'] = [p for p in exclude_patterns if p.strip()]

        self.settings['text_scanned_extensions'] = {
            ext: var.get() for ext, var in self.settings_vars['text_scanned_extensions_vars'].items()
        }
        self.settings['ocr_scanned_extensions'] = {
            ext: var.get() for ext, var in self.settings_vars['ocr_scanned_extensions_vars'].items()
        }

        if hasattr(self, 'scan_path_var'): self.settings['scan_path'] = self.scan_path_var.get()
        if hasattr(self, 'directory_list_file_var'): self.settings['directory_list_file'] = self.directory_list_file_var.get()
        if hasattr(self, 'net_ip_range_var'): self.settings['ip_range'] = self.net_ip_range_var.get()

    def _apply_settings_to_ui(self):
        logger.debug("Применение настроек к UI...")
        keys_to_exclude = ['text_scanned_extensions_vars', 'ocr_scanned_extensions_vars']
        for key, var in self.settings_vars.items():
            if key in self.settings and key not in keys_to_exclude:
                var.set(self.settings[key])

        for ext, var in self.settings_vars['text_scanned_extensions_vars'].items():
            var.set(self.settings.get('text_scanned_extensions', {}).get(ext, True))

        for ext, var in self.settings_vars['ocr_scanned_extensions_vars'].items():
            var.set(self.settings.get('ocr_scanned_extensions', {}).get(ext, True))

        self.text_patterns_text.delete("1.0", tk.END)
        self.text_patterns_text.insert(tk.END, "\n".join(self.settings.get('text_grif_patterns', [])))
        self.ocr_patterns_text.delete("1.0", tk.END)
        self.ocr_patterns_text.insert(tk.END, "\n".join(self.settings.get('ocr_grif_patterns', [])))

        if hasattr(self, 'exclude_patterns_text'):
            self.exclude_patterns_text.delete("1.0", tk.END)
            self.exclude_patterns_text.insert(tk.END, "\n".join(self.settings.get('exclude_path_patterns', [])))

        if hasattr(self, 'scan_path_var'): self.scan_path_var.set(self.settings.get('scan_path', ''))
        if hasattr(self, 'directory_list_file_var'): self.directory_list_file_var.set(self.settings.get('directory_list_file', ''))
        if hasattr(self, 'net_ip_range_var'): self.net_ip_range_var.set(self.settings.get('ip_range', ''))

        self._apply_logging_settings()
        self.populate_scheduler_tree()

    def save_configuration_as(self):
        filepath = filedialog.asksaveasfilename(
            title="Сохранить профиль конфигурации",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            defaultextension=".json",
            initialdir=os.getcwd()
        )
        if not filepath: return

        self._gather_all_ui_settings()
        _save_config_to_file(self.settings, filepath)
        logger.info(f"Конфигурация сохранена в файл: {filepath}")
        messagebox.showinfo("Успешно", f"Конфигурация успешно сохранена в:\n{filepath}")
        self.current_config_file_var.set(os.path.basename(filepath))

    def load_configuration_from_file(self):
        filepath = filedialog.askopenfilename(
            title="Загрузить профиль конфигурации",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=os.getcwd()
        )
        if not filepath: return

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                loaded_settings = json.load(f)
            defaults = get_default_settings()
            defaults.update(loaded_settings)
            self.settings = defaults
            self._apply_settings_to_ui()
            logger.info(f"Конфигурация успешно загружена из файла: {filepath}")
            messagebox.showinfo("Успешно", "Конфигурация успешно загружена.")
            self.current_config_file_var.set(os.path.basename(filepath))
        except (json.JSONDecodeError, IOError, TypeError) as e:
            logger.error(f"Не удалось загрузить конфигурацию из {filepath}: {e}")
            messagebox.showerror("Ошибка загрузки", f"Не удалось загрузить или прочитать файл конфигурации.\n\nОшибка: {e}")

    def restore_default_settings(self):
        if messagebox.askyesno("Подтверждение", "Вы уверены, что хотите сбросить все настройки до заводских значений?"):
            self.settings = get_default_settings()
            self._apply_settings_to_ui()
            logger.info("Настройки сброшены к значениям по умолчанию.")
            self.save_all_settings(silent=True)
            self.current_config_file_var.set("по умолчанию (config.json)")

    def start_scan(self):
        self.save_all_settings(silent=True)
        scan_mode = self.scan_mode_var.get()
        if scan_mode == 'single':
            self._start_scan_single_path()
        elif scan_mode == 'list':
            self._start_scan_from_list()
        elif scan_mode == 'network':
            self._start_network_scan()

    def stop_scan(self):
        if self.current_scan_thread and self.current_scan_thread.is_alive():
            self.current_stop_event.set()
            logger.info("Попытка остановить сканирование...")
            self.update_status("Остановка сканирования...")
        else:
            self.set_controls_state(is_scanning=False)

    def open_report(self):
        scan_mode = self.scan_mode_var.get()
        path_to_open = ""
        if scan_mode == 'network':
            path_to_open = self.network_report_dir
        else:
            path_to_open = self.report_path
        self._platform_specific_open(path_to_open)

    def _clear_results(self):
        scan_mode = self.scan_mode_var.get()
        if scan_mode == 'network':
            self.network_tree.delete(*self.network_tree.get_children())
            self.found_network_results.clear()
            logger.info("Результаты сканирования сети очищены.")
        else:
            self.results_tree.delete(*self.results_tree.get_children())
            self.scan_results.clear()
            logger.info("Результаты сканирования документов очищены.")
        self.update_status("Результаты очищены.")

    def _start_scan_single_path(self, is_scheduled=False, scheduled_job=None):
        if not self.scan_lock.acquire(blocking=False):
            msg = "Сканирование уже запущено."
            logger.warning(f"Попытка запуска (single): {msg}")
            if not is_scheduled: messagebox.showwarning("Внимание", msg, parent=self)
            return

        try:
            scan_path = self.scan_path_var.get()
            if is_scheduled and scheduled_job:
                scan_path = scheduled_job.get('target')

            scan_path = os.path.normpath(scan_path)

            if not os.path.isdir(scan_path):
                msg = f"Путь для сканирования не является директорией: {scan_path}"
                if not is_scheduled: messagebox.showerror("Ошибка пути", msg)
                logger.error(msg)
                self.scan_lock.release()
                return

            if not is_scheduled:
                self.settings['scan_path'] = scan_path

            self._clear_results()
            self.current_stop_event.clear()
            self.set_controls_state(is_scanning=True)

            snapshot_settings = self.settings.copy()
            snapshot_config = Config(snapshot_settings)

            self.current_scan_thread = threading.Thread(
                target=self.run_doc_scan_thread,
                args=([scan_path], snapshot_config, is_scheduled, scheduled_job),
                daemon=True
            )
            self.current_scan_thread.start()
        except Exception:
            self.scan_lock.release()
            raise

    def _start_scan_from_list(self, is_scheduled=False, scheduled_job=None):
        if not self.scan_lock.acquire(blocking=False):
            msg = "Сканирование уже запущено."
            logger.warning(f"Попытка запуска (list): {msg}")
            if not is_scheduled: messagebox.showwarning("Внимание", msg, parent=self)
            return

        try:
            directories_to_scan = []
            if self.accessible_directories and not is_scheduled:
                logger.info("Используется список предварительно проверенных доступных каталогов.")
                directories_to_scan = self.accessible_directories
            else:
                logger.info("Предварительная проверка не использовалась. Чтение полного списка из файла.")
                filepath = self.directory_list_file_var.get()
                if is_scheduled and scheduled_job:
                    filepath = scheduled_job.get('target')
                if not filepath or not os.path.exists(filepath):
                    msg = f"Файл со списком каталогов не найден: {filepath}"
                    if not is_scheduled: messagebox.showerror("Ошибка", msg)
                    logger.warning(msg)
                    self.scan_lock.release()
                    return
                directories_to_scan = read_directories_from_file(filepath)

            if not directories_to_scan:
                msg = "Список каталогов для сканирования пуст."
                if not is_scheduled: messagebox.showwarning("Пустой список", msg)
                logger.info(msg)
                self.scan_lock.release()
                return

            self._clear_results()
            self.current_stop_event.clear()
            self.set_controls_state(is_scanning=True)

            snapshot_settings = self.settings.copy()
            snapshot_config = Config(snapshot_settings)

            self.current_scan_thread = threading.Thread(
                target=self.run_doc_scan_thread,
                args=(directories_to_scan, snapshot_config, is_scheduled, scheduled_job),
                daemon=True
            )
            self.current_scan_thread.start()
        except Exception:
            self.scan_lock.release()
            raise

    def _start_network_scan(self, is_scheduled=False, scheduled_job=None):
        if not self.scan_lock.acquire(blocking=False):
            msg = "Сканирование уже запущено."
            logger.warning(f"Попытка запуска (network): {msg}")
            if not is_scheduled: messagebox.showwarning("Внимание", msg, parent=self)
            return

        try:
            ip_ranges = []
            if is_scheduled and scheduled_job:
                ip_ranges.append(scheduled_job.get('target'))
            else:
                input_mode = self.network_input_mode_var.get()
                if input_mode == 'manual':
                    ip_range = self.net_ip_range_var.get()
                    if not ip_range:
                        messagebox.showerror("Ошибка", "Не указан диапазон IP для сканирования.", parent=self)
                        self.scan_lock.release()
                        return
                    ip_ranges.append(ip_range)
                    self.settings['ip_range'] = ip_range
                else:
                    filepath = self.network_list_file_var.get()
                    if not filepath or not os.path.exists(filepath):
                        messagebox.showerror("Ошибка", "Файл со списком сетей не выбран или не существует.", parent=self)
                        self.scan_lock.release()
                        return
                    ip_ranges = read_networks_from_file(filepath)
                    if not ip_ranges:
                        messagebox.showinfo("Пустой список", "Файл со списком сетей пуст.", parent=self)
                        self.scan_lock.release()
                        return
                    self.settings['network_list_file'] = filepath

            self._clear_results()
            self.current_stop_event.clear()
            self.set_controls_state(is_scanning=True)

            self.current_scan_thread = threading.Thread(
                target=self._run_network_scan_thread,
                args=(ip_ranges, is_scheduled, scheduled_job),
                daemon=True
            )
            self.current_scan_thread.start()
        except Exception:
            self.scan_lock.release()
            raise

    def _file_producer_thread(self, paths_to_scan: List[str], file_queue: queue.Queue, stop_event: threading.Event, exclude_regex: Optional[re.Pattern], config_snapshot: Config):
        try:
            use_filter = config_snapshot.get('use_date_filter', False)
            days = config_snapshot.get('date_filter_days', 30)

            text_exts = {ext for ext, enabled in config_snapshot.get('text_scanned_extensions', {}).items() if enabled}
            ocr_exts = {ext for ext, enabled in config_snapshot.get('ocr_scanned_extensions', {}).items() if enabled}
            enabled_extensions = tuple(text_exts.union(ocr_exts))

            if not enabled_extensions:
                logger.warning("Поток-сборщик файлов остановлен: не выбрано ни одного расширения.")
                return

            found_count = 0
            last_status_update = time.monotonic()
            for path in paths_to_scan:
                if stop_event.is_set(): break
                for root, _, files in os.walk(path, topdown=True):
                    if stop_event.is_set(): break
                    for f in files:
                        if stop_event.is_set(): break
                        if f.lower().endswith(enabled_extensions) and not f.startswith('~$') and not f.startswith('._'):
                            filepath = os.path.normpath(os.path.join(root, f))
                            if exclude_regex and exclude_regex.search(filepath):
                                logger.debug(f"Пропущен (фильтр исключений): {filepath}")
                                continue
                            
                            should_add = True
                            if use_filter:
                                try:
                                    mod_time = datetime.fromtimestamp(os.path.getmtime(filepath))
                                    if mod_time < (datetime.now() - timedelta(days=days)):
                                        should_add = False
                                except OSError:
                                    should_add = False
                            
                            if should_add:
                                file_queue.put(filepath)
                                found_count += 1
                                now = time.monotonic()
                                if now - last_status_update > 0.3:
                                    last_status_update = now
                                    self.after(0, self.update_status, f"Поиск файлов... Найдено: {found_count}")
        except Exception as e:
            logger.error(f"Ошибка в потоке-сборщике файлов: {e}", exc_info=True)
        finally:
            logger.info(f"Поток-сборщик файлов завершил работу. Найдено всего: {found_count}")
            self.after(0, self.update_status, f"Поиск завершен. Найдено всего файлов: {found_count}")
            file_queue.put(None)

    def _update_gui_with_result(self, results_list: List[Dict]):
        if not self.results_tree.winfo_exists():
            return
        row_count = len(self.results_tree.get_children())
        last_inserted_iid = None
        for result in results_list:
            tag = 'oddrow' if row_count % 2 != 0 else 'evenrow'
            row_count += 1

            def format_display_path(res):
                return f"{res.get('archive_path')} -> {res.get('filename')}" if res.get('archive_path') else res.get('path', '')

            display_values = (
                format_display_path(result),
                result.get('text_grifs_display', ''),
                result.get('ocr_grifs_display', ''),
                result.get('owner', 'N/A'),
                result.get('size_mb', 'N/A')
            )
            new_iid = self.results_tree.insert("", "end", iid=result['path'], tags=(tag,), values=display_values)
            last_inserted_iid = new_iid
        if last_inserted_iid:
            self.results_tree.see(last_inserted_iid)

    def run_doc_scan_thread(self, paths_to_scan: List[str], config_snapshot: Config, is_scheduled=False, scheduled_job=None):
        start_time = time.monotonic()
        try:
            self._load_hash_cache()
            config = config_snapshot

            if config.ocr_enabled and not config.tesseract_cmd:
                logger.warning("OCR включен, но Tesseract не настроен. Сканирование без OCR.")

            self.after(0, lambda: self.progress_bar.config(mode='indeterminate'))
            self.after(0, lambda: self.progress_bar.start(10))

            file_queue = queue.Queue(maxsize=5000)
            exclude_regex = config.exclude_path_regex
            producer = threading.Thread(target=self._file_producer_thread, args=(paths_to_scan, file_queue, self.current_stop_event, exclude_regex, config), daemon=True)
            producer.start()

            processed_count = 0
            max_workers = self.settings.get('max_workers', os.cpu_count() or 4)

            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {}
                producer_finished = False
                last_status_update_time = 0
                max_gui_rows = 1000
                results_added_to_gui = 0
                
                while not producer_finished or futures:
                    if self.current_stop_event.is_set():
                        for f in futures:
                            f.cancel()
                        break

                    try:
                        while len(futures) < max_workers * 2:
                            try:
                                filepath = file_queue.get_nowait()
                                if filepath is None:
                                    producer_finished = True
                                    break
                                future = executor.submit(process_file, filepath, config, self.hash_cache, self.hash_lock, self.current_stop_event)
                                futures[future] = filepath
                            except queue.Empty:
                                if not producer.is_alive():
                                    producer_finished = True
                                break
                    except Exception as e:
                        logger.error(f"Ошибка при отправке файла в пул: {e}")

                    if futures:
                        done, _ = concurrent.futures.wait(futures.keys(), timeout=0.03, return_when=concurrent.futures.FIRST_COMPLETED)
                        for future in done:
                            original_filepath = futures.pop(future)
                            try:
                                results_tuple = future.result()
                                results_list, file_hash, scanner_name = results_tuple
                                norm_path = os.path.normpath(original_filepath)
                                if file_hash:
                                    with self.hash_lock:
                                        self.hash_cache[norm_path] = file_hash

                                if results_list:
                                    self.scan_results.extend(results_list)
                                    if self.settings.get('show_results_in_gui', True):
                                        if results_added_to_gui < max_gui_rows:
                                            self.after(0, self._update_gui_with_result, results_list)
                                            results_added_to_gui += 1
                                        elif results_added_to_gui == max_gui_rows:
                                            self.after(0, self.update_status, "Лимит отображения в GUI достигнут.")
                                            results_added_to_gui += 1
                            except Exception as e:
                                logger.error(f"Ошибка при получении результата для {original_filepath}: {e}")
                            finally:
                                processed_count += 1
                                current_time = time.monotonic()
                                if current_time - last_status_update_time > 0.3:
                                    last_status_update_time = current_time
                                    if producer_finished:
                                        total_val = len(futures) + processed_count
                                        self.after(0, self.progress_bar.config, {'mode': 'determinate', 'maximum': total_val, 'value': processed_count})
                                        self.after(0, self.update_status, f"Обработка: {processed_count}/{total_val} ({os.path.basename(original_filepath)})")
                                    else:
                                        self.after(0, self.update_status, f"Найдено... Обработано {processed_count}")
                    else:
                        time.sleep(0.04)

            if self.current_stop_event.is_set():
                logger.info("Сканирование остановлено пользователем.")
                self.after(0, self.update_status, "Сканирование остановлено.")
            else:
                logger.info("Сканирование завершено.")
                if self.scan_results:
                    self.after(0, self.update_status, "Генерация отчета в фоновом режиме...")
                    report_output_dir = self.settings.get('report_output_dir', os.path.expanduser("~"))
                    os.makedirs(report_output_dir, exist_ok=True)
                    scan_type = self.scan_mode_var.get()
                    report_suffix = f"{scan_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    base_filepath = os.path.join(report_output_dir, f"grif_scan_report_{report_suffix}")

                    report_thread = threading.Thread(
                        target=self._generate_reports_in_background,
                        args=(
                            self.scan_results.copy(),
                            base_filepath,
                            processed_count,
                            paths_to_scan,
                            is_scheduled
                        ),
                        daemon=True
                    )
                    report_thread.start()
                else:
                    self.after(0, self.update_status, "Сканирование завершено. Документы с грифами не найдены.")
        except Exception as e:
            logger.critical(f"Критическая ошибка в потоке сканирования: {e}", exc_info=True)
            self.after(0, self.update_status, f"Ошибка: {e}")
            if not is_scheduled:
                messagebox.showerror("Критическая ошибка", f"Произошла ошибка: {e}\nСм. логи.")
        finally:
            self.after(0, self.progress_bar.stop)
            self.after(0, lambda: self.progress_bar.config(mode='determinate', value=self.progress_bar['maximum']))
            if self.settings.get('play_sound_on_finish', True):
                self.after(0, self.bell)
            cleanup_thread = threading.Thread(target=self._cleanup_in_background, daemon=True)
            cleanup_thread.start()
            if self.settings.get('use_hash_cache'):
                logger.info("Запуск фонового сохранения кэша...")
                with self.hash_lock:
                    cache_copy = self.hash_cache.copy()
                cache_thread = threading.Thread(
                    target=self._save_hash_cache_in_background,
                    args=(cache_copy,),
                    daemon=True
                )
                cache_thread.start()
            duration_str = format_duration(time.monotonic() - start_time)
            total_findings = len(self.scan_results)
            final_msg = f"Сканирование завершено за {duration_str}. Всего обработано файлов: {processed_count}. Найдено объектов с нарушениями: {total_findings}"
            logger.info(f"Итог сканирования: {final_msg}")
            self.after(0, self.update_status, final_msg)
            self.after(0, self.set_controls_state, False)
            self.scan_lock.release()

    def _save_hash_cache_in_background(self, hash_cache_copy: Dict):
        if not self.settings.get('use_hash_cache'):
            logger.debug("Фоновое сохранение кэша пропущено (режим кэширования отключен)")
            return
        data_to_save = {
            "config_hash": self.current_config_hash,
            "scanned_files": hash_cache_copy
        }
        try:
            with open(CACHE_FILE, 'w', encoding='utf-8') as f:
                json.dump(data_to_save, f, ensure_ascii=False)
            logger.info(f"Фоновое сохранение кэша успешно завершено. Записей: {len(hash_cache_copy)}.")
        except IOError as e:
            logger.error(f"Не удалось сохранить файл кэша в фоновом режиме: {e}")

    def _run_network_scan_thread(self, ip_ranges: List[str], is_scheduled=False, scheduled_job=None):
        start_time = time.monotonic()
        try:
            logger.info("Запуск сканирования сети...")
            self.after(0, self._update_network_tree)

            all_ips_to_scan = []
            for range_str in ip_ranges:
                all_ips_to_scan.extend(self._parse_ip_range(range_str))

            ips_to_scan = list(set(all_ips_to_scan))

            if not ips_to_scan:
                msg = f"Не удалось распознать ни одного IP-адреса из указанных диапазонов/файла."
                logger.error(msg)
                self.after(0, self.update_status, msg)
                return

            total_ips = len(ips_to_scan)
            self.after(0, lambda: self.progress_bar.config(value=0, maximum=total_ips))
            logger.info(f"Найдено {total_ips} уникальных IP адресов для проверки.")

            max_workers = self.settings.get('network_scan_threads', 100)
            current_processed_ips = 0
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_ip = {executor.submit(self._check_and_enumerate_shares, str(ip)): ip for ip in ips_to_scan}

                for future in concurrent.futures.as_completed(future_to_ip):
                    if self.current_stop_event.is_set():
                        for f in future_to_ip:
                            f.cancel()
                        break

                    ip = future_to_ip[future]
                    try:
                        fqdn, unc_paths = future.result()
                        if unc_paths or fqdn not in ["N/A", "Не удалось разрешить", "Ошибка разрешения"]:
                            result_item = {'ip': str(ip), 'fqdn': fqdn, 'resources': unc_paths}
                            self.found_network_results.append(result_item)
                            logger.info(f"Хост {ip} (FQDN: {fqdn}) обработан. Найдено ресурсов: {len(unc_paths)}")
                    except concurrent.futures.CancelledError:
                        logger.info(f"Задача для IP {ip} остановлена.")
                    except Exception as e:
                        logger.error(f"Ошибка проверки IP {ip}: {e}")

                    current_processed_ips += 1
                    status_text = f"Проверка: {str(ip)} ({current_processed_ips}/{total_ips})"
                    self.after(0, lambda val=current_processed_ips, text=status_text: (self.progress_bar.config(value=val), self.status_label.config(text=text)))

            self.after(0, self._update_network_tree)

            if self.current_stop_event.is_set():
                msg = "Сканирование сети остановлено пользователем."
            else:
                msg = "Сканирование сети завершено."

            logger.info(msg)
            self.after(0, self.update_status, msg)

            total_found_shares = sum(len(item.get('resources', [])) for item in self.found_network_results)
            self.after(0, self.update_log, f"Найдено хостов: {len(self.found_network_results)}, Общее кол-во ресурсов: {total_found_shares}")

            summary_lines = [
                f"Сканирование сети завершено.",
                f"Проверенный диапазон: {', '.join(ip_ranges)}",
                f"Найдено хостов с ресурсами или FQDN: {len(self.found_network_results)}\n"
            ]
            share_results = []
            if self.found_network_results:
                for item in self.found_network_results:
                    for resource in item.get('resources', []):
                        summary_lines.append(f"- Ресурс: {resource} (IP: {item['ip']}, FQDN: {item['fqdn']})")
                        share_results.append({'resource': resource, 'ip': item['ip'], 'fqdn': item['fqdn']})

            summary = "\n".join(summary_lines)

            report_paths_for_email = []
            output_dir = self.settings.get('report_output_dir', os.path.expanduser("~"))
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            self.network_report_dir = output_dir

            if share_results:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                unc_with_ip = [res['resource'] for res in share_results]
                if unc_with_ip:
                    report_path_ip = os.path.join(output_dir, f"network_scan_results_ip_{timestamp}.txt")
                    try:
                        with open(report_path_ip, 'w', encoding='utf-8') as f:
                            f.write("\n".join(unc_with_ip))
                        logger.info(f"Отчет с UNC-путями по IP-адресам сохранен в {report_path_ip}")
                        report_paths_for_email.append(report_path_ip)
                    except Exception as e:
                        logger.error(f"Не удалось сохранить отчет по IP: {e}")

                share_results_with_fqdn = [res for res in share_results if res.get('fqdn') and res['fqdn'] not in ["N/A", "Не удалось разрешить", "Ошибка разрешения"]]
                unc_with_fqdn = [f"\\\\{res['fqdn']}\\{res['resource'].split(os.sep)[-1]}" for res in share_results_with_fqdn]
                if unc_with_fqdn:
                    report_path_fqdn = os.path.join(output_dir, f"network_scan_results_fqdn_{timestamp}.txt")
                    try:
                        with open(report_path_fqdn, 'w', encoding='utf-8') as f:
                            f.write("\n".join(unc_with_fqdn))
                        logger.info(f"Отчет с UNC-путями по FQDN сохранен в файл: {report_path_fqdn}")
                        report_paths_for_email.append(report_path_fqdn)
                    except Exception as e:
                        logger.error(f"Не удалось сохранить отчет по FQDN: {e}")

            send_email_report(self.settings, summary, attachments=report_paths_for_email)
        except Exception as e:
            logger.critical(f"Критическая ошибка при сканировании сети: {e}", exc_info=True)
            if not is_scheduled:
                self.after(0, lambda: messagebox.showerror("Critical Error", f"Критическая ошибка при сканировании сети: {e}\nСмотри логи для детализации."))
        finally:
            duration_str = format_duration(time.monotonic() - start_time)
            log_msg = f"Сканирование сети выполнено за: {duration_str}."
            logger.info(log_msg)
            self.after(0, self.set_controls_state, False)
            self.scan_lock.release()

    def _update_network_tree(self):
        self.network_tree.delete(*self.network_tree.get_children())
        for host_info in self.found_network_results:
            resources = host_info.get('resources', [])
            if resources:
                for resource_path in resources:
                    self.network_tree.insert("", "end", iid=resource_path, values=(
                        resource_path,
                        host_info.get('ip', 'N/A'),
                        host_info.get('fqdn', 'N/A'),
                        "Не проверено"
                    ))
            else:
                iid = f"host_only_{host_info.get('ip')}"
                self.network_tree.insert("", "end", iid=iid, values=(
                    "Нет доступных ресурсов",
                    host_info.get('ip', 'N/A'),
                    host_info.get('fqdn', 'N/A'),
                    "N/A"
                ))

    def _check_and_enumerate_shares(self, ip: str) -> Tuple[str, List[str]]:
        if is_port_open(ip, 445, timeout=0.5):
            logger.info(f"Хост {ip} доступен (открыт порт 445). Запрашиваем у хоста его ресурсы...")
            fqdn = "N/A"
            if self.settings.get('network_resolve_fqdn'):
                fqdn = _resolve_fqdn_with_timeout(ip, timeout=2)
            unc_paths = enumerate_shares(ip)
            return fqdn, unc_paths
        return "N/A", []

    def _parse_ip_range(self, range_str: str) -> List[ipaddress.IPv4Address]:
        range_str = range_str.strip()
        try:
            if '-' in range_str:
                start_ip_str, end_ip_str = [p.strip() for p in range_str.split('-', 1)]
                start_ip = ipaddress.ip_address(start_ip_str)
                try:
                    end_ip = ipaddress.ip_address(end_ip_str)
                except ValueError:
                    ip_parts = str(start_ip).split('.')
                    if len(ip_parts) == 4:
                        end_ip = ipaddress.ip_address(f"{ip_parts[0]}.{ip_parts[1]}.{ip_parts[2]}.{end_ip_str}")
                    else:
                        raise
                return [ipaddress.IPv4Address(ip) for ip in range(int(start_ip), int(end_ip) + 1)]
            elif '/' in range_str:
                return list(ipaddress.ip_network(range_str, strict=False).hosts())
            else:
                return [ipaddress.ip_address(range_str)]
        except ValueError:
            logger.info(f"'{range_str}' не является IP-адресом или диапазоном. Попытка разрешить как имя хоста...")
            try:
                ip = socket.gethostbyname(range_str)
                logger.info(f"Имя '{range_str}' успешно разрешено в IP-адрес: {ip}")
                return [ipaddress.ip_address(ip)]
            except socket.gaierror:
                logger.error(f"Не удалось распознать или разрешить '{range_str}' как IP-адрес, диапазон или имя хоста.")
                return []
        except Exception as e:
            logger.error(f"Неизвестная ошибка при парсинге IP-диапазона '{range_str}': {e}")
            return []

    def start_discovered_shares_check(self):
        if not self.scan_lock.acquire(blocking=False):
            messagebox.showwarning("Внимание", "Другая операция сканирования или проверки уже запущена.", parent=self)
            return

        try:
            share_iids = self.network_tree.get_children()
            resources_to_check = []
            for iid in share_iids:
                if self.network_tree.item(iid, "values")[0] != "Нет доступных ресурсов":
                    resources_to_check.append(iid)

            if not resources_to_check:
                messagebox.showinfo("Нет ресурсов", "Не найдено сетевых ресурсов для проверки.", parent=self)
                self.scan_lock.release()
                return

            self.current_stop_event.clear()
            self.set_controls_state(is_scanning=True)
            self.current_scan_thread = threading.Thread(
                target=self._run_discovered_shares_check_thread,
                args=(resources_to_check,),
                daemon=True
            )
            self.current_scan_thread.start()
        except Exception:
            self.scan_lock.release()
            raise

    def _run_discovered_shares_check_thread(self, resources_to_check: List[str]):
        start_time = time.monotonic()
        try:
            logger.info(f"Начало проверки доступности для {len(resources_to_check)} сетевых ресурсов.")
            self.after(0, self.update_status, "Проверка доступности сетевых ресурсов...")
            self.after(0, lambda: self.progress_bar.config(value=0, maximum=len(resources_to_check)))

            for i, resource_path in enumerate(resources_to_check):
                if self.current_stop_event.is_set():
                    logger.warning("Проверка доступности ресурсов отменена пользователем.")
                    break

                self.after(0, self.update_status, f"Проверка: {os.path.basename(resource_path)}")
                _, is_accessible, status_msg = check_directory_accessibility(resource_path)

                log_msg = f"Проверка доступности '{resource_path}': {status_msg}"
                if is_accessible:
                    logger.info(log_msg)
                else:
                    logger.warning(log_msg)

                self.after(0, self._update_share_status_in_tree, resource_path, status_msg, is_accessible)
                self.after(0, lambda val=i + 1: self.progress_bar.config(value=val))

            duration_str = format_duration(time.monotonic() - start_time)
            final_msg = f"Проверка доступности ресурсов завершена за {duration_str}."
            if self.current_stop_event.is_set():
                final_msg = "Проверка доступности ресурсов отменена."

            logger.info(final_msg)
            self.after(0, self.update_status, final_msg)
        finally:
            self.after(0, self.set_controls_state, False)
            self.scan_lock.release()

    def _update_share_status_in_tree(self, iid: str, status_msg: str, is_accessible: bool):
        if not self.network_tree.exists(iid):
            return
        tag = "accessible" if is_accessible else "inaccessible"
        current_values = list(self.network_tree.item(iid, "values"))
        current_values[3] = status_msg
        self.network_tree.item(iid, values=tuple(current_values), tags=(tag,))

    def save_unc_paths_to_file(self):
        if not self.found_network_results:
            messagebox.showwarning("Нет данных", "Нет найденных сетевых ресурсов для сохранения.", parent=self.scan_tab)
            return

        dialog = SaveNetworkPathsDialog(self)
        choice = dialog.result

        if not choice:
            return

        all_share_results = [
            {'resource': res_path, 'ip': host['ip'], 'fqdn': host['fqdn']}
            for host in self.found_network_results
            for res_path in host.get('resources', [])
        ]

        if not all_share_results:
            messagebox.showinfo("Нет ресурсов", "Не найдено доступных для сохранения сетевых ресурсов (только хосты без общих папок).", parent=self.scan_tab)
            return

        share_results_to_save = []
        if choice == 'all':
            share_results_to_save = all_share_results
            logger.info("Пользователь выбрал сохранение ВСЕХ найденных путей.")
        elif choice == 'accessible':
            accessible_iids = self.network_tree.tag_has("accessible")
            accessible_paths = set(accessible_iids)
            share_results_to_save = [res for res in all_share_results if res['resource'] in accessible_paths]
            logger.info(f"Пользователь выбрал сохранение ТОЛЬКО ДОСТУПНЫХ путей. Найдено доступных: {len(share_results_to_save)}.")
            if not share_results_to_save:
                messagebox.showinfo("Нет доступных ресурсов", "Не найдено ресурсов, отмеченных как 'Доступен'. Проведите проверку доступности.", parent=self.scan_tab)
                return

        filepath = filedialog.asksaveasfilename(
            title="Сохранить отчеты по сети (укажите базовое имя)",
            filetypes=[("Text files", "*.txt")],
            defaultextension=".txt",
            initialfile=f"network_scan_{choice}",
            parent=self.scan_tab
        )
        if not filepath:
            return

        base_path, _ = os.path.splitext(filepath)
        saved_files = []

        unc_with_ip = [item['resource'] for item in share_results_to_save]
        if unc_with_ip:
            report_path_ip = f"{base_path}_ip.txt"
            try:
                with open(report_path_ip, 'w', encoding='utf-8') as f:
                    f.write("\n".join(unc_with_ip))
                saved_files.append(report_path_ip)
                logger.info(f"UNC пути по IP сохранены в файл: {report_path_ip}")
            except Exception as e:
                logger.error(f"Не удалось сохранить файл {report_path_ip}: {e}")
                messagebox.showerror("Ошибка сохранения (IP)", f"Не удалось сохранить файл:\n{e}", parent=self.scan_tab)

        share_results_with_fqdn = [
            item for item in share_results_to_save
            if item.get('fqdn') and item['fqdn'] not in ["N/A", "Не удалось разрешить", "Ошибка разрешения"]
        ]
        unc_with_fqdn = [f"\\\\{item['fqdn']}\\{item['resource'].split(os.sep)[-1]}" for item in share_results_with_fqdn]
        if unc_with_fqdn:
            report_path_fqdn = f"{base_path}_fqdn.txt"
            try:
                with open(report_path_fqdn, 'w', encoding='utf-8') as f:
                    f.write("\n".join(unc_with_fqdn))
                saved_files.append(report_path_fqdn)
                logger.info(f"UNC пути по FQDN сохранены в файл: {report_path_fqdn}")
            except Exception as e:
                logger.error(f"Не удалось сохранить файл {report_path_fqdn}: {e}")
                messagebox.showerror("Ошибка сохранения (FQDN)", f"Не удалось сохранить файл:\n{e}", parent=self.scan_tab)

        if saved_files:
            messagebox.showinfo("Успешно", f"Отчеты успешно сохранены:\n" + "\n".join(saved_files), parent=self.scan_tab)

    def _platform_specific_open(self, path_to_open: str):
        if not path_to_open or not os.path.exists(path_to_open):
            messagebox.showwarning("Путь не найден", f"Не удалось найти указанный путь:\n{path_to_open}", parent=self)
            return
        try:
            os.startfile(os.path.normpath(path_to_open))
            logger.info(f"Попытка открыть: {path_to_open}")
        except AttributeError:
            import subprocess
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            try:
                subprocess.call([opener, os.path.normpath(path_to_open)])
                logger.info(f"Попытка открыть через '{opener}': {path_to_open}")
            except Exception as sub_e:
                error_msg = f"Не удалось автоматически открыть путь.\nПуть: {os.path.normpath(path_to_open)}\nОшибка: {sub_e}"
                logger.error(error_msg)
                messagebox.showerror("Ошибка", error_msg, parent=self)
        except Exception as e:
            error_msg = f"Не удалось открыть путь: {e}"
            logger.error(error_msg)
            messagebox.showerror("Ошибка", error_msg, parent=self)

    def on_result_double_click(self, event):
        if self.current_scan_thread and self.current_scan_thread.is_alive():
            return "break"
        selection = self.results_tree.selection()
        if not selection:
            return
        selection_id = selection[0]
        selected_result = next((res for res in self.scan_results if res['path'] == selection_id), None)
        if not selected_result:
            return
        path_to_open = selected_result.get('archive_path') or selected_result.get('path')
        self._platform_specific_open(path_to_open)

    def show_result_context_menu(self, event):
        if self.current_scan_thread and self.current_scan_thread.is_alive():
            return "break"
        iid = self.results_tree.identify_row(event.y)
        if iid:
            self.results_tree.selection_set(iid)
            context_menu = tk.Menu(self, tearoff=0)
            context_menu.add_command(label="📂 Открыть файл/архив", command=self._open_selected_file)
            context_menu.add_command(label="📁 Открыть папку с файлом/архивом", command=self._open_selected_file_location)
            context_menu.add_separator()
            selected_result = next((res for res in self.scan_results if res['path'] == iid), None)
            is_whitelisted = False
            if selected_result:
                path_to_check = selected_result.get('archive_path') or selected_result.get('path')
                norm_path = os.path.normpath(path_to_check)
                with self.hash_lock:
                    is_whitelisted = norm_path in self.hash_cache
            if is_whitelisted:
                context_menu.add_command(label="❌ Удалить из белого списка", command=self._remove_selected_from_whitelist)
            else:
                context_menu.add_command(label="✅ Добавить в белый список", command=self._add_selected_to_whitelist)
            context_menu.post(event.x_root, event.y_root)

    def _open_selected_file(self):
        selection = self.results_tree.selection()
        if not selection: return
        selection_id = selection[0]
        selected_result = next((res for res in self.scan_results if res['path'] == selection_id), None)
        if selected_result:
            path_to_open = selected_result.get('archive_path') or selected_result.get('path')
            self._platform_specific_open(path_to_open)

    def _open_selected_file_location(self):
        selection = self.results_tree.selection()
        if not selection: return
        selection_id = selection[0]
        selected_result = next((res for res in self.scan_results if res['path'] == selection_id), None)
        if selected_result:
            path_for_dir = selected_result.get('archive_path') or selected_result.get('path')
            directory = os.path.dirname(path_for_dir)
            self._platform_specific_open(directory)

    def _add_selected_to_whitelist(self):
        selection = self.results_tree.selection()
        if not selection: return
        selection_id = selection[0]
        selected_result = next((res for res in self.scan_results if res['path'] == selection_id), None)
        if not selected_result: return
        path_to_whitelist = selected_result.get('archive_path') or selected_result.get('path')
        norm_path = os.path.normpath(path_to_whitelist)
        file_hash = get_file_hash(path_to_whitelist)
        if file_hash:
            with self.hash_lock:
                self.hash_cache[norm_path] = file_hash
            self._save_hash_cache_in_background(self.hash_cache.copy())
            for iid in self.results_tree.get_children(''):
                res = next((r for r in self.scan_results if r['path'] == iid), None)
                if res and (res.get('archive_path') or res.get('path')) == norm_path:
                    self.results_tree.item(iid, tags=("whitelisted",))
            logger.info(f"Файл/архив '{norm_path}' вручную добавлен в белый список.")
        else:
            logger.error(f"Не удалось рассчитать хэш для добавления в белый список: {norm_path}")

    def _remove_selected_from_whitelist(self):
        selection = self.results_tree.selection()
        if not selection: return
        selection_id = selection[0]
        selected_result = next((res for res in self.scan_results if res['path'] == selection_id), None)
        if not selected_result: return
        path_to_unwhitelist = selected_result.get('archive_path') or selected_result.get('path')
        norm_path = os.path.normpath(path_to_unwhitelist)
        with self.hash_lock:
            if norm_path in self.hash_cache:
                del self.hash_cache[norm_path]
        self._save_hash_cache_in_background(self.hash_cache.copy())
        for iid in self.results_tree.get_children(''):
            res = next((r for r in self.scan_results if r['path'] == iid), None)
            if res and (res.get('archive_path') or res.get('path')) == norm_path:
                self.results_tree.item(iid, tags=())
        logger.info(f"Файл/архив '{norm_path}' удален из белого списка.")

    def update_log(self, message: str):
        def _update():
            if self.log_text.winfo_exists():
                self.log_text.configure(state='normal')
                self.log_text.insert(tk.END, message + '\n')
                self.log_text.configure(state='disabled')
                self.log_text.see(tk.END)
        self.after(0, _update)

    def update_status(self, message: str):
        if self.status_label.winfo_exists():
            self.status_label.config(text=message)

    def _set_children_state(self, parent_widget, state):
        if parent_widget == self.dir_list_tree or parent_widget == self.net_list_tree:
            return
        for child in parent_widget.winfo_children():
            if isinstance(child, (ttk.Scrollbar, ttk.Sizegrip)):
                continue
            try:
                child.config(state=state)
            except tk.TclError:
                self._set_children_state(child, state)

    def set_controls_state(self, is_scanning: bool):
        scan_state = "disabled" if is_scanning else "normal"
        stop_state = "normal" if is_scanning else "disabled"
        self._update_traffic_light(is_scanning)

        self.start_button.config(state=scan_state)
        self.stop_button.config(state=stop_state)
        self.clear_button.config(state=scan_state)

        self._set_children_state(self.single_scan_controls, scan_state)
        self._set_children_state(self.list_scan_controls, scan_state)
        self._set_children_state(self.network_scan_controls, scan_state)

        self.check_shares_button.config(state=scan_state)
        self.save_unc_button.config(state=scan_state)

        mode_frame = self.scan_tab.winfo_children()[0]
        self._set_children_state(mode_frame, scan_state)

        self.notebook.tab(self.settings_tab, state=scan_state)

        if not is_scanning:
            scan_mode = self.scan_mode_var.get()
            if scan_mode in ['single', 'list']:
                report_button_state = 'normal' if self.report_path else 'disabled'
                self.report_button.config(state=report_button_state)
            elif scan_mode == 'network':
                has_results = bool(self.found_network_results)
                self.check_shares_button.config(state='normal' if has_results else 'disabled')
                self.save_unc_button.config(state='normal' if has_results else 'disabled')
                report_button_state = 'normal' if self.network_report_dir else 'disabled'
                self.report_button.config(state=report_button_state)
        else:
            self.report_button.config(state='disabled')
            self.check_shares_button.config(state='disabled')
            self.save_unc_button.config(state='disabled')

    def _apply_logging_settings(self, *args):
        level = self.settings_vars['log_level'].get()
        self.settings['log_level'] = level
        setup_logging(level, self.gui_logger_handler)
        logger.info(f"Уровень логирования обновлен до {level}.")

    def check_dependencies(self):
        msgs = []
        if not PPTX_AVAILABLE:
            msgs.append("Обработка презентаций: python-pptx не найден. Чтение .pptx будет невозможно (pip install python-pptx).")
        if not XLRD_AVAILABLE:
            msgs.append("Чтение XLS: xlrd не найден. Чтение .xls будет ограничено (pip install xlrd).")
        if not TESSERACT_AVAILABLE:
            msgs.append("OCR: Tesseract-OCR не найден. Функции OCR будут недоступны.")
        else:
            if not self.config.tesseract_cmd:
                msgs.append("OCR: Tesseract-OCR найден, но путь к tesseract.exe не настроен. Проверьте настройки.")

        if RARFILE_AVAILABLE:
            unrar_tool = self.config.unrar_cmd
            if not unrar_tool:
                msgs.append("Обработка RAR: Утилита UnRAR не найдена. Укажите путь в 'Настройки -> Интеграции и OCR' или установите WinRAR (с добавлением в PATH).")
        else:
            msgs.append("Обработка RAR: Библиотека 'rarfile' не установлена (pip install rarfile). Сканирование .rar невозможно.")

        if not TIKA_AVAILABLE:
            msgs.append("Apache Tika: Библиотека 'tika' не найдена. Извлечение текста через Tika будет невозможно (pip install tika).")

        if not EXCHANGELIB_AVAILABLE:
            msgs.append("Отправка почты через Exchange: exchangelib не найден. Эта функция будет недоступна (pip install exchangelib).")

        if sys.platform == "win32":
            pywin32_missing_features = []
            if not WIN_COM_AVAILABLE:
                pywin32_missing_features.append("конвертация .doc/.xls/.ppt")
            if not WIN32NET_AVAILABLE:
                pywin32_missing_features.append("поиск сетевых ресурсов")
            if not WIN32SECURITY_AVAILABLE:
                pywin32_missing_features.append("определение владельца файла")
            if pywin32_missing_features:
                features_str = ", ".join(pywin32_missing_features)
                msgs.append(f"Windows: pywin32 не найден. Недоступные функции: {features_str}. Для установки выполните: pip install pywin32")
        else:
            if not pwd:
                msgs.append("POSIX: модуль 'pwd' не найден. Определение владельца файла будет недоступно.")
            msgs.append("Вы используете не Windows. Конвертация .doc/.xls/.ppt и автоматический поиск сетевых ресурсов будут недоступны.")

        if msgs:
            msg = "\n".join(msgs)
            logger.warning(f"Проблемы с зависимостями:\n{msg}")
            messagebox.showwarning("Проверка зависимостей", f"Обнаружены проблемы:\n\n{msg}\n\nНекоторые функции могут работать некорректно или быть недоступны.")

    def _make_treeview_sortable(self, treeview, cols):
        for col in cols:
            treeview.heading(col, command=lambda _col=col: self._sort_treeview_column(treeview, _col, False))

    def _setup_text_field_bindings(self, text_widget):
        context_menu = tk.Menu(text_widget, tearoff=0)

        def _execute_command(command):
            try:
                text_widget.event_generate(f'<<{command}>>')
            except tk.TclError:
                if command == 'Paste':
                    try:
                        text_widget.insert(tk.INSERT, self.clipboard_get())
                    except tk.TclError:
                        pass

        context_menu.add_command(label="Вырезать", command=lambda: _execute_command('Cut'))
        context_menu.add_command(label="Копировать", command=lambda: _execute_command('Copy'))
        context_menu.add_command(label="Вставить", command=lambda: _execute_command('Paste'))

        def show_context_menu(event):
            try:
                if text_widget.tag_ranges("sel"):
                    context_menu.entryconfig("Вырезать", state=tk.NORMAL)
                    context_menu.entryconfig("Копировать", state=tk.NORMAL)
                else:
                    context_menu.entryconfig("Вырезать", state=tk.DISABLED)
                    context_menu.entryconfig("Копировать", state=tk.DISABLED)
            except tk.TclError:
                pass
            try:
                if self.clipboard_get():
                    context_menu.entryconfig("Вставить", state=tk.NORMAL)
                else:
                    context_menu.entryconfig("Вставить", state=tk.DISABLED)
            except tk.TclError:
                context_menu.entryconfig("Вставить", state=tk.DISABLED)
            context_menu.tk_popup(event.x_root, event.y_root)

        text_widget.bind("<Control-x>", lambda e: _execute_command('Cut'))
        text_widget.bind("<Control-c>", lambda e: _execute_command('Copy'))
        text_widget.bind("<Control-v>", lambda e: _execute_command('Paste'))
        text_widget.bind("<Button-3>", show_context_menu)

    def _sort_treeview_column(self, tv, col, reverse):
        try:
            def get_sort_key(item):
                value = tv.set(item, col)
                if col == 'size_mb':
                    try:
                        return float(value.replace(' MB', ''))
                    except (ValueError, TypeError):
                        return 0.0
                try:
                    return float(value)
                except (ValueError, TypeError):
                    return str(value).lower()
            data_list = [(get_sort_key(k), k) for k in tv.get_children('')]
            data_list.sort(key=lambda t: t[0], reverse=reverse)
            for index, (val, k) in enumerate(data_list):
                tv.move(k, '', index)
            tv.heading(col, command=lambda _col=col: self._sort_treeview_column(tv, _col, not reverse))
        except tk.TclError:
            logger.debug(f"Ошибка сортировки колонки {col} (вероятно, виджет был уничтожен).")

    def start_scheduler_thread(self):
        self.scheduler_thread = threading.Thread(target=self._scheduler_loop, daemon=True)
        self.scheduler_thread.start()

    def _scheduler_loop(self):
        while not self.scheduler_stop_event.is_set():
            try:
                if self.scan_lock.locked():
                    logger.debug("Планировщик: пропуск проверки, т.к. выполняется другая операция.")
                else:
                    now = datetime.now()
                    jobs_to_run = []
                    for job in list(self.settings.get('scheduler_jobs', [])):
                        if not job.get('enabled', False):
                            continue
                        next_run_time = self.get_next_run_time(job)
                        if next_run_time and now >= next_run_time:
                            jobs_to_run.append(job)
                    if jobs_to_run:
                        job_to_run = jobs_to_run[0]
                        self.after(0, self._trigger_scheduled_scan, job_to_run)
            except Exception as e:
                logger.error(f"Ошибка в цикле планировщика: {e}")
            self.scheduler_stop_event.wait(60)

    def _trigger_scheduled_scan(self, job: Dict):
        scan_type = job.get('scan_type')
        job_id = job.get('id')
        logger.info(f"Планировщик: Запущена задача '{job_id}' (type: '{scan_type}').")
        self.notebook.select(self.scan_tab)
        self.scan_mode_var.set(scan_type)
        self._on_scan_mode_change()
        if scan_type == 'single':
            self.scan_path_var.set(job.get('target', ''))
        elif scan_type == 'list':
            self.directory_list_file_var.set(job.get('target', ''))
        elif scan_type == 'network':
            self.network_input_mode_var.set('manual')
            self._on_network_input_mode_change()
            self.net_ip_range_var.set(job.get('target', ''))
        scan_function = {
            'single': self._start_scan_single_path,
            'list': self._start_scan_from_list,
            'network': self._start_network_scan
        }.get(scan_type)
        if scan_function:
            for j in self.settings['scheduler_jobs']:
                if j.get('id') == job_id:
                    j['last_run'] = datetime.now().isoformat()
                    break
            self.save_all_settings(silent=True)
            scan_function(is_scheduled=True, scheduled_job=job)
        else:
            logger.error(f"Неизвестный тип сканирования в задаче планировщика: {scan_type}")

    def on_closing(self):
        try:
            if hasattr(self, 'tika_manager'):
                self.tika_manager.stop_server()
        except:
            pass

        if self.scan_lock.locked():
            if messagebox.askyesno("Выход", "Операция сканирования все еще выполняется. Остановить и выйти?", parent=self):
                self.stop_scan()
                self.scheduler_stop_event.set()
                self.after(200, self.destroy)
        else:
            logger.info("Приложение закрывается.")
            self.scheduler_stop_event.set()
            self.destroy()


class SaveNetworkPathsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.transient(parent)
        self.title("Выбор сохранения")
        self.result = None

        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(expand=True, fill="both")

        ttk.Label(main_frame, text="Какие сетевые ресурсы вы хотите сохранить?").pack(pady=(0, 15))

        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(pady=5)

        ttk.Button(buttons_frame, text="💾 Сохранить все найденные", command=lambda: self.set_result('all')).pack(fill='x', pady=3)
        ttk.Button(buttons_frame, text="✅ Сохранить только доступные", command=lambda: self.set_result('accessible')).pack(fill='x', pady=3)
        ttk.Button(buttons_frame, text="Отмена", command=lambda: self.set_result(None)).pack(fill='x', pady=(10, 0))

        self.update_idletasks()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        w = self.winfo_width()
        h = self.winfo_height()
        x = parent_x + (parent_w // 2) - (w // 2)
        y = parent_y + (parent_h // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

        self.grab_set()
        self.wait_window(self)

    def set_result(self, result):
        self.result = result
        self.destroy()


class SchedulerJobEditor(tk.Toplevel):
    def __init__(self, parent, settings, on_save_callback, job_data=None):
        super().__init__(parent)
        self.transient(parent)
        self.title("Редактор задачи планировщика")
        self.settings = settings
        self.on_save = on_save_callback
        self.job_data = job_data or {}

        self.scan_type_map = {"Один каталог": "single", "По списку": "list", "Сеть": "network"}
        self.scan_type_map_rev = {v: k for k, v in self.scan_type_map.items()}

        self.create_widgets()
        self.load_data()

        self.update_idletasks()
        parent_x = self.master.winfo_rootx()
        parent_y = self.master.winfo_rooty()
        parent_width = self.master.winfo_width()
        parent_height = self.master.winfo_height()
        win_width = self.winfo_width()
        win_height = self.winfo_height()
        x = parent_x + (parent_width // 2) - (win_width // 2)
        y = parent_y + (parent_height // 2) - (win_height // 2)
        self.geometry(f"+{x}+{y}")
        self.grab_set()
        self.wait_window(self)

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill="both", expand=True)

        self.enabled_var = tk.BooleanVar()
        ttk.Checkbutton(main_frame, text="Задача включена", variable=self.enabled_var).grid(row=0, column=0, columnspan=2, sticky='w', pady=5)

        ttk.Label(main_frame, text="Тип сканирования:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.scan_type_var = tk.StringVar()
        self.scan_type_combo = ttk.Combobox(main_frame, textvariable=self.scan_type_var, values=list(self.scan_type_map.keys()), state='readonly')
        self.scan_type_combo.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        self.scan_type_combo.bind("<<ComboboxSelected>>", self.on_scan_type_change)

        self.target_label = ttk.Label(main_frame, text="Цель:")
        self.target_label.grid(row=2, column=0, sticky='w', padx=5, pady=5)

        self.target_frame = ttk.Frame(main_frame)
        self.target_frame.grid(row=2, column=1, sticky='ew')
        self.target_var = tk.StringVar()
        self.target_entry = ttk.Entry(self.target_frame, textvariable=self.target_var, width=40)
        self.target_entry.pack(side='left', fill='x', expand=True, padx=(5,0))
        self.target_browse_button = ttk.Button(self.target_frame, text="...", width=3, command=self.browse_target)
        self.target_browse_button.pack(side='left', padx=(5,0))

        schedule_frame = ttk.LabelFrame(main_frame, text="Расписание", padding="10")
        schedule_frame.grid(row=3, column=0, columnspan=2, sticky='ew', pady=10)

        self.schedule_mode_var = tk.StringVar()
        self.interval_vars = {'minutes': tk.IntVar(value=60)}
        self.daily_vars = {'time': tk.StringVar(value="02:00")}

        ttk.Radiobutton(schedule_frame, text="Интервал:", variable=self.schedule_mode_var, value="interval").grid(row=0, column=0, sticky='w')
        ttk.Spinbox(schedule_frame, from_=1, to=1440, textvariable=self.interval_vars['minutes'], width=8).grid(row=0, column=1, sticky='w')
        ttk.Label(schedule_frame, text="минут").grid(row=0, column=2, sticky='w', padx=5)

        ttk.Radiobutton(schedule_frame, text="Ежедневно в:", variable=self.schedule_mode_var, value="daily").grid(row=1, column=0, sticky='w', pady=(5,0))
        ttk.Entry(schedule_frame, textvariable=self.daily_vars['time'], width=10).grid(row=1, column=1, sticky='w', columnspan=2, pady=(5,0))

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(20, 0))
        ttk.Button(button_frame, text="Сохранить", command=self.save_data).pack(side='left', padx=10)
        ttk.Button(button_frame, text="Отмена", command=self.destroy).pack(side='left', padx=10)

    def on_scan_type_change(self, event=None):
        scan_type = self.scan_type_map.get(self.scan_type_var.get())
        if scan_type == 'network':
            self.target_label.config(text="Диапазон/Хост:")
            self.target_browse_button.config(state='disabled')
        elif scan_type == 'single':
            self.target_label.config(text="Каталог:")
            self.target_browse_button.config(state='normal')
        elif scan_type == 'list':
            self.target_label.config(text="Файл со списком:")
            self.target_browse_button.config(state='normal')
        self.target_var.set("")

    def browse_target(self):
        scan_type = self.scan_type_map.get(self.scan_type_var.get())
        if scan_type == 'single':
            path = filedialog.askdirectory(parent=self)
            if path: self.target_var.set(os.path.normpath(path))
        elif scan_type == 'list':
            path = filedialog.askopenfilename(parent=self, filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
            if path: self.target_var.set(os.path.normpath(path))

    def load_data(self):
        if not self.job_data:
            self.enabled_var.set(True)
            self.scan_type_combo.set(list(self.scan_type_map.keys())[0])
            self.schedule_mode_var.set("interval")
            self.on_scan_type_change()
            return
        self.enabled_var.set(self.job_data.get("enabled", False))
        scan_type_key = self.job_data.get("scan_type", "single")
        self.scan_type_combo.set(self.scan_type_map_rev.get(scan_type_key))
        self.on_scan_type_change()
        self.target_var.set(self.job_data.get("target", ""))
        self.schedule_mode_var.set(self.job_data.get("schedule_mode", "interval"))
        self.interval_vars['minutes'].set(self.job_data.get("interval_minutes", 60))
        self.daily_vars['time'].set(self.job_data.get("daily_time", "02:00"))

    def save_data(self):
        if not self.target_var.get():
            messagebox.showerror("Ошибка", "Цель сканирования не может быть пустой.", parent=self)
            return
        try:
            dt_time.fromisoformat(self.daily_vars['time'].get())
        except (ValueError, TypeError):
            messagebox.showerror("Ошибка", "Неверный формат времени для ежедневного запуска. Используйте ЧЧ:ММ.", parent=self)
            return
        new_job_data = {
            "id": self.job_data.get('id', str(uuid.uuid4())),
            "enabled": self.enabled_var.get(),
            "scan_type": self.scan_type_map.get(self.scan_type_var.get()),
            "target": self.target_var.get(),
            "schedule_mode": self.schedule_mode_var.get(),
            "interval_minutes": self.interval_vars['minutes'].get(),
            "daily_time": self.daily_vars['time'].get(),
            "last_run": self.job_data.get('last_run')
        }
        job_found = False
        for i, job in enumerate(self.settings.get('scheduler_jobs', [])):
            if job.get('id') == new_job_data['id']:
                self.settings['scheduler_jobs'][i] = new_job_data
                job_found = True
                break
        if not job_found:
            if 'scheduler_jobs' not in self.settings:
                self.settings['scheduler_jobs'] = []
            self.settings['scheduler_jobs'].append(new_job_data)
        self.on_save()
        self.destroy()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    initialize_basic_logging()
    app = App()
    app.mainloop()