import os
import logging
import ctypes
from pathlib import Path
import psutil
import time
import keyboard


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

INSTALL_DIR = Path("C:/Program Files/NekoBox")
DESKTOP_DIR = Path(os.path.join(os.environ["USERPROFILE"], "Desktop"))
START_MENU_DIR = Path(os.path.join(os.environ["APPDATA"], "Microsoft/Windows/Start Menu/Programs"))
REPO_MAIN = "MatsuriDayo/nekoray"
REPO_CONFIG = "leviofanh/w3vpn_neko_autoinstall"
TASK_NAME = "NekoBoxAutostart"
PROCESS_NAME = "nekoray.exe"


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def kill_all_related_processes():
    try:
        process_names = [PROCESS_NAME, "nekoray.exe", "nekoray_core.exe"]

        killed = False
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                proc_name = proc.info['name'].lower()
                if any(p.lower() in proc_name for p in process_names):
                    proc.kill()
                    killed = True
                    logging.info(f"Завершен процесс: {proc_name}")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        if killed:
            time.sleep(1)
        return True
    except Exception as e:
        logging.error(f"Ошибка при завершении процессов: {e}")
        return False


def wait_for_enter():
    while True:
        if keyboard.is_pressed("enter"):
            break

