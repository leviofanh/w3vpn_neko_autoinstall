import subprocess
import sys
import requests
import shutil
import tempfile
from src.win.common import *


def download_routes_file(repo):
    try:
        url = f"https://raw.githubusercontent.com/{repo}/main/config/routes_box/Default"
        logging.info(f"Загрузка правил маршрутизации с {url}")
        response = requests.get(url)
        response.raise_for_status()
        return response.content
    except Exception as e:
        logging.error(f"Ошибка загрузки правил маршрутизации: {e}")
        return None


def update_routes():
    try:
        if not INSTALL_DIR.exists():
            logging.error("NekoBox не установлен")
            return False

        routes_file = INSTALL_DIR / "config" / "routes_box" / "Default"
        if not routes_file.parent.exists():
            routes_file.parent.mkdir(parents=True, exist_ok=True)
            logging.info("Создана директория для правил маршрутизации")

        if not kill_all_related_processes():
            logging.error("Не удалось остановить процессы NekoBox")
            return False

        new_routes = download_routes_file(REPO_CONFIG)
        if new_routes is None:
            return False

        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(new_routes)
            temp_path = Path(temp_file.name)

        if routes_file.exists():
            try:
                routes_file.unlink()
                logging.info("Удален старый файл правил маршрутизации")
            except Exception as e:
                logging.error(f"Ошибка удаления старого файла: {e}")
                temp_path.unlink()
                return False

        try:
            shutil.copy2(temp_path, routes_file)
            logging.info("Установлен новый файл правил маршрутизации")
        except Exception as e:
            logging.error(f"Ошибка установки нового файла: {e}")
            temp_path.unlink()
            return False

        temp_path.unlink()

        exe_path = INSTALL_DIR / PROCESS_NAME
        if exe_path.exists():
            subprocess.Popen([str(exe_path)])
            logging.info("NekoBox успешно перезапущен")
        else:
            logging.error(f"Файл '{PROCESS_NAME}' не найден")
            return False

        logging.info("Обновление правил маршрутизации успешно завершено")
        return True

    except Exception as e:
        logging.error(f"Ошибка при обновлении правил маршрутизации: {e}")
        return False


if __name__ == "__main__":
    try:
        if not is_admin():
            logging.error("Пожалуйста, перезапустите под правами администратора")
            print("Нажмите ENTER для выхода...", flush=True)
            wait_for_enter()
            sys.exit(1)

        success = update_routes()
        if not success:
            logging.warning("Обновление правил маршрутизации завершено с ошибками")
            print("Нажмите ENTER для выхода...", flush=True)
            wait_for_enter()
            sys.exit(1)

    except Exception as e:
        logging.error(f"Необработанная ошибка: {e}")
        print("Нажмите ENTER для выхода...", flush=True)
        wait_for_enter()
        sys.exit(1)
    finally:
        print("Нажмите ENTER для выхода...", flush=True)
        wait_for_enter()