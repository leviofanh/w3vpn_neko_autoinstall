import sys
import shutil
import subprocess
from src.win.common import *


def force_delete_directory(directory):
    try:
        if not directory.exists():
            return True

        index_html = directory / "config" / "dashboard" / "index.html"
        if index_html.exists():
            try:
                os.chmod(str(index_html), 0o777)
                os.remove(str(index_html))
                logging.info("Удален index.html")
            except Exception as e:
                logging.error(f"Не удалось удалить index.html: {e}")
                return False

        try:
            shutil.rmtree(directory)
            logging.info("Папка установки успешно удалена")
            return True
        except Exception as e:
            logging.error(f"Ошибка удаления папки: {e}")
            return False

    except Exception as e:
        logging.error(f"Ошибка при удалении папки {directory}: {e}")
        return False


def remove_scheduled_task():
    try:
        subprocess.run(['schtasks', '/Delete', '/TN', TASK_NAME, '/F'], check=False)
        logging.info("Удалена задача планировщика")
    except Exception as e:
        logging.error(f"Ошибка удаления задачи планировщика: {e}")


def uninstall():
    try:
        if not kill_all_related_processes():
            logging.error("Не удалось завершить все процессы. Попробуйте закрыть приложение вручную.")
            return False

        if INSTALL_DIR.exists():
            if not force_delete_directory(INSTALL_DIR):
                logging.error("Не удалось удалить папку установки")
                return False

        for shortcut in ["NekoBox.lnk"]:
            for path in [DESKTOP_DIR / shortcut, START_MENU_DIR / shortcut]:
                if path.exists():
                    try:
                        os.remove(path)
                        logging.info(f"Удалён ярлык: {path}")
                    except Exception as e:
                        logging.error(f"Ошибка удаления ярлыка {path}: {e}")

        remove_scheduled_task()

        logging.info("Удаление успешно завершена")
        return True

    except Exception as e:
        logging.error(f"Ошибка при удалении: {e}")
        return False


if __name__ == "__main__":
    try:
        if not is_admin():
            logging.error("Пожалуйста, перезапустите под правами администратора.")
            print("Нажмите ENTER для выхода...", flush=True)
            wait_for_enter()
            sys.exit(1)

        success = uninstall()
        if not success:
            logging.warning("Удаление завершено с ошибками")
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