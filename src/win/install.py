import sys
import requests
import zipfile
import shutil
import subprocess
import tempfile
from datetime import datetime
from src.win.common import *


def create_scheduled_task(exe_path):
    try:
        xml_content = f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>{datetime.now().isoformat()}</Date>
    <Author>{os.environ.get('USERNAME')}</Author>
    <URI>\\{TASK_NAME}</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{os.environ.get('USERDOMAIN')}\\{os.environ.get('USERNAME')}</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <GroupId>S-1-5-32-544</GroupId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{exe_path}</Command>
      <WorkingDirectory>{os.path.dirname(exe_path)}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xml') as xml_file:
            xml_file.write(xml_content.encode('utf-16'))
            xml_path = xml_file.name

        subprocess.run([
            'schtasks',
            '/Create',
            '/TN', TASK_NAME,
            '/XML', xml_path,
            '/F'
        ], check=True)

        os.unlink(xml_path)
        logging.info(f"Создана задача планировщика для автозапуска: {TASK_NAME}")

    except Exception as e:
        logging.error(f"Ошибка создания задачи планировщика: {e}")


def create_shortcut(target_path, shortcut_path):
    try:
        from win32com.client import Dispatch
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(str(shortcut_path))
        shortcut.Targetpath = str(target_path)
        shortcut.WorkingDirectory = str(target_path.parent)
        shortcut.save()

        with open(shortcut_path, 'rb') as file:
            data = bytearray(file.read())
            data[0x15] = data[0x15] | 0x20

        with open(shortcut_path, 'wb') as file:
            file.write(data)

        logging.info(f"Создан ярлык: {shortcut_path}")
    except Exception as e:
        logging.error(f"Ошибка создания ярлыка: {e}")


def download_file(url, dest_path):
    logging.info(f"Загрузка файла с {url} в {dest_path}")
    response = requests.get(url, stream=True)
    response.raise_for_status()
    with open(dest_path, 'wb') as file:
        shutil.copyfileobj(response.raw, file)


def find_latest_release(repo):
    logging.info(f"Установка последней версии nekobox {repo}")
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    response = requests.get(url)
    response.raise_for_status()
    release_data = response.json()
    for asset in release_data.get('assets', []):
        if asset['name'].endswith("windows64.zip"):
            return asset['browser_download_url']
    raise ValueError("Не удалось найти установщик nekobox")


def handle_index_html(config_dir):
    try:
        index_html = config_dir / "dashboard" / "index.html"
        if index_html.exists():
            try:
                os.chmod(str(index_html), 0o777)
                os.remove(str(index_html))
                logging.info("Удален index.html")
                return True
            except Exception as e:
                logging.error(f"Не удалось удалить index.html: {e}")
                return False
        return True
    except Exception as e:
        logging.error(f"Ошибка при обработке index.html: {e}")
        return False


def setup_application(is_update=False):
    try:
        exe_path = INSTALL_DIR / PROCESS_NAME
        if exe_path.exists():
            ctypes.windll.kernel32.SetFileAttributesA(str(exe_path).encode('utf-8'), 0x20)

            if not is_update:
                create_shortcut(exe_path, DESKTOP_DIR / "NekoBox.lnk")
                create_shortcut(exe_path, START_MENU_DIR / "NekoBox.lnk")
                create_scheduled_task(str(exe_path))

            subprocess.Popen([str(exe_path)])
            logging.info("Nekobox успешно установлен и запущен.")
        else:
            logging.error(f"Файл '{PROCESS_NAME}' не найден.")
    except Exception as e:
        logging.error(f"Ошибка при настройке приложения: {e}")


def update():
    try:
        logging.info("Начало обновления Nekobox...")

        if not kill_all_related_processes():
            logging.error("Не удалось завершить процессы. Обновление прервано.")
            return False

        config_dir = INSTALL_DIR / "config"
        if not handle_index_html(config_dir):
            logging.error("Не удалось обработать index.html. Обновление прервано.")
            return False

        temp_config_dir = None
        if config_dir.exists():
            temp_config_dir = Path(tempfile.mkdtemp()) / "config"
            shutil.copytree(config_dir, temp_config_dir)
            logging.info("Конфигурация сохранена во временную папку")

        for item in INSTALL_DIR.iterdir():
            if item.name != "config":
                if item.is_file():
                    item.unlink()
                else:
                    shutil.rmtree(item)

        release_url = find_latest_release(REPO_MAIN)
        zip_path = INSTALL_DIR / "release.zip"
        download_file(release_url, zip_path)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_path)

            nested_dir = temp_path / "nekoray"
            if nested_dir.exists():
                for item in nested_dir.iterdir():
                    if item.name != "config":
                        shutil.move(str(item), str(INSTALL_DIR))
            else:
                for item in temp_path.iterdir():
                    if item.name != "config":
                        shutil.move(str(item), str(INSTALL_DIR))

        if temp_config_dir and temp_config_dir.exists():
            if config_dir.exists():
                shutil.rmtree(config_dir)
            shutil.copytree(temp_config_dir, config_dir)
            shutil.rmtree(temp_config_dir.parent)
            logging.info("Конфигурация восстановлена")

        if zip_path.exists():
            os.remove(zip_path)

        setup_application(is_update=True)
        logging.info("Обновление успешно завершено")
        return True

    except Exception as e:
        logging.error(f"Ошибка при обновлении: {e}")
        return False


def install():
    try:
        if INSTALL_DIR.exists():
            logging.info("Обнаружена существующая установка Nekobox. Обновление...")
            return update()

        INSTALL_DIR.mkdir(parents=True, exist_ok=True)

        release_url = find_latest_release(REPO_MAIN)
        zip_path = INSTALL_DIR / "release.zip"
        download_file(release_url, zip_path)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_path)

            nested_dir = temp_path / "nekoray"
            if nested_dir.exists():
                for item in nested_dir.iterdir():
                    shutil.move(str(item), str(INSTALL_DIR))
            else:
                for item in temp_path.iterdir():
                    shutil.move(str(item), str(INSTALL_DIR))

        config_url = f"https://github.com/{REPO_CONFIG}/archive/refs/heads/main.zip"
        config_zip_path = INSTALL_DIR / "config.zip"
        download_file(config_url, config_zip_path)

        with zipfile.ZipFile(config_zip_path, 'r') as zip_ref:
            for member in zip_ref.namelist():
                file_name = Path(member).parts[1:]
                target_path = INSTALL_DIR / Path(*file_name)

                if member.endswith('/'):
                    target_path.mkdir(parents=True, exist_ok=True)
                else:
                    with zip_ref.open(member) as source, open(target_path, 'wb') as target:
                        shutil.copyfileobj(source, target)

        setup_application()

        os.remove(zip_path)
        os.remove(config_zip_path)
        return True

    except Exception as e:
        logging.error(f"Ошибка установки: {e}")
        return False


if __name__ == "__main__":
    try:
        if not is_admin():
            logging.error("Пожалуйста, перезапустите под правами администратора.")
            print("Нажмите ENTER для выхода...", flush=True)
            wait_for_enter()
            sys.exit(1)

        install()
    except Exception as e:
        logging.error(f"Необработанная ошибка: {e}")
        print("Нажмите ENTER для выхода...", flush=True)
        wait_for_enter()
    finally:
        print("Нажмите ENTER для выхода...", flush=True)
        wait_for_enter()
