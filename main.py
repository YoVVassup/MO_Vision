import zipfile
from pathlib import Path
import os
import subprocess
import sys
import psutil
import time

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def remove():
    if not os.path.exists(f'{cwd}\\INI\\'):  # проверка на существование каталога INI
        os.mkdir(f'{cwd}\\INI\\')
    if os.path.exists(BattleClient_path):
        os.remove(BattleClient_path)
    if os.path.exists(MentalOmegaMaps_path):
        os.remove(MentalOmegaMaps_path)
    if os.path.exists(artmo_path):
        os.remove(artmo_path)
    if os.path.exists(fan_art_path):
        os.remove(fan_art_path)
    if os.path.exists(soundmo_path):
        os.remove(soundmo_path)
    OriginalFiles = zipfile.ZipFile(f'{cwd}\\Vision_ini.zip')
    OriginalFiles.extract('INI/BattleClient.ini', f'{cwd}')
    OriginalFiles.extract('INI/MentalOmegaMaps.ini', f'{cwd}')
    OriginalFiles.close()


cwd = resource_path(Path.cwd())
BattleClient_path = f'{cwd}\\INI\\BattleClient.ini'
MentalOmegaMaps_path = f'{cwd}\\INI\\MentalOmegaMaps.ini'
artmo_path = f'{cwd}\\artmo.ini'
fan_art_path = f'{cwd}\\fan_art.ini'
soundmo_path = f'{cwd}\\soundmo.ini'

remove()  # очистка к первоначальному состоянию

Vision_zip = zipfile.ZipFile(f'{cwd}\\Vision_ini.zip')
Vision_zip.extract('BattleClient.ini', f'{cwd}\\INI\\')
Vision_zip.extract('MentalOmegaMaps.ini', f'{cwd}\\INI\\')
Vision_zip.extract('artmo.ini', f'{cwd}')
Vision_zip.extract('fan_art.ini', f'{cwd}')
Vision_zip.extract('soundmo.ini', f'{cwd}')
Vision_zip.close()

if os.path.exists(f'{cwd}\\MentalOmegaClient.exe'):  # проверка на существование exe
    game_process = subprocess.Popen('MentalOmegaClient.exe')
    game_process.wait()
    while "clientdx.exe" in (p.name() for p in psutil.process_iter()):
        time.sleep(1)

remove()
exit()

