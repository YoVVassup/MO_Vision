# -*- coding: utf-8 -*-
# =====================
# Автор проекта "MO_Vision": YoWassup
# github_project：MO_Vision
# url：https://github.com/YoVVassup/MO_Vision
# Автор проекта "MO_loading"：Rivendell2898
# url：https://github.com/Rivendell2898/MO_loading [использовался для добавления возможности загрузки видеоэкранов]
# Этот код не предназначен для коммерческого использования.
# =====================

import os
import sys
from pathlib import Path
import win32gui
import win32con
import subprocess
import tkinter
import tkinter.messagebox
import shutil
import time
import win32file
from win32file import *
from win32api import GetSystemMetrics
from PyQt5 import QtMultimedia
from PyQt5.QtCore import QUrl
import random
import zipfile
import psutil


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# Инициализировать путь
cwd = resource_path(Path.cwd())

os.environ['PYTHON_VLC_MODULE_PATH'] = f'{cwd}\\vlc-3.0.16'
# Устанока пути к библиотеке VLC перед "import vlc".
import vlc

BattleClient_path = f'{cwd}\\INI\\BattleClient.ini'
MentalOmegaMaps_path = f'{cwd}\\INI\\MentalOmegaMaps.ini'
artmo_path = f'{cwd}\\artmo.ini'
fan_art_path = f'{cwd}\\fan_art.ini'
soundmo_path = f'{cwd}\\soundmo.ini'
fan_soundmo_path = f'{cwd}\\fan_soundmo.ini'

# Инициализировать выбор загрузочного экрана.
load_lst = [f.path for f in os.scandir(f'{cwd}\\LoadScreen') if f.is_dir()]
get_catalog = random.choice(load_lst)

# Загрузить аудио к видео.
file = QUrl.fromLocalFile(get_catalog + '\\audio_muvi.wav')
content = QtMultimedia.QMediaContent(file)
wav_player = QtMultimedia.QMediaPlayer()
wav_player.setMedia(content)
wav_player.setVolume(100)

# Загрузить аудио к выбору в меню.
file1 = QUrl.fromLocalFile(f'{cwd}\\gamecreated.wav')
content1 = QtMultimedia.QMediaContent(file1)
wav_player_fin = QtMultimedia.QMediaPlayer()
wav_player_fin.setMedia(content1)
wav_player_fin.setVolume(100)


def is_open(filename):
    try:
        # Получить процесс.
        vHandle = win32file.CreateFile(filename, GENERIC_READ, 0, None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
        if int(vHandle) == INVALID_HANDLE_VALUE:
            return True
        else:
            return False

    except Exception:
        remove()
        return True


# Видео-плеер
class Player:

    def __init__(self, *args):
        if args:
            instance = vlc.Instance(*args)
            self.media = instance.media_player_new()
        else:
            self.media = vlc.MediaPlayer()

    # Установка URL-адрес или локального пути к файлу для воспроизведения, каждый вызов будет перезагружать его.
    def set_uri(self, uri):
        self.media.set_mrl(uri)

    # Воспроизведение возвращает "0" в случае успеха, "-1" в случае неудачи.
    def play(self, path=None):
        if path:
            self.set_uri(path)
            return self.media.play()
        else:
            return self.media.play()

    def stop(self):
        return self.media.stop()

    # Освобождение ресурсов.
    def release(self):
        return self.media.release()

    # Воспроизведение.
    def is_playing(self):
        return self.media.is_playing()

    # Прошедшее время, возвращается в миллисекундах.
    def get_time(self):
        return self.media.get_time()

    # Общая длина аудио и видео, возвращается в миллисекундах.
    def get_length(self):
        return self.media.get_length()

    # Текущий прогресс воспроизведения. Возвращает число с плавающей запятой в диапазоне от "0,0" до "1,0".
    def get_position(self):
        return self.media.get_position()

    def set_time(self, ms):
        return self.media.set_time(ms)

    # Получить текущую скорость воспроизведения файла.
    def get_rate(self):
        return self.media.get_rate()

    def set_fullscreen(self):
        return self.media.set_fullscreen(1)

    def video_set_logo_string(self, str_param):
        return self.media.video_set_logo_string(str_param)

    def get_hwnd(self):
        return self.media.get_hwnd()

    # Установить соотношение сторон.
    def set_ratio(self):
        ratio = f'{GetSystemMetrics(0)}:{GetSystemMetrics(1)}'
        self.media.video_set_scale(
            0)  # Должен быть установлен на 0, иначе ширина и высота экрана не смогут быть изменены
        self.media.video_set_aspect_ratio(ratio)


def show_err():
    root = tkinter.Tk()
    root.withdraw()
    tkinter.messagebox.showerror('Сообщение',
                                 'Ошибка при открытии "MO Vision", попробуйте открыть его с правами Администратора. ' +
                                 'Но, если эта ошибка появилась еще раз, это может быть связано с тем, что файл ' +
                                 'клиента занят,или отсутствует файл "clientdx.exe" в папке Resources, а так же это' +
                                 'может быть ошибкой выполнения программы или попыткой запустить клиент "MO Vision" ' +
                                 'с флешки или USB-HDD/SSD.')


def show_info():
    root = tkinter.Tk()
    root.withdraw()
    show_flag = tkinter.messagebox.askyesno('Вы действительно хотите запустить "MO Vision"?',
                                            'Обнаружено, что клиент Mental Omega запущен, рекомендуется завершить ' +
                                            'активный или фоновый процесс программы клиента перед очередным ' +
                                            'открытием. Вы хотите принудительно открыть клиент "MO Vision"?')
    if show_flag == 0:
        exit(0)


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
    if os.path.exists(fan_soundmo_path):
        os.remove(fan_soundmo_path)
    OriginalFiles = zipfile.ZipFile(f'{cwd}\\Vision_ini.zip')
    OriginalFiles.extract('INI/BattleClient.ini', f'{cwd}')
    OriginalFiles.extract('INI/MentalOmegaMaps.ini', f'{cwd}')
    OriginalFiles.extract('chaoticimpulse.wma', f'{cwd}\\Resources')
    OriginalFiles.close()


if "__main__" == __name__:

    remove()  # Очистка к первоначальному состоянию.

    Vision_zip = zipfile.ZipFile(f'{cwd}\\Vision_ini.zip')
    Vision_zip.extract('BattleClient.ini', f'{cwd}\\INI\\')
    Vision_zip.extract('MentalOmegaMaps.ini', f'{cwd}\\INI\\')
    Vision_zip.extract('artmo.ini', f'{cwd}')
    Vision_zip.extract('fan_art.ini', f'{cwd}')
    Vision_zip.extract('soundmo.ini', f'{cwd}')
    Vision_zip.extract('fan_soundmo.ini', f'{cwd}')
    Vision_zip.close()

    # Начало инициализации и аудиомониторинга.
    start_time = time.time()

    # Инициализация.
    player = Player()
    player.set_ratio()
    player.set_fullscreen()

    # Предотвратить запуск "MO Client" в фоновом режиме.
    try:
        os.system('taskkill /f /im %s' % 'clientdx.exe')
    except:
        remove()
        show_info()
        pass

    # Смена стиля.
    try:
        os.unlink(f'{cwd}\\Resources\\chaoticimpulse.wma')
    except:
        pass
    shutil.copy2(get_catalog + '\\chaoticimpulse.wma', f'{cwd}\\Resources\\chaoticimpulse.wma')

    time.sleep(0.1)

    # Открыть МО.
    try:
        file = subprocess.Popen(f'{cwd}\\Resources\\clientdx.exe')
        pass
    except Exception as e:
        remove()
        show_err()
        file.kill()

    # Определить окно.
    FrameClass = 'WindowsForms10.Window.8.app.0.1ca0192_r6_ad1'
    FrameTitle = 'MO Client'
    hwnd = win32gui.FindWindow(None, FrameTitle)

    playerClass = 'IME'
    playerTile = 'Default IME'
    player_hwnd = win32gui.FindWindow(None, playerTile)

    i = 0
    j = 1
    # Свернуть окно и воспроизвести видео.
    while i < 1000:
        if hwnd:
            if j == 1:
                time.sleep(0.01)  # Предотвратить несвоевременное сворачивание окна.
                player.play(get_catalog + '\\muvi.mp4')
            flag = win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
            if flag:
                break
            time.sleep(0.01)
            i = i + 1
            j = 0

        else:
            hwnd = win32gui.FindWindow(FrameClass, FrameTitle)
            time.sleep(0.01)
            i = i + 1

    # Если время истекло, сообщить об ошибке.
    if i >= 1000:
        show_err()
        exit(0)

    j = 1
    flag = 0
    cnt = 0
    wav_player.play()

    # Предотвратить завершение процесса проигрывания.
    while True:
        time.sleep(0.1)  # Предотвращение слишком быстрого сбоя цикла while.
        end_time = time.time()

        # Если окно плеера не найдено.
        if player_hwnd == 0:
            player_hwnd = win32gui.FindWindow(None, playerTile)
        # Если окно найдено.
        else:
            try:
                win32gui.BringWindowToTop(player_hwnd)
                win32gui.SetForegroundWindow(player_hwnd)
                j = 0
            except Exception as e:
                pass

        # Для воспроизведения видео.
        if player.get_position() >= 0.95:
            wav_player.stop()
            time.sleep(1)
            win32gui.CloseWindow(player_hwnd)
            break

        if (end_time - start_time) >= 10:
            wav_player.stop()
            time.sleep(1)
            win32gui.CloseWindow(player_hwnd)
            break

        if cnt % 20 == 0 and cnt > 1:
            if not is_open(f'{cwd}\\Resources\\OptionsWindow.ini'):
                wav_player.stop()
                wav_player_fin.play()
                player.set_time(13000)  # Видео должно воспроизводиться, когда окно закрыто, иначе оно зависнет.
                time.sleep(1)
                wav_player_fin.stop()
                time.sleep(1)
                win32gui.CloseWindow(player_hwnd)
                break
        cnt = cnt + 1

    time.sleep(1)
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    # Ожидание конца.
    try:
        win32gui.DestroyWindow(player_hwnd)
    except:
        while "clientdx.exe" in (p.name() for p in psutil.process_iter()):
            time.sleep(1)
        remove()
        pass
