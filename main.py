#!/usr/bin/python
# -*- coding: utf-8 -*-
# =====================
# Автор проекта "MO_Vision": YoWassup
# github_project:MO_Vision
# url:https://github.com/YoVVassup/MO_Vision
# Автор проекта "MO_loading":Rivendell2898
# url:https://github.com/Rivendell2898/MO_loading [использовался для добавления возможности загрузки видеоэкранов]
# Этот код не предназначен для коммерческого использования.
# =====================

import os
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
import configparser


def resource_path(relative_path):
    try:
        base_path = os.path.dirname(os.path.abspath(__file__))
    except (Exception,):
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# Инициализировать путь
cwd = resource_path(Path.cwd())

os.environ['PYTHON_VLC_MODULE_PATH'] = f'{cwd}\\vlc-3.0.16'
try:
    # Устанока пути к библиотеке VLC перед "import vlc".
    import vlc
except (Exception,):
    raise

BattleClient_path = f'{cwd}\\INI\\BattleClient.ini'
MentalOmegaMaps_path = f'{cwd}\\INI\\MentalOmegaMaps.ini'
artmo_path = f'{cwd}\\artmo.ini'
fan_art_path = f'{cwd}\\fan_art.ini'
soundmo_path = f'{cwd}\\soundmo.ini'
fan_soundmo_path = f'{cwd}\\fan_soundmo.ini'
gamemd_path = f'{cwd}\\gamemd.exe'
Syringe_path = f'{cwd}\\Syringe.exe'
ra2mo_ini_path = f'{cwd}\\RA2MO.ini'

# Заглушки
file = ''
wav_player = ''
wav_player_fin = ''

# Инициализировать выбор загрузочного экрана.
load_lst = [f.path for f in os.scandir(f'{cwd}\\LoadScreen') if f.is_dir()]
if load_lst:
    get_catalog = random.choice(load_lst)
else:
    get_catalog = ''


def show_err_compat_not_applicable():
    root = tkinter.Tk()
    root.withdraw()
    tkinter.messagebox.showerror('Сообщение',
                                 'Режим совместимости не может быть установлен для gamemd.exe и Syringe.exe,' +
                                 ' по неизвестной причине.')


def show_err_compat_failure():
    root = tkinter.Tk()
    root.withdraw()
    tkinter.messagebox.showerror('Сообщение',
                                 'Файлы gamemd.exe и Syringe.exe, возможно, отсутствуют в директории игры.')


# Выставление режима совместимости для gamemd.exe и Syringe.exe
def reg_add(path):
    if os.path.exists(path):
        process_compat = subprocess.Popen(r'REG.EXE ADD "HKLM\SOFTWARE\Microsoft\Windows ' +
                                          r'NT\CurrentVersion\AppCompatFlags\Layers" ' +
                                          '/v "{}" /t REG_SZ /d "WINXPSP3 RUNASADMIN" /f'.format(path))
        result_proc = process_compat.wait()
        if result_proc != 0:
            show_err_compat_not_applicable()
            exit(0)
    else:
        show_err_compat_failure()
        exit(0)


reg_add(gamemd_path)
reg_add(Syringe_path)

if get_catalog != '':
    # Загрузить аудио к видео.
    file1 = QUrl.fromLocalFile(get_catalog + '\\audio_muvi.wav')
    content = QtMultimedia.QMediaContent(file1)
    wav_player = QtMultimedia.QMediaPlayer()
    wav_player.setMedia(content)
    wav_player.setVolume(100)

    # Загрузить аудио готовности игры.
    file2 = QUrl.fromLocalFile(f'{cwd}\\gamecreated.wav')
    content1 = QtMultimedia.QMediaContent(file2)
    wav_player_fin = QtMultimedia.QMediaPlayer()
    wav_player_fin.setMedia(content1)
    wav_player_fin.setVolume(100)

# Захват конфигурационного файла
if os.path.exists(gamemd_path):
    config = configparser.ConfigParser()
    config.read(ra2mo_ini_path)  # читаем конфиг
    window_conf = config.get('Video', 'BorderlessWindowedClient', fallback='True')
else:
    window_conf = 'True'


def is_open(filename):
    try:
        # Получить процесс.
        vhandle = win32file.CreateFile(filename, GENERIC_READ, 0, None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
        if int(vhandle) == INVALID_HANDLE_VALUE:
            return True
        else:
            return False

    except (Exception,):
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
        return self.media.video_set_logo_string(str_param, 0)

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
    original_files = zipfile.ZipFile(f'{cwd}\\Vision_ini.zip')
    original_files.extract('INI/BattleClient.ini', f'{cwd}')
    original_files.extract('INI/MentalOmegaMaps.ini', f'{cwd}')
    original_files.extract('chaoticimpulse.wma', f'{cwd}\\Resources')
    original_files.close()


def kill_client():
    try:
        if "clientdx.exe" in (p.name() for p in psutil.process_iter()):
            os.system('taskkill /f /im %s' % 'clientdx.exe')
    except (Exception,):
        remove()
        show_info()
        pass


def open_client():
    global file
    try:
        file = subprocess.Popen(f'{cwd}\\Resources\\clientdx.exe')
        pass
    except (Exception,):
        remove()
        show_err()
        file.kill()


if "__main__" == __name__:

    remove()  # Сброс всех предыдущих параметров.

    vision_zip = zipfile.ZipFile(f'{cwd}\\Vision_ini.zip')
    vision_zip.extract('BattleClient.ini', f'{cwd}\\INI\\')
    vision_zip.extract('MentalOmegaMaps.ini', f'{cwd}\\INI\\')
    vision_zip.extract('artmo.ini', f'{cwd}')
    vision_zip.extract('fan_art.ini', f'{cwd}')
    vision_zip.extract('soundmo.ini', f'{cwd}')
    vision_zip.extract('fan_soundmo.ini', f'{cwd}')
    vision_zip.close()

    if get_catalog != '':

        # Начало инициализации и аудиомониторинга.
        start_time = time.time()

        # Инициализация.
        player = Player()
        player.set_ratio()
        player.set_fullscreen()

        # Завершить запуск "MO Client", при его предыдущем некоректоном закрытии.
        kill_client()

        # Смена стиля.
        try:
            os.unlink(f'{cwd}\\Resources\\chaoticimpulse.wma')
        except (Exception,):
            pass
        shutil.copy2(get_catalog + '\\chaoticimpulse.wma', f'{cwd}\\Resources\\chaoticimpulse.wma')

        time.sleep(0.1)

        # Открыть МО.
        open_client()

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
                except (Exception,):
                    pass

            # Для воспроизведения видео.
            if player.get_position() == 1:
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

        # Обработка окна меню.
        if window_conf == 'True':
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)

        # Ожидание конца.
        try:
            win32gui.DestroyWindow(player_hwnd)
        except (Exception,):
            while "clientdx.exe" in (p.name() for p in psutil.process_iter()):
                time.sleep(1)
            remove()
            pass
    else:
        # Завершить запуск "MO Client", при его предыдущем некоректоном закрытии.
        kill_client()

        # Открыть МО.
        open_client()

        # Ожидание конца.
        while "clientdx.exe" in (p.name() for p in psutil.process_iter()):
            time.sleep(1)
        remove()
