# MO_Vision

[![N|Solid](https://i.ibb.co/yFBZZqJ/mo.gif)](http://mentalomega.com/)

MO_Vision - это дополнительное расширение контента фанатского мода Mental Omega для RA2YR.

- Позволяет делать замену глобальных .ini файлов без перезаписи оригинальных файлов.
- Добовляет кастомные видеозагрузочные экраны.
- ✨  Не влияет  на оригинал  Mental Omega✨

## Сборка

Для сборки проекта используйте Python 3.6 либо новее.

Клонируйте репозиторий себе или скачайте его в формате .zip. 
Если это требуется, то разархивируте его. 
Запустите проект.
Доустановите библиотеки `pathlib`,`PyQt5`,`psutil`,`python-vlc`.

```sh
pip install pathlib
```
```sh
pip install PyQt5
```
```sh
pip install psutil
```
```sh
pip install python-vlc
```
Так же потребуется установка `winapi` для python, чтобы избежать возможных проблем с несовместимостью - перейдите на репозиторий проекта **[pywin32](https://github.com/mhammond/pywin32)**, на основании битности и версии установленой версии python скачайте `.exe` файл и установите его.

Для того чтобы лучше ориентироваться в версиях я использую библиотеку `pyinstaller_versionfile`.

```sh
pip install pyinstaller-versionfile
```

Для легкой настройки сборки поставте `auto-py-to-exe`.

```sh
pip install auto-py-to-exe
```
Затем, чтобы запустить ее, выполните в терминале следующее:

```sh
auto-py-to-exe
```
Вы можите самостоятельно сконфигурировать настройки упаковки.

Моя конфигурация сборки:

```sh
pyinstaller --noconfirm --onefile --windowed "C:/'Указать путь до иконки'/vision.ico" --name "MO Vision" --version-file "C:/'Указать путь до файла-конфигуратора'/versionfile.txt" "C:/'Указать путь до скрипта'/main.py"
```
Нажмите `[Конвертировать .py в .exe]`. Готово.

## Сборка +

Вы также можите самостоятельно сконфигурировать параметры версии файла с помощью `version.py`, просто прописав желаемые параметры в файл и запустив его. Таким образом вы получите новый файл `versionfile.txt`.

## Установка
Просто скопируйте получившийся `.exe`,`Vision_ini.zip`,`gamecreated.wav`,`RA2MO.ini` а так же папки `vlc-3.0.18-x32`, `vlc-3.0.18-x64` и `LoadScreen` в папку с игрой. Теперь, запуская игру с такого лаунчера вы будете использовать кастомные `.ini`, которые хранятся в `Vision_ini.zip`. Если по какой-либо причине вам захочется отключить видеозаставки, то просто переместите куда-нибудь или удалите все папки внутри `LoadScreen`. Работаспособность видеозаставок проверена и работает на Windows 8.1 x32, Windows 10 x32, Windows 8.1 x64, Windows 10 x64, Windows 11 x64. В остальных системах отсутствуют две необходимые для работы dll - `api-ms-win-core-winrt-l1-1-0.dll`, `api-ms-win-core-winrt-string-l1-1-0.dll`. Работаспособность без видеозаставок является такой же как и у лаунчера "MO Client".

## Лицензия

**Free Software, Hell Yeah!**