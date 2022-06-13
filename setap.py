from setuptools import setup
import platform
from glob import glob

SETUP_DICT = {

    'name': 'Программа формирования производственных заданий',
    'version': '1.0',
    'description': 'Программа формирования производственных заданий',
    'author': 'Ivan Metliaev',
    'author_email': 'ivan.metliaev.helper@gmail.com',

    'data_files': (
        ('', glob(r'C:\Windows\SYSTEM32\msvcp100.dll')),
        ('', glob(r'C:\Windows\SYSTEM32\msvcr100.dll')),
        ('platforms', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\plugins\platforms\qwindows.dll')),
        ('icons', ['images/manager.png']),
        ('sqldrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\plugins\sqldrivers\qsqlite.dll')),
        ('qtcoredrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Core.dll')),
        ('qtguidrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Gui.dll')),
        ('qtwidgetdrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Widgets.dll')),
    ),
    'windows': [{'script': 'main_awe_v0.9.py'}],
    'options': {
        'py2exe': {
            'includes': ["lxml._elementpath", "PyQt5.QtCore", "PyQt5.QtGui", "PyQt5.QtWidgets", "db_connect", "config", "images", "excel_writer"],
        },
    }
}
if platform.system() == 'Windows':
    import py2exe
    SETUP_DICT['windows'] = [{
        'Name': 'Ivan Metliaev',
        'product_name': 'Программа формирования производственных заданий',
        'version': '1.3',
        'description': 'Программа cоздана Метляевым Иваном специально для ООО "Тентовые Конструкции"',
        'copyright': '© 2022, ivan.metliaev.helper@gmail.com. All Rights Reserved',
        'script': 'main_awe_v0.9.py',
        'icon_resources': [(0, r'taskmanager.ico')]
    }]
    SETUP_DICT['zipfile'] = None

setup(**SETUP_DICT)

