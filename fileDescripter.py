import os
import sys

def get_resource_path(relative_path):
    """ PyInstaller가 생성한 임시 폴더에서 리소스의 절대 경로를 얻습니다. """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
