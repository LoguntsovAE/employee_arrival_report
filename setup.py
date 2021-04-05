from distutils.util import execute
from cx_Freeze import setup, Executable


executeables = [Executable('main.py')]

setup(
    name='Report',
    version='0.1',
    description='Отчёт о приходе сотрудников',
    executables=executeables
)
