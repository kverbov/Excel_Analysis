from excel import Excel
from config import Config
import os


book=Excel()
os.chdir(Config.path)

for f in os.listdir():
    book.sortFile(f, r'..\rez.txt')
print('Обработка завершена. Файл сохранен REZ.TXT')
