import random
import time
import requests
from bs4 import BeautifulSoup
import openpyxl
import os
import sys
import config


class Word:  # for working Words
    TYPE = None


class FileWork:  # save load compare file methods
    def __init__(self, path_old_words):
        self.path_old_words = path_old_words
        self.set_old_words = None

    def load_bank_words(self):
        try:
            with open(self.path_old_words, 'r') as filehandle:  # подгрузка множества со словами из файла
                self.set_old_words = set(current_place.rstrip() for current_place in filehandle.readlines())
            filehandle.close()
        except Exception as exll:
            print(exll, 'File old_words.txt isn`t founded change config.txt', sep='\n')
            os.system('pause')
            exit(1)

    def save_bank_words(self):  # Функция выгружает множество слов из файла
        with open(self.set_old_words, 'w') as filehandle:
            filehandle.writelines("%s\n" % place for place in self.set_old_words)
        filehandle.close()
        print('File old_words.txt update')


class Tools:

    @staticmethod
    def clear_output_file():
        with open('Output.txt', 'w') as filehandle:
            pass
        filehandle.close()


current_directory = os.getcwd()

if __name__ == '__main__':
    Bank_File = FileWork(path_old_words='old_words.txt')
    Bank_File.load_bank_words()
