import random
import time
import requests
from bs4 import BeautifulSoup
import openpyxl
import os
import config


class Word:  # for working Words
    excel_list_name = config.excel_list_name
    excel_file_name = config.excel_file_name
    enable_stacking = None
    enable_dictionary_wooordhunt_ru = None
    enable_dictionary_cambridge_org = None
    sheet = None

    def __init__(self, word, string):
        self.word = word
        self.string = string

    @classmethod
    def processing(cls):
        wd = openpyxl.load_workbook(filename=Word.excel_file_name)
        Word.sheet = wd[Word.excel_list_name]

        count_empty_string = 0
        for string in range(config.start_cell, config.end_cell):

            value_current = Word.sheet['A' + str(string)].value
            value_previous = Word.sheet['A' + str(string - 1)].value
            value_next = Word.sheet['A' + str(string + 1)].value
            value_next2 = Word.sheet['A' + str(string + 2)].value

            if value_current in Bank_File.set_old_words:
                continue
            Bank_File.set_old_words.add(value_current)
            if value_current == value_previous:  # If woord already met
                continue
            if count_empty_string > 50:  # Если 50 пустых строк то выход из цикла
                break
            if count_empty_string == None:  # Выход из цикла если он дошел до пустых строк
                count_empty_string += 1
                continue
            if value_current != value_next and value_current != value_next2:  # Слово неравно след. просто запись всех значений
                Word_exm = Word(value_current, string)
                Word_exm.add_case_one_word()
                continue
            if value_current == value_next or value_current == value_next2:  #
                Word_exm = Word(value_current, string)
                Word_exm.add_case_several_words()
                continue

    def add_case_one_word(self):
        a_0 = self.word
        b_0 = Word.sheet['B' + str(self.string)].value
        c_0 = Word.sheet['C' + str(self.string)].value
        d_0 = Word.sheet['D' + str(self.string)].value

        if not a_0:
            a_0 = ' '
        if not b_0:
            b_0 = ' '
        if not c_0:
            c_0 = ' '
        if not d_0:
            d_0 = ' '

        example = self.get_examples()

        with open("Output.txt", "a", encoding="utf8") as file:
            file.write(f"{a_0.lower()} * {b_0} * {c_0} * {d_0} * {example}\n")

    def add_case_several_words(self):

        a_0 = self.word
        b_0 = Word.sheet['B' + str(self.string)].value
        c_0 = Word.sheet['C' + str(self.string)].value
        d_0 = Word.sheet['D' + str(self.string)].value

        if not a_0:
            a_0 = ' '
        if not b_0:
            b_0 = ' '
        if not c_0:
            c_0 = ' '
        if not d_0:
            d_0 = ' '

        b_array = []
        c_array = []
        d_array = []
        b_array.append(b_0)
        c_array.append(c_0)
        d_array.append(d_0)

        repeat = 0
        for j in range(1, 20):

            string = j + self.string

            a_1 = Word.sheet['A' + str(string)].value
            b_1 = Word.sheet['B' + str(string)].value
            c_1 = Word.sheet['C' + str(string)].value
            d_1 = Word.sheet['D' + str(string)].value
            if not a_1:
                a_1 = ' '
            if not b_1:
                b_1 = ' '
            if not c_1:
                c_1 = ' '
            if not d_1:
                d_1 = ' '

            if a_0 != a_1:
                repeat += 1
                if repeat > 3:
                    break
                continue
            b_array.append(b_1)
            c_array.append(c_1)
            d_array.append(d_1)

        example = self.get_examples()

        b_ex = ", ".join(sorted(b_array))
        c_ex = "<br>".join(c_array)
        d_ex = "<br>".join(d_array)

        with open("Output.txt", "a", encoding="utf8") as file:
            file.write(f"{self.word} * {b_ex} * {c_ex} * {d_ex} * {example}\n")
        file.close()

    def get_examples(self):
        self.example_list = []

        if Word.enable_dictionary_wooordhunt_ru:
            self.get_examples_from_wooordhunt_ru()
        if Word.enable_dictionary_cambridge_org:
            self.get_examples_from_cambridge_org()

        example = "<br>".join(self.example_list)
        return example

    def get_examples_from_wooordhunt_ru(self):
        url = f"https://wooordhunt.ru/word/{self.word}"
        headers = {
            "Accept": "*/*",
            "User-Agent": "Mozilla/5.0 (iPad; CPU OS 11_0 like Mac OS X) AppleWebKit/604.1.34 (KHTML, like Gecko) Version/11.0 Mobile/15A5341f Safari/604.1"
        }
        req = requests.get(url, headers=headers)
        src = req.text
        soup = BeautifulSoup(src, "lxml")
        tr = soup.find_all(class_="block phrases")
        tr = tr[0].text
        lk = tr.split('  ')
        for i in lk:
            i = i.strip(' ')
            self.example_list.append(i)
        wordexamples = soup.find_all(class_="ex_o")
        wordexamples2 = soup.find_all(class_="ex_t human")
        for key in range(30):
            try:
                example1 = wordexamples[key].text + wordexamples2[key].text
                self.example_list.append(example1)
            except:
                continue

        time.sleep(random.randrange(1, 4))

    def get_examples_from_cambridge_org(self):
        url = "https://dictionary.cambridge.org/dictionary/english-russian/" + self.word
        headers = {
            "Accept": "*/*",
            "User-Agent": "Mozilla/5.0 (iPad; CPU OS 11_0 like Mac OS X) AppleWebKit/604.1.34 (KHTML, like Gecko) Version/11.0 Mobile/15A5341f Safari/604.1"
        }
        req = requests.get(url, headers=headers)
        src = req.text
        soup = BeautifulSoup(src, "lxml")
        tr = soup.find_all(class_="deg")
        for i in tr:
            i = i.text.strip()
            self.example_list.append(i)

        time.sleep(random.randrange(1, 2))


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
    Word.processing()
