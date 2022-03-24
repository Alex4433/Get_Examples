import random
import time
from functools import wraps

import requests
from bs4 import BeautifulSoup
import openpyxl
import os
import config


def logging(func):
    call_count = dict()

    @wraps(func)
    def wrapper(*args, **kwargs):
        wrapper.count += 1
        res = func(*args, **kwargs)
        call_count[func.__name__] = wrapper.count
        print(func.__name__, args[1], call_count[func.__name__], sep='\n')
        return res

    wrapper.count = 0
    return wrapper


class Word:  # for working Words
    excel_list_name = config.excel_list_name
    excel_file_name = config.excel_file_name
    enable_stacking = None
    enable_dictionary_wooordhunt_ru = config.dictionary_wooordhunt_ru
    enable_dictionary_cambridge_org = config.dictionary_cambridge_org
    sheet = None

    def __init__(self, word, string):
        self.word = word
        self.string = string

    @classmethod
    def processing(cls):
        try:
            wd = openpyxl.load_workbook(filename=Word.excel_file_name)
        except:
            print('Excel file isn`t founded')
            raise (AttributeError)
        try:
            Word.sheet = wd[Word.excel_list_name]
        except:
            print('Excel list name isn`t correct or founded')
            raise (AttributeError)

        count_empty_string = 0
        for string in range(config.start_cell, config.end_cell):

            value_current = Word.sheet['A' + str(string)].value
            value_previous = Word.sheet['A' + str(string - 1)].value
            value_next = Word.sheet['A' + str(string + 1)].value
            value_next2 = Word.sheet['A' + str(string + 2)].value

            if value_current in Bank_File.set_old_words:
                continue
            Bank_File.set_old_words.add(value_current)
            if value_current == value_previous:  # If word`ve met already
                continue
            if count_empty_string > 50:  # if 50 empty string than exit
                break
            if count_empty_string == None:  #
                count_empty_string += 1
                continue
            if value_current != value_next and value_current != value_next2:  #
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

        example = self.get_examples(a_0)

        try:
            with open("Output.txt", "a", encoding="utf8") as file:
                file.write(f"{a_0.lower()} * {b_0} * {c_0} * {d_0} {b_0} * {example}\n")
        except:
            print(a_0, 'faile output')

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

        example = self.get_examples(a_0)

        b_ex = ", ".join(sorted(b_array))
        c_ex = "<br>".join(c_array)
        d_ex = []
        for i in range(len(d_array)):
            d_ex.append(d_array[i])
            d_ex.append(' ')
            d_ex.append(b_array[i])
            d_ex.append('<br>')
        d_ex = ''.join(d_ex)

        try:
            with open("Output.txt", "a", encoding="utf8") as file:
                file.write(f"{self.word.lower()} * {b_ex} * {c_ex} * {d_ex} * {example}\n")
        except:
            print(self.word, 'faile output')

    @logging
    def get_examples(self, *args):
        self.example_list_phrases_wooordhunt = []
        self.example_list_sentence_wooordhunt = []
        self.example_list_sentence_cambridge = []

        if Word.enable_dictionary_wooordhunt_ru:
            try:
                self.get_examples_from_wooordhunt_ru()
            except:
                print(self.word, 'fail getting wooordhunt')
        if Word.enable_dictionary_cambridge_org:
            try:
                self.get_examples_from_cambridge_org()
            except:
                print(self.word, 'fail getting cambridge')
        # setting sequence examples
        try:
            ex = self.example_list_phrases_wooordhunt + self.example_list_sentence_wooordhunt + self.example_list_sentence_cambridge
            example = "<br>".join(ex)
        except:
            print(self.word, 'example crash all')
            return ' '
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
            self.example_list_phrases_wooordhunt.append(i)
        word_examples = soup.find_all(class_="ex_o")
        word_examples_2 = soup.find_all(class_="ex_t human")
        for key in range(30):
            try:
                example1 = word_examples[key].text + word_examples_2[key].text
                self.example_list_sentence_wooordhunt.append(example1)
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
            self.example_list_sentence_cambridge.append(i)

        time.sleep(random.randrange(1, 2))


class FileWork:  # save load compare file methods
    def __init__(self, path_old_words):
        self.path_old_words = path_old_words
        self.set_old_words = None

    def load_bank_words(self):
        try:
            with open(self.path_old_words, 'r') as filehandle:  # Load file old words
                self.set_old_words = set(current_place.rstrip() for current_place in filehandle.readlines())
            filehandle.close()
        except Exception as exll:
            print(exll, 'File old_words.txt isn`t founded change config.txt', sep='\n')
            os.system('pause')
            exit(1)

    def save_bank_words(self):  # Update file old words
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


if __name__ == '__main__':
    Tools.clear_output_file()
    Bank_File = FileWork(path_old_words='old_words.txt')
    Bank_File.load_bank_words()
    Word.processing()
    Bank_File.save_bank_words()
    os.system('pause')
