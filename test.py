import nltk
import pymorphy2
import string
import math
import os
import pandas as pd
from collections import Counter
morph = pymorphy2.MorphAnalyzer() ## для нормализации слова
part = 0.1                        ## часть слов, которая будет отображена в excel файле
excel_file_name = 'output.xlsx'   ## имя excel файла
stopwords = nltk.corpus.stopwords.words('russian') + ['–','«','»','’'] ## лист стоп-слов


## функция для обработки строки в текстовом файле
## принимает строку, возвращает лист с нормализованными словами
def process_line(line):
    tokens = nltk.word_tokenize(line)  ## разбить строку на токены
    tokens = [i for i in tokens if i not in string.punctuation]  ## уберем знаки пунктуации
    tokens = [i for i in tokens if i not in stopwords]           ## уберем стоп-слова
    return [morph.parse(token)[0].normal_form for token in tokens] ## сделаем нормализацию

    
## функция для обработки тестового файла
## принимает путь к файлу, возвращает лист с нормализованными словами
def process_file(path_to_file):
    words = []
    print('Processing txt file...')
    with open(path_to_file, 'r') as f:
        for line in f:
            words += process_line(line.lower())
    return words


def main():
    path_to_file = input("Enter path to txt file: ")
    try:
        data = process_file(path_to_file)
    except FileNotFoundError:
        print('File not found!')
        return
    words_count = Counter(data) ## подсчитаем частотность слов
    most_common = words_count.most_common(math.ceil(len(words_count)*part)) ## отберем самые популярные из них(вместе с кол-вом вхождений)
    most_common_words = [i[0] for i in most_common]  ## полуим только слова
    pd.DataFrame(most_common_words).to_excel(excel_file_name, header=False, index=False) ## создаем excel файл
    print('File {} was created'.format(os.path.abspath(excel_file_name)))


main()

