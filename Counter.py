from openpyxl import load_workbook
from openpyxl import Workbook
import re


def matcher_slice(keywords, cell_data):
    sen = cell_data.lower()
    for key in keywords:
        if re.search(key, sen):
            return True
    return False


def matcher_split(keywords, cell_data):
    clean_sentance = re.sub(r'[^\w]', ' ', cell_data.lower())
    word_list = clean_sentance.split()
    for key in keywords:
        if key in word_list:
            return True
    return False


class Counter():
    def __init__(self, case_list=None):
        if case_list:
            self.cases = (load_workbook(str(case_list))).active

    def keyword_list(self):
        with open('keyword_single.txt') as keywords_sig:
            kw_single_list = [str(kw)
                              for kw in keywords_sig.readline().split(', ')]

        with open('keyword_double.txt') as keywords_dou:
            kw_double_list = [str(kw)
                              for kw in keywords_dou.readline().split(', ')]
        return kw_single_list, kw_double_list

    def counter(self):
        current = 0
        total_amount = 0
        kw_single_list, kw_double_list = self.keyword_list()
        for tc in self.cases.iter_rows(max_col=3, values_only=True):
            current += 1
            detail = []
            for i in tc:
                if i is None:
                    detail.append('none')
                else:
                    detail.append(i)

            print('iterate case {}/1499'.format(str(current)))
            if matcher_slice(kw_double_list, detail[1]) or matcher_slice(kw_double_list, detail[2]) or matcher_split(kw_single_list, detail[1]) or matcher_split(kw_single_list, detail[2]):
                total_amount += 1

        print('ValueCAN-related case: {}'.format(str(total_amount)))
        print('Which is {}%'.format(str(round(total_amount * 100 / 1499, 2))))


k = Counter('MY22_1499.xlsx')
k.counter()
