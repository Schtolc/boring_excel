#!/usr/bin/env python3

import argparse
import os
import time

import openpyxl


def fill_price(price_st, price_our, workbook, max_rows):
    start = time.perf_counter()
    rows_processed = 0

    for article in price_our['E5':'E{}'.format(max_rows)]:
        art = str(article[0].value).strip()
        if art == 'None':
            break
        for article2 in price_st['B4':'B{}'.format(max_rows)]:
            art2 = str(article2[0].value).strip()
            if art2 == 'None':
                break
            if art == art2:
                price_our['G' + str(article[0].row)].value = str(price_st['B' + str(article2[0].row)].value).strip()
                price_our['H' + str(article[0].row)].value = str(price_st['C' + str(article2[0].row)].value).strip()
                price_our['I' + str(article[0].row)].value = str(price_st['D' + str(article2[0].row)].value).strip()
                tmp = str(price_st['E' + str(article2[0].row)].value).strip()
                price_our['J' + str(article[0].row)].value = '' if tmp == 'None' else tmp
                break

        rows_processed += 1
        if rows_processed % 200 == 0:
            print('[{:0.4f}s] Price fill: processed {} rows from our price.'.
                  format(time.perf_counter() - start, rows_processed))

    print('Done price fill')
    workbook.save('result.xlsx')


def find_diff(price_st, price_our, max_rows):
    start = time.perf_counter()
    rows_processed = 0

    with open('diff.txt', 'w+') as diff:
        for article in price_st['B4':'B{}'.format(max_rows)]:
            art_st = str(article[0].value).strip()
            if art_st == 'None':
                break
            found = False
            for article2 in price_our['E5':'E{}'.format(max_rows)]:
                art_our = str(article2[0].value).strip()
                if art_our == 'None':
                    break
                if art_st == art_our:
                    found = True
            if not found:
                diff.write(art_st + '\n')

            rows_processed += 1
            if rows_processed % 200 == 0:
                print('[{:0.4f}s] Find diff: processed {} rows from st price.'.
                      format(time.perf_counter() - start, rows_processed))

    print('Done find diff')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('--path', dest='path', required=True, help='путь до файла')
    parser.add_argument('--st', dest='st_name', default='Прайс St', help='название листа с прайсом ST')
    parser.add_argument('--our', dest='our_name', default='Прайс наш', help='название листа с нашим прайсом')
    parser.add_argument('--max-rows', dest='max_rows', default=10000, type=int,
                        help='максимальное количество артиклов в двух прайсах')
    args = parser.parse_args()

    wb = openpyxl.load_workbook(os.path.expanduser(args.path))
    st = wb[args.st_name]
    our = wb[args.our_name]

    fill_price(st, our, wb, args.max_rows)
    find_diff(st, our, args.max_rows)
