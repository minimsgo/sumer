import sys
from os import listdir
from os.path import isfile, join

from functools import reduce

from xlrd import open_workbook
from xlwt import Workbook

from core import add_books, book2cells

if __name__ == '__main__':

  if len(sys.argv) < 2:
    print('请输出文件夹名称')
    sys.exit(0)

  # 输入参数:文件夹名称
  dir = sys.argv[1]

  files = [dir + '/' + f for f in listdir(dir) if isfile(join(dir, f))]

  books = [open_workbook(file) for file in files]

  result = reduce(add_books, [book2cells(book) for book in books])

  output = Workbook()

  for s in range(len(result)):
    sheet = output.add_sheet(books[0].sheets()[s].name)
    for r in range(len(result[s])):
      for c in range(len(result[s][r])):
        sheet.write(r, c, result[s][r][c])

  output.save(dir + '.xls')

