def book2cells(book):
    rows_array = [sheet.get_rows() for sheet in book.sheets()]
    return map(lambda i: map(lambda j: map(lambda c: c.value, j), i), rows_array)

def add_books(book0, book1):
  return list(map(add_sheets, book0, book1))

def add_sheets(s0, s1):
  return list(map(add_rows, s0, s1))

def add_rows(r0, r1):
  return list(map(add_cells, r0, r1))

def add_cells(c0, c1):
  if type(c0) == float and type(c1) == float:
    return c0 + c1
  else:
    return c1
