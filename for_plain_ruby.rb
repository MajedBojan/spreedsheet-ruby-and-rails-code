require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'

# will create new worksheet
book = Spreadsheet::Workbook.new

sheet1 = book.create_worksheet
sheet1.name = 'Test spreadsheet'

sheet1.row(0).concat %w{Name Country speacialist}

row = sheet1.row(1)
row[0] = 'Majed Bojan'
row[1] =  'Yemen'
row[2] = 'Fullstack Developer'

row = sheet1.row(2)
row[0] = 'Ali Sheiba'
row[1] =  'Yemen'
row[2] = 'Fullstack Developer'

row = sheet1.row(3)
row[0] = 'Mohammed Balfaqi'
row[1] =  'Yemen'
row[2] = 'ROR Developer'

row = sheet1.row(4)
row[0] = 'Mohammed Basalah'
row[1] =  'Yemen'
row[2] = 'ROR Developer'

row = sheet1.row(5)
row[0] = 'Mohammed Aljefry'
row[1] =  'Yemen'
row[2] = 'Fullstack Developer'

sheet1.row(0).height = 18

format = Spreadsheet::Format.new :color => :blue,
                                 :weight => :bold,
                                 :size => 18
sheet1.row(0).default_format = format

bold = Spreadsheet::Format.new :weight => :bold

book.write '/home/bojan/Desktop/spreedsheet.xls'
