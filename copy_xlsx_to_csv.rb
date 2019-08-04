require 'roo'
require 'spreadsheet'

var1 = '/Users/admin/script/example.xlsx'
var2 = File.basename(var1)

workbook = Roo::Spreadsheet.open '/your_path/script_excel/example.xlsx'
worksheets = workbook.sheets

book = Spreadsheet::Workbook.new
sheet = book.create_worksheet
num_rows = workbook.sheet("sheet1").last_row

puts "Reading: #{var1}"
puts "Processing..."
sleep(1)
puts "#{num_rows} rows"

i = 1
loop do
  i+=1
  if i < 30
    if workbook.sheet("sheet1").cell(1,2) == "Name" && workbook.sheet("sheet1").cell(1,4) == "Country"
      var3 = workbook.sheet("sheet1").column(2)
      var4 = workbook.sheet("sheet1").column(4)
    end
  else
    break
  end

  y = 1
  loop do
    sheet.row(0).default_format = Spreadsheet::Format.new(weight: :bold)
    sheet.row(0).replace(["NAME", "COUNTRY", "FILE NAME"])
    sheet.row(y).replace([var3[y], var4[y], var2])
    y+=1
    if y == num_rows
      break
    end
  end
end

puts "File name: #{var2}"

spreadsheet = StringIO.new
export_file_path = "new_report.csv"

book.write export_file_path
puts "Processing..."
sleep(0.5)
puts "Done"
