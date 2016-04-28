# Running Roo + Spreadsheet b/c Spreadsheet doesn't like xlsx
require 'roo'
require 'spreadsheet'

Dir.chdir(File.dirname(__FILE__))

spreadsheets = Roo::Spreadsheet.open('./keyword-list.xlsx')
keys_we_track = spreadsheets.sheet(0)
book = Spreadsheet::Workbook.new
sheet1 = book.create_worksheet :name => 'by-volume'
sheet2 = book.create_worksheet :name => 'by-cpc'
format = Spreadsheet::Format.new :color => :blue, :weight => :bold
our_keys = []
keys_we_track.each do |k|
    our_keys << k[0].delete(' ')
end
by_vol = spreadsheets.sheet(1)
by_cpc = spreadsheets.sheet(2)

by_vol.each_with_index do |row, index|
    inc_flag = false
    col0_concat = row[0].delete(' ')
    inc_flag = our_keys.include?(col0_concat)

    if (inc_flag == true)
        sheet1.row(index).push(row[0]).set_format(0, format)
        sheet1.row(index).insert 1, row[1]
        sheet1.row(index).insert 2, row[2]
        sheet1.row(index).insert 3, row[3]
        sheet1.row(index).insert 4, row[4]
        sheet2.row(index).push(row[0]).set_format(0, format)
        sheet2.row(index).insert 1, row[1]
        sheet2.row(index).insert 2, row[2]
        sheet2.row(index).insert 3, row[3]
        sheet2.row(index).insert 4, row[4]
   else (inc_flag == false)
        sheet1.row(index).push(row[0])
        sheet1.row(index).insert 1, row[1]
        sheet1.row(index).insert 2, row[2]
        sheet1.row(index).insert 3, row[3]
        sheet1.row(index).insert 4, row[4]
        sheet2.row(index).push(row[0])
        sheet2.row(index).insert 1, row[1]
        sheet2.row(index).insert 2, row[2]
        sheet2.row(index).insert 3, row[3]
        sheet2.row(index).insert 4, row[4]
    end
end

book.write './keyword-list-new.xls'
print("Cool dude, your script ran without error.  check out the new file at keyword-list-new.xls\n")
