require 'roo'
require 'spreadsheet'

Dir.chdir(File.dirname(__FILE__))

# Running Roo + Spreadsheet b/c Spreadsheet doesn't like xlsx
spreadsheets = Roo::Spreadsheet.open('./keyword-list.xlsx')
keys_we_track = spreadsheets.sheet(0)
our_keys = []
book = Spreadsheet::Workbook.new
sheet1 = book.create_worksheet :name => 'by-volume'
sheet2 = book.create_worksheet :name => 'by-cpc'
format = Spreadsheet::Format.new :color => :blue,
                                 :weight => :bold

keys_we_track.each do |k|
    # strip all white space for comparison sake. ya never know what's going to be in an xlsx cell
    our_keys << k[0].delete(' ')
end

by_vol = spreadsheets.sheet(1)
by_vol.each_with_index do |vol, index|
    # 3.  check whether or vol is in the array our_keys
    col1 = "x" 
    if !vol[0].nil?
        col1 = vol[0]
    end
    col2 = "x"
    if !vol[1].nil?
        col2 = vol[1]
    end
    col3 = "x"
    if !vol[2].nil?
        col3 = vol[2]
    end
    col4 = "x"
    if !vol[3].nil?
        col4 = vol[3]
    end
    col5 = "x"
    if !vol[4].nil?
        col5 = vol[4]
    end

    inc_flag = false
    col1_concat = col1.delete(' ')
    inc_flag = our_keys.include?(col1_concat)

    if (inc_flag == true)
        # 4.  if 450_by_vol is in our_keys list, indicate that it's being tracked.
        sheet1.row(index).push(col1).set_format(0, format)
        sheet1.row(index).insert 1, col2
        sheet1.row(index).insert 2, col3
        sheet1.row(index).insert 3, col4
        sheet1.row(index).insert 4, col5
   else (inc_flag == false)
        # 5.  if 450_by_vol isn't in our keys list, print without any style editing.
        sheet1.row(index).push(col1)
        sheet1.row(index).insert 1, col2
        sheet1.row(index).insert 2, col3
        sheet1.row(index).insert 3, col2
        sheet1.row(index).insert 4, col4
    end
end

# 6.  do the same as above but for the top 450 CPC entries!

by_cpc = spreadsheets.sheet(2)
by_cpc.each_with_index do |vol, index|
    # 3.  check whether or vol is in the array our_keys
    col1 = "x" 
    if !vol[0].nil?
        col1 = vol[0]
    end
    col2 = "x"
    if !vol[1].nil?
        col2 = vol[1]
    end
    col3 = "x"
    if !vol[2].nil?
        col3 = vol[2]
    end
    col4 = "x"
    if !vol[3].nil?
        col4 = vol[3]
    end
    col5 = "x"
    if !vol[4].nil?
        col5 = vol[4]
    end

    inc_flag = false
    col1_concat = col1.delete(' ')
    inc_flag = our_keys.include?(col1_concat)

    if (inc_flag == true)
        # 4.  if 450_by_vol is in our_keys list, indicate that it's being tracked.
        sheet2.row(index).push(col1).set_format(0, format)
        sheet2.row(index).insert 1, col2
        sheet2.row(index).insert 2, col3
        sheet2.row(index).insert 3, col4
        sheet2.row(index).insert 4, col5
   else (inc_flag == false)
        # 5.  if 450_by_vol isn't in our keys list, print without any style editing.
        sheet2.row(index).push(col1)
        sheet2.row(index).insert 1, col2
        sheet2.row(index).insert 2, col3
        sheet2.row(index).insert 3, col2
        sheet2.row(index).insert 4, col4
    end
end

book.write './keyword-list-new.xls'
print("Cool dude, your script ran without error.  check out the new file at keyword-list-new.xls\n")
