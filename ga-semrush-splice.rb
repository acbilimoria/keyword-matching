# Running Roo + Spreadsheet b/c Spreadsheet doesn't like xlsx
require 'roo'
require 'spreadsheet'

Dir.chdir(File.dirname(__FILE__))

ga_data = Roo::Spreadsheet.open('./ga-data.xlsx')
sem_data = Roo::Spreadsheet.open('./sem-rush-data.xlsx')
result = Spreadsheet::Workbook.new
sheet1 = result.create_worksheet :name => 'sheet1'

sem_rush_kywds = []

sem_queries = []
ga_queries = []

ga_data.each do |gad|
    print(gad)
    ga_queries << gad[0].delete(' ')
end

sem_data.each do |semd|
    print(semd)
    sem_queries << semd[0].delete(' ')
end

sem_data.each_with_index do |sem_row, index|
    ga_data.each do |ga_row|
        s = sem_row[0].to_s.delete(' ')
        g = ga_row[0].to_s.delete(' ')[0...-1]
        final_row = ""

        if g == s
            final_row = sem_row + ga_row
            new_row = sheet1.row(index)
            new_row.replace final_row
            break
        else g != s
            final_row = sem_row
            new_row = sheet1.row(index)
            new_row.replace final_row
       end
    end
end

result.write './result.xls'
print("\nCool dude, your script ran without any errors.\n")
