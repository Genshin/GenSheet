require 'roo'
require 'writeexcel'
require 'odf/spreadsheet'

class GenSheet
  def roo_to_xls(workbook, filename)
    puts "converting sheet"
    outbook = WriteExcel.new(filename)

    workbook.each_with_pagename do |name, sheet|
      outsheet =  outbook.add_worksheet(name)

      sheet.each_with_index do |row, y|
        row.each_with_index do |cell, x|
          if cell != nil
            font = sheet.font(y, x)
            if font != nil
              format = outbook.add_format(
                :italic => 1#font.italic
              )
              #format.set_color(font.color)
              outsheet.write(y, x, cell, format)
            else
              outsheet.write(y, x, cell)
            end
          end
        end

        #outsheet.write_row(y, 0, row)
      end

      sheet.parse do |row|
        puts row
      end


    end

    outbook.close
  end

end
