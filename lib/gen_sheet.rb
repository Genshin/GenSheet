require 'roo'
require 'writeexcel'
require 'odf/spreadsheet'
require 'pry'

class GenSheet
  def roo_to_xls(workbook, filename)
    puts "converting sheet"
    outbook = WriteExcel.new(filename)

    workbook.each_with_pagename do |name, sheet|
      outsheet =  outbook.add_worksheet(name)

      getfonts =  workbook.instance_variable_get('@fonts')
      sheet.each_with_index do |row, y|
        row.each_with_index do |cell, x|
          if cell != nil
            #font = sheet.font(y, x)
            @getfont = getfonts[name][[y + 1, x + 1]]
            format = outbook.add_format(
              :color => @getfont.color,
              #:encoding => @getfont.encoding,
              #:escapement => @getfont.escapement,
              #:family => @getfont.family,
              :italic => @getfont.italic ? 1 : 0,
              :font => @getfont.name,
              :outline => @getfont.outline,
              #:previous_fast_key => @getfont.previous_fast_key,
              :shadow => @getfont.shadow,
              :size => @getfont.size,
              :strikeout => @getfont.strikeout,
              :underline => @getfont.underline == :none ? 0 : 1,
              :bold => @getfont.weight > 400 ? 1 : 0    # weightが普通だと400、boldだと700になるようなのでとりあえず
            )
            outsheet.write(y, x, cell, format)

            # if font != nil
              # format = outbook.add_format(
                # :italic => 1#font.italic
              # )
              # #format.set_color(font.color)
              # outsheet.write(y, x, cell, format)
            # else
              # outsheet.write(y, x, cell)
            # end
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
