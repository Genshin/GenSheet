require 'roo'
require 'writeexcel'
require 'odf/spreadsheet'

class GenSheet

  @roo

  def to_xls(filename)
#    outbook = WriteExcel.new(filename)
#
#    @workbook = @roo.instance_variable_get('@workbook')
#    @formats = @workbook.instance_variable_get('@formats')
#    @worksheets = @workbook.instance_variable_get('@worksheets')
#
#    i = 0
#    @roo.each_with_pagename do |name, sheet|
#      outsheet =  outbook.add_worksheet(name)
#
#      # mergeしてあるセルの取り出し
#      @merge = @worksheets[i].instance_variable_get('@merged_cells')
#      if @merge != nil
#        format_merged = outbook.add_format(:align => 'merge')
#        for num in 0..@merge.size - 1
#          outsheet.merge_range(@merge[num][0], @merge[num][2], @merge[num][1], @merge[num][3], '', format_merged)
#        end
#      end
#
#      # フォント設定の取り出し
#      getfonts = @roo.instance_variable_get('@fonts')
#      sheet.each_with_index do |row, y|
#        row.each_with_index do |cell, x|
#          if cell != nil
#            # 各フォント設定の取り出し
#            @getfont = getfonts[name][[y + 1, x + 1]]
#            #@getformat = @formats[(x + 1) * (y + 1) - 1]   # fontもborderも入ってるけどもインデックスがわからない
#            format = outbook.add_format(
#              #:bottom => 1,      #
#              #:top => 1,         # border
#              #:left => 1,        #
#              #:right => 1,       #
#              :color => @getfont.color,
#              :italic => @getfont.italic ? 1 : 0,
#              :font => @getfont.name,
#              :outline => @getfont.outline,
#              :shadow => @getfont.shadow,
#              :size => @getfont.size,
#              :strikeout => @getfont.strikeout,
#              :underline => @getfont.underline == :none ? 0 : 1,  # アンダーラインの種類は4つだけどとりあえず
#              :bold => @getfont.weight > 400 ? 1 : 0              # weightが普通だと400、boldだと700になるようなのでとりあえず
#              #:encoding => @getfont.encoding,                    #
#              #:escapement => @getfont.escapement,                # fontにまとめて入っていたけど
#              #:family => @getfont.family,                        # どれに対応するのか..
#              #:previous_fast_key => @getfont.previous_fast_key,  #
#            )
#            # mergeしてあるセルの場合フォーマット追加
#            for num in 0..@merge.size - 1
#              if y == @merge[num][0] && x == @merge[num][2]
#                format.set_align('center')
#                format.set_valign('vcenter')
#              end
#            end
#
#            outsheet.write(y, x, cell, format)
#
#          end
#        end
#      end
#
#      sheet.parse do |row|
#        puts row
#      end
#
#    i += 1
#    end
#
#    outbook.close
  end

  def to_ods(filename)
#    puts "converting sheet"
#    outbook = ODF::SpreadSheet.new
#
#    workbook.each_with_pagename do |name, sheet|
#      # シート作成
#      ob_table = outbook.table name
#
#      sheet.each_with_index do |row, y|
#        # 行作成
#        ob_row = ob_table.row
#
#        row.each_with_index do |cell, x|
#          # スタイル作成
#          outbook.style 'font-style', :family => :cell do
#            #property :text, 'font-weight' => 'bold', 'color' => '#ff0000'
#          end
#
#          # セル作成、スタイル適用
#          ob_cell = ob_row.cell(cell, :style => 'font-style')
#        end
#      end
#    end
#
#    outbook.write_to filename
  end

  def initialize(roo)
    @roo = roo
  end
end
