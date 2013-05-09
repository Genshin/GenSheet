require 'roo'
require 'writeexcel'
require 'odf/spreadsheet'

class GenSheet

  @roo

  def to_xls(filename)
    @outbook = WriteExcel.new(filename)

    workbook = @roo.instance_variable_get('@workbook')
    formats = workbook.instance_variable_get('@formats')
    worksheets = workbook.instance_variable_get('@worksheets')

    outsheets = Array.new
    originalsheets = Array.new
    names = Array.new

    # 出力シート準備、元シート分解
    @roo.each_with_pagename do |name, sheet|
      # FIXME MySheetが上書きされる…
      outsheets << @outbook.add_worksheet(name)
      originalsheets << sheet
      names << name
    end

    # mergeしてあるセルのセット
    for i in 0..worksheets.size - 1
      merge = worksheets[i].instance_variable_get('@merged_cells')
      if merge != nil
        set_mergedcell(outsheets[i], merge)
      end
    end

    # シートのセット
    fonts = @roo.instance_variable_get('@fonts')
    for i in 0..originalsheets.size - 1
      set_sheet(outsheets[i], originalsheets[i], names[i], merge, fonts)
    end

    # 出力
    for i in 0..originalsheets.size - 1
      originalsheets[i].parse do |row|
        puts row
      end
    end

    @outbook.close
  end

  def set_mergedcell(outsheet, merge)
    format = @outbook.add_format(:align => 'merge')
    for j in 0..merge.size - 1
      outsheet.merge_range(merge[j][0], merge[j][2], merge[j][1], merge[j][3], '', format)
    end
  end

  def set_sheet(outsheet, originalsheet, name, merge, fonts)
    #@getformat = @formats[(x + 1) * (y + 1) - 1]   # fontもborderも入ってるけどもインデックスがわからない
    originalsheet.each_with_index do |row, y|
      row.each_with_index do |cell, x|
        # XXX なぜかMySheetの情報が消えるので一時的に挿入
        break if !fonts.key?(name)
        font = fonts[name][[y + 1, x + 1]]
        set_cell(outsheet, cell, x, y, font, merge)
      end
    end
  end

  def set_cell(outsheet, cell, x, y, font, merge)
    return if cell == nil

    # 各フォントフォーマットのセット
    format = set_format(font)

    # mergeしてあるセルの場合フォーマット追加
    for i in 0..merge.size - 1
      if y == merge[i][0] && x == merge[i][2]
        set_mergedcell_format(format)
      end
    end

    outsheet.write(y, x, cell, format)
  end

  def set_format(font)
    format = @outbook.add_format(
      #:bottom => 1,      #
      #:top => 1,         # border
      #:left => 1,        #
      #:right => 1,       #
      :color => font.color,
      :italic => font.italic ? 1 : 0,
      :font => font.name,
      :outline => font.outline,
      :shadow => font.shadow,
      :size => font.size,
      :strikeout => font.strikeout,
      :underline => font.underline == :none ? 0 : 1,  # アンダーラインの種類は4つだけどとりあえず
      :bold => font.weight > 400 ? 1 : 0              # weightが普通だと400、boldだと700になるようなのでとりあえず
      #:encoding => font.encoding,                    #
      #:escapement => font.escapement,                # fontにまとめて入っていたけど
      #:family => font.family,                        # どれに対応するのか..
      #:previous_fast_key => font.previous_fast_key,  #
    )

    return format
  end

  def set_mergedcell_format(format)
    format.set_align('center')
    format.set_valign('vcenter')
  end

  def to_ods(filename)
    puts "converting sheet"
    outbook = ODF::SpreadSheet.new

    @roo.each_with_pagename do |name, sheet|
      # シート作成
      ob_table = outbook.table name

      sheet.each_with_index do |row, y|
        # 行作成
        ob_row = ob_table.row

        row.each_with_index do |cell, x|
          # スタイル作成
          outbook.style 'font-style', :family => :cell do
            #property :text, 'font-weight' => 'bold', 'color' => '#ff0000'
          end

          # セル作成、スタイル適用
          ob_cell = ob_row.cell(cell, :style => 'font-style')
        end
      end
    end

    outbook.write_to filename
  end

  def initialize(roo)
    @roo = roo
  end

end