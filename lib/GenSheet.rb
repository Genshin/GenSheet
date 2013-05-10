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

    # 出力シート準備、元シート分解
    outsheets = Array.new
    originalsheets = Array.new
    names = Array.new
    pre_sheet(outsheets, originalsheets, names)

    # mergeしてあるセルのセット
    merges = Array.new
    set_mergedcell(worksheets, outsheets, merges)

    # シートのセット
    set_sheet(outsheets, originalsheets, names, merges)

    # # 出力
    # for i in 0..originalsheets.size - 1
      # originalsheets[i].parse do |row|
        # puts row
      # end
    # end

    @outbook.close
  end

  def pre_sheet(outsheets, originalsheets, names)
    @roo.each_with_pagename do |name, sheet|
      # FIXME MySheetが上書きされる…
      outsheets << @outbook.add_worksheet(name)
      originalsheets << sheet.clone
      names << name
    end
  end

  def set_mergedcell(worksheets, outsheets, merges)
    for i in 0..worksheets.size - 1
      merges << worksheets[i].instance_variable_get('@merged_cells')
      if merges[i] != nil && merges[i] != []
        format = @outbook.add_format(:align => 'merge')
        for j in 0..merges.size - 1
          temp = merges[i]
          outsheets[i].merge_range(temp[j][0], temp[j][2], temp[j][1], temp[j][3], '', format)
        end
      end
    end
  end

  def set_sheet(outsheets, originalsheets, names, merges)
    #@getformat = @formats[(x + 1) * (y + 1) - 1]   # fontもborderも入ってるけどもインデックスがわからない
    fonts = @roo.instance_variable_get('@fonts')
    for i in 0..originalsheets.size - 1
      originalsheets[i].each_with_index do |row, y|
        row.each_with_index do |cell, x|
          font = fonts[names[i]][[y + 1, x + 1]]
          set_cell(outsheets[i], cell, x, y, font, merges[i])
        end
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
      :color => font.instance_variable_get('@color'),
      :italic => font.instance_variable_get('@italic') ? 1 : 0,
      :font => font.instance_variable_get('@name'),
      :outline => font.instance_variable_get('@outline'),
      :shadow => font.instance_variable_get('@shadow'),
      :size => font.instance_variable_get('@size'),
      :strikeout => font.instance_variable_get('@strikeout'),
      :underline => font..instance_variable_get('@underline') == :none ? 0 : 1,  # アンダーラインの種類は4つだけどとりあえず
      :bold => font.instance_variable_get('@weight') > 400 ? 1 : 0              # weightが普通だと400、boldだと700になるようなのでとりあえず
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