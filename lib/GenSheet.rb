require 'roo'
require 'writeexcel'
require 'odf/spreadsheet'

class GenSheet

  @roo

  def to_xls(filename)
    @outbook_xls = WriteExcel.new(filename)

    workbook = @roo.instance_variable_get('@workbook')
    formats = workbook.instance_variable_get('@formats')
    worksheets = workbook.instance_variable_get('@worksheets')

    # 出力シート準備、元シート分解
    outsheets = Array.new
    originalsheets = Array.new
    names = Array.new
    pre_sheet_xls(outsheets, originalsheets, names)

    # mergeしてあるセルのセット
    merges = Array.new
    set_mergedcell_sheet(worksheets, outsheets, merges)

    # シートのセット
    set_sheet_xls(outsheets, originalsheets, names, merges)

    # # 出力
    # for i in 0..originalsheets.size - 1
      # originalsheets[i].parse do |row|
        # puts row
      # end
    # end

    @outbook_xls.close
  end

  def pre_sheet_xls(outsheets, originalsheets, names)
    @roo.each_with_pagename do |name, sheet|
      outsheets << @outbook_xls.add_worksheet(name)
      originalsheets << sheet.clone
      names << name
    end
  end

  def set_mergedcell_sheet(worksheets, outsheets, merges)
    for i in 0..worksheets.size - 1
      merges << worksheets[i].instance_variable_get('@merged_cells')
      if merges[i] != nil && merges[i] != []
        set_mergedcell(outsheets[i], merges[i])
      end
    end
  end

  def set_mergedcell(outsheet, merge)
    format = @outbook_xls.add_format(:align => 'merge')
    for j in 0..merge.size - 1
      outsheet.merge_range(merge[j][0], merge[j][2], merge[j][1], merge[j][3], '', format)
    end
  end

  def set_sheet_xls(outsheets, originalsheets, names, merges)
    #@getformat = @formats[(x + 1) * (y + 1) - 1]   # fontもborderも入ってるけどもインデックスがわからない
    fonts = @roo.instance_variable_get('@fonts')
    for i in 0..originalsheets.size - 1
      originalsheets[i].each_with_index do |row, y|
        row.each_with_index do |cell, x|
          font = fonts[names[i]][[y + 1, x + 1]]
          set_cell_xls(outsheets[i], cell, x, y, font, merges[i])
        end
      end
    end
  end

  def set_cell_xls(outsheet, cell, x, y, font, merge)
    # cellが空の場合は何もしない
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
    format = @outbook_xls.add_format(
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
    @outbook_ods = ODF::SpreadSheet.new

    tables = Array.new
    sheets = Array.new
    pre_sheet_ods(tables, sheets)

    @outbook_ods.write_to filename
  end

  def pre_sheet_ods(tables, sheets)
    @roo.each_with_pagename do |name, sheet|
      # シート作成
      tables << (@outbook_ods.table name)
      sheets << sheet.clone
    end

    set_sheet_ods(tables, sheets)
  end

  def set_sheet_ods(tables, sheets)
    for i in 0..sheets.size - 1
      sheets[i].each_with_index do |row, y|
        # 行作成
        ob_row = tables[i].row
        set_cell_ods(row, ob_row)
      end
    end
  end

  def set_cell_ods(row, ob_row)
    row.each_with_index do |cell, x|
      # スタイル作成
      @outbook_ods.style 'font-style', :family => :cell do
        #property :text, 'font-weight' => 'bold', 'color' => '#ff0000'
      end
      # セル作成、スタイル適用
      ob_cell = ob_row.cell(cell, :style => 'font-style')
    end
  end

  def initialize(roo)
    @roo = roo
  end

end