require 'roo'
require 'writeexcel'
require 'odf/spreadsheet'

class Roo::GenericSpreadsheet

  public

  def to_ods(filename = nil)
    ods = ODF::SpreadSheet.new
    _workbook_to_ods(ods)
    ods.write_to(filename) unless filename.nil?
    return ods
  end

  private

  def _workbook_to_ods(ods)
    tables, sheets = [], []
    self.each_with_pagename do |name, sheet|
      # シート作成
      tables << (ods.table name)
      sheets << sheet.clone
    end

    _set_ods_sheets(ods, tables, sheets)
    return ods
  end

  def _set_ods_sheets(ods, tables, sheets)
    for i in 0..sheets.size - 1
      sheets[i].each_with_index do |row, y|
        # 行作成
        ob_row = tables[i].row
        _set_cell_ods(ods, row, ob_row)
      end
    end
  end

  def _set_cell_ods(ods, row, ob_row)
    row.each_with_index do |cell, x|
      # スタイル作成
      ods.style 'font-style', family: :cell do
        # property :text, 'font-weight' => 'bold', 'color' => '#ff0000'
      end
      # セル作成、スタイル適用
      ob_row.cell(cell, style: 'font-style')
    end
  end

  public

  def to_xls(filename = nil)
    xls = WriteExcel.new(filename)

    # 出力シート準備、元シート分解
    outsheets, originalsheets, names = [], [], []
    _workbook_to_xls(xls, outsheets, originalsheets, names)

    # mergeしてあるセルのセット
    merges = []
    _set_sheet_mergedcells(xls, outsheets, merges)

    # シートのセット
    _set_cell_formating(xls, outsheets, originalsheets, names, merges)

    xls.close
    return xls
  end

  private

  def _workbook_to_xls(xls, outsheets, originalsheets, names)
    self.each_with_pagename do |name, sheet|
      outsheets << xls.add_worksheet(name)
      originalsheets << sheet.clone
      names << name
    end
  end

  def _set_sheet_mergedcells(xls, outsheets, merges)
    worksheets = @workbook.instance_variable_get('@worksheets')
    for i in 0..worksheets.size - 1
      merges << worksheets[i].instance_variable_get('@merged_cells')
      if merges[i] != nil && merges[i] != []
        _set_mergedcell(xls, outsheets[i], merges[i])
      end
    end
  end

  def _set_mergedcell(xls, outsheet, merge)
    format = xls.add_format(align: 'merge')
    for j in 0..merge.size - 1
      outsheet.merge_range(merge[j][0], merge[j][2],
                           merge[j][1], merge[j][3], '', format)
    end
  end

  def _set_cell_formating(xls, outsheets, originalsheets, names, merges)
    # fontもborderも入ってるけどもインデックスがわからない
    for i in 0..originalsheets.size - 1
      originalsheets[i].each_with_index do |row, y|
        row.each_with_index do |cell, x|
          font = @fonts[names[i]][[y + 1, x + 1]]
          _set_cell_xls(xls, outsheets[i], cell, x, y, font, merges[i])
        end
      end
    end
  end

  def _set_cell_xls(xls, outsheet, cell, x, y, font, merge)
    # cellが空の場合は何もしない
    return if cell == nil

    # 各フォントフォーマットのセット
    format = _set_format(xls, font)

    # mergeしてあるセルの場合フォーマット追加
    for i in 0..merge.size - 1
      set_mergedcell_format(format) if y == merge[i][0] && x == merge[i][2]
    end

    outsheet.write(y, x, cell, format)
  end

  def _set_format(xls, font)
    format = xls.add_format(
      # :bottom => 1,      #
      # :top => 1,         # border
      # :left => 1,        #
      # :right => 1,       #
      color:    font.instance_variable_get('@color'),
      italic:   font.instance_variable_get('@italic') ? 1 : 0,
      font:     font.instance_variable_get('@name'),
      outline:  font.instance_variable_get('@outline'),
      shadow:   font.instance_variable_get('@shadow'),
      size:     font.instance_variable_get('@size'),
      strikeout:  font.instance_variable_get('@strikeout'),
      # アンダーラインの種類は4つだけどとりあえず
      underline:  font..instance_variable_get('@underline') == :none ? 0 : 1,
      # weightが普通だと400、boldだと700になるようなのでとりあえず
      bold: font.instance_variable_get('@weight') > 400 ? 1 : 0
      # :encoding => font.encoding,                    #
      # :escapement => font.escapement,                # fontにまとめて入っていたけど
      # :family => font.family,                        # どれに対応するのか..
      # :previous_fast_key => font.previous_fast_key,  #
    )

    return format
  end

  def set_mergedcell_format(format)
    format.set_align('center')
    format.set_valign('vcenter')
  end
end

class GenSheet < Roo::Spreadsheet
end
