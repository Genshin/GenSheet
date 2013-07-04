require 'writeexcel'
require 'odf/spreadsheet'

module GenSheetExporters
  def to_ods(filename = nil)
    ods = ODF::SpreadSheet.new

    _workbook_to_ods(ods)
    ods.write_to(filename) unless filename.nil?

    return ods
  end

  def to_xls(filename = nil)
    xls = WriteExcel.new(filename)

    # Prepare sheets & get original sheet's data
    outsheets, originalsheets, names = [], [], []
    _workbook_to_xls(xls, outsheets, originalsheets, names)

    _set_data_to_xls(xls, outsheets, originalsheets, names)

    xls.close

    return xls
  end

  private

  def _workbook_to_ods(ods)
    tables, sheets = [], []

    # Create sheets
    self.each_with_pagename do |name, sheet|
      tables << (ods.table name)
      sheets << sheet.clone
    end

    _set_ods_sheets(ods, tables, sheets)

    return ods
  end

  def _set_ods_sheets(ods, tables, sheets)
    (0..sheets.size - 1).each do |i|
      sheets[i].each_with_index do |row, y|
        # Create rows
        ob_row = tables[i].row
        _set_cell_ods(ods, row, ob_row)
      end
    end
  end

  def _set_cell_ods(ods, row, ob_row)
    row.each_with_index do |cell, x|
      # Create styles
      ods.style 'font-style', family: :cell do
        # property :text, 'font-weight' => 'bold', 'color' => '#ff0000'
      end
      # Create cell & add style
      ob_row.cell(cell, style: 'font-style')
    end
  end

  def _workbook_to_xls(xls, outsheets, originalsheets, names)
    self.each_with_pagename do |name, sheet|
      outsheets << xls.add_worksheet(name)
      originalsheets << sheet.clone
      names << name
    end
  end

  def _set_data_to_xls(xls, outsheets, originalsheets, names)
    # Take out the merged data that is separate
    worksheets = @workbook.instance_variable_get('@worksheets')
    merges = _get_merge_data(worksheets) unless worksheets.nil?

    # Set the data in each sheet
    (0..outsheets.size - 1).each do |i|
      _set_merge_data(xls, outsheets[i], merges[i]) unless merges.nil?
      cells = _get_cell_data(originalsheets[i], names[i])
      _get_format_data(xls, cells, merges[i]) unless merges.nil?
      _set_cell_data(outsheets[i], cells)
    end
  end

  def _get_merge_data(worksheets)
    merges = []

    (0..worksheets.size - 1).each do |i|
      merges << worksheets[i].instance_variable_get('@merged_cells')
    end

    return merges
  end

  def _set_merge_data(xls, outsheet, merge)
    format = xls.add_format(align: 'merge')

    (0..merge.size - 1).each do |i|
      outsheet.merge_range(merge[i][0], merge[i][2],
                           merge[i][1], merge[i][3], '', format)
    end
  end

  def _get_cell_data(originalsheet, name)
    cells = []

    originalsheet.each_with_index do |row, y|
      row.each_with_index do |cell, x|
        font = @fonts.nil? ? nil : @fonts[name][[y + 1, x + 1]]
        cell = CellData.new(x, y, cell, font)
        cells << cell
      end
    end

    return cells
  end

  def _get_format_data(xls, cells, merge)
    cells.each do |cell|
      format = _set_format(xls, cell.format) unless cell.format.nil?

      # add format when merged cell
      (0..merge.size - 1).each do |i|
        set_mergedcell_format(format) if cell.y == merge[i][0] &&
                                         cell.x == merge[i][2]
      end

      cell.format = format
    end
  end

  def _set_cell_data(outsheet, cells)
    cells.each do |cell|
      outsheet.write(cell.y, cell.x, cell.data, cell.format)
    end
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
      # Underline's style have 4 pattern
      underline:  font..instance_variable_get('@underline') == :none ? 0 : 1,
      # Weight normal is 400, bold is 700
      bold: font.instance_variable_get('@weight') > 400 ? 1 : 0
      # :encoding => font.encoding,
      # :escapement => font.escapement,
      # :family => font.family,
      # :previous_fast_key => font.previous_fast_key,
    )

    return format
  end

  def set_mergedcell_format(format)
    format.set_align('center')
    format.set_valign('vcenter')
  end
end

# x position, y position, cell's data, cell's format
class CellData
  attr_reader :x, :y, :data, :format
  attr_writer :format

  def initialize(x, y, data, format)
    @x = x
    @y = y
    @data = data
    @format = format
  end
end