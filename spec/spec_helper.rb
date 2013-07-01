require_relative '../lib/GenSheet.rb'

def open_xls
  xls = GenSheet::Spreadsheet.open('./spec/files/template.xls')
  xls.inspect
  return xls
end

def open_ods
  ods = GenSheet::Spreadsheet.open('./spec/files/template.ods')
  ods.inspect
  return ods
end
