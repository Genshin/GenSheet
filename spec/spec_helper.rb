require_relative '../lib/GenSheet.rb'

def open_xls
  xls = GenSheet.open('./spec/files/template.xls')
  xls.inspect
  return xls
end

def open_ods
  ods = GenSheet.open('./spec/files/template.ods')
  ods.inspect
  return ods
end
