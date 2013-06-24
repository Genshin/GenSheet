require_relative '../lib/GenSheet.rb'

@xls
@ods

def open_xls()
  @xls = GenSheet.open('./spec/files/template.xls')
  @xls.inspect
end

def open_ods()
  @ods = GenSheet.open('./spec/files/template.ods')
  @ods.inspect
end
