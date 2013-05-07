require 'roo'
require_relative '../lib/GenSheet.rb'

@xls
@ods
@gens

def open_xls()
  @xls = Roo::Spreadsheet.open('./spec/files/template.xls')
  @xls.inspect
end

def open_ods()
  @ods = Roo::Spreadsheet.open('./spec/files/template.ods')
  @ods.inspect
end

def create_gens_from_xls()
  open_xls()
  @gens = GenSheet.new(@xls)
end

def create_gens_from_ods()
  open_ods()
  @gens = GenSheet.new(@ods)
end
