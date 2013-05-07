require 'spec_helper'

@xls
describe 'open xls with roo' do
  it 'opens and inspects an xls file' do
    @xls = Roo::Spreadsheet.open('./spec/files/template.xls')
    @xls.inspect
  end
end

@gens

describe '#new' do
  it 'creates a new GenSheet object' do
    @gens = GenSheet.new
  end
end

describe '#roo_to_xls' do
  it 'creates an xls file' do
    @gens.roo_to_xls(@xls, './spec/files/out.xls')
  end
end
