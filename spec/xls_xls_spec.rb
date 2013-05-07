require 'spec_helper'

def open_xls()
  @xls = Roo::Spreadsheet.open('./spec/files/template.xls')
  @xls.inspect
end

@xls
describe 'open xls with roo' do
  it 'opens and inspects an xls file' do
    open_xls()
  end
end

@gens

def create_gens()
  open_xls()
  @gens = GenSheet.new(@xls)
end

describe '#new' do
  it 'creates a new GenSheet object' do
    create_gens()
  end
end

describe '#to_xls' do
  it 'creates an xls file' do
    create_gens()
    @gens.to_xls('./spec/files/out.xls')
  end
end
