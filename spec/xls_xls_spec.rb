require 'spec_helper'

describe 'open xls with roo' do
  it 'opens and inspects an xls file' do
    open_xls()
  end
end

describe '#new' do
  it 'creates a new GenSheet object from Roo XLS' do
    create_gens_from_xls()
  end
end

describe '#to_xls' do
  it 'creates an xls file' do
    create_gens_from_xls()
    @gens.to_xls('./spec/files/xls_out.xls')
  end
end
