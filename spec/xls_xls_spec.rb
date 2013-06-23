require 'spec_helper'

describe 'opens an xls file' do
  it 'opens and inspects an xls file' do
    open_xls()
  end
end

describe '#to_xls' do
  it 'creates an xls file' do
    create_gens_from_xls()
    @xls.to_xls('./spec/files/xls_out.xls')
  end
end
