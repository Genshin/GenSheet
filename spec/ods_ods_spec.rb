require 'spec_helper'

describe 'opens an ods file' do
  it 'opens and inspects an xls file' do
    open_ods()
  end
end

describe '#to_ods' do
  it 'creates an ods file' do
    open_ods()
    @ods.to_ods('./spec/files/ods_out.ods')
  end
end
