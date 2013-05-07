require 'spec_helper'

describe 'open ods with roo' do
  it 'opens and inspects an xls file' do
    open_ods()
  end
end

describe '#new' do
  it 'creates a new GenSheet object from Roo ODS' do
    create_gens_from_ods()
  end
end

describe '#to_ods' do
  it 'creates an ods file' do
    create_gens_from_ods()
    @gens.to_ods('./spec/files/ods_out.ods')
  end
end
