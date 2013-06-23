require 'spec_helper'

describe '#to_xls' do
  it 'creates an xls file from an ods base' do
    create_gens_from_ods()
    @ods.to_xls('./spec/files/ods_out.xls')
  end
end
