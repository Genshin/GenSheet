require 'spec_helper'

describe '#to_ods' do
  it 'creates an ods file from an xls base' do
    create_gens_from_xls()
    @gens.to_ods('./spec/files/xls_out.ods')
  end
end
