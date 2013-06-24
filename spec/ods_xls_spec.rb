require 'spec_helper'

describe 'ods #to_xls' do
  it 'creates an xls file from an ods base' do
    open_ods()
    @ods.to_xls('./spec/files/ods_out.xls')
  end
end
