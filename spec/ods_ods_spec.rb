require 'spec_helper'

describe 'opens an ods file' do
  it 'opens and inspects an ods file' do
    open_ods
  end
end

describe 'ods #to_ods' do
  it 'creates an ods file' do
    open_ods.to_ods('./spec/files/ods_out.ods')
  end
end
