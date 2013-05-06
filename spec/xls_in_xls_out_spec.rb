require 'spec_helper'

@xls

describe import_xls do
  it 'opens the specified file in a Roo object' do
    @xls = Roo::Excel.new('./files/template.xls')
  end
  it 'runs inspect and returns not nil' do
    @xls.inspect
  end
end

@gens

describe '#new' do
  it 'creates a new GenSheet object' do
    @gens = GenSheet.new
  end
  it 'creates an xls file' do
    @gens.roo_to_xls(xls, "./files/out.xls")
  end
end
