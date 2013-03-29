require 'test/unit'
require 'roo'
require '../lib/gen_sheet.rb'

class TestGenSheet < Test::Unit::TestCase

  def setup
  end

  def test_xls_template
    xls = Roo::Spreadsheet.open('./files/template.xls')
    assert_not_nil(xls, "Roo could not open the sample XLS sheet.")
    assert(xls.inspect, "XLS inspection fails.")
    gens = GenSheet.new
    gens.roo_to_xls(xls, "out.xls")
  end

  def test_ods_template
    ods = Roo::Spreadsheet.open('./files/template.ods')
    assert_not_nil(ods, "Roo could not open the sample ODS sheet.")
    assert(ods.inspect, "ODS inspection fails.")
    gens = GenSheet.new
    gens.roo_to_ods(ods, "out.ods")
  end

end

