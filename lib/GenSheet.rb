require 'roo'
require 'GenSheetExporters'

class GenSheet
  class << self
    def open(file)
      sheet =Roo::Spreadsheet.open(file)
      sheet.extend GenSheetExporters
    end
  end
end
