require_relative 'spec_helper'

describe "Worksheets that failed to compile" do
worksheet_1_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="9">
      <c r="R9">
        <f t="shared" ref="R9:R68" si="0">I9+J9+K9+P9+Q9</f>
        <v>582.62431955664806</v>
      </c>
    </row>
    <row r="10">
      <c r="R10">
        <f t="shared" si="0"/>
        <v>436.4944796177661</v>
      </c>
    </row>
  </sheetData>
</worksheet>
END

worksheet_1_ruby =<<END
# coding: utf-8
# Outputs
class Sheet1 < Spreadsheet
  def r9; @r9 ||= i9+j9+k9+p9+q9; end
  def r10; @r10 ||= i10+j10+k10+p10+q10; end
end

END
it "2050 failed sheet should compile" do
  SheetNames.instance['Sheet2'] = 'sheet2'
  # SharedStrings.instance[24] = 'A shared string'
  worksheet = Worksheet.new(Nokogiri::XML(worksheet_1_xml))
  worksheet.workbook = mock(:workbook,:indirects_used => true)
  worksheet.name = "sheet1"
  worksheet.to_ruby.should == worksheet_1_ruby
end
end