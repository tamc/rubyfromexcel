require_relative 'spec_helper'

describe Worksheet do  

SheetNames.instance['Outputs'] = 'sheet1'

simple_worksheet_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" mc:Ignorable="mv" mc:PreserveAttributes="mv:*"><dimension ref="A1:A2"/><sheetViews><sheetView view="pageLayout" workbookViewId="0"/></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="13"/><sheetData><row r="1" spans="1:1"><c r="A1" t="str"><f>IF(A2="Hello","hello",Sheet2!B4)</f><v>hello</v></c></row><row r="2" spans="1:1"><c r="A2" t="s"><v>24</v></c></row></sheetData><sheetCalcPr fullCalcOnLoad="1"/><phoneticPr fontId="2" type="noConversion"/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><pageSetup paperSize="10" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="1" OnePage="0" WScale="0"/></ext></extLst></worksheet>
END

simple_worksheet_ruby =<<END
# coding: utf-8
# Outputs
class Sheet1 < Spreadsheet
  def a1; @a1 ||= excel_if(excel_comparison(a2,"==","Hello"),"hello",sheet2.b4); end
  def a2; "A shared string"; end
end

END

simple_worksheet_test =<<END
# coding: utf-8
require_relative '../spreadsheet'
# Outputs
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell a1 should equal "hello"' do
    sheet1.a1.should == "hello"
  end

end

END

it 'should create a ruby file from xml, complete with properly interpreted cells' do
  SheetNames.instance['Sheet2'] = 'sheet2'
  SharedStrings.instance[24] = 'A shared string'
  worksheet = Worksheet.new(Nokogiri::XML(simple_worksheet_xml))
  worksheet.workbook = mock(:workbook,:indirects_used => true)
  worksheet.name = 'sheet1'
  worksheet.to_ruby.should == simple_worksheet_ruby
  worksheet.to_test.should == simple_worksheet_test
end

named_reference_simple_worksheet_ruby =<<END
# coding: utf-8
# Outputs
class Sheet1 < Spreadsheet
  def a1; @a1 ||= excel_if(excel_comparison(a2,"==","Hello"),"hello",sheet2.b4); end
  def a2; "A shared string"; end
  def reference_one; sheet2.a1; end
end

END

it 'should add any named references that apply just to this sheet' do
  SheetNames.instance['Sheet2'] = 'sheet2'
  SharedStrings.instance[24] = 'A shared string'
  worksheet = Worksheet.new(Nokogiri::XML(simple_worksheet_xml))
  worksheet.workbook = mock(:workbook,:indirects_used => true)  
  worksheet.named_references['reference_one'] = 'sheet2.a1'
  worksheet.name = 'sheet1'
  worksheet.to_ruby.should == named_reference_simple_worksheet_ruby  
end

shared_formula_worksheet_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" mc:Ignorable="mv" mc:PreserveAttributes="mv:*"><dimension ref="A1:A10"/><sheetViews><sheetView tabSelected="1" view="pageLayout" workbookViewId="0"><selection activeCell="B5" sqref="B5"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="13"/><sheetData><row r="1" spans="1:1"><c r="A1"><v>1</v></c></row><row r="2" spans="1:1"><c r="A2"><f>A1*2</f><v>2</v></c></row><row r="3" spans="1:1"><c r="A3"><f t="shared" ref="A3:A9" si="0">A2*2</f><v>4</v></c></row><row r="4" spans="1:1"><c r="A4"><f t="shared" si="0"/><v>8</v></c></row><row r="5" spans="1:1"><c r="A5"><f t="shared" si="0"/><v>16</v></c></row><row r="6" spans="1:1"><c r="A6"><f t="shared" si="0"/><v>32</v></c></row><row r="7" spans="1:1"><c r="A7"><f t="shared" si="0"/><v>64</v></c></row><row r="8" spans="1:1"><c r="A8"><f t="shared" si="0"/><v>128</v></c></row><row r="9" spans="1:1"><c r="A9"><f t="shared" si="0"/><v>256</v></c></row><row r="10" spans="1:1"><c r="A10"><f>A9*2</f><v>512</v></c></row></sheetData><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><pageSetup paperSize="10" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="1" OnePage="0" WScale="0"/></ext></extLst></worksheet>
END
shared_formula_worksheet_ruby =<<END
# coding: utf-8
# Outputs
class Sheet1 < Spreadsheet
  def a1; 1.0; end
  def a2; @a2 ||= a1*2.0; end
  def a3; @a3 ||= a2*2.0; end
  def a4; @a4 ||= a3*2.0; end
  def a5; @a5 ||= a4*2.0; end
  def a6; @a6 ||= a5*2.0; end
  def a7; @a7 ||= a6*2.0; end
  def a8; @a8 ||= a7*2.0; end
  def a9; @a9 ||= a8*2.0; end
  def a10; @a10 ||= a9*2.0; end
end

END

shared_formula_worksheet_test =<<END
# coding: utf-8
require_relative '../spreadsheet'
# Outputs
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell a2 should equal 2.0' do
    sheet1.a2.should be_within(0.2).of(2.0)
  end

  it 'cell a3 should equal 4.0' do
    sheet1.a3.should be_within(0.4).of(4.0)
  end

  it 'cell a4 should equal 8.0' do
    sheet1.a4.should be_within(0.8).of(8.0)
  end

  it 'cell a5 should equal 16.0' do
    sheet1.a5.should be_within(1.6).of(16.0)
  end

  it 'cell a6 should equal 32.0' do
    sheet1.a6.should be_within(3.2).of(32.0)
  end

  it 'cell a7 should equal 64.0' do
    sheet1.a7.should be_within(6.4).of(64.0)
  end

  it 'cell a8 should equal 128.0' do
    sheet1.a8.should be_within(12.8).of(128.0)
  end

  it 'cell a9 should equal 256.0' do
    sheet1.a9.should be_within(25.6).of(256.0)
  end

  it 'cell a10 should equal 512.0' do
    sheet1.a10.should be_within(51.2).of(512.0)
  end

end

END

it 'should create a ruby file from xml, dealing correctly with shared formulas' do
  worksheet = Worksheet.new(Nokogiri::XML(shared_formula_worksheet_xml))
  worksheet.name = 'sheet1'
  worksheet.workbook = mock(:workbook,:indirects_used => true)  
  worksheet.to_ruby.should == shared_formula_worksheet_ruby
  worksheet.to_test.should == shared_formula_worksheet_test
end

array_formula_worksheet_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" mc:Ignorable="mv" mc:PreserveAttributes="mv:*"><dimension ref="A1:B5"/><sheetViews><sheetView tabSelected="1" view="pageLayout" workbookViewId="0"><selection activeCell="B6" sqref="B6"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="13"/><sheetData><row r="1" spans="1:2"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="B1:B5">A1:A5</f><v>1</v></c></row><row r="2" spans="1:2"><c r="A2"><v>2</v></c><c r="B2"><v>2</v></c></row><row r="3" spans="1:2"><c r="A3"><v>3</v></c><c r="B3"><v>3</v></c></row><row r="4" spans="1:2"><c r="A4"><v>4</v></c><c r="B4"><v>4</v></c></row><row r="5" spans="1:2"><c r="A5"><v>5</v></c><c r="B5"><v>5</v></c></row></sheetData><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><pageSetup paperSize="10" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="1" OnePage="0" WScale="0"/></ext></extLst></worksheet>
END

array_formula_worksheet_ruby =<<END
# coding: utf-8
# Outputs
class Sheet1 < Spreadsheet
  def a1; 1.0; end
  def b1_array; @b1_array ||= a('a1','a5'); end
  def b1; @b1 ||= b1_array.array_formula_offset(0,0); end
  def a2; 2.0; end
  def b2; @b2 ||= b1_array.array_formula_offset(1,0); end
  def a3; 3.0; end
  def b3; @b3 ||= b1_array.array_formula_offset(2,0); end
  def a4; 4.0; end
  def b4; @b4 ||= b1_array.array_formula_offset(3,0); end
  def a5; 5.0; end
  def b5; @b5 ||= b1_array.array_formula_offset(4,0); end
end

END

array_formula_worksheet_test =<<END
# coding: utf-8
require_relative '../spreadsheet'
# Outputs
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell b1 should equal 1.0' do
    sheet1.b1.should be_within(0.1).of(1.0)
  end

  it 'cell b2 should equal 2.0' do
    sheet1.b2.should be_within(0.2).of(2.0)
  end

  it 'cell b3 should equal 3.0' do
    sheet1.b3.should be_within(0.30000000000000004).of(3.0)
  end

  it 'cell b4 should equal 4.0' do
    sheet1.b4.should be_within(0.4).of(4.0)
  end

  it 'cell b5 should equal 5.0' do
    sheet1.b5.should be_within(0.5).of(5.0)
  end

end

END

it 'should create a ruby file from xml, dealing correctly with array formulas' do
  worksheet = Worksheet.new(Nokogiri::XML(array_formula_worksheet_xml))
  worksheet.name = 'sheet1'
  worksheet.workbook = mock(:workbook,:indirects_used => true)
  worksheet.to_ruby.should == array_formula_worksheet_ruby
  worksheet.to_test.should == array_formula_worksheet_test
end

table_xml =<<-END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="FirstTable" displayName="FirstTable" ref="B3:D7" totalsRowCount="1">
  <autoFilter ref="B3:D6"/>
  <tableColumns count="3">
    <tableColumn id="1" name="ColA" totalsRowLabel="Total"/>
    <tableColumn id="2" name="ColB" totalsRowFunction="custom" dataDxfId="1">
      <calculatedColumnFormula>FirstTable[[#This Row],[ColA]]+FirstTable[[#This Row],[ColC]]</calculatedColumnFormula>
      <totalsRowFormula>AVERAGE(FirstTable[ColB])</totalsRowFormula>
    </tableColumn>
    <tableColumn id="3" name="ColC" totalsRowFunction="count" dataDxfId="0">
      <calculatedColumnFormula>3*FirstTable[[#This Row],[ColA]]</calculatedColumnFormula>
    </tableColumn>
  </tableColumns>
  <tableStyleInfo name="TableStyleLight16" showFirstColumn="1" showLastColumn="1" showRowStripes="1" showColumnStripes="1"/>
</table>
END

table_worksheet_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
  <dimension ref="B3:H10"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="H4" sqref="H4"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
  <sheetData>
    <row r="3" spans="2:8" x14ac:dyDescent="0.25">
      <c r="B3" t="s">
        <v>0</v>
      </c>
      <c r="C3" t="s">
        <v>1</v>
      </c>
      <c r="D3" t="s">
        <v>2</v>
      </c>
      <c r="F3" t="str">
        <f>FirstTable[[#Headers],[ColB]]</f>
        <v>ColB</v>
      </c>
      <c r="H3">
        <f>+MATCH("ColC",FirstTable[#Headers])</f>
        <v>3</v>
      </c>
    </row>
    <row r="4" spans="2:8" x14ac:dyDescent="0.25">
      <c r="B4">
        <v>1</v>
      </c>
      <c r="C4">
        <f>FirstTable[[#This Row],[ColA]]+FirstTable[[#This Row],[ColC]]</f>
        <v>4</v>
      </c>
      <c r="D4">
        <f>3*FirstTable[[#This Row],[ColA]]</f>
        <v>3</v>
      </c>
      <c r="F4">
        <f>SUM(FirstTable[#This Row])</f>
        <v>8</v>
      </c>
      <c r="H4">
        <f>MATCH(8,FirstTable[[#All],[ColB]])</f>
        <v>3</v>
      </c>
    </row>
    <row r="5" spans="2:8" x14ac:dyDescent="0.25">
      <c r="B5">
        <v>2</v>
      </c>
      <c r="C5">
        <f>FirstTable[[#This Row],[ColA]]+FirstTable[[#This Row],[ColC]]</f>
        <v>8</v>
      </c>
      <c r="D5">
        <f>3*FirstTable[[#This Row],[ColA]]</f>
        <v>6</v>
      </c>
      <c r="F5">
        <f>FirstTable[[#This Row],[ColB]]</f>
        <v>8</v>
      </c>
    </row>
    <row r="6" spans="2:8" x14ac:dyDescent="0.25">
      <c r="B6">
        <v>3</v>
      </c>
      <c r="C6">
        <f>FirstTable[[#This Row],[ColA]]+FirstTable[[#This Row],[ColC]]</f>
        <v>12</v>
      </c>
      <c r="D6">
        <f>3*FirstTable[[#This Row],[ColA]]</f>
        <v>9</v>
      </c>
    </row>
    <row r="7" spans="2:8" x14ac:dyDescent="0.25">
      <c r="B7" t="s">
        <v>3</v>
      </c>
      <c r="C7">
        <f>AVERAGE(FirstTable[ColB])</f>
        <v>8</v>
      </c>
      <c r="D7">
        <f>SUBTOTAL(103,FirstTable[ColC])</f>
        <v>3</v>
      </c>
    </row>
    <row r="10" spans="2:8" x14ac:dyDescent="0.25">
      <c r="B10">
        <f>SUM(FirstTable[ColA])</f>
        <v>6</v>
      </c>
      <c r="C10">
        <f>FirstTable[[#Totals],[ColB]]</f>
        <v>8</v>
      </c>
      <c r="D10">
        <f>FirstTable[[#Totals],[ColC]]</f>
        <v>3</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <tableParts count="1">
    <tablePart r:id="rId1"/>
  </tableParts>
</worksheet>
END

table_worksheet_shared_strings =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>ColA</t></si><si><t>ColB</t></si><si><t>ColC</t></si><si><t>Total</t></si></sst>
END

table_worksheet_ruby =<<END
# coding: utf-8
# Outputs
class Sheet1 < Spreadsheet
  def b3; "ColA"; end
  def c3; "ColB"; end
  def d3; "ColC"; end
  def f3; @f3 ||= sheet1.c3; end
  def h3; @h3 ||= +3.0; end
  def b4; 1.0; end
  def c4; @c4 ||= sheet1.b4+sheet1.d4; end
  def d4; @d4 ||= 3.0*sheet1.b4; end
  def f4; @f4 ||= sum(sheet1.a('b4','d4')); end
  def h4; @h4 ||= match(8.0,sheet1.a('c3','c7')); end
  def b5; 2.0; end
  def c5; @c5 ||= sheet1.b5+sheet1.d5; end
  def d5; @d5 ||= 3.0*sheet1.b5; end
  def f5; @f5 ||= sheet1.c5; end
  def b6; 3.0; end
  def c6; @c6 ||= sheet1.b6+sheet1.d6; end
  def d6; @d6 ||= 3.0*sheet1.b6; end
  def b7; "Total"; end
  def c7; @c7 ||= average(sheet1.a('c4','c6')); end
  def d7; @d7 ||= subtotal(103.0,sheet1.a('d4','d6')); end
  def b10; @b10 ||= sum(sheet1.a('b4','b6')); end
  def c10; @c10 ||= sheet1.c7; end
  def d10; @d10 ||= sheet1.d7; end
end

END

table_worksheet_test =<<END
# coding: utf-8
require_relative '../spreadsheet'
# Outputs
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell f3 should equal "ColB"' do
    sheet1.f3.should == "ColB"
  end

  it 'cell h3 should equal 3.0' do
    sheet1.h3.should be_within(0.30000000000000004).of(3.0)
  end

  it 'cell c4 should equal 4.0' do
    sheet1.c4.should be_within(0.4).of(4.0)
  end

  it 'cell d4 should equal 3.0' do
    sheet1.d4.should be_within(0.30000000000000004).of(3.0)
  end

  it 'cell f4 should equal 8.0' do
    sheet1.f4.should be_within(0.8).of(8.0)
  end

  it 'cell h4 should equal 3.0' do
    sheet1.h4.should be_within(0.30000000000000004).of(3.0)
  end

  it 'cell c5 should equal 8.0' do
    sheet1.c5.should be_within(0.8).of(8.0)
  end

  it 'cell d5 should equal 6.0' do
    sheet1.d5.should be_within(0.6000000000000001).of(6.0)
  end

  it 'cell f5 should equal 8.0' do
    sheet1.f5.should be_within(0.8).of(8.0)
  end

  it 'cell c6 should equal 12.0' do
    sheet1.c6.should be_within(1.2000000000000002).of(12.0)
  end

  it 'cell d6 should equal 9.0' do
    sheet1.d6.should be_within(0.9).of(9.0)
  end

  it 'cell c7 should equal 8.0' do
    sheet1.c7.should be_within(0.8).of(8.0)
  end

  it 'cell d7 should equal 3.0' do
    sheet1.d7.should be_within(0.30000000000000004).of(3.0)
  end

  it 'cell b10 should equal 6.0' do
    sheet1.b10.should be_within(0.6000000000000001).of(6.0)
  end

  it 'cell c10 should equal 8.0' do
    sheet1.c10.should be_within(0.8).of(8.0)
  end

  it 'cell d10 should equal 3.0' do
    sheet1.d10.should be_within(0.30000000000000004).of(3.0)
  end

end

END

it 'should create a ruby file from xml, dealing correctly with tables' do
  worksheet = Worksheet.new(Nokogiri::XML(table_worksheet_xml))
  Table.from_xml(worksheet,Nokogiri::XML(table_xml).root)
  worksheet.workbook = mock(:workbook,:indirects_used => true, :worksheets => { 'sheet1' => worksheet})
  SharedStrings.instance.clear
  SharedStrings.instance.load_strings_from_xml(Nokogiri::XML(table_worksheet_shared_strings))
  worksheet.name = 'sheet1'
  worksheet.to_ruby.should == table_worksheet_ruby
  worksheet.to_test.should == table_worksheet_test
end

table_worksheet_relationships =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>
END

it 'should have a Worksheet#from_file method that also loads the tables' do
  { '/usr/local/excel/xl/worksheets/sheet1.xml' => table_worksheet_xml, 
    '/usr/local/excel/xl/worksheets/_rels/sheet1.xml.rels' => table_worksheet_relationships,
    '/usr/local/excel/xl/tables/table1.xml' => table_xml
  }.each do |filename,xml| 
    File.should_receive(:open).with(filename).and_yield(StringIO.new(xml))
  end
  File.should_receive(:exist?).with('/usr/local/excel/xl/worksheets/_rels/sheet1.xml.rels').and_return(true)
  
  Table.tables.clear
  worksheet = Worksheet.from_file('/usr/local/excel/xl/worksheets/sheet1.xml')
  SharedStrings.instance.clear
  SharedStrings.instance.load_strings_from_xml(Nokogiri::XML(table_worksheet_shared_strings))
  worksheet.workbook = mock(:workbook,:indirects_used => true, :worksheets => { 'sheet1' => worksheet})  
  worksheet.name = 'sheet1'
  worksheet.to_ruby.should == table_worksheet_ruby
end
end