require_relative 'spec_helper'

describe Workbook do

workbook_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9114"/>
  <workbookPr defaultThemeVersion="124226"/>
  <bookViews>
    <workbookView xWindow="240" yWindow="90" windowWidth="9435" windowHeight="2070"/>
  </bookViews>
  <sheets>
    <sheet name="First sheet" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
    <sheet name="Sheet3" sheetId="3" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="OneAnd2">Sheet2!$A$2:$A$3</definedName>
    <definedName name="Reference.2">Sheet2!$A$3</definedName>
    <definedName name="ReferenceOne" localSheetId="1">'First sheet'!$A$1</definedName>
    <definedName name="ReferenceOne">Sheet2!$A$2</definedName>
  </definedNames>
  <calcPr calcId="144315"/>
</workbook>
END

workbook_relationship_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
END

simple_worksheet_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" mc:Ignorable="mv" mc:PreserveAttributes="mv:*">
  <dimension ref="A1:A2"/>
  <sheetViews>
    <sheetView view="pageLayout" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr baseColWidth="10" defaultRowHeight="13"/>
  <sheetData>
    <row r="1" spans="1:1">
      <c r="A1" t="str">
        <f>IF(A2="Hello","hello",Sheet2!B4)</f>
        <v>hello</v>
      </c>
    </row>
    <row r="2" spans="1:1">
      <c r="A2" t="s">
        <v>0</v>
      </c>
    </row>
  </sheetData>
  <sheetCalcPr fullCalcOnLoad="1"/>
  <phoneticPr fontId="2" type="noConversion"/>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
  <pageSetup paperSize="10" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/>
  <extLst>
    <ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main">
      <mx:PLV Mode="1" OnePage="0" WScale="0"/>
    </ext>
  </extLst>
</worksheet>
END

simple_worksheet_ruby =<<END
# coding: utf-8
# First sheet
class Sheet1 < Spreadsheet
  def a1; @a1 ||= excel_if(excel_comparison(a2,"==","Hello"),"hello",sheet2.b4); end
  def a2; "Hello"; end
end

END

shared_strings_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <t>Hello</t>
    <phoneticPr fontId="3" type="noConversion"/>
  </si>
</sst>
END

workbook_ruby =<<END
# coding: utf-8
require 'rubyfromexcel'

class Spreadsheet
  include RubyFromExcel::ExcelFunctions

  def initialize
    @worksheet_names = {"First sheet"=>"sheet1", "Sheet2"=>"sheet2", "Sheet3"=>"sheet3"}
    @workbook_tables = {}
  end

  def oneand2; sheet2.a('a2','a3'); end
  def reference_2; sheet2.a3; end
  def referenceone; sheet2.a2; end
end

Dir[File.join(File.dirname(__FILE__),"sheets/","sheet*.rb")].each {|f| Spreadsheet.autoload(File.basename(f,".rb").capitalize,f)}
END

workbook_no_indirects_ruby =<<END
# coding: utf-8
require 'rubyfromexcel'

class Spreadsheet
  include RubyFromExcel::ExcelFunctions

end

Dir[File.join(File.dirname(__FILE__),"sheets/","sheet*.rb")].each {|f| Spreadsheet.autoload(File.basename(f,".rb").capitalize,f)}
END


it "should load its relationships, and therefore create the requisite worksheets and shared strings" do
  SharedStrings.instance.clear
  Table.tables.clear
  SheetNames.instance.clear
  {
    '/usr/local/excel/xl/workbook.xml' => workbook_xml,
    '/usr/local/excel/xl/_rels/workbook.xml.rels' => workbook_relationship_xml,
    '/usr/local/excel/xl/worksheets/sheet1.xml' => simple_worksheet_xml,
    '/usr/local/excel/xl/worksheets/sheet2.xml' => simple_worksheet_xml,
    '/usr/local/excel/xl/worksheets/sheet3.xml' => simple_worksheet_xml,
    '/usr/local/excel/xl/sharedStrings.xml' => shared_strings_xml,
    '/usr/local/excel/xl/worksheets/_rels/sheet1.xml.rels' => '',
    '/usr/local/excel/xl/worksheets/_rels/sheet2.xml.rels' => '',
    '/usr/local/excel/xl/worksheets/_rels/sheet3.xml.rels' => ''
  }.each do |filename,xml|
    File.should_receive(:open).with(filename).and_yield(StringIO.new(xml))
  end
  File.stub!(:exist?).and_return(true)
  workbook = Workbook.new('/usr/local/excel/xl/workbook.xml')
  workbook.worksheets['sheet1'].to_ruby.should == simple_worksheet_ruby
  SheetNames.instance['First sheet'].should == 'sheet1'
  workbook.to_ruby.should == workbook_no_indirects_ruby
  workbook.indirects_used = true
  workbook.to_ruby.should == workbook_ruby
end
  
end

describe Worksheet, "in pruning mode" do
pruning_workbook_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr date1904="1" showInkAnnotation="0" autoCompressPictures="0"/><bookViews><workbookView xWindow="-20" yWindow="-20" windowWidth="20760" windowHeight="14600" tabRatio="500"/></bookViews><sheets><sheet name="Outputs" sheetId="1" r:id="rId1"/><sheet name="Calcs" sheetId="2" r:id="rId2"/><sheet name="Inputs" sheetId="3" r:id="rId3"/></sheets><definedNames><definedName name="In_result">Inputs!$A$3</definedName></definedNames><calcPr calcId="130407" concurrentCalc="0"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:ArchID Flags="2"/></ext></extLst></workbook>
END

pruning_workbook_relationships_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>
END

pruning_shared_strings_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="6"><si><t>Result</t><phoneticPr fontId="1" type="noConversion"/></si><si><t>Input</t><phoneticPr fontId="1" type="noConversion"/></si><si><t>Not in result</t><phoneticPr fontId="1" type="noConversion"/></si><si><t>In result</t><phoneticPr fontId="1" type="noConversion"/></si><si><t>Inputs</t><phoneticPr fontId="1" type="noConversion"/></si><si><t>Doesn't depend on an input</t><phoneticPr fontId="1" type="noConversion"/></si></sst>
END

pruning_output_sheet_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" mc:Ignorable="mv" mc:PreserveAttributes="mv:*"><dimension ref="A1:B4"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B5" sqref="B5"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="13"/><sheetData><row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>Calcs!A1</f><v>121</v></c></row><row r="2" spans="1:2"><c r="A2" t="s"><v>1</v></c><c r="B2"><f>Inputs!A1</f><v>99</v></c></row><row r="3" spans="1:2"><c r="B3" t="str"><f ca="1">Calcs!C9</f><v>In result</v></c></row><row r="4" spans="1:2"><c r="B4" t="str"><f>Calcs!C13</f><v>Doesn't depend on an input</v></c></row></sheetData><sheetCalcPr fullCalcOnLoad="1"/><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="0" OnePage="0" WScale="0"/></ext></extLst></worksheet>
END

pruning_calc_sheet_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" mc:Ignorable="mv" mc:PreserveAttributes="mv:*"><dimension ref="A1:D13"/><sheetViews><sheetView view="pageLayout" workbookViewId="0"><selection activeCell="C14" sqref="C14"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="13"/><sheetData><row r="1" spans="1:4"><c r="A1"><f>SUM(A2:A7)</f><v>121</v></c></row><row r="2" spans="1:4"><c r="A2"><v>1</v></c><c r="C2" t="s"><v>2</v></c></row><row r="3" spans="1:4"><c r="A3"><v>2</v></c><c r="C3"><f>Inputs!A1*10</f><v>990</v></c></row><row r="4" spans="1:4"><c r="A4"><f>Inputs!A1</f><v>99</v></c></row><row r="5" spans="1:4"><c r="A5"><v>4</v></c></row><row r="6" spans="1:4"><c r="A6"><f>C6</f><v>10</v></c><c r="C6"><v>10</v></c></row><row r="7" spans="1:4"><c r="A7"><v>5</v></c></row><row r="8" spans="1:4"><c r="C8" t="s"><v>4</v></c></row><row r="9" spans="1:4"><c r="C9" t="str"><f ca="1">INDIRECT("'"&amp;C8&amp;"'!In_result")</f><v>In result</v></c></row><row r="10" spans="1:4"><c r="C10" t="str"><f ca="1">INDIRECT("'"&amp;C8&amp;"'!A2")</f><v>Not in result</v></c></row><row r="13" spans="1:4"><c r="C13" t="str"><f>D13</f><v>Doesn't depend on an input</v></c><c r="D13" t="s"><v>5</v></c></row></sheetData><sheetCalcPr fullCalcOnLoad="1"/><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><pageSetup paperSize="10" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="1" OnePage="0" WScale="0"/></ext></extLst></worksheet>
END

pruning_input_sheet_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" mc:Ignorable="mv" mc:PreserveAttributes="mv:*"><dimension ref="A1:A3"/><sheetViews><sheetView view="pageLayout" workbookViewId="0"><selection activeCell="A3" sqref="A3"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="13"/><sheetData><row r="1" spans="1:1"><c r="A1"><v>99</v></c></row><row r="2" spans="1:1"><c r="A2" t="s"><v>2</v></c></row><row r="3" spans="1:1"><c r="A3" t="s"><v>3</v></c></row></sheetData><sheetCalcPr fullCalcOnLoad="1"/><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><pageSetup paperSize="10" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/><extLst><ext xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" uri="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="1" OnePage="0" WScale="0"/></ext></extLst></worksheet>
END

pruning_calc_sheet_ruby_no_prune =<<END
# coding: utf-8
# Calcs
class Sheet2 < Spreadsheet
  def a1; @a1 ||= sum(a('a2','a7')); end
  def a2; 1.0; end
  def c2; "Not in result"; end
  def a3; 2.0; end
  def c3; @c3 ||= sheet3.a1*10.0; end
  def a4; @a4 ||= sheet3.a1; end
  def a5; 4.0; end
  def a6; @a6 ||= c6; end
  def c6; 10.0; end
  def a7; 5.0; end
  def c8; "Inputs"; end
  def c9; @c9 ||= sheet3.a3; end
  def c10; @c10 ||= sheet3.a2; end
  def c13; @c13 ||= d13; end
  def d13; "Doesn't depend on an input"; end
end

END

pruning_calc_sheet_ruby_output_prune =<<END
# coding: utf-8
# Calcs
class Sheet2 < Spreadsheet
  def a1; @a1 ||= sum(a('a2','a7')); end
  def a2; 1.0; end
  def a3; 2.0; end
  def a4; @a4 ||= sheet3.a1; end
  def a5; 4.0; end
  def a6; @a6 ||= c6; end
  def c6; 10.0; end
  def a7; 5.0; end
  def c8; "Inputs"; end
  def c9; @c9 ||= sheet3.a3; end
  def c13; @c13 ||= d13; end
  def d13; "Doesn't depend on an input"; end
end

END

pruning_calc_sheet_ruby_input_and_output_prune =<<END
# coding: utf-8
# Calcs
class Sheet2 < Spreadsheet
  def a1; @a1 ||= sum(a('a2','a7')); end
  def a2; 1.0; end
  def a3; 2.0; end
  def a4; @a4 ||= sheet3.a1; end
  def a5; 4.0; end
  def a6; 10.0; end
  def a7; 5.0; end
  def c8; "Inputs"; end
  def c9; @c9 ||= sheet3.a3; end
end

END

before do
  SharedStrings.instance.clear
  Table.tables.clear
  SheetNames.instance.clear
  {
    '/usr/local/excel/xl/workbook.xml' => pruning_workbook_xml,
    '/usr/local/excel/xl/_rels/workbook.xml.rels' => pruning_workbook_relationships_xml,
    '/usr/local/excel/xl/worksheets/sheet1.xml' => pruning_output_sheet_xml,
    '/usr/local/excel/xl/worksheets/sheet2.xml' => pruning_calc_sheet_xml,
    '/usr/local/excel/xl/worksheets/sheet3.xml' => pruning_input_sheet_xml,
    '/usr/local/excel/xl/sharedStrings.xml' => pruning_shared_strings_xml,
    '/usr/local/excel/xl/worksheets/_rels/sheet1.xml.rels' => '',
    '/usr/local/excel/xl/worksheets/_rels/sheet2.xml.rels' => '',
    '/usr/local/excel/xl/worksheets/_rels/sheet3.xml.rels' => ''
  }.each do |filename,xml|
    File.should_receive(:open).with(filename).and_yield(StringIO.new(xml))
  end
  File.stub!(:exist?).and_return(true)
  @workbook = Workbook.new('/usr/local/excel/xl/workbook.xml')
end

it "should work normally when no output sheets are identified" do
  @workbook.total_cells.should == 24
  @workbook.worksheets['sheet2'].to_ruby.should == pruning_calc_sheet_ruby_no_prune
end

it "should prune any cells that aren't needed for the output sheet calculations when output sheets have been specified" do
  @workbook.work_out_dependencies
  @workbook.prune_cells_not_needed_for_output_sheets('Outputs')
  @workbook.worksheets['sheet2'].to_ruby.should == pruning_calc_sheet_ruby_output_prune
end

it "should convert cells to values where they don't depend on inputs, and then prune" do
  @workbook.work_out_dependencies
  @workbook.convert_cells_to_values_when_independent_of_input_sheets('Inputs')
  @workbook.prune_cells_not_needed_for_output_sheets('Outputs')
  @workbook.worksheets['sheet2'].to_ruby.should == pruning_calc_sheet_ruby_input_and_output_prune
  $DEBUG = false
end

end