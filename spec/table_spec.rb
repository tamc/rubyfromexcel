require_relative 'spec_helper'

describe Table do
  
  def table
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
    @worksheet ||= mock(:worksheet,:to_s => 'sheet1')
    @table ||= Table.from_xml(@worksheet, Nokogiri::XML(table_xml).root)
  end
  
  it "should know its name" do
    table.name.should == 'FirstTable'
  end
  
  it "should know the area (returned in #all) that it covers" do
    table.all.start_cell.to_s.should == 'b3'
    table.all.end_cell.to_s.should == 'd7'
  end
  
  it "should know the area (returned in #data) covered by its data" do
    table.data.start_cell.to_s.should == 'b4'
    table.data.end_cell.to_s.should == 'd6'
  end  
  
  it "should know the area covered by its headers" do
    table.headers.start_cell.to_s.should == 'b3'
    table.headers.end_cell.to_s.should == 'd3'
  end
  
  it "should know the area covered by its totals" do
    table.totals.start_cell.to_s.should == 'b7'
    table.totals.end_cell.to_s.should == 'd7'
  end
  
  it "should know the area covered by each of its columns (excluding the headers and the totals)" do
    table.column('ColA').start_cell.to_s.should == 'b4'
    table.column('ColA').end_cell.to_s.should == 'b6'
    table.column('ColC').start_cell.to_s.should == 'd4'
    table.column('ColC').end_cell.to_s.should == 'd6'    
  end
  
  it "should be able to translate structured references into normal references" do
    table.reference_for('ColA').start_cell.to_s.should == 'b4'
    table.reference_for('ColA').end_cell.to_s.should == 'b6'
    table.reference_for('#All').start_cell.to_s.should == 'b3'
    table.reference_for('#All').end_cell.to_s.should == 'd7'
    table.reference_for('#Data').start_cell.to_s.should == 'b4'
    table.reference_for('#Data').end_cell.to_s.should == 'd6'
    table.reference_for('#Headers').start_cell.to_s.should == 'b3'
    table.reference_for('#Headers').end_cell.to_s.should == 'd3'
    table.reference_for('#Totals').start_cell.to_s.should == 'b7'
    table.reference_for('#Totals').end_cell.to_s.should == 'd7'    
  end
  
  it "should be able to translate structured references that are intersections into normal references" do
    table.reference_for('[#Totals],[ColA]').to_s.should == 'sheet1.b7'
    table.reference_for('[#Headers],[ColC]').to_s.should == 'sheet1.d3'
    table.reference_for('[#All],[ColB]').start_cell.to_s.should == 'c3'
    table.reference_for('[#All],[ColB]').end_cell.to_s.should == 'c7'
  end
  
  it "should be able to translate structured references that include the special variable #This Row" do
    table.reference_for('[#This Row],[ColA]', Reference.new('e5')).to_s.should == 'sheet1.b5'
    table.reference_for('#This Row',Reference.new('e5')).start_cell.to_s.should == 'b5'
    table.reference_for('#This Row',Reference.new('e5')).end_cell.to_s.should == 'd5'
  end
  
  it "should be able to translate structured references from cells that are in the table, treating them as if they had indicated [#This row] as part of the reference" do
    table.reference_for('ColA', Reference.new('e5',@worksheet)).to_s.should == "sheet1.a('b4','b6')"
    table.reference_for('ColA', Reference.new('c5',@worksheet)).to_s.should == 'sheet1.b5'
  end
  
  it "should be able to access structured references by providing the table name to Table.reference_for(table_name,structured_reference,referring_cell)" do
    table
    Table.reference_for('FirstTable','#Data').start_cell.to_s.should == 'b4'
    Table.reference_for('Not a known table','#Data').should == ":ref"
    Table.reference_for('FirstTable','[#This Row],[ColA]', Reference.new('e5')).to_s.should == 'sheet1.b5'
  end
  
  it "should be able to deal with a column range, returning a range" do
    table
    Table.reference_for('FirstTable','[#This Row],[ColA]:[ColC]', Reference.new('e5')).to_s.should == "sheet1.a('b5','d5')"
  end
  
  it "#inspect should create a string that can be used to recreate the table" do
    table.inspect.should == %Q{'Table.new(sheet1,"FirstTable","B3:D7",["ColA", "ColB", "ColC"],1)'}
  end
    
end