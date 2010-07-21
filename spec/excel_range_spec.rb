require_relative 'spec_helper'

class TestSheet
  include RubyFromExcel::ExcelFunctions
  
  {
    :a => ['a1',  'a2', 'a3', nil , 'a5'],
    :b => ['b1',  nil,  nil,  'b4', 'b5'],
    :c => ['c1',  'c2', 'c3', 'c4', 'c5']
  }.each do |column,row_values|
    row_values.each_with_index do |value,row|
      next unless value
      class_eval "def #{column}#{row+1}; '#{value}'; end"
    end
  end

  def to_s
    'sheet1'
  end

end

describe Area do
  it "should return the values in an area" do
    Area.new(TestSheet.new,'a1','a3').to_a.should == ['a1','a2','a3']
    Area.new(TestSheet.new,'a1','b1').to_a.should == ['a1','b1']
  end
  
  it "should return references in its area" do
    Area.new(TestSheet.new,'a1','a3').to_reference_enum.to_a.should == ['a1','a2','a3']
    Area.new(TestSheet.new,'r9','r11').to_reference_enum.to_a.should == ['r9','r10','r11']
    Area.new(TestSheet.new,'zx10','zz10').to_reference_enum.to_a.should == ['zx10','zy10','zz10']
  end
  
  it "should respond to array_formula_index(row_index,column_index) by returning the reference at those coordinates" do
    Area.new(TestSheet.new,'a1','c5').array_formula_index(0,0).to_ruby.should == 'a1'
    Area.new(TestSheet.new,'a1','c5').array_formula_index(1,1).to_ruby.should == 'b2'
    Area.new(TestSheet.new,'a1','c5').array_formula_index(10,10).to_ruby.should == :ref
  end
  
  it "should respond to array_formula_offset(row_index,column_index) by returning the value at those coordinates (origin is 0,0)" do
    Area.new(TestSheet.new,'a1','c5').array_formula_offset(0,0).should == 'a1'
    Area.new(TestSheet.new,'a1','c5').array_formula_offset(1,1).should == 0.0
    Area.new(TestSheet.new,'a1','c5').array_formula_offset(10,10).should == :ref    
  end
  
  it "if the area is a single column, then it should ignore the column index when responding to array_formula_index(row_index,column_index)" do
      Area.new(TestSheet.new,'a1','a5').array_formula_index(0,1).to_ruby.should == 'a1'
      Area.new(TestSheet.new,'a1','a5').array_formula_index(1,1).to_ruby.should == 'a2'
      Area.new(TestSheet.new,'a1','a5').array_formula_index(6,1).to_ruby.should == :ref
  end

  it "if the area is a single row, then it should ignore the row index when responding to array_formula_index(row_index,column_index)" do
      Area.new(TestSheet.new,'a1','c1').array_formula_index(1,0).to_ruby.should == 'a1'
      Area.new(TestSheet.new,'a1','c1').array_formula_index(1,1).to_ruby.should == 'b1'
      Area.new(TestSheet.new,'a1','c1').array_formula_index(6,4).to_ruby.should == :ref
  end  
  
  it "#row(index) returns a subset of the area for just that row (index=0 is the top row, index=-1 is the last row)" do
    Area.new(TestSheet.new,'a1','c5').row(0).start_cell.to_s.should == 'a1'
    Area.new(TestSheet.new,'a1','c5').row(0).end_cell.to_s.should == 'c1'
    Area.new(TestSheet.new,'a1','c5').row(-1).start_cell.to_s.should == 'a5'
    Area.new(TestSheet.new,'a1','c5').row(-1).end_cell.to_s.should == 'c5'
  end
  
  it "#rows(start_index,end_index) returns a subset of the area for those rows (index=0 is the top row, index=-1 is the last row)" do
    Area.new(TestSheet.new,'a1','c5').rows(0,-1).start_cell.to_s.should == 'a1'
    Area.new(TestSheet.new,'a1','c5').rows(0,-1).end_cell.to_s.should == 'c5'
    Area.new(TestSheet.new,'a1','c5').rows(1,-2).start_cell.to_s.should == 'a2'
    Area.new(TestSheet.new,'a1','c5').rows(1,-2).end_cell.to_s.should == 'c4'
  end
  
  it "#column(index) returns a subset of the area for just that column (index=0 is the first column)" do
    Area.new(TestSheet.new,'a1','c5').column(0).start_cell.to_s.should == 'a1'
    Area.new(TestSheet.new,'a1','c5').column(0).end_cell.to_s.should == 'a5'
    Area.new(TestSheet.new,'a1','c5').column(2).start_cell.to_s.should == 'c1'
    Area.new(TestSheet.new,'a1','c5').column(2).end_cell.to_s.should == 'c5'
  end
  
  it "#to_s returns the area in the form of the excel function used to create it" do
    Area.new(TestSheet.new,'a1','c5').to_s.should == "sheet1.a('a1','c5')"
  end
  
  it "#to_excel_matrix returns the area in the form of a matrix, with methods to allow excel array formula manipulations" do
    em = Area.new(TestSheet.new,'a1','c5').to_excel_matrix
    em.rows.should == 5
    em.columns.should == 3
    em.values.should == [['a1','b1','c1'],['a2',0.0,'c2'],['a3',0.0,'c3'],[0.0,'b4','c4'],['a5','b5','c5']]
  end
  
end

describe Area, "when area refers to just a single cell" do
  
  it "should duck type as if it were just that cell for use in arithmetic, string joins etc" do
     (Area.new(TestSheet.new,'a1','a1') + "-cell").should == "a1-cell"
     (1 + Area.new(TestSheet.new,'b3','b3') + 10).should == 11
  end
  
end

describe Columns do
  it "should return all the defined values in a single column" do
    Columns.new(TestSheet.new,'c','c').each.to_a.should ==  ['c1',  'c2', 'c3', 'c4', 'c5']
  end
  
  it "should cope with undefined values where they are in the middle of the column" do
    Columns.new(TestSheet.new,'a','a').each.to_a.should == ['a1','a2','a3',0.0,'a5']
  end
  
  it "should return all the defined values in several columns as a single array" do
    Columns.new(TestSheet.new,'a','c').each.to_a.should == ['a1','a2','a3',0.0,'a5','b1',0.0,0.0,'b4','b5','c1',  'c2', 'c3', 'c4', 'c5']
  end
  
  it "should cope with references to columns that aren't defined at all" do
    Columns.new(TestSheet.new,'aa','zz').each.to_a.should == []
  end

  it "should respond to array_formula_index(row_index,column_index) by returning the reference for those coordinates" do
    Columns.new(TestSheet.new,'a','c').array_formula_index(0,0).should == 'a1'
    Columns.new(TestSheet.new,'a','c').array_formula_index(1,1).should == 'b2'
    Columns.new(TestSheet.new,'a','c').array_formula_index(10,10).should == :na
  end
  
  it "if the area is a single column, then it should ignore the column index when responding to array_formula_index(row_index,column_index)" do
    Columns.new(TestSheet.new,'c','c').array_formula_index(0,1).should == 'c1'
    Columns.new(TestSheet.new,'c','c').array_formula_index(1,1).should == 'c2'
    Columns.new(TestSheet.new,'c','c').array_formula_index(6,6).should == 'c7'
  end    
    
end

describe Rows do 
  it "should return all the defined values in a single row" do
    Rows.new(TestSheet.new,1,1).each.to_a.should ==  ['a1','b1','c1']
  end
  
  it "should cope with undefined values where they are in the middle of the row" do
    Rows.new(TestSheet.new,2,2).each.to_a.should == ['a2',0.0,'c2']
  end
  
  it "should return all the defined values in several columns as a single array" do
    Rows.new(TestSheet.new,1,4).each.to_a.should == ['a1','b1','c1','a2',0.0,'c2','a3',0.0,'c3',0.0,'b4','c4']
  end
   
  it "should cope with references to rows that aren't defined at all" do
    Rows.new(TestSheet.new,100,105).each.to_a.should == []
  end

  it "should respond to array_formula_index(row_index,column_index) by returning the reference at those coordinates" do
    Rows.new(TestSheet.new,1,4).array_formula_index(0,0).should == 'a1'
    Rows.new(TestSheet.new,1,4).array_formula_index(1,1).should == 'b2'
    Rows.new(TestSheet.new,1,4).array_formula_index(10,10).should == :na
  end
  
  it "if the area is a single row, then it should ignore the row index when responding to array_formula_index(row_index,column_index)" do
    Rows.new(TestSheet.new,1,1).array_formula_index(1,0).should == 'a1'
    Rows.new(TestSheet.new,1,1).array_formula_index(1,1).should == 'b1'
    Rows.new(TestSheet.new,1,1).array_formula_index(6,6).should == 'g1'
  end
end