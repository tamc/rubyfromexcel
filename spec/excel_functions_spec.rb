require_relative 'spec_helper'

class FunctionTest
  extend ExcelFunctions
  def self.a1; 10.0; end
  def self.a2; 100.0; end
  def self.b1; 'Pear'; end
  def self.b2; 'Bear'; end
  def self.b3; 'Apple'; end
  def self.zx10; 2.0; end
  def self.zy10; 3.0; end
  def self.zz10; 4.0; end
end

describe "set" do
  it "should set the value of existing cells" do
    class ValueSetTest
      include ExcelFunctions
      def a1; 10.0; end
    end
    v = ValueSetTest.new
    v.a1.should == 10.0
    v.set('a1',99.0)
    v.a1.should == 99.0
    v.set('a1',nil)
    v.a1.should == 10.0
  end
end

describe "round" do
  it "should round numbers correctly" do
    FunctionTest.round(1.1,0).should == 1.0
    FunctionTest.round(1.5,0).should == 2.0
    FunctionTest.round(1.56,1).should == 1.6
  end
end

describe "roundup" do
  it "should round numbers up correctly" do
    FunctionTest.roundup(1.0,0).should == 1.0
    FunctionTest.roundup(1.1,0).should == 2.0
    FunctionTest.roundup(1.5,0).should == 2.0
    FunctionTest.roundup(1.53,1).should == 1.6
  end
end

describe "rounddown" do
  it "should round numbers up correctly" do
    FunctionTest.rounddown(1.0,0).should == 1.0
    FunctionTest.rounddown(1.1,0).should == 1.0
    FunctionTest.rounddown(1.5,0).should == 1.0
    FunctionTest.rounddown(1.53,1).should == 1.5
  end
end

describe "mod" do
  it "should return the remainder of a number" do
    FunctionTest.mod(10,3).should == 1.0
    FunctionTest.mod(10,5).should == 0.0
    FunctionTest.mod(1.1,1).should be_close(0.1,0.01)
  end
end

describe "sum" do
  it "should total areas correctly" do
    FunctionTest.sum(1,2,3).should == 6
    FunctionTest.sum(FunctionTest.a('a1','a2'),3).should == 113    
  end
end

describe "choose" do
  it "should return whichever argument is matched by the first argument" do
    FunctionTest.choose(1,1,2,3,4).should == 1
    FunctionTest.choose(2,1,2,3,4).should == 2
    FunctionTest.choose(3,1,2,3,4).should == 3
  end
end

describe "abs" do
  it "should return the absolute value of the input" do
    FunctionTest.abs(1.0).should == 1
    FunctionTest.abs(-1.0).should == 1.0
  end
end

describe "sumif" do
  it "should only sum values in the area that meet the criteria" do
    FunctionTest.sumif(FunctionTest.a('a1','a3'),10.0).should == 10.0
    FunctionTest.sumif(FunctionTest.a('b1','b3'),'Bear',FunctionTest.a('a1','a2')).should == 100.0
  end
  
  it "should understand >0 type criteria" do
    FunctionTest.sumif(FunctionTest.a('a1','a3'),">0").should == 110.0
    FunctionTest.sumif(FunctionTest.a('a1','a3'),">10").should == 100.0
    FunctionTest.sumif(FunctionTest.a('a1','a3'),"<100").should == 10.0
  end
  
end

describe "countif" do
  it "should only count values in the area that meet the criteria" do
    FunctionTest.countif(FunctionTest.a('a1','a3'),10.0).should == 1.0
    FunctionTest.sumif(FunctionTest.a('b1','b3'),'Bear',FunctionTest.a('a1','a2')).should == 100.0
  end
  
  it "should understand >0 type criteria" do
    FunctionTest.countif(FunctionTest.a('a1','a3'),">0").should == 2.0
    FunctionTest.countif(FunctionTest.a('a1','a3'),">10").should == 1.0
    FunctionTest.countif(FunctionTest.a('a1','a3'),"<100").should == 1.0
  end
  
end

describe "sumifs" do
  it "should only sum values that meet all of the criteria" do
    FunctionTest.sumifs(FunctionTest.a('a1','a3'),FunctionTest.a('a1','a3'),10.0,FunctionTest.a('b1','b3'),'Bear').should == 0.0
    FunctionTest.sumifs(FunctionTest.a('a1','a3'),FunctionTest.a('a1','a3'),10.0,FunctionTest.a('b1','b3'),'Pear').should == 10.0
  end
  
  it "should work when single cells are given where ranges expected" do
    FunctionTest.sumifs(0.143897265452564, "CAR", "CAR", "FCV", "FCV").should == 0.143897265452564
  end
end

describe "sumproduct" do
  it "should multiply together and then sum the elements in row or column areas given as arguments" do
    FunctionTest.sumproduct(FunctionTest.a('a1','a3'),FunctionTest.a('zx10','zz10')).should == 320.0
  end

  it "should return :value when miss-matched array sizes" do
    FunctionTest.sumproduct(FunctionTest.a('a1','a4'),FunctionTest.a('zx10','zz10')).should == :value
  end
  
end

describe "count" do
  it "should count the number of numeric values in an area" do
    FunctionTest.count(1,"two",FunctionTest.a('a1','a3')).should == 3    
  end
end

describe "counta" do
  it "should count the number of numeric or text values in an area" do
    FunctionTest.counta(1,"two",FunctionTest.a('a1','a3')).should == 4    
  end
end

describe "average" do
  it "should calculate the mean of its arguments" do
    FunctionTest.average(1,2,3).should == 2
    FunctionTest.average(1,"two",FunctionTest.a('a1','a3')).should == 111.0/3.0    
  end
end

describe "subtotal" do
  it "should calculate averages, counts, countas, sums depending on first argument" do
    FunctionTest.subtotal(1.0,1,"two",FunctionTest.a('a1','a3')).should == 111.0/3.0 # Average
    FunctionTest.subtotal(2.0,1,"two",FunctionTest.a('a1','a3')).should == 3 # count
    FunctionTest.subtotal(3.0,1,"two",FunctionTest.a('a1','a3')).should == 4 # counta
    FunctionTest.subtotal(9.0,1,"two",FunctionTest.a('a1','a3')).should == 111 # sum

    FunctionTest.subtotal(101.0,1,"two",FunctionTest.a('a1','a3')).should == 111.0/3.0 # Average
    FunctionTest.subtotal(102.0,1,"two",FunctionTest.a('a1','a3')).should == 3 # count
    FunctionTest.subtotal(103.0,1,"two",FunctionTest.a('a1','a3')).should == 4 # counta
    FunctionTest.subtotal(109.0,1,"two",FunctionTest.a('a1','a3')).should == 111 # sum    
  end
end

describe "match" do
  it "should return the index of the first match of the first argument in the area" do
    FunctionTest.match(10.0,FunctionTest.a('a1','a2'),0.0).should == 1
    FunctionTest.match(100.0,FunctionTest.a('a1','a2'),0.0).should == 2
    FunctionTest.match(1000.0,FunctionTest.a('a1','a2'),0.0).should == :na
    FunctionTest.match('bEAr',FunctionTest.a('b1','b3'),0.0).should == 2
    FunctionTest.match(1000.0,FunctionTest.a('a1','a2'),1.0).should == 2
    FunctionTest.match(1.0,FunctionTest.a('a1','a2'),1.0).should == :na
    FunctionTest.match('Care',FunctionTest.a('b1','b3'),-1.0).should == 1  
    FunctionTest.match('Zebra',FunctionTest.a('b1','b3'),-1.0).should == :na
    FunctionTest.match('a',FunctionTest.a('b1','b3'),-1.0).should == 2
    # FunctionTest.match("v.02",["p.01","V.01","v.05"],"false").should == 2
  end
end

describe "index" do
  it "should return the value at the row and column number given" do
    FunctionTest.index(FunctionTest.a('a1','b1'),2.0).should == "Pear"
    FunctionTest.index(FunctionTest.a('a1','a3'),2.0).should == 100.0
    FunctionTest.index(FunctionTest.a('b1','b3'),2.0).should == "Bear"
    FunctionTest.index(FunctionTest.a('a1','b3'),1.0,2.0).should == "Pear"
    FunctionTest.index(FunctionTest.a('a1','b3'),2.0,1.0).should == 100.0
    FunctionTest.index(FunctionTest.a('a1','b3'),3.0,1.0).should == 0.0
    FunctionTest.index(FunctionTest.a('a1','b3'),3.0,3.0).should == :ref
    FunctionTest.index(FunctionTest.a('a1','b3'),3.0).should == :ref
    FunctionTest.index(FunctionTest.a('a1','b3'),1.0,0.0).to_a.should == [10.0,"Pear"]
    FunctionTest.index(FunctionTest.a('a1','b3'),0.0,2.0).to_a.should == ["Pear","Bear","Apple"]    
  end
end

describe "max" do
  it "should return the argument with the greatest value" do
    FunctionTest.max(1,"two",FunctionTest.a('a1','a3')).should == 100    
  end
  it "should return an error if any of its inputs are errors" do
    FunctionTest.max(1,"two",FunctionTest.a('a1','a3'),:ref).should == :ref
  end
  
end

describe "min" do
  it "should return the argument with the smallest value" do
    FunctionTest.min(1000,"two",FunctionTest.a('a1','a3')).should == 10    
  end
end

describe "na" do
  it "should return an na error" do
    FunctionTest.na().should == :na    
  end
end

describe "iserr" do
  it "should return true if passed a symbol" do
    FunctionTest.iserr(:na).should == true
    FunctionTest.iserr(:ref).should == true
  end
  
  it "should return true if passed nan or infinity" do
    FunctionTest.iserr(0.0/0.0).should == true
    FunctionTest.iserr(10.0/0.0).should == true
  end
  
  it "should return false if passed anything else" do
    FunctionTest.iserr(123).should == false
  end
end

describe "excel_if" do
  it "should return its second argument if its first argument is true" do
    FunctionTest.excel_if(true,:second,:third).should == :second
  end
  
  it "should return its third argument if its first argument is false" do
    FunctionTest.excel_if(false,:second,:third).should == :third
  end
  
  it "should have a default value of false for its second argument" do
    FunctionTest.excel_if(true,:second).should == :second
    FunctionTest.excel_if(false,:second).should == false
  end
end

describe "iferror" do
  it "should return its second value if there is an error in the first" do
    FunctionTest.iferror(FunctionTest.index(FunctionTest.a('a1','b3'),3.0,1.0),"Not found").should == 0.0
    FunctionTest.iferror(FunctionTest.index(FunctionTest.a('a1','b3'),3.0,3.0),"Not found").should == "Not found"
    FunctionTest.iferror(0.0/0.0,"Zero division").should == "Zero division"
  end
end

describe "excel_and" do
  it "should return true if all its arguments are true" do
    FunctionTest.excel_and(true,true,true) == true
  end
  
  it "should return false if any of its arguments are false" do
    FunctionTest.excel_and(true,false,true) == false
  end
end

describe "excel_or" do
  it "should return false if all of its arguments are false" do
    FunctionTest.excel_or(false,false,false) == false
  end
  
  it "should return true if any of its arguments are true" do
    FunctionTest.excel_or(false,false,true) == true
    FunctionTest.excel_or(true,true,true) == true
  end
end

describe "left" do
  it "should return the left n characters from a string" do
    FunctionTest.left("ONE").should == "O"
    FunctionTest.left("ONE",1).should == "O"
    FunctionTest.left("ONE",3).should == "ONE"
  end
end

describe "find" do
  it "should find the first occurrence of one string in another" do
    FunctionTest.find("one","onetwothree").should == 1
    FunctionTest.find("one","twoonethree").should == 4
    FunctionTest.find("one","twoonthree").should == :value
  end
  
  it "should find the first occurrence of one string in another after a given index" do
    FunctionTest.find("one","onetwothree",1).should == 1
    FunctionTest.find("one","twoonethree",5).should == :value
    FunctionTest.find("one","oneone",2).should == 4
  end
end

describe "pmt" do
  it "should calculate the monthly payment required for a given principal, interest rate and loan period" do
    FunctionTest.pmt(0.1,10,100).should be_within(0.01).of(-16.27)
    FunctionTest.pmt(0.0123,99.1,123.32).should be_within(0.01).of(-2.159)
    FunctionTest.pmt(0,2,10).should be_within(0.01).of(-5)
  end
end

describe "npv" do
  it "should calculate the discounted value of future cash flows" do
    FunctionTest.npv(0.1,-10000,3000,4200,6800).should be_within(0.01).of(1188.44)
    FunctionTest.npv(0.08,8000,9200,10000,12000,14500).should be_within(0.01).of(1922.06+40000)
    FunctionTest.npv(0.08,8000,9200,10000,12000,14500,-9000).should be_within(0.01).of(-3749.47+40000)
    FunctionTest.npv(0.1,FunctionTest.a('a1','a2'),20).should  be_within(0.01).of(106.76)
  end
end

describe "text" do
  it "should round a numeric value to a number of decimal places and then return a string if in the format text(3.1415,0), otherwise not implemented" do
    FunctionTest.text(3.1415,0).should == "3"
    FunctionTest.text(3.1415,1).should == "3.1"
  end
end


describe "excel comparisons" do
  
  it "should carry out comparisons in the usual way" do
    FunctionTest.excel_comparison(10,"==",5).should == false
    FunctionTest.excel_comparison(10,"<=",5).should == false
    FunctionTest.excel_comparison(10,">=",5).should == true
    FunctionTest.excel_comparison(10,"<",5).should == false
    FunctionTest.excel_comparison(10,">",5).should == true
    FunctionTest.excel_comparison(10,"==",10).should == true
    FunctionTest.excel_comparison(10,"!=",10).should == false
    FunctionTest.excel_comparison(10,"<=",10).should == true
    FunctionTest.excel_comparison(10,">=",10).should == true
    FunctionTest.excel_comparison(10,"<",10).should == false
    FunctionTest.excel_comparison(10,">",10).should == false 
  end

  it "should test for equality, ignoring string case" do        
    FunctionTest.excel_comparison("A","==","a").should == true
  end
end

describe "ability to respond to empty cell references" do
  it "should return 0 if a reference is made to an empty cell" do
    FunctionTest.a23.should == 0.0
  end
  
  it "should return an object that is kind_of?(Empty) if a reference is made to an empty cell" do
    FunctionTest.a23.should be_kind_of(Empty)
  end
end

describe "ability to respond to the a, r and c methods for creating area references" do
  it "should return an Area object for a(start_cell,end_cell)" do
    FunctionTest.a('a1','b10').should be_kind_of(Area)
  end
  
  it "should return a Columns object for c(start_column_as_text,end_column_as_text)" do
    FunctionTest.c('a','b').should be_kind_of(Columns)
  end
  
  it "should return a Rows object for r(start_row_number,end_row_number)" do
    FunctionTest.r(1,10).should be_kind_of(Rows)
  end
end

class FunctionTest2
  include RubyFromExcel::ExcelFunctions
  def initialize
    @worksheet_names = {'first sheet'=>'sheet1'}
    @workbook_tables = {"FirstTable"=>'Table.new(sheet1,"FirstTable","A1:C3",["ColA", "ColB", "ColC"],1)'}
  end
  def sheet1
    self
  end
  
  def name; "sheet1"; end
    
  def a1; "Cell A1"; end
  def a2; "Middle A2"; end
  def a3; "Total A3"; end
  def to_s; 'sheet1'; end
end

describe "ability to return a sheet from the full excel name using s('excel name')" do
  it "should return a string for known sheets" do
    sheet = FunctionTest2.new
    sheet.s('first sheet').should == sheet
  end
end

describe "ability to return the table for a given name using t('table name')" do
  it "should return a string for known sheets" do
    FunctionTest2.new.t('FirstTable').reference_for('[#Totals],[ColA]').to_s.should == 'sheet1.a3'
  end
  
  it "should return a lowercase form of the class name as variable_name for compatibility with Table" do
    FunctionTest.variable_name.should == "class"
  end
end

describe "indirect method" do
  it "should deal with the simple case where the indirect is pointing to a sheet or cell" do
    FunctionTest2.new.indirect("A$1$").should == "Cell A1"
  end
  
  it "should also work with sheet name references" do
    FunctionTest2.new.indirect("'first sheet'!A$1$").should == "Cell A1"
  end
  
  it "should also work with table references" do
    FunctionTest2.new.indirect("FirstTable[[#Totals],[ColA]]").should == "Total A3"
    FunctionTest2.new.indirect("FirstTable[[#This Row],[ColA]]",'B2').should == "Middle A2"
  end
  
end

describe "boolean values" do
  it "should be possible to treat true values as 1.0 when using them in arithmetic" do
    (true*5.0).should == 5.0
    (true+1.0).should == 2.0
    (true/5.0).should == 0.2
    (true-2.0).should == -1.0
    (5.0*true).should == 5.0
    (1.0+true).should == 2.0
    (5.0/true).should == 5.0
    (2.0-true).should == 1.0
  end
  
  it "should be possible to treat false values as 0.0 when using them in arithmetic" do
    (false*5.0).should == 0.0
    (false+1.0).should == 1.0
    (false/5.0).should == 0.0
    (false-2.0).should == -2.0    
    (5.0*false).should == 0.0
    (1.0+false).should == 1.0
    (5.0/false).should == (5.0/0.0)
    (2.0-false).should == 2.0
  end
end

describe "symbols" do
  it "should be able to have infix operators applied to them" do
    (-:one).should == :one
  end
  
  it "should be able to participate in arithmetic" do
    (1 + :na).should == :na
    (1 - :na).should == :na
    (1 * :na).should == :na
    (1 / :na).should == :na
    (:na + 1).should == :na
    (:na - 1).should == :na
    (:na * 1).should == :na
    (:na / 1).should == :na
  end
end

describe "Strings", "should be able to operate as numbers if they appear to be numbers" do
  it "should be able to have infix operators applied to them" do
    (-"10").should == -10.0
  end
  
  it "should be able to participate in arithmetic" do
    (1 + "10").should == 11.0
    (1 - "10").should == -9.0
    (1 * "10").should == 10.0
    (1 / "10").should == 0.1
    ("10" + 1).should == 11.0
    ("10" + "20").should == 30.0
    ("10" + " apples").should == "10 apples"
    ("10" - 1).should == 9.0
    ("10" * 2).should == 20.0
    ("10" / 1).should == 10.0
  end
end

class CacheFuncitonTest
  attr_accessor :method_executed
  include RubyFromExcel::ExcelFunctions
  def a23
    @method_executed = true
    "one"
  end
  
  def other_function
    @method_executed = true
    "one"
  end
end

describe "array formula operations" do
  
  it "m(expression) should create an ExcelMatrixCollection, carry out an excel_map using the associated block and return an ExcelMatrix" do
    m = FunctionTest.m("one",[[3]],[[:a,:b,:c],[1,2,3]],ExcelMatrix.new([:a,:b,:c])) { |i,j,k| "#{i}#{j}#{k}" }
    m.should be_kind_of(ExcelMatrix)
    m.values.should == [["one3a", "one3b", "one3c"], ["one31", "one32", "one33"]]
  end
  
end