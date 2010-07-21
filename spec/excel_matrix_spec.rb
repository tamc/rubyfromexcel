require_relative 'spec_helper'

class TestSheet2
  
  def variable_name
    'sheet1'
  end

  def a1; 1.0; end
  def a2; 2.0; end
  def a3; 3.0; end
  
end

describe ExcelMatrixCollection do
  
  it "can be created with a series of values" do
    emc = ExcelMatrixCollection.new("one",[[3]],[[:a,:b,:c],[1,2,3]],ExcelMatrix.new([:a,:b,:c]))
  end
  
  it "can enumerate the values in the collection, making sure that each matrix is the size of the largest" do
    emc = ExcelMatrixCollection.new("one",[[3]],[[:a,:b,:c],[1,2,3]],ExcelMatrix.new([:a,:b,:c]))
    emc.to_a.should == [["one",3,:a,:a],["one",3,:b,:b],["one",3,:c,:c],["one",3,1,:a],["one",3,2,:b],["one",3,3,:c]]
  end
  
  it "can map the values to a new excel matrix of he size of the largest argument" do
    emc = ExcelMatrixCollection.new("one",[[3]],[[:a,:b,:c],[1,2,3]],ExcelMatrix.new([:a,:b,:c]))
    emc.matrix_map do |a,b,c,d,e|
      1
    end.values.should == [[1,1,1],[1,1,1]]
  end
  
  it "if the mapped ExcelMatrix is a single cell, then transforms into that" do
    result = ExcelMatrixCollection.new([[3]]).matrix_map do |r1|
      Area.new(TestSheet2.new,'a1','a3')
    end
    result.should be_kind_of(Area)
    result.to_a.should == [1.0,2.0,3.0]
    result.array_formula_offset(1,0).should == 2.0
    result = ExcelMatrixCollection.new([[3]]).matrix_map do |r1|
      r1 + 1
    end
    result.should  == 4    
  end
  
end

describe ExcelMatrix do
  
  it "can be created from a single value" do
    em = ExcelMatrix.new(3)
    em.rows.should == 1
    em.columns.should == 1
    em.values.should == [[3]]
  end
  
  it "can be created from an array of arrays" do
    em = ExcelMatrix.new([[:a,:b,:c],[1,2,3]])
    em.rows.should == 2
    em.columns.should == 3
    em.values.should == [[:a,:b,:c],[1,2,3]]
  end
  
  it "can be created from a single array" do
    em = ExcelMatrix.new([:a,:b,:c])
    em.rows.should == 1
    em.columns.should == 3
    em.values.should == [[:a,:b,:c]]
  end
    
  it "can have its number of rows expanded" do
    em = ExcelMatrix.new([:a,:b,:c])
    em.rows.should == 1
    em.add_rows!(2)
    em.rows.should == 3
    em.values.should == [[:a,:b,:c],[:a,:b,:c],[:a,:b,:c]]
  end
  
  it "can have its number of columns expanded" do
    em = ExcelMatrix.new([[:a],[1]])
    em.columns.should == 1
    em.add_columns!(2)
    em.columns.should == 3
    em.values.should == [[:a,:a,:a],[1,1,1]]
  end
    
  it "responds to #array_formula_offset(row_index,column_index) with origin 0,0" do
    em = ExcelMatrix.new([[1,2,3],[4,5,6]])
    em.array_formula_offset(0,0).should == 1
    em.array_formula_offset(1,1).should == 5
  end
end