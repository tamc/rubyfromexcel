require_relative 'spec_helper'

describe Reference do
  it "converts excel A1 style references to ruby method calls" do
    Reference.new("A1").to_ruby.should == "a1"
    Reference.new("$AA$103").to_ruby.should == "aa103"
    Reference.new("ZZ1").to_ruby.should == "zz1"
  end
  
  it "offsets references if required, dealing with any absolute constraints" do
    Reference.new("A1").shift!([1,1]).to_ruby.should == "b2"
    Reference.new("A$1").shift!([1,1]).to_ruby.should == "b1"
    Reference.new("$A1").shift!([1,1]).to_ruby.should == "a2"
    Reference.new("$A$1").shift!([1,1]).to_ruby.should == "a1"
  end
  
  it "manages wrap on column numbering when offsetting" do
    Reference.new("Z1").shift!([1,1]).to_ruby.should == "aa2"
    Reference.new("ZZ1").shift!([1,1]).to_ruby.should == "aaa2"  
  end
  
  it "deals with a nil value as an offset, treating it as offset [0,0]" do
    Reference.new("A1").shift!(nil).to_ruby.should == "a1"
  end
  
  it "converts columns to integers" do
    Reference.column_to_integer("A").should == 1
    Reference.column_to_integer("B").should == 2
    Reference.column_to_integer("Z").should == 26
    Reference.column_to_integer("AA").should == 27
    Reference.column_to_integer("AB").should == 28
  end
  
  it "converts integers to columns" do
    Reference.integer_to_column(1).should == "a"
    Reference.integer_to_column(2).should == "b"
    Reference.integer_to_column(26).should == "z"
    Reference.integer_to_column(27).should == "aa"
    Reference.integer_to_column(28).should == "ab"
  end
  
  it "raises an exception if asked for a column number < 1 (Excel column numbering starts at 1)" do
    lambda { Reference.integer_to_column(0) }.should raise_error(Exception)
  end
  
  it "can return the column_number and row_number for a reference" do
    Reference.new("AB1").column_number.should == 28
    Reference.new("AB1").row_number.should == 1
  end
  
  it "converts column and row numbers into ruby versions of cell references" do
    Reference.ruby_for(28,1).should == "ab1"
    Reference.ruby_for(16384,16384).should == "xfd16384"
    Reference.ruby_for(26,10).should == "z10"
    Reference.ruby_for(701,10).should == "zy10"
    Reference.ruby_for(702,10).should == "zz10"
    Reference.ruby_for(703,10).should == "aaa10"
    Reference.ruby_for(704,10).should == "aab10"
    Reference.ruby_for(1378,10).should == "azz10"
    Reference.ruby_for(2054,10).should == "bzz10"
  end
  
  it "returns its value as a ruby_method style reference when to_s is called" do
    Reference.new("$AB$1").to_s.should == "ab1"
  end
  
  it "returns an array with [row_offset,column_offset] when subtracted from another reference" do
    (Reference.new("A1") - Reference.new("A1")).should == [0,0]
    (Reference.new("B2") - Reference.new("A1")).should == [1,1]
    (Reference.new("AA10") - Reference.new("Z9")).should == [1,1]
  end
end
