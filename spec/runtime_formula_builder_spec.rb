# encoding: utf-8

require_relative 'spec_helper'

describe RuntimeFormulaBuilder, "For parsing indirect references" do
  
  before(:each) do
    @worksheet = mock(:worksheet)
    @builder = RuntimeFormulaBuilder.new(@worksheet)
  end
  
  def ruby_for(formula)
    ast = Formula.parse(formula)
    ast.visit(@builder)
  end
  # These are the new conversions
  it "should convert sheet names to ruby methods" do
    ruby_for("asheetname!A3:B20").should == "s('asheetname').a('a3','b20')"
    ruby_for("'a long sheet name with spaces'!A3:B20").should == "s('a long sheet name with spaces').a('a3','b20')"
  end
  
  it "should convert table names to references" do
    table = mock(:table)
    table.should_receive('reference_for').with('[#totals],[Description]',nil).and_return("cell reference")
    @worksheet.should_receive('t').with('Vectors').and_return(table)
    ruby_for("Vectors[[#totals],[Description]]").should == "cell reference"
  end  
  
  # All these are the standard conversions
  it "Should convert simple references to ruby method calls" do
    ruby_for("AA1").should == "aa1"
    ruby_for("$AA$100").should == "aa100"
    ruby_for("A$1").should == "a1"
    ruby_for("$A1").should == "a1"
    ruby_for("I9+J9+K9+P9+Q9").should == "i9+j9+k9+p9+q9"
  end
  
  it "Should convert area references to a(start,end) functions" do
    ruby_for('$A$1:BZ$2000').should == "a('a1','bz2000')"
  end
  
  it "Should convert column references to c(start,end) functions" do
    ruby_for('A:ZZ').should == "c('a','zz')"
  end
  
  it "should convert row references to r(start,end) functions" do
    ruby_for('1:1000').should == 'r(1,1000)'
  end
  
  it "should convert named references to ruby methods" do
    ruby_for("SUM(OneAnd2)").should == "sum(one_and2)"
    ruby_for("ReferenceOne").should == "reference_one"
    ruby_for("Reference.2").should == "reference_2"
    ruby_for("阿三").should == "阿三"
    ruby_for("-($AG70+$X70)*EF.NaturalGas.N2O").should == "-(ag70+x70)*ef_natural_gas_n2o"
  end

end