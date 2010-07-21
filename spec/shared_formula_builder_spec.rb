require_relative 'spec_helper'

describe SharedFormulaBuilder, "Formulas with shared_formula_offsets" do
  
  before(:each) do
    @builder = SharedFormulaBuilder.new
    @builder.shared_formula_offset = [1,1]
  end
  
  def ruby_for(formula)
    ast = Formula.parse(formula)
    ast.visit(@builder)
  end
  
  it "should move individual references appropriately" do
    ruby_for("A1").should == "b2"
    ruby_for("A$1").should == "b1"
    ruby_for("$A1").should == "a2"
    ruby_for("$A$1").should == "a1"
  end
  
  it "should move several references in the same formula" do
    ruby_for("A1+B2").should == "b2+c3"
    ruby_for("A$1+B$1").should == "b1+c1"
    ruby_for("$A1+$B1").should == "a2+b2"
    ruby_for("$A$1+$B$1").should == "a1+b1"
  end
  
end