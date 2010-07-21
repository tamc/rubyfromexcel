require_relative 'spec_helper'

describe SingleCellArrayFormulaBuilder, "Single cell array formulas" do
  
  before(:each) do
    @builder = SingleCellArrayFormulaBuilder.new
  end
  
  def ruby_for(formula)
    ast = Formula.parse(formula)
    ast.visit(@builder)
  end
  
  it "should wrap operations in an array_operation(left,operation,right) methods" do
    ruby_for("SUM(F$437:F$449/$M$150:$M$162*($B454=$F$149:$N$149)*($F$150:$N$162))").should == "sum(m(a('f437','f449'),a('m150','m162'),(m(b454,a('f149','n149')) { |r1,r2| r1==r2 }),(a('f150','n162'))) { |r1,r2,r3,r4| r1/r2*r3*r4 })"
  end
  
    
end