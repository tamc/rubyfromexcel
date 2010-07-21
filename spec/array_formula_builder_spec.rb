require_relative 'spec_helper'

describe ArrayFormulaBuilder, "Formulas with shared_formula_offsets" do
  
  before(:each) do
    @builder = ArrayFormulaBuilder.new
  end
  
  def ruby_for(formula)
    ast = Formula.parse(formula)
    ast.visit(@builder)
  end
  
  it "should leave individual references alone" do
    ruby_for("A1").should == "a1"
    ruby_for("A$1").should == "a1"
    ruby_for("$A1").should == "a1"
    ruby_for("$A$1").should == "a1"
  end
  
  it "should replace range references with individual cell references" do
    ruby_for("A:C").should == "c('a','c')"
    ruby_for("1:10").should == "r(1,10)"
    ruby_for("IF(A1:A5>0,A1:A5,1+2)").should == "m(m(a('a1','a5'),0.0) { |r1,r2| r1>r2 },a('a1','a5'),m(1.0,2.0) { |r1,r2| r1+r2 }) { |r1,r2,r3| excel_if(r1,r2,r3) }"
    ruby_for("B10:B20+B5").should == "m(a('b10','b20'),b5) { |r1,r2| r1+r2 }"
  end
  
  it "should not replace range references where they are expected by the formula" do
    ruby_for("INDEX($F$16:$I$19, ,MATCH($E$8, $F$15:$I$15, 0))").should == "m(0.0,m(e8,0.0) { |r1,r2| match(r1,a('f15','i15'),r2) }) { |r1,r2| index(a('f16','i19'),r1,r2) }"
    ruby_for("SUMIF($F$16:$I$19,C8)").should == "m(c8) { |r1| sumif(a('f16','i19'),r1) }"
    ruby_for("SUMIF($F$16:$I$19,C8,$Q$16:$R$19)").should == "m(c8) { |r1| sumif(a('f16','i19'),r1,a('q16','r19')) }"
    ruby_for("SUMIFS($F$16:$F$19,$G$16:$G$19,1.0,$H$16:$H$19,Q1:Q4)").should == "m(1.0,a('q1','q4')) { |r1,r2| sumifs(a('f16','f19'),a('g16','g19'),r1,a('h16','h19'),r2) }"
  end
    
end