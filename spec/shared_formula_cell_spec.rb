require_relative 'spec_helper'

describe SharedFormulaCell do
  
  before do
    @cell = SharedFormulaCell.new(
      mock(:worksheet,:class_name => 'Sheet1', :to_s => 'sheet1'),
      Nokogiri::XML('<c r="W8" s="49"><f t="shared" si="2"/><v>1</v></c>').root
    )
    @cell.shared_formula = Formula.parse("$A$1+B2")
    @cell.shared_formula_offset = [1,1]  
  end
  
  it "it is given a pre-parsed formula and offsets its references by the amount of its shared_formula_offset" do
    @cell.to_ruby.should == "def w8; @w8 ||= a1+c3; end\n"
    @cell.to_test.should == "it 'cell w8 should equal 1.0' do\n  sheet1.w8.should be_close(1.0,0.1)\nend\n\n"
  end
  
  it "knows what it depends on" do
    @cell.work_out_dependencies
    @cell.dependencies.should == ['sheet1.a1','sheet1.c3']
  end
end
