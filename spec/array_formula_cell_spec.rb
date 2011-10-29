require_relative 'spec_helper'

describe ArrayFormulaCell do
  
#  <c r="B3"><f t="array" ref="B3:E6">B2:E2+A3:A6</f><v>2</v></c>

it "it is given a value cell and a pre-parsed formula and picks out values from its array references according to array_formula_offset" do
  value_cell = ValueCell.new(mock('worksheet',:name => 'sheet1',:to_s => 'sheet1'),Nokogiri::XML('<c r="D6"><v>7</v></c>').root)
  
  cell = ArrayFormulaCell.from_other_cell(value_cell)
  cell.array_formula_reference = "b3_array"
  cell.array_formula_offset = [1,1]
  cell.to_ruby.should == "def d6; @d6 ||= b3_array.array_formula_offset(1,1); end\n"
  cell.to_test.should == "it 'cell d6 should equal 7.0' do\n  sheet1.d6.should be_within(0.7000000000000001).of(7.0)\nend\n\n"
end

end
