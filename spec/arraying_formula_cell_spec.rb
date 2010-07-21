require_relative 'spec_helper'

describe ArrayingFormulaCell do
  
#  <c r="B3"><f t="array" ref="B3:E6">B2:E2+A3:A6</f><v>2</v></c>

  before do
    @worksheet = mock(:worksheet,:class_name => 'Sheet1', :to_s => 'sheet1')

    @arraying_cell = ArrayingFormulaCell.new(
      @worksheet,
      Nokogiri::XML('<c r="B3"><f t="array" ref="B3:C3">D2:E2+A3:A6</f><v>2</v></c>').root
    )
    @value_cell = ValueCell.new(
      @worksheet,
      Nokogiri::XML('<c r="C3"><v>7</v></c>').root
    )
    @worksheet.should_receive(:cell).with('c3').and_return(@value_cell)
    @worksheet.should_receive(:replace_cell) do |reference,new_cell|
      reference.should == 'c3'
      new_cell.to_ruby.should == "def c3; @c3 ||= b3_array.array_formula_offset(0,1); end\n"
    end
    @arraying_cell.alter_other_cells_if_required
  end
  
  it "it is given a pre-parsed formula and picks out values from its array references according to array_formula_offset" do
    @arraying_cell.to_ruby.should == "def b3_array; @b3_array ||= m(a('d2','e2'),a('a3','a6')) { |r1,r2| r1+r2 }; end\ndef b3; @b3 ||= b3_array.array_formula_offset(0,0); end\n"
  end
  
  it "should know its dependencies, and also apply them to the cells that it arrays with" do
    array_cell = mock(:array_cell)
    array_cell.should_receive(:dependencies=).with(["sheet1.a3", "sheet1.a4", "sheet1.a5", "sheet1.a6", "sheet1.d2", "sheet1.e2"])
    @worksheet.should_receive(:cell).with('c3').and_return(array_cell)
    @arraying_cell.work_out_dependencies
    @arraying_cell.dependencies.should == ["sheet1.a3", "sheet1.a4", "sheet1.a5", "sheet1.a6", "sheet1.d2", "sheet1.e2"]
  end
  
end
