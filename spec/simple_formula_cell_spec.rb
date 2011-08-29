require_relative 'spec_helper'

describe SimpleFormulaCell, "when result is a number" do
  
  before do
    @cell = SimpleFormulaCell.new(
      mock(:worksheet,:class_name => 'Sheet1', :to_s => 'sheet1'),
      Nokogiri::XML("<c r=\"C3\"><f>1+2</f><v>3</v></c>").root
    )
  end
  
  it "should use the FormulaPeg to create ruby code for a formula and turn that into a method" do
    @cell.to_ruby.should == "def c3; @c3 ||= 1.0+2.0; end\n"
  end
  
  it "should create a test for the ruby code" do
    @cell.to_test.should == %Q{it 'cell c3 should equal 3.0' do\n  sheet1.c3.should be_within(0.30000000000000004).of(3.0)\nend\n\n}
  end
  
  it "has a method that says it cannot be replaced with its value" do
    @cell.can_be_replaced_with_value?.should be_false
  end
  
end

describe SimpleFormulaCell, "when result is a boolean" do
  
  before do
    @cell = SimpleFormulaCell.new(
      mock(:worksheet,:class_name => 'Sheet1', :to_s => 'sheet1'),
      Nokogiri::XML('<c r="C3" t="b"><f>AND(1,2)</f><v>TRUE</v></c>').root
    )
  end
  
  it "should use the FormulaPeg to create ruby code for a formula and turn that into a method" do
    @cell.to_ruby.should == "def c3; @c3 ||= excel_and(1.0,2.0); end\n"
  end
  
  it "should create a test for the ruby code" do
    @cell.to_test.should == %Q{it 'cell c3 should equal true' do\n  sheet1.c3.should == true\nend\n\n}
  end
  
end

describe SimpleFormulaCell, "when result is a string" do
  
  before do
    @cell = SimpleFormulaCell.new(
      mock(:worksheet,:class_name => 'Sheet1', :to_s => 'sheet1'),
      Nokogiri::XML('<c r="C3" t="str"><f>"Hello "&amp;A1</f><v>Hello Bob</v></c>').root
    )
  end
  
  it "should use the FormulaPeg to create ruby code for a formula and turn that into a method" do
    @cell.to_ruby.should == %Q{def c3; @c3 ||= "Hello "+a1.to_s; end\n}
  end
  
  it "should create a test for the ruby code" do
    @cell.to_test.should == %Q{it 'cell c3 should equal "Hello Bob"' do\n  sheet1.c3.should == "Hello Bob"\nend\n\n}
  end
  
end

describe SimpleFormulaCell, "can list the cells upon which it depends" do
  
  before do
    @cell = SimpleFormulaCell.new(
        mock(:worksheet,:class_name => 'Sheet1', :to_s => 'sheet1'),
        Nokogiri::XML("<c r=\"C3\" t=\"str\"><f>A1</f><v>Hello Bob</v></c>").root
      )
  end
  
  it "in simple cases knows what it depends upon" do
    @cell.work_out_dependencies
    @cell.dependencies.should == ['sheet1.a1']
  end

end  