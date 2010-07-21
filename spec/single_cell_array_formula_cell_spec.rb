require_relative 'spec_helper'

describe SingleCellArrayFormulaCell do
  
  before do
    @cell = SingleCellArrayFormulaCell.new(
      mock(:worksheet,:class_name => 'Sheet1', :to_s => 'sheet1'),
      Nokogiri::XML("<c r=\"C3\"><f t=\"array\" ref=\"C3\">SUM(F$437:F$449/$M$150:$M$162*($B454=$F$149:$N$149)*($F$150:$N$162))</f><v>3</v></c>").root
    )
  end
  
  it "should use the FormulaPeg to create ruby code for a formula and turn that into a method" do
    @cell.to_ruby.should == "def c3; @c3 ||= sum(m(a('f437','f449'),a('m150','m162'),(m(b454,a('f149','n149')) { |r1,r2| r1==r2 }),(a('f150','n162'))) { |r1,r2,r3,r4| r1/r2*r3*r4 }); end\n"
  end
  
  it "should create a test for the ruby code" do
    @cell.to_test.should == %Q{it 'cell c3 should equal 3.0' do\n  sheet1.c3.should be_close(3.0,0.3)\nend\n\n}
  end
  
  it "in can list the cells upon which it depends" do
    @cell.work_out_dependencies
    dependencies = @cell.dependencies
    dependencies.include?('sheet1.g151').should be_true
  end
end