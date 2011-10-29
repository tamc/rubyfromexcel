require_relative 'spec_helper'

describe SharingFormulaCell do
  it "it is given a pre-parsed formula and offsets its references by the amount of its shared_formula_offset" do
    worksheet = mock(:worksheet,:name => 'sheet1',:class_name => 'Sheet1',:to_s => 'sheet1')
    sharing_cell = SharingFormulaCell.new(
      worksheet,
      Nokogiri::XML('<c r="W7" s="49"><f t="shared" ref="W7:W8" si="2">A1+B2</f><v>1</v></c>').root
    )
    shared_cell = SharedFormulaCell.new(
      worksheet,
      Nokogiri::XML('<c r="W8" s="49"><f t="shared" si="2"></f><v>2</v></c>').root
      )
    worksheet.should_receive(:cell).with('w8').and_return(shared_cell)
    
    sharing_cell.alter_other_cells_if_required
    sharing_cell.to_ruby.should == "def w7; @w7 ||= a1+b2; end\n"
    shared_cell.to_ruby.should == "def w8; @w8 ||= a2+b3; end\n"
  end
  
  it "knows what it depends upon" do
    worksheet = mock(:worksheet,:name => 'sheet1', :class_name => 'Sheet1',:to_s => 'sheet1')
    sharing_cell = SharingFormulaCell.new(
      worksheet,
      Nokogiri::XML('<c r="W7" s="49"><f t="shared" ref="W7:W8" si="2">A1+B2</f><v>1</v></c>').root
    )
    shared_cell = SharedFormulaCell.new(
      worksheet,
      Nokogiri::XML('<c r="W8" s="49"><f t="shared" si="2"></f><v>2</v></c>').root
      )
    worksheet.should_receive(:cell).with('w8').and_return(shared_cell)
    
    sharing_cell.alter_other_cells_if_required
    sharing_cell.work_out_dependencies
    shared_cell.work_out_dependencies
    sharing_cell.dependencies.should == ['sheet1.a1','sheet1.b2']
    shared_cell.dependencies.should == ['sheet1.a2','sheet1.b3']
  end
  
  it "it should cope if it is sharing only with itself" do
    worksheet = mock(:worksheet,:name => 'sheet1', :class_name => 'Sheet1',:to_s => 'sheet1')
    sharing_cell = SharingFormulaCell.new(
      worksheet,
      Nokogiri::XML('<c r="W7" s="49"><f t="shared" ref="W7" si="2">A1+B2</f><v>1</v></c>').root
    )
    
    sharing_cell.alter_other_cells_if_required
    sharing_cell.to_ruby.should == "def w7; @w7 ||= a1+b2; end\n"
  end
  
  it "it should cope if it is sharing to a formula that doesn't exist" do
    worksheet = mock(:worksheet,:name => 'sheet1', :class_name => 'Sheet1',:to_s => 'sheet1')
    sharing_cell = SharingFormulaCell.new(
      worksheet,
      Nokogiri::XML('<c r="W7" s="49"><f t="shared" ref="W7:W8" si="2">A1+B2</f><v>1</v></c>').root
    )
    worksheet.should_receive(:cell).with('w8').and_return(nil)
    
    sharing_cell.alter_other_cells_if_required
    sharing_cell.to_ruby.should == "def w7; @w7 ||= a1+b2; end\n"
  end
  
  it "it should cope if it is sharing with a cell that has its own formula" do
    worksheet = mock(:worksheet,:name => 'sheet1', :class_name => 'Sheet1',:to_s => 'sheet1')
    sharing_cell = SharingFormulaCell.new(
      worksheet,
      Nokogiri::XML('<c r="W7" s="49"><f t="shared" ref="W7:W8" si="2">A1+B2</f><v>1</v></c>').root
    )
    shared_cell = SimpleFormulaCell.new(
      worksheet,
      Nokogiri::XML('<c r="W8" s="49"><f>AND(1,2)</f><v>1</v></c>').root
      )
    worksheet.should_receive(:cell).with('w8').and_return(shared_cell)
    
    sharing_cell.alter_other_cells_if_required
    sharing_cell.to_ruby.should == "def w7; @w7 ||= a1+b2; end\n"
    shared_cell.to_ruby.should == "def w8; @w8 ||= excel_and(1.0,2.0); end\n"
  end
end
