require_relative 'spec_helper'

describe SharedFormulaDependencyBuilder do
  
  before(:each) do
    SheetNames.instance.clear
    SheetNames.instance['Other Sheet'] = 'sheet2'
    @workbook = mock(:workbook, :named_references => {'named_cell' => 'sheet2.z10', 'named_cell2' => "sheet2.a('z10','ab10')"})
    @worksheet1 = mock(:worksheet, :to_s => 'sheet1', :workbook => @workbook, :named_references => {'named_cell' => 'sheet1.a1'})
    @worksheet2 = mock(:worksheet, :to_s => 'sheet2', :workbook => @workbook, :named_references => {})
    @workbook.stub!(:worksheets => {'sheet1' => @worksheet1, 'sheet2' => @worksheet2 })
    @cell = mock(:cell,:worksheet => @worksheet1, :reference => Reference.new('c30',@worksheet1))
    @builder = SharedFormulaDependencyBuilder.new(@cell)
    @builder.shared_formula_offset = [1,1]
  end
  
  def ruby_for(formula)
    ast = Formula.parse(formula)
    ast.visit(@builder)
  end
  
  it "should offset single cell dependencies appropriately" do
     ruby_for("A1").should == ['sheet1.b2']
     ruby_for("$A1").should == ['sheet1.a2']
     ruby_for("A$1").should == ['sheet1.b1']
     ruby_for("$A$1").should == ['sheet1.a1']
   end

   it "should offset dependences on other worksheets" do
     ruby_for("A1+'Other Sheet'!A1").should == ['sheet1.b2','sheet2.b2']
   end

   it "should offset dependences in ranges" do
     ruby_for("A1:A3").should == ['sheet1.b2','sheet1.b3','sheet1.b4']
   end

   it "should not offset dependencies in named_cells" do
     ruby_for("named_cell").should == ['sheet1.a1']
   end
   
   it "should not offset dependencies created by table references" do
     Table.new(@worksheet2,'Vectors','a1:b2',['ColA','Description'],1)     
     ruby_for("Vectors[#all]").should == ["sheet2.a1","sheet2.a2",'sheet2.b1','sheet2.b2']
     Table.new(@worksheet1,'Vectors','a1:d41',['ColA','Description','ColC'],1)
     ruby_for("[Description]").should == ["sheet1.b30"]
   end
    
end