require_relative 'spec_helper'

describe DependencyBuilder do
  
  before(:each) do
    SheetNames.instance.clear
    SheetNames.instance['Other Sheet'] = 'sheet2'
    @workbook = mock(:workbook, :named_references => {'named_cell' => 'sheet2.z10', 'named_cell2' => "sheet2.a('z10','ab10')"})
    @worksheet1 = mock(:worksheet, :to_s => 'sheet1', :workbook => @workbook, :named_references => {'named_cell' => 'sheet1.a1'})
    @worksheet2 = mock(:worksheet, :to_s => 'sheet2', :workbook => @workbook, :named_references => {})
    @workbook.stub!(:worksheets => {'sheet1' => @worksheet1, 'sheet2' => @worksheet2 })
    @cell = mock(:cell,:worksheet => @worksheet1, :reference => Reference.new('c30',@worksheet1))
    @builder = DependencyBuilder.new(@cell)
  end
  
  def dependencies_for(formula)
    ast = Formula.parse(formula)
    ast.visit(@builder)
  end
  
  it "should know about single dependencies to referred cells" do
     dependencies_for("A1").should == ['sheet1.a1']
   end

   it "should know about multiple dependencies to referred cells " do
     dependencies_for("A1+(B2*A1)").should == ['sheet1.a1','sheet1.b2']
   end

   it "should know about dependences to other worksheets" do
     dependencies_for("A1+'Other Sheet'!A1").should == ['sheet1.a1','sheet2.a1']
   end

   it "should know about dependences in ranges" do
     dependencies_for("A1:A3").should == ['sheet1.a1','sheet1.a2','sheet1.a3']
   end

   it "should know about dependences in ranges on other sheets" do
     dependencies_for("'Other Sheet'!A1:A3").should == ['sheet2.a1','sheet2.a2','sheet2.a3']
   end

   it "should know about dependencies in named ranges" do
     dependencies_for("named_cell").should == ['sheet1.a1']
   end

   it "should know about dependencies in named ranges on other sheets" do
     dependencies_for("'Other Sheet'!named_cell").should == ['sheet2.z10']
     dependencies_for("'Other Sheet'!named_cell2").should == ['sheet2.aa10','sheet2.ab10','sheet2.z10']
   end
   
   it "should know about dependencies created by table references" do
     Table.new(@worksheet2,'Vectors','a1:b2',['ColA','Description'],1)     
     dependencies_for("Vectors[#all]").should == ["sheet2.a1","sheet2.a2",'sheet2.b1','sheet2.b2']
   end
   
   it "should know about dependencies created by table references provided without table names" do
     Table.new(@worksheet1,'Vectors','a1:c41',['ColA','Description','ColC'],1)
     dependencies_for("[Description]").should == ["sheet1.b30"]
   end
   
   it "should know about dependencies created by indirect formulae" do
      Table.new(@worksheet2,'IndirectVectors','a1:b2',['ColA','Description'],1)
      dependencies_for('INDIRECT("IndirectVectors[#all]")').should == ["sheet2.a1","sheet2.a2",'sheet2.b1','sheet2.b2']
   end
  
   it "and be able to deal with indirect formulae that call upon other vectors" do
      Table.new(@worksheet2,'IndirectVectors2','a1:b10',['ColA','Description'],1)
      @worksheet1.should_receive(:cell).with('c1').and_return(mock(:cell,:value_for_including => 'ColA',:can_be_replaced_with_value? => true))
      dependencies_for('INDIRECT("IndirectVectors2["&C1&"]")').should == ["sheet1.c1","sheet2.a2", "sheet2.a3", "sheet2.a4", "sheet2.a5", "sheet2.a6", "sheet2.a7", "sheet2.a8", "sheet2.a9"]
   end  
    
end