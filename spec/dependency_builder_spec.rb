require_relative 'spec_helper'

describe DependencyBuilder do
  
  before(:each) do
    SheetNames.instance.clear
    SheetNames.instance['Other Sheet'] = 'sheet2'
    @workbook = mock(:workbook, :named_references => {'named_cell' => 'sheet2.z10', 'named_cell2' => "sheet2.a('z10','ab10')"})
    @worksheet1 = mock(:worksheet, :to_s => 'sheet1', :workbook => @workbook, :named_references => {'named_cell' => 'sheet1.a1','this_year' => 'sheet1.a1'})
    @worksheet2 = mock(:worksheet, :to_s => 'sheet2', :workbook => @workbook, :named_references => {'year_matrix' => "sheet1.a('a20','a22')" })
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
   
   it "should deal with indirect formulae that don't point to anywhere useful, returning a #REF!" do
     dependencies_for('INDIRECT("IndirectVectors[#all]")').should == []
     @workbook.should_receive(:indirects_used=).with(true)
     dependencies_for("INDIRECT('Other Sheet'!A1)").should == ['sheet2.a1']
   end
  
   it "and be able to deal with indirect formulae that call upon other vectors" do
      Table.new(@worksheet2,'IndirectVectors2','a1:b10',['ColA','Description'],1)
      @worksheet1.should_receive(:cell).with('c1').and_return(mock(:cell,:value_for_including => 'colA',:can_be_replaced_with_value? => true))
      @worksheet1.should_receive(:cell).with('c2').and_return(mock(:cell,:value_for_including => 'indirectvectors2',:can_be_replaced_with_value? => true))
      dependencies_for('INDIRECT(C2&"["&C1&"]")').should == ["sheet1.c1","sheet1.c2","sheet2.a2", "sheet2.a3", "sheet2.a4", "sheet2.a5", "sheet2.a6", "sheet2.a7", "sheet2.a8", "sheet2.a9"]
   end
   
   it "should be able to deal with awkward combinations of tables and indirects" do
     @worksheet1.should_receive(:cell).with('c102').twice.and_return(mock(:cell,:value_for_including => 'XVI.a',:can_be_replaced_with_value? => true))
     @worksheet1.should_receive(:cell).with('a1').and_return(mock(:cell,:value_for_including => '2050',:can_be_replaced_with_value? => true))
     Table.new(@worksheet1,'XVI.a.Outputs','a1:c41',['2050','Description','Vector'],1)
     dependencies_for(%q|IFERROR(INDEX(INDIRECT($C102&".Outputs["&this.Year&"]"), MATCH(Z$5, INDIRECT($C102&".Outputs[Vector]"), 0)), 0)|).should == ["sheet1.a1", "sheet1.a30", "sheet1.c102", "sheet1.c30", "sheet1.z5"]
     Table.new(@worksheet1,'xvi.a.inputs','a1:c41',['2050','Description','Vector'],1)
     @worksheet1.should_receive(:cell).with('c1').twice.and_return(mock(:cell,:value_for_including => 'Other Sheet',:can_be_replaced_with_value? => true))
     dependencies_for(%q|INDIRECT("'"&XVI.a.Inputs[#Headers]&"'!Year.Matrix")|).should == ["sheet1.a20", "sheet1.a21", "sheet1.a22", "sheet1.c1"] 
     dependencies_for(%q|MATCH([Vector], INDIRECT("'"&XVI.a.Inputs[#Headers]&"'!Year.Matrix"), 0)|).should == ["sheet1.a20", "sheet1.a21", "sheet1.a22", "sheet1.c1", "sheet1.c30"]
   end
   
   it "Should be able to deal with a new combination of sumproduct, table and indirect that is causing a bug" do
     workbook = mock(:workbook)
     worksheet = mock(:worksheet,:name => "sheet13", :to_s => 'sheet13', :workbook => workbook)
     workbook.stub!(:worksheets => {'sheet13' => worksheet})
     worksheet.should_receive(:cell).with('d386').and_return(mock(:cell,:value_for_including => 'PM10',:can_be_replaced_with_value? => true))
     worksheet.should_receive(:cell).with('g385').and_return(mock(:cell,:value_for_including => '2010',:can_be_replaced_with_value? => true))
     cell = mock(:cell,:worksheet => worksheet, :reference => Reference.new('f386',worksheet))
     @builder.formula_cell = cell
     Table.new(worksheet,"EF.I.a.PM10","C90:O94",["Code", "Module", "Vector", "2007", "2010", "2015", "2020", "2025", "2030", "2035", "2040", "2045", "2050"],0)
     dependencies_for('SUMPRODUCT(G$378:G$381,INDIRECT("EF.I.a."&$D386&"["&G$385&"]"))').should ==  ["sheet1.d386", "sheet1.g378", "sheet1.g379", "sheet1.g380", "sheet1.g381", "sheet1.g385", "sheet13.g91", "sheet13.g92", "sheet13.g93", "sheet13.g94"]
   end
end