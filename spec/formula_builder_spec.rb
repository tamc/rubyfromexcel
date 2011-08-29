require_relative 'spec_helper'

describe FormulaBuilder do
  
  before(:each) do
    @builder = FormulaBuilder.new
    SheetNames.instance['asheetname'] = 'sheet1'
    SheetNames.instance['a long sheet name with spaces'] = 'sheet2'
  end
  
  def ruby_for(formula)
    ast = Formula.parse(formula)
    ast.visit(@builder)
  end
  
  it "Should convert simple references to ruby method calls" do
    ruby_for("AA1").should == "aa1"
    ruby_for("$AA$100").should == "aa100"
    ruby_for("A$1").should == "a1"
    ruby_for("$A1").should == "a1"
    ruby_for("I9+J9+K9+P9+Q9").should == "i9+j9+k9+p9+q9"
  end
  
  it "Should convert simple arithmetic, converting numbers to floats" do
    ruby_for("(1+3)*(12+13/2.0E-12)").should == "(1.0+3.0)*(12.0+13.0/2.0e-12)"
    ruby_for("((1+3)*(12+13/2.0))").should == "((1.0+3.0)*(12.0+13.0/2.0))"
  end
  
  it "Should convert percentages appropriately to their float equivalents" do
    ruby_for("1+3%").should == "1.0+0.03"
    ruby_for("1+103.12%").should == "1.0+1.0312000000000001" # Uh oh
  end
  
  it "Should convert powers to their ruby equivalent" do
    ruby_for("1^12").should == "1.0**12.0"
  end
  
  it "Should join strings together, adding to_s " do
    ruby_for('"Hello"&"world"').should == '"Hello"+"world"'
    ruby_for('"Hello "&A1').should == '"Hello "+(a1).to_s'
    ruby_for('AA1&"GW"').should == '(aa1).to_s+"GW"'
    ruby_for('"GW"&ISERR($AA$1)').should == '"GW"+(iserr(aa1)).to_s'
    ruby_for('"GW"&23/15').should == '"GW"+(23.0/15.0).to_s'
  end
  
  it "Should convert area references to a(start,end) functions" do
    ruby_for('$A$1:BZ$2000').should == "a('a1','bz2000')"
  end
  
  it "Should convert column references to c(start,end) functions" do
    ruby_for('A:ZZ').should == "c('a','zz')"
  end
  
  it "should convert row references to r(start,end) functions" do
    ruby_for('1:1000').should == 'r(1,1000)'
  end
  
  it "should leave strings as they are, even if they look like formulas" do
    ruby_for('"A1+3*2"&$A3').should == '"A1+3*2"+(a3).to_s'
  end
  
  it "should properly escape strings where necessary" do
    ruby_for(%q{"String with 'quotes' in it"}).should == '"String with \'quotes\' in it"'
  end

  it "should convert double sets of quotes (\"\") into single sets (\")" do
    ruby_for('"A ""quote"""').should == '"A \"quote\""'
  end

  
  it "should convert sheet names to ruby methods" do
    ruby_for("asheetname!A3:B20").should == "sheet1.a('a3','b20')"
    ruby_for("'a long sheet name with spaces'!A3:B20").should == "sheet2.a('a3','b20')"
    
  end
  
  it "should ignore external references in sheet names (because they are not implemented yet)" do
    ruby_for("'[1]a long sheet name with spaces'!A3:B20").should == "sheet2.a('a3','b20')"
  end
  
  it "should convert named references to ruby methods and inline their values" do
    worksheet = mock(:worksheet)
    workbook = mock(:workbook)
    @builder.formula_cell = mock(:cell,:worksheet => worksheet)
    worksheet.should_receive(:named_references).and_return({"one_and2"=>'sheet1.a(\'a1\',\'f10\')'})
    ruby_for("SUM(OneAnd2)").should == "sum(sheet1.a('a1','f10'))"
    worksheet.should_receive(:named_references).and_return({})
    worksheet.should_receive(:workbook).and_return(workbook)
    workbook.should_receive(:named_references).and_return({"reference_one" => "sheet10.a1"})
    ruby_for("ReferenceOne").should == "sheet10.a1"
    worksheet.should_receive(:named_references).and_return({"one_and2"=>'sheet1.a(\'a1\',\'f10\')'})
    worksheet.should_receive(:workbook).and_return(workbook)
    workbook.should_receive(:named_references).and_return({"reference_one" => "sheet10.a1"})
    ruby_for("Reference.2").should == ":name"
    worksheet.should_receive(:named_references).and_return({"one_and2"=>'sheet1.a(\'a1\',\'f10\')'})
    worksheet.should_receive(:workbook).and_return(workbook)
    workbook.should_receive(:named_references).and_return({"reference_one" => "sheet10.a1","ef_natural_gas_n2o"=> "sheet10.a1"})
    ruby_for("-($AG70+$X70)*EF.NaturalGas.N2O").should == "-(ag70+x70)*sheet10.a1"
  end
  
  it "Edge case: it should convert formulae that refer to an error on a particular sheet made in a named reference definition" do
    @builder.formula_cell = nil
    ruby_for("Control!#REF!").should == ":name"
  end
  
  it "should convert table names to references" do
    Table.new(mock(:worksheet,:name => 'sheet1',:to_s => 'sheet1'),'Vectors','a1:c41',['ColA','Description','ColC'],1)
    ruby_for("Vectors[Description]").should == "sheet1.a('b2','b40')"
    ruby_for("Vectors[#all]").should == "sheet1.a('a1','c41')"
    ruby_for("Vectors[#totals]").should == "sheet1.a('a41','c41')"
    ruby_for("Vectors[[#totals],[Description]]").should == "sheet1.b41"
    @builder.formula_cell = mock(:cell,:reference => Reference.new('f30'))
    ruby_for("Vectors[[#This Row],[Description]]").should == "sheet1.b30"
    ruby_for("Vectors[[#This Row],[Description]:[ColC]]").should == "sheet1.a('b30','c30')"
  end
  
  it "should convert table names inside indirects" do
    workbook = mock(:workbook, :named_references => {'named_cell' => 'sheet2.z10', 'named_cell2' => "sheet2.a('z10','ab10')"})
    worksheet1 = mock(:worksheet,:name => "sheet1", :to_s => 'sheet1', :workbook => workbook, :named_references => {'named_cell' => 'sheet1.a1','this_year' => 'sheet1.a1'})
    worksheet2 = mock(:worksheet,:name => "sheet2", :to_s => 'sheet2', :workbook => workbook, :named_references => {})
    workbook.stub!(:worksheets => {'sheet1' => worksheet1, 'sheet2' => worksheet2 })
    worksheet1.should_receive(:cell).with('c102').twice.and_return(mock(:cell,:value_for_including => 'XVI.a',:can_be_replaced_with_value? => true))
    worksheet1.should_receive(:cell).with('a1').and_return(mock(:cell,:value_for_including => '2050',:can_be_replaced_with_value? => true))
    cell = mock(:cell,:worksheet => worksheet1, :reference => Reference.new('c30',worksheet1))    
    @builder.formula_cell = cell
    Table.new(worksheet1,'XVI.a.Outputs','a1:c41',['2050','Description','Vector'],1)
    
    ruby_for('INDIRECT($C102&".Outputs["&this.Year&"]")').should == "sheet1.a30"
    ruby_for('INDIRECT($C102&".Outputs[Vector]")').should == "sheet1.c30"
  end
  
  it "should ignore external references, assuming that they point to an internal equivalent" do
    formula_with_external = "INDEX([1]!Modules[Module], MATCH($C5, [1]!Modules[Code], 0))"
    Table.new(mock(:worksheet,:to_s => 'sheet1',:name => "sheet1",),'Modules','a1:c41',['Module','Code','ColC'],1)
    ruby_for(formula_with_external).should == "index(sheet1.a('a2','a40'),match(c5,sheet1.a('b2','b40'),0.0))"
  end
  
  it "should convert unkown table names to :ref" do
    ruby_for("Unknown[Not likely]").should == ":ref"
  end
  
  it "should convert unqualified table names to references" do
    sheet = mock(:worksheet,:name => "sheet1",:to_s => 'sheet1')
    Table.new(sheet,'Vectors','a1:c41',['ColA','Description','ColC'],1)
    @builder.formula_cell = mock(:cell,:reference => Reference.new('c30',sheet), :worksheet => sheet)
    ruby_for("[Description]").should == "sheet1.b30"
    ruby_for("[#all]").should == "sheet1.a('a1','c41')"
    ruby_for("[#headers]").should == "sheet1.c1"
    ruby_for("[#totals]").should == "sheet1.c41"
    ruby_for("[[#totals],[Description]]").should == "sheet1.b41"
    ruby_for("[[#This Row],[Description]]").should == "sheet1.b30"
  end
    
  it "should convert booleans to their ruby equivalents" do
    ruby_for("TRUE*FALSE").should == "true*false"
  end
  
  it "should convert the not function to !()" do
    ruby_for("NOT(TRUE)").should == "!(true)"
  end
  
  it "should throw an exception if trying to convert an unknown function" do
    lambda { ruby_for("NEWEXCELFUNCTION(TRUE,FALSE)") }.should raise_error(ExcelFunctionNotImplementedError)
  end
    
  it "should convert clashing excel function names to excel_name variants" do
    ruby_for("IF(TRUE,FALSE,TRUE)").should == "excel_if(true,false,true)"
    ruby_for("AND(TRUE,FALSE)").should == "excel_and(true,false)"
    ruby_for("OR(TRUE,FALSE)").should == "excel_or(true,false)"
  end
  
  it "should add the cell reference as a second argument to indirect() functions, setting workbook.indirects_used to true" do
    worksheet = mock(:worksheet)
    workbook = mock(:workbook)    
    worksheet.should_receive(:workbook).and_return(workbook)
    cell = mock(:cell,:value => 'ASD',:can_be_replaced_with_value? => false)
    worksheet.should_receive(:cell).with('a1').and_return(cell)
    workbook.should_receive(:indirects_used=).with(true)
    @builder.formula_cell = mock(:cell,:reference => Reference.new('f30'),:worksheet => worksheet)
    ruby_for('INDIRECT("ONE"&A1)').should == 'indirect("ONE"+(a1).to_s,\'f30\')'
  end
  
  it "should attempt to interpret indirect functions where that is appropriate" do
    worksheet = mock(:worksheet, :to_s => 'sheet1',:name => "sheet1")
    @builder.formula_cell = mock(:cell,:reference => Reference.new('f30',worksheet),:worksheet => worksheet)
    ruby_for('INDIRECT("A1")').should == "a1"
    ruby_for('INDIRECT("A"&"1")').should == "a1"
    ruby_for('INDIRECT("A"&1)').should == "a1"

    SheetNames.instance['sheet100'] = 'sheet100'    
    cell = mock(:cell,:value_for_including => 'sheet100',:can_be_replaced_with_value? => true)
    worksheet.should_receive(:cell).with('a1').and_return(cell)
    ruby_for('INDIRECT(A1&"!A1")').should == "sheet100.a1"

    SheetNames.instance['sheet100'] = 'sheet100'    
    workbook = mock(:workbook)    
    worksheet.should_receive(:workbook).and_return(workbook)
    workbook.should_receive(:worksheets).and_return({'sheet100' => worksheet})
    worksheet.should_receive(:cell).with('a1').and_return(cell)
    ruby_for('INDIRECT(sheet100!A1&"!A1")').should == "sheet100.a1"
    
    worksheet.should_receive(:named_references).and_return({"this_year" => 'sheet1.a1'})
    worksheet.should_receive(:workbook).and_return(workbook)
    workbook.should_receive(:worksheets).and_return({'sheet1' => worksheet})
    worksheet.should_receive(:cell).with('a1').and_return(nil)
    ruby_for('INDIRECT(this.Year & "!A1")').should == ":name"
    
    Table.new(worksheet,'Vectors','a1:c41',['ColA','Description','ColC'],1)
    worksheet.should_receive(:workbook).and_return(workbook)
    workbook.should_receive(:worksheets).and_return({'sheet1' => worksheet})
    worksheet.should_receive(:cell).with('b30').and_return(cell)
    ruby_for('INDIRECT(Vectors[[#This Row],[Description]] & "!A1")').should == "sheet100.a1"

    Table.new(worksheet,'Vectors','a1:f41',['ColA','Description','ColC','ColD','ColE','ColF'],1)
    worksheet.should_receive(:workbook).and_return(workbook)
    workbook.should_receive(:worksheets).and_return({'sheet1' => worksheet})
    worksheet.should_receive(:cell).with('f1').and_return(cell)
    ruby_for('INDIRECT([#Headers]& "!A1")').should == "sheet100.a1"
    
    worksheet.should_receive(:cell).with('c120').and_return(cell)
    ruby_for('INDIRECT($C120&".Outputs[Vector]")').should == ":ref"
  end
  
  
  it "should convert comparators into a function, so that can cope with excels version of equality" do
    ruby_for('IF(A1=3,"A","B")').should == 'excel_if(excel_comparison(a1,"==",3.0),"A","B")'
  end

  it "should convert complex formulas" do
    # SheetNames.instance['DUKES 09 (2.5)'] = 'sheet100'
    # SheetNames.instance['DUKES 09 (1.2)'] = 'sheet101'
    # ruby_for("(-'DUKES 09 (2.5)'!$B$25*1000000*Constants.GCV.Coke)+(-'DUKES 09 (2.5)'!$C$25*1000000*Constants.GCV.CokeBreeze)+('DUKES 09 (1.2)'!$B$22*Unit.ktoe)").should == "(-sheet100.b25*1000000.0*constants_gcv_coke)+(-sheet100.c25*1000000.0*constants_gcv_coke_breeze)+(sheet101.b22*unit_ktoe)"
    # complex_formula = %q{-(INDEX(INDIRECT(BI$19&"!Year.Matrix"),MATCH("Subtotal.Supply",INDIRECT(BI$19&"!Year.Modules"),0),MATCH("V.04",INDIRECT(BI$19&"!Year.Vectors"),0))+INDEX(INDIRECT(BI$19&"!Year.Matrix"),MATCH("Subtotal.Consumption",INDIRECT(BI$19&"!Year.Modules"),0),MATCH("V.04",INDIRECT(BI$19&"!Year.Vectors"),0)))}
    # # ruby_for(complex_formula).should == "-(index(indirect(bi19.to_s+\"!Year.Matrix\",''),match(\"Subtotal.Supply\",indirect(bi19.to_s+\"!Year.Modules\",''),0.0),match(\"V.04\",indirect(bi19.to_s+\"!Year.Vectors\",''),0.0))+index(indirect(bi19.to_s+\"!Year.Matrix\",''),match(\"Subtotal.Consumption\",indirect(bi19.to_s+\"!Year.Modules\",''),0.0),match(\"V.04\",indirect(bi19.to_s+\"!Year.Vectors\",''),0.0)))"
    # ruby_for("MAX(MIN(F121, -F22),0)").should == "max(min(f121,-f22),0.0)"
  end
  
  
  it "should replace MATCH() with its answer where it depends only on cells that can be replaced with their values" do
    worksheet = mock(:worksheet, :to_s => 'sheet1')
    a1 = mock(:cell,:reference => Reference.new('a1',worksheet),:worksheet => worksheet,:value_for_including => 'A', :can_be_replaced_with_value? => true)
    a2 = mock(:cell,:reference => Reference.new('a2',worksheet),:worksheet => worksheet,:value_for_including => 'A', :can_be_replaced_with_value? => true)
    a3 = mock(:cell,:reference => Reference.new('a3',worksheet),:worksheet => worksheet,:value_for_including => 'B', :can_be_replaced_with_value? => true)
    worksheet.should_receive(:cell).with('a1').exactly(3).times.and_return(a1)
    worksheet.should_receive(:cell).with('a2').exactly(2).times.and_return(a2)
    worksheet.should_receive(:cell).with('a3').exactly(2).times.and_return(a3)
    @builder.formula_cell = mock(:cell,:reference => Reference.new('f30',worksheet),:worksheet => worksheet)
    ruby_for('MATCH("A",A1:A3)').should == "2.0"
    ruby_for('MATCH("X",A1:A3,FALSE)').should == "na"
    ruby_for('MATCH("A",A1:A3,FALSE)').should == "1.0"
  end
  
  it "should replace INDEX() with a cell reference where it depends only on cells that can be replaced with their values" do
    worksheet = mock(:worksheet, :to_s => 'sheet1')
    @builder.formula_cell = mock(:cell,:reference => Reference.new('f30',worksheet),:worksheet => worksheet)
    ruby_for('INDEX(A1:A3,3.0)').should == "sheet1.a3"
    ruby_for('INDEX(A1:A3,2.0,1.0)').should == "sheet1.a2"
  end
  
  it "should put multiple items as single arguments in a function" do
    ruby_for("IF(A23>=(1.0+B38),1.0,2.0)").should == "excel_if(excel_comparison(a23,\">=\",(1.0+b38)),1.0,2.0)"
    ruby_for("MAX(F60+(G$59-F$59)*G38,0)").should == "max(f60+(g59-f59)*g38,0.0)"
  end

  it "should convert formulas with null arguments, replacing the null with 0.0" do
    spaced4 = "SUMIFS(INDEX($G$62:$J$73, , MATCH($E$11, $G$61:$J$61, 0)), $C$62:$C$73, $C195, $D$62:$D$73, $D195)"
    ruby_for(spaced4).should == "sumifs(index(a('g62','j73'),0.0,match(e11,a('g61','j61'),0.0)),a('c62','c73'),c195,a('d62','d73'),d195)"
  end
end
