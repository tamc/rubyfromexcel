require_relative 'spec_helper'

require 'textpeg2rubypeg'
text_peg = IO.readlines(File.join(File.dirname(__FILE__),'..','lib','formulae','parse','formula_peg.txt')).join
ast = TextPeg.parse(text_peg)
# puts ast.to_ast
# exit
builder = TextPeg2RubyPeg.new
ruby = ast.visit(builder)
Kernel.eval(ruby)

describe Formula do
  
  def check(text)
    puts
    e = Formula.new
    e.parse(text)
    e.pretty_print_cache(true)
    puts
  end
  
  it "returns formulas" do
    Formula.parse('1+1').to_ast.first.should == :formula
  end
  
  it "returns cells" do
    Formula.parse('$A$1').to_ast.should == [:formula,[:cell,'$A$1']]
    Formula.parse('A1').to_ast.should == [:formula,[:cell,'A1']]    
    Formula.parse('$A1').to_ast.should == [:formula,[:cell,'$A1']]    
    Formula.parse('A$1').to_ast.should == [:formula,[:cell,'A$1']]    
    Formula.parse('AAA1123').to_ast.should == [:formula,[:cell,'AAA1123']]
    Formula.parse("IF(a23>=(1.0+b38),1.0,2.0)") == [:formula, [:function, "IF", [:comparison, [:cell, "a23"], [:comparator, ">="], [:brackets, [:arithmetic, [:number, "1.0"], [:operator, "+"], [:cell, "b38"]]]], [:number, "1.0"], [:number, "2.0"]]]
  end
  
  it "returns areas" do
    Formula.parse('$A$1:$Z$1').to_ast.should == [:formula,[:area,'$A$1','$Z$1']]
    Formula.parse('A1:$Z$1').to_ast.should == [:formula,[:area,'A1','$Z$1']]    
  end
  
  it "returns row ranges" do
    Formula.parse('$1:$1000').to_ast.should == [:formula,[:row_range,'$1','$1000']]
    Formula.parse('1000:1').to_ast.should == [:formula,[:row_range,'1000','1']]    
  end
  
  it "returns column ranges" do
    Formula.parse('$C:$AZ').to_ast.should == [:formula,[:column_range,'$C','$AZ']]
    Formula.parse('C:AZ').to_ast.should == [:formula,[:column_range,'C','AZ']]    
  end
  
  it "returns references to other sheets" do
    Formula.parse('sheet1!$A$1').to_ast.should == [:formula,[:sheet_reference,'sheet1',[:cell,'$A$1']]]    
    Formula.parse('sheet1!$A$1:$Z$1').to_ast.should == [:formula,[:sheet_reference,'sheet1',[:area,'$A$1','$Z$1']]]
    Formula.parse('sheet1!$1:$1000').to_ast.should == [:formula,[:sheet_reference,'sheet1',[:row_range,'$1','$1000']]]
    Formula.parse('sheet1!$C:$AZ').to_ast.should == [:formula,[:sheet_reference,'sheet1',[:column_range,'$C','$AZ']]]
  end
  
  it "returns references to other sheets with extended names" do
    Formula.parse("'sheet 1'!$A$1").to_ast.should == [:formula,[:quoted_sheet_reference,'sheet 1',[:cell,'$A$1']]]    
    Formula.parse("'sheet 1'!$A$1:$Z$1").to_ast.should == [:formula,[:quoted_sheet_reference,'sheet 1',[:area,'$A$1','$Z$1']]]
    Formula.parse("'sheet 1'!$1:$1000").to_ast.should == [:formula,[:quoted_sheet_reference,'sheet 1',[:row_range,'$1','$1000']]]
    Formula.parse("'sheet 1'!$C:$AZ").to_ast.should == [:formula,[:quoted_sheet_reference,'sheet 1',[:column_range,'$C','$AZ']]]
    Formula.parse("'2007.0'!Year.Matrix").to_ast.should == [:formula, [:quoted_sheet_reference, "2007.0", [:named_reference, "Year.Matrix"]]]
  end
  
  it "returns numbers" do
    Formula.parse("1").to_ast.should == [:formula,[:number,'1']]
    Formula.parse("103.287").to_ast.should == [:formula,[:number,'103.287']]
    Formula.parse("-1.0E-27").to_ast.should == [:formula,[:number,'-1.0E-27']]
  end

  it "returns percentages" do
    Formula.parse("1%").to_ast.should == [:formula,[:percentage,'1']]
    Formula.parse("103.287%").to_ast.should == [:formula,[:percentage,'103.287']]
    Formula.parse("-1.0%").to_ast.should == [:formula,[:percentage,'-1.0']]
  end
  
  it "returns strings" do
    Formula.parse('"A handy string"').to_ast.should == [:formula,[:string,"A handy string"]]
    Formula.parse('"$A$1"').to_ast.should == [:formula,[:string,"$A$1"]]  
  end
  
  it "returns string joins" do
    Formula.parse('"A handy string"&$A$1').to_ast.should == [:formula,[:string_join,[:string,"A handy string"],[:cell,'$A$1']]]
    Formula.parse('$A$1&"A handy string"').to_ast.should == [:formula,[:string_join,[:cell,'$A$1'],[:string,"A handy string"]]]
    Formula.parse('$A$1&"A handy string"&$A$1').to_ast.should == [:formula,[:string_join,[:cell,'$A$1'],[:string,"A handy string"],[:cell,'$A$1'],]]
    Formula.parse('$A$1&$A$1&$A$1').to_ast.should == [:formula,[:string_join,[:cell,'$A$1'],[:cell,'$A$1'],[:cell,'$A$1'],]]
    Formula.parse('"GW"&ISERR($AA$1)').to_ast.should == [:formula,[:string_join,[:string,'GW'],[:function,'ISERR',[:cell,'$AA$1']]]]
  end
  
  it "returns numeric operations" do
    Formula.parse('$A$1+$A$2+1').to_ast.should == [:formula,[:arithmetic,[:cell,'$A$1'],[:operator,"+"],[:cell,'$A$2'],[:operator,"+"],[:number,'1']]]
    Formula.parse('$A$1-$A$2-1').to_ast.should == [:formula,[:arithmetic,[:cell,'$A$1'],[:operator,"-"],[:cell,'$A$2'],[:operator,"-"],[:number,'1']]]
    Formula.parse('$A$1*$A$2*1').to_ast.should == [:formula,[:arithmetic,[:cell,'$A$1'],[:operator,"*"],[:cell,'$A$2'],[:operator,"*"],[:number,'1']]]
    Formula.parse('$A$1/$A$2/1').to_ast.should == [:formula,[:arithmetic,[:cell,'$A$1'],[:operator,"/"],[:cell,'$A$2'],[:operator,"/"],[:number,'1']]            ]
    Formula.parse('$A$1^$A$2^1').to_ast.should == [:formula,[:arithmetic,[:cell,'$A$1'],[:operator,"^"],[:cell,'$A$2'],[:operator,"^"],[:number,'1']]]
  end
  
  it "returns expressions in brackets" do
    Formula.parse('($A$1+$A$2)').to_ast.should == [:formula,[:brackets,[:arithmetic,[:cell,'$A$1'],[:operator,"+"],[:cell,'$A$2']]]]
    Formula.parse('($A$1+$A$2)+2').to_ast.should == [:formula, [:arithmetic, [:brackets, [:arithmetic, [:cell,'$A$1'], [:operator,"+"], [:cell,'$A$2']]], [:operator,"+"], [:number,'2']]]
    Formula.parse('($A$1+$A$2)+(2+(1*2))').to_ast.should == [:formula, [:arithmetic, [:brackets, [:arithmetic, [:cell,'$A$1'], [:operator,"+"], [:cell,'$A$2']]], [:operator,"+"], [:brackets, [:arithmetic, [:number,'2'], [:operator,'+'] ,[:brackets, [:arithmetic, [:number,'1'], [:operator,'*'], [:number,'2']]]]]]]  
  end
  
  it "returns comparisons" do
    Formula.parse('$A$1>$A$2').to_ast.should  == [:formula,[:comparison,[:cell,'$A$1'],[:comparator,">"],[:cell,'$A$2']]]
    Formula.parse('$A$1<$A$2').to_ast.should  == [:formula,[:comparison,[:cell,'$A$1'],[:comparator,"<"],[:cell,'$A$2']]]
    Formula.parse('$A$1=$A$2').to_ast.should  == [:formula,[:comparison,[:cell,'$A$1'],[:comparator,"="],[:cell,'$A$2']]]
    Formula.parse('$A$1>=$A$2').to_ast.should == [:formula,[:comparison,[:cell,'$A$1'],[:comparator,">="],[:cell,'$A$2']]]
    Formula.parse('$A$1<=$A$2').to_ast.should == [:formula,[:comparison,[:cell,'$A$1'],[:comparator,"<="],[:cell,'$A$2']]]
    Formula.parse('$A$1<>$A$2').to_ast.should == [:formula,[:comparison,[:cell,'$A$1'],[:comparator,"<>"],[:cell,'$A$2']]]
  end
  
  it "returns functions" do
    Formula.parse('PI()').to_ast.should  == [:formula,[:function,'PI']]
    Formula.parse('ERR($A$1)').to_ast.should  == [:formula,[:function,'ERR',[:cell,'$A$1']]]
    Formula.parse('SUM($A$1,sheet1!$1:$1000,1)').to_ast.should  == [:formula,[:function,'SUM',[:cell,'$A$1'],[:sheet_reference,'sheet1',[:row_range,'$1','$1000']],[:number,'1']]]
    Formula.parse('IF(A2="Hello","hello",sheet1!B4)').to_ast.should == [:formula, [:function, "IF", [:comparison, [:cell, "A2"], [:comparator, "="], [:string, "Hello"]], [:string, "hello"], [:sheet_reference, "sheet1", [:cell, "B4"]]]]
  end
  
  it "returns fully qualified structured references (i.e., Table[column])" do
    Formula.parse('DeptSales[Sale Amount]').to_ast.should  == [:formula,[:table_reference,'DeptSales','Sale Amount']]
    Formula.parse('DeptSales[[#Totals],[ColA]]').to_ast.should == [:formula,[:table_reference,'DeptSales','[#Totals],[ColA]']]
    Formula.parse('IV.Outputs[Vector]').to_ast.should == [:formula,[:table_reference,'IV.Outputs','Vector']]
    Formula.parse("I.b.Outputs[2007.0]").to_ast.should == [:formula,[:table_reference,'I.b.Outputs','2007.0']]
    Formula.parse("INDEX(Global.Assumptions[Households], MATCH(F$321,Global.Assumptions[Year], 0))").to_ast.should == [:formula, [:function, "INDEX", [:table_reference, "Global.Assumptions", "Households"], [:function, "MATCH", [:cell, "F$321"], [:table_reference, "Global.Assumptions", "Year"], [:number, "0"]]]]
    Formula.parse("MAX(-SUM(I.a.Inputs[2007])-F$80,0)").to_ast.should == [:formula, [:function, "MAX", [:arithmetic, [:prefix, "-", [:function, "SUM", [:table_reference, "I.a.Inputs", "2007"]]], [:operator, "-"], [:cell, "F$80"]], [:number, "0"]]]
  end
  
  it "returns booleans" do
    Formula.parse("TRUE*FALSE").to_ast.should == [:formula,[:arithmetic,[:boolean_true],[:operator,'*'],[:boolean_false]]]
  end
  
  it "returns prefixes (+/-)" do
    Formula.parse("-(3-1)").to_ast.should == [:formula, [:prefix, "-", [:brackets, [:arithmetic, [:number, "3"], [:operator, "-"], [:number, "1"]]]]]
  end
  
  it "returns local structured references (i.e., [column])" do
    Formula.parse('[Sale Amount]').to_ast.should  == [:formula,[:local_table_reference,'Sale Amount']]
  end
  
  it "returns named references" do
    Formula.parse('EF.NaturalGas.N2O').to_ast.should == [:formula,[:named_reference,'EF.NaturalGas.N2O']]
  end
  
  it "returns infix modifiers in edge cases" do
    complex = "(-'DUKES 09 (2.5)'!$B$25)"
    Formula.parse(complex).to_ast.should == [:formula, [:brackets, [:prefix, "-", [:quoted_sheet_reference, "DUKES 09 (2.5)", [:cell, "$B$25"]]]]]
    complex2 = %q{-(INDEX(INDIRECT(BI$19&"!Year.Matrix"),MATCH("Subtotal.Supply",INDIRECT(BI$19&"!Year.Modules"),0),MATCH("V.04",INDIRECT(BI$19&"!Year.Vectors"),0))+
INDEX(INDIRECT(BI$19&"!Year.Matrix"),MATCH("Subtotal.Consumption",INDIRECT(BI$19&"!Year.Modules"),0),MATCH("V.04",INDIRECT(BI$19&"!Year.Vectors"),0)))}
    Formula.parse(complex2).to_ast.should == [:formula, [:prefix, "-", [:brackets, [:arithmetic, [:function, "INDEX", [:function, "INDIRECT", [:string_join, [:cell, "BI$19"], [:string, "!Year.Matrix"]]], [:function, "MATCH", [:string, "Subtotal.Supply"], [:function, "INDIRECT", [:string_join, [:cell, "BI$19"], [:string, "!Year.Modules"]]], [:number, "0"]], [:function, "MATCH", [:string, "V.04"], [:function, "INDIRECT", [:string_join, [:cell, "BI$19"], [:string, "!Year.Vectors"]]], [:number, "0"]]], [:operator, "+"], [:function, "INDEX", [:function, "INDIRECT", [:string_join, [:cell, "BI$19"], [:string, "!Year.Matrix"]]], [:function, "MATCH", [:string, "Subtotal.Consumption"], [:function, "INDIRECT", [:string_join, [:cell, "BI$19"], [:string, "!Year.Modules"]]], [:number, "0"]], [:function, "MATCH", [:string, "V.04"], [:function, "INDIRECT", [:string_join, [:cell, "BI$19"], [:string, "!Year.Vectors"]]], [:number, "0"]]]]]]]
    Formula.parse("MAX(MIN(F121, -F22),0)").to_ast.should == [:formula, [:function, "MAX", [:function, "MIN", [:cell, "F121"], [:prefix, "-", [:cell, "F22"]]], [:number, "0"]]]
  end
  
  it "returns formulas with spaces" do
    spaced = "(13.56 / 96935) * EF.IndustrialCoal.CO2 * GWP.CH4"
    Formula.parse(spaced).to_ast.should == [:formula, [:arithmetic, [:brackets, [:arithmetic, [:number, "13.56"], [:operator, "/"], [:number, "96935"]]], [:operator, "*"], [:named_reference, "EF.IndustrialCoal.CO2"], [:operator, "*"], [:named_reference, "GWP.CH4"]]]
    spaced2 ='"per " & Preferences.EnergyUnits'
    Formula.parse(spaced2).to_ast.should == [:formula, [:string_join, [:string, "per "],[:named_reference, "Preferences.EnergyUnits"]]]
    spaced3 = ' 0.00403/$F$76'
    Formula.parse(spaced3).to_ast.should == [:formula, [:arithmetic, [:number, "0.00403"], [:operator, "/"], [:cell, "$F$76"]]]
    spaced4 = "SUMIFS(INDEX($G$62:$J$73, , MATCH($E$11, $G$61:$J$61, 0)), $C$62:$C$73, $C195, $D$62:$D$73, $D195)"
    Formula.parse(spaced4).to_ast.should == [:formula, [:function, "SUMIFS", [:function, "INDEX", [:area, "$G$62", "$J$73"], [:null], [:function, "MATCH", [:cell, "$E$11"], [:area, "$G$61", "$J$61"], [:number, "0"]]], [:area, "$C$62", "$C$73"], [:cell, "$C195"], [:area, "$D$62", "$D$73"], [:cell, "$D195"]]]
  end
  
end