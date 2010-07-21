module RubyFromExcel
  class SingleCellArrayFormulaBuilder < ArrayFormulaBuilder
  
    def formula(*expressions)
      expressions.map { |e| e.visit(self) }.join
    end
    
  end
end