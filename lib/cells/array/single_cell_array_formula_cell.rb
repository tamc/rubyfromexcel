require_relative 'single_cell_array_formula_builder'

module RubyFromExcel
  class SingleCellArrayFormulaCell < FormulaCell
    
    def ruby_value
      ruby_code = ast.visit(SingleCellArrayFormulaBuilder.new(self))
      ruby_code = "(#{ruby_code}).array_formula_offset(0,0)" if ruby_code =~ /^m\(.*\}$/
      "@#{reference.to_ruby} ||= #{ruby_code}"
    end
  
  end
end