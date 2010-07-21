require_relative 'single_cell_array_formula_builder'

module RubyFromExcel
  class SingleCellArrayFormulaCell < FormulaCell
    
    def ruby_value
      "@#{reference.to_ruby} ||= #{ast.visit(SingleCellArrayFormulaBuilder.new(self))}"
    end
  
  end
end