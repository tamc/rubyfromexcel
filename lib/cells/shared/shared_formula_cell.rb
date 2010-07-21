require_relative 'shared_formula_builder'

module RubyFromExcel
  class SharedFormulaCell < FormulaCell
    attr_accessor :shared_formula
    attr_accessor :shared_formula_offset
    
    def parse_formula
      # No
    end
  
    def work_out_dependencies
      self.dependencies ||= shared_formula.visit(SharedFormulaDependencyBuilder.new(self,shared_formula_offset))
    end
  
    def ruby_value
      "@#{reference.to_ruby} ||= #{shared_formula.visit(SharedFormulaBuilder.new(self, shared_formula_offset))}"
    end
  end
end