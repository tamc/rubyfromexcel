module RubyFromExcel
  class SharedFormulaBuilder < FormulaBuilder
    attr_accessor :shared_formula_offset
  
    def initialize(formula_cell = nil, shared_formula_offset = nil)
      super formula_cell
      self.shared_formula_offset = shared_formula_offset
    end
  
    def cell(reference)
      Reference.new(reference).shift!(shared_formula_offset).to_ruby
    end
  
  end
end