module RubyFromExcel

  class RuntimeFormulaBuilder < FormulaBuilder
  
    attr_accessor :worksheet, :refering_cell_reference
  
    def initialize(worksheet, refering_cell_reference = nil)
      self.worksheet = worksheet
      self.refering_cell_reference = refering_cell_reference
    end
    
    def sheet_reference(sheet_name,reference)
      "s('#{sheet_name}').#{reference.visit(self)}"
    end
  
    alias :quoted_sheet_reference :sheet_reference
  
    def table_reference(table_name,structured_reference)
      table = worksheet.t(table_name)
      table.respond_to?(:reference_for) ? table.reference_for(structured_reference,refering_cell_reference).to_s : table
    end
  
    def named_reference(name)
      name.to_method_name
    end
  
    def indirect_function(text_formula)
      "indirect(#{text_formula.visit(self)},'#{formula_cell && formula_cell.reference}')"
    end
  
  end
end