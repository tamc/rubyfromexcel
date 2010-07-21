module RubyFromExcel
  class ArrayFormulaCell < FormulaCell
  
    def ArrayFormulaCell.from_other_cell(cell)
      afc = ArrayFormulaCell.new(cell.worksheet)
      afc.reference = cell.reference
      afc.xml_value = cell.xml_value
      afc.xml_type =  cell.xml_type
      afc
    end
  
    attr_accessor :array_formula_reference
    attr_accessor :array_formula_offset
  
    def parse_xml(xml)
      # No
    end
  
    def work_out_dependencies
      # No
    end
  
    def ruby_value
      "@#{reference.to_ruby} ||= #{array_formula_reference}.array_formula_offset(#{array_formula_offset.join(',')})"
    end
  end
end