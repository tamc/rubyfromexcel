require_relative 'array_formula_builder'

module RubyFromExcel
  class ArrayingFormulaCell < FormulaCell
  
    attr_accessor :array_range
    attr_accessor :array_formula_reference
    attr_accessor :array_formula_offset

    def parse_xml(xml)
      super
      self.array_range = Area.new(worksheet,*xml.at_css("f")['ref'].split(':'))
    end

    def alter_other_cells_if_required
     self.array_formula_offset = [0,0]
     self.array_formula_reference = reference.to_ruby + "_array"
     array_formula_from_this_cell_onto_range
    end
  
    def array_formula_from_this_cell_onto_range
      each_array_formula do |array_formula_reference,value_cell|
        array_cell = ArrayFormulaCell.from_other_cell(value_cell)
        array_formula_onto_cell array_cell
        worksheet.replace_cell(array_formula_reference,array_cell)
      end
    end
  
    def array_formula_onto_cell(cell)
      cell.array_formula_reference = self.array_formula_reference
      cell.array_formula_offset = offset_from(cell)
    end
  
    def offset_from(cell)
      cell.reference - self.reference
    end
  
    def work_out_dependencies
      super
      each_array_formula do |array_formula_reference,array_cell|
        next unless array_cell
        array_cell.dependencies = self.dependencies
      end
    end
  
    def each_array_formula
      array_range.to_reference_enum.each do |array_formula_reference|
        next if array_formula_reference == reference.to_s
        yield array_formula_reference, worksheet.cell(array_formula_reference)
      end
    end
  
    def to_ruby(r = RubyScriptWriter.new)
      r.put_simple_method array_formula_reference, "@#{array_formula_reference} ||= #{ruby_array_value}"
      r.put_simple_method reference.to_ruby, ruby_value
      r.to_s
    end
  
    def ruby_value
      "@#{reference.to_ruby} ||= #{array_formula_reference}.array_formula_offset(#{array_formula_offset.join(',')})"
    end
  
    def ruby_array_value
      ast.visit(ArrayFormulaBuilder.new(self))
    end
  end
end