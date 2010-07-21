module RubyFromExcel
  class SharingFormulaCell < FormulaCell
  
    attr_accessor :sharing_range
  
    def parse_xml(xml)
      super
      sharing_range = xml.at_css("f")['ref']
      sharing_range = "#{sharing_range}:#{sharing_range}" unless sharing_range =~ /:/
      self.sharing_range = Area.new(worksheet,*sharing_range.split(':'))
    end
  
    def alter_other_cells_if_required
      share_formula
    end
  
    def share_formula
      sharing_range.to_reference_enum.each do |shared_formula_reference|
        next if shared_formula_reference == reference.to_s
        share_formula_with_cell  worksheet.cell(shared_formula_reference)
      end
    end
  
    def share_formula_with_cell(cell)
      return unless cell
      return unless cell.is_a?(SharedFormulaCell)
      cell.shared_formula = self.ast
      cell.shared_formula_offset = offset_from(cell)
    end
  
    def offset_from(cell)
      cell.reference - self.reference
    end

  end
end