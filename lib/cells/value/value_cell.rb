module RubyFromExcel
  class ValueCell < Cell
  
    def self.for(other_cell)
      vc = ValueCell.new(other_cell.worksheet)
      vc.reference = other_cell.reference
      vc.xml_value = other_cell.xml_value
      vc.xml_type = other_cell.xml_type
      vc
    end
  
    def ruby_value
      value
    end
  
    def can_be_replaced_with_value?
      true
    end
  
    def to_test( r = RubySpecWriter.new)
      nil # i.e., don't bother with tests for simple_values
    end  
  end
end