module RubyFromExcel
  class FormulaCell < Cell
    attr_accessor :ast, :original_formula
  
    def parse_xml(xml)
      super
      self.original_formula = xml.at_css("f").content
      self.ast = Formula.parse(original_formula)
    end
  
    def work_out_dependencies
      self.dependencies ||= ast.visit(DependencyBuilder.new(self))
    end
    
    def ruby_value
      "@#{reference.to_ruby} ||= #{ast.visit(FormulaBuilder.new(self))}"
    end
    
    def debug
      RubyFromExcel.debug(:cells,"#{worksheet.name}.#{reference} -> #{original_formula.inspect} -> #{ast.inspect} -> #{xml_value} (#{xml_type}) -> #{value_for_including.inspect}")
    end
    
  end
end