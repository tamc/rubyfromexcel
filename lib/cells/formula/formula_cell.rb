module RubyFromExcel
  class FormulaCell < Cell
    attr_accessor :ast
  
    def parse_xml(xml)
      super
      self.ast = Formula.parse(xml.at_css("f").content)
    end
  
    def work_out_dependencies
      self.dependencies ||= ast.visit(DependencyBuilder.new(self))
    end
    
    def ruby_value
      "@#{reference.to_ruby} ||= #{ast.visit(FormulaBuilder.new(self))}"
    end
  end
end