module RubyFromExcel

  class SharedFormulaDependencyBuilder < DependencyBuilder
    attr_accessor :shared_formula_offset
  
    def initialize(formula_cell = nil, shared_formula_offset = nil)
      super formula_cell
      self.shared_formula_offset = shared_formula_offset
    end
  
    def cell(reference)
      self.dependencies << Reference.new(reference,worksheet).shift!(shared_formula_offset).to_ruby(true)
      Reference.new(reference).to_ruby
    end
  
    def function(name,*args)
      if name == "INDIRECT"
        args.first.visit(self)
        dependencies_for SharedFormulaBuilder.new(formula_cell,shared_formula_offset).indirect_function(args.first)
      else
        args.each { |a| a.visit(self) }
      end
    end
  
    alias area_without_offset area
  
    def area(start_area,end_area)
      self.dependencies.concat Area.new(worksheet,Reference.new(start_area).shift!(shared_formula_offset).to_ruby,Reference.new(end_area).shift!(shared_formula_offset).to_ruby).dependencies
    end
  
    def dependencies_for(full_reference)
      return [] unless full_reference  =~ /(sheet\d+)\.(.*)/
      sheet_name, reference = $1, $2
      using_worksheet(sheet_name) do
        case reference
        when /a\('(.*?)','(.*?)'\)/; area_without_offset($1,$2)
        else; self.dependencies << full_reference
        end
      end    
    end
  
  end
end