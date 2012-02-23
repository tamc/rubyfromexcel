module RubyFromExcel
  class DependencyBuilder
    attr_accessor :formula_cell, :dependencies, :worksheet
  
    def initialize(formula_cell = nil)
      self.formula_cell = formula_cell
      self.worksheet = formula_cell.worksheet
    end
  
    def formula(*expressions)
      self.dependencies = []
      expressions.each do |e| 
        e.visit(self)
      end
      self.dependencies.uniq.sort
    end
  
    def sheet_reference(sheet_name,reference)
      sheet_name = $1 if sheet_name.to_s =~ /^(\d+)\.0+$/
      puts "Warning, #{formula_cell} refers to an external workbook in '#{sheet_name}'" if sheet_name =~ /^\[\d+\]/
      using_worksheet(SheetNames.instance[sheet_name]) do
        reference.visit(self)
      end
    end
  
    alias quoted_sheet_reference sheet_reference
  
    def named_reference(name)
      dependencies_for reference_for_name(name)
    end
  
    def dependencies_for(full_reference)
      return [] unless full_reference  =~ /^(sheet\d+)\.(.*)$/
      sheet_name, reference = $1, $2
      using_worksheet(sheet_name) do
        case reference
        when /^a\('(.*?)','(.*?)'\)$/; area($1,$2)
        else; self.dependencies << full_reference
        end
      end    
    end
  
    def table_reference(table_name,structured_reference)
      dependencies_for Table.reference_for(table_name,structured_reference,formula_cell && formula_cell.reference).to_s
    end
  
    def local_table_reference(structured_reference)
      dependencies_for Table.reference_for_local_reference(formula_cell,structured_reference).to_s
    end
  
    def cell(reference)
      self.dependencies << Reference.new(reference,worksheet).to_ruby(true)
      Reference.new(reference).to_ruby
    end
  
    def area(start_area,end_area)
      self.dependencies.concat Area.new(worksheet,Reference.new(start_area).to_ruby,Reference.new(end_area).to_ruby).dependencies
    end
  
    def function(name,*args)
      if name == "INDIRECT"
        args.first.visit(self)
        reference_for_indirect = FormulaBuilder.new(formula_cell).indirect_function(args.first)
        # puts "INDIRECT REFERENCE: #{[args.first.inspect, reference_for_indirect]}" # if reference_for_indirect.to_s.start_with?(":")
        d = dependencies_for(reference_for_indirect)
      else
        args.each { |a| a.visit(self) }
      end
    end
  
    def reference_for_name(name)
      worksheet.named_references[name.downcase] || 
      workbook.named_references[name.downcase] || 
      (raise Exception.new("#{name} in #{formula_cell} not found"))
    end
  
    def workbook
      formula_cell.worksheet.workbook
    end
  
    def using_worksheet(sheet_name)
      original_worksheet = self.worksheet
      self.worksheet = workbook.worksheets[sheet_name] || (raise Exception.new("#{sheet_name} in #{formula_cell} not found"))
      yield
      self.worksheet = original_worksheet
    end
  
  end
end