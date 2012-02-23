module TerminalNode
  def to_method_name
    self.gsub(/([a-z])([A-Z])/,'\1_\2').downcase.gsub(/[^\p{word}]/,'_')
  end  
end

module RubyFromExcel
  class ExcelFunctionNotImplementedError < Exception
  end

  class DependsOnCalculatedFormulaError < Exception
  end

  class FunctionCompiler
    include RubyFromExcel::ExcelFunctions
    attr_accessor :worksheet
  
    def initialize(worksheet)
      self.worksheet = worksheet
    end
  
    def method_missing(method,*arguments, &block)
      return super unless arguments.empty?
      return super unless block == nil
      return find_or_create_worksheet(method.to_s) if method.to_s =~ /sheet\d+/
      return super unless method.to_s =~ /[a-z]+\d+/
      cell = worksheet.cell(method.to_s)
      return 0.0.extend(Empty) unless cell
      raise DependsOnCalculatedFormulaError.new unless cell.can_be_replaced_with_value?
      cell.value_for_including
    end
  
    def find_or_create_worksheet(worksheet_name)
      @worksheets ||= {}
      return @worksheets[worksheet_name] if @worksheets.has_key?(worksheet_name)
      new_worksheet = FunctionCompiler.new(worksheet.workbooks.worksheets[worksheet_name])
      @worksheets[worksheet_name] = new_worksheet
      new_worksheet
    end
  
    def to_s
      worksheet.to_s
    end
  
  end

  class FormulaBuilder
  
    attr_accessor :formula_cell
  
    def initialize(formula_cell = nil)
      self.formula_cell = formula_cell
    end
  
    def formula(*expressions)
      expressions.map { |e| e.visit(self) }.join
    end
  
    def number(number_as_text)
      number_as_text.to_f
    end
  
    def percentage(percentage_as_text)
      (percentage_as_text.to_f/100).to_s
    end
  
    def brackets(*expression)
      "(#{expression.map{ |e| e.visit(self)}.join})"
    end
  
    def named_reference(name, worksheet = nil)
      worksheet ||= formula_cell ? formula_cell.worksheet : nil
      return ":name" unless worksheet
      worksheet.named_references[name.to_method_name] || 
      worksheet.workbook.named_references[name.to_method_name] ||
      ":name"
    end
  
    def cell(reference)
      Reference.new(reference).to_ruby
    end
  
    def area(start_area,end_area)
      "a('#{cell(start_area)}','#{cell(end_area)}')"
    end

    def column_range(start_area,end_area)
      "c('#{cell(start_area)}','#{cell(end_area)}')"
    end

    def row_range(start_area,end_area)
      "r(#{cell(start_area)},#{cell(end_area)})"
    end
  
    def external_reference(external_reference_number,remainder_of_reference)
      puts "Warning, external references not supported (#{formula_cell}) #{remainder_of_reference}"
      remainder_of_reference.visit(self)
    end
  
    def sheet_reference(sheet_name,reference)
      sheet_name = $1 if sheet_name.to_s =~ /^(\d+)\.0+$/
      if sheet_name =~ /^\[\d+\]/
        puts "Warning, #{formula_cell} refers to an external workbook in '#{sheet_name}'"
        sheet_name.gsub!(/^\[\d+\]/,'')
      end
      if reference.type == :named_reference
        return ":name" unless formula_cell
        worksheet = formula_cell.worksheet.workbook.worksheets[SheetNames.instance[sheet_name]]
        # raise Exception.new("#{sheet_name.inspect} not found in #{SheetNames.instance} and therefore in #{formula_cell.worksheet.workbook.worksheets.keys}") unless worksheet
        return ":ref" unless worksheet
        named_reference(reference.first,worksheet)
      else
        "#{SheetNames.instance[sheet_name]}.#{reference.visit(self)}"
      end
    end
  
    def table_reference(table_name,structured_reference)
      Table.reference_for(table_name,structured_reference,formula_cell && formula_cell.reference).to_s
    end
  
    def local_table_reference(structured_reference)
      Table.reference_for_local_reference(formula_cell,structured_reference).to_s
    end
  
    alias :quoted_sheet_reference :sheet_reference
  
    OPERATOR_CONVERSIONS = { '^' => '**' }
  
    def operator(excel_operator)
      OPERATOR_CONVERSIONS[excel_operator] || excel_operator
    end
  
    def string_join(*strings)
      strings.map { |s| s.type == :string ? s.visit(self) : "(#{s.visit(self)}).to_s"}.join('+')
    end
  
    def string(string_text)
      string_text.gsub('""','"').inspect
    end
    
    def function(name,*args)
      raise ExcelFunctionNotImplementedError.new("#{name}(#{args})") unless self.respond_to?("#{name.downcase}_function")
      self.send("#{name.downcase}_function",*args)
    end
  
    def self.excel_function(name,name_to_use_in_ruby = name)
      define_method("#{name}_function") do |*args|
        standard_function name_to_use_in_ruby, args
      end
    end
  
    excel_function :and, :excel_and
    excel_function :average
    excel_function :count
    excel_function :counta
    excel_function :choose
    excel_function :abs
    excel_function :find
    excel_function :if, :excel_if
    excel_function :iferror
    excel_function :iserr
    excel_function :left
    excel_function :max
    excel_function :min
    excel_function :na
    excel_function :not, '!'
    excel_function :or, :excel_or
    excel_function :sum
    excel_function :sumif
    excel_function :sumifs
    excel_function :subtotal
    excel_function :sumproduct
    excel_function :round
    excel_function :roundup
    excel_function :rounddown
    excel_function :mod
    excel_function :pmt
    excel_function :npv
    excel_function :countif
    excel_function :text
  
    def standard_function(name_to_use_in_ruby,args)
      "#{name_to_use_in_ruby}(#{args.map {|a| a.visit(self) }.join(',')})"
    end
  
    def index_function(*args)
      attempt_to_calculate_index(*args) ||
      standard_function("index",args)
    end
  
    def match_function(*args)
      attempt_to_calculate_match(*args) ||
      standard_function("match",args)
    end
  
    def attempt_to_calculate_index(lookup_array,row_number,column_number = :ignore)
      lookup_array = range_for(lookup_array)
      row_number = single_value_for(row_number)
      column_number = single_value_for(column_number) unless column_number == :ignore
      return nil unless lookup_array
      return nil unless row_number
      return nil unless column_number
      if column_number == :ignore
        ref = FunctionCompiler.new(formula_cell.worksheet).calculate_index_formula(lookup_array,row_number,nil,:index_reference)
      else
        ref = FunctionCompiler.new(formula_cell.worksheet).calculate_index_formula(lookup_array,row_number,column_number,:index_reference)
      end
      return nil unless ref
      return nil if ref.is_a?(Symbol)
      return ref.to_ruby(true)
    rescue DependsOnCalculatedFormulaError
      return nil
    end
  
    def attempt_to_calculate_match(lookup_value,lookup_array,match_type = :ignore)
        lookup_value = single_value_for(lookup_value)
        lookup_array = range_for(lookup_array)
        match_type = single_value_for(match_type) unless match_type == :ignore
        return nil unless lookup_value
        return nil unless lookup_array
        return nil if match_type == nil
        result = nil
        if match_type == :ignore
          result = FunctionCompiler.new(formula_cell.worksheet).match(lookup_value,lookup_array).to_f
        else
          result = FunctionCompiler.new(formula_cell.worksheet).match(lookup_value,lookup_array,match_type)
        end
        result.respond_to?(:to_f) ? result.to_f : result
      rescue DependsOnCalculatedFormulaError
        return nil
    end
  
    def single_value_for(ast)
      return nil unless ast.respond_to?(:visit)
      ast = ast.visit(self)
      return true if ast == "true"
      return false if ast == "false"
      return ast if ast.is_a?(Numeric)
      return ast if ast =~ /^[0-9.]+$/
      return $1 if ast =~ /^"([^"]*)"$/
      return nil unless formula_cell
      return nil unless formula_cell.worksheet
      return nil unless ast =~ /^(sheet\d+)?\.?([a-z]+\d+)$/
      cell = $1 ? formula_cell.worksheet.workbook.worksheets[$1].cell($2) : formula_cell.worksheet.cell($2)
      return nil unless cell
      return nil unless cell.can_be_replaced_with_value?
      cell.value_for_including
    end
  
    def range_for(ast)
      return nil unless ast.respond_to?(:visit)
      return nil unless formula_cell
      return nil unless formula_cell.worksheet
      ast = ast.visit(self)
      return nil unless ast =~ /^(sheet\d+)?\.?a\('([a-z]+\d+)','([a-z]+\d+)'\)$/
      worksheet = $1 ? formula_cell.worksheet.workbook.worksheets[$1] : formula_cell.worksheet
      FunctionCompiler.new(worksheet).a($2,$3)
    end
  
    def indirect_function(text_formula)
      attempt_to_parse_indirect(text_formula) || (formula_cell.worksheet.workbook.indirects_used = true; "indirect(#{text_formula.visit(self)},'#{formula_cell && formula_cell.reference}')")
    end
  
    def attempt_to_parse_indirect(text_formula)
      #puts "Attempting to parse indirect #{text_formula.inspect}"
      return parse_and_visit(text_formula.first) if text_formula.type == :string
      return nil unless text_formula.type == :string_join
      reformated_indirect = text_formula.map do |non_terminal|
        if non_terminal.respond_to?(:type)
          case non_terminal.type
          when :string, :number
            non_terminal
          when :cell
            cell = formula_cell.worksheet.cell(non_terminal.visit(self))
            if cell
              return nil unless cell.can_be_replaced_with_value?
              cell.value_for_including
            else
              ""
            end
          when :sheet_reference, :named_reference, :table_reference, :local_table_reference
            reference = non_terminal.visit(self)
            # puts reference
            return nil unless reference =~ /^(sheet\d+)\.([a-z]+\d+)$/
            cell = formula_cell.worksheet.workbook.worksheets[$1].cell($2)
            if cell
              return nil unless cell.can_be_replaced_with_value?
              cell.value_for_including
            else
              ""
            end
          else
            return nil
          end
        else
          non_terminal
        end
      end
      # puts "Reformatted indirect: #{reformated_indirect.join}"
      parse_and_visit(reformated_indirect.join)
    end
  
    def parse_and_visit(text)
      ast = Formula.parse(text)
      # p [text,ast]
      return ":name" unless ast
      ast.visit(self.class.new(formula_cell))
    end
  
    def comparison(left,comparator,right)
      "excel_comparison(#{left.visit(self)},\"#{comparator.visit(self)}\",#{right.visit(self)})"
    end
  
    def arithmetic(*strings)
      strings.map { |s| s.visit(self) }.join
    end
  
    def prefix(prefix,thing)
     "#{prefix.visit(self)}#{thing.visit(self)}"
    end
  
    COMPARATOR_CONVERSIONS = {'=' => '==' }
  
    def comparator(string)
      COMPARATOR_CONVERSIONS[string] || string
    end
  
    def boolean_true
      "true"
    end
  
    def boolean_false
      "false"
    end
  
    def null
      0.0
    end

  end
end