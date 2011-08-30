class Numeric
  def number_like?; true; end
  def array_formula_offset(i,j); self; end
end

class String
  def +@; number_like? ? self.to_f : self; end
  def -@; number_like? ? -self.to_f : self; end
  alias old_plus +
  alias old_multiply *
  def +(other) (other.number_like? && number_like?) ? self.to_f + other : self.old_plus(other) end
  def -(other) (other.number_like? && number_like?) ? self.to_f - other : :value end
  def *(other) (other.number_like? && number_like?) ? self.to_f * other : self.old_multiply(other) end
  def /(other) (other.number_like? && number_like?) ? self.to_f / other : :value end
  def coerce(other); number_like? ? self.to_f.coerce(other) : [other,:value]; end
  def number_like?
    self =~ /^\d+\.?\d*$/
  end
end

class Symbol
  def +@; self; end
  def -@; self; end
  def coerce(other)
    return [other,self] if other.is_a?(Symbol)
    [self,self]
  end
  def +(other) self end
  def -(other) self end
  def *(other) self end
  def /(other) self end
  def **(other) self end
  def to_ruby(include_worksheet = false); self; end
end


class TrueClass
  def +(other); 1.0 + other; end
  def -(other); 1.0 - other; end
  def *(other); 1.0 * other; end
  def /(other); 1.0 / other; end
  def coerce(other); 1.0.coerce(other); end
end

class FalseClass
  def +(other); 0.0 + other; end
  def -(other); 0.0 - other; end
  def *(other); 0.0 * other; end
  def /(other); 0.0 / other; end
  def coerce(other); 0.0.coerce(other); end
end

module RubyFromExcel
  
  module Empty
    # Used to indicate empty cells
  end

  module ExcelFunctions
  
    def result_cache
      @result_cache ||= {}
    end
  
    def formula_cache
      @formula_cache ||= {}
    end
  
    def recalculate
      formula_cache.clear
      result_cache.clear
    end
    
    def set(cell,value)
      instance_variable_name = "@#{cell}"
      unless instance_variable_defined?(instance_variable_name)
        self.class.class_eval do
          if method_defined?(cell)
            alias_method "old_#{cell}", cell
            define_method(cell) do
              instance_variable_get(instance_variable_name) || self.send("old_#{cell}")
            end
          end
        end
      end
      instance_variable_set(instance_variable_name, value)
    end
    
    def method_missing(method,*arguments, &block)
      return super unless arguments.empty?
      return super unless block == nil
      return find_or_create_worksheet(method.to_s) if method.to_s =~ /sheet\d+/
      return super unless method.to_s =~ /[a-z]+\d+/
      0.0.extend(Empty)
    end
  
    def find_or_create_worksheet(worksheet_name)
      @worksheets ||= {variable_name => self}
      return @worksheets[worksheet_name] if @worksheets.has_key?(worksheet_name)
      worksheet = self.class.class_eval("#{worksheet_name.capitalize}.new")
      worksheet.instance_variable_set("@worksheets",@worksheets)
      worksheet.instance_variable_set("@formula_cache",formula_cache)
      @worksheets[worksheet_name] = worksheet
      worksheet
    end
    
    def text(number,format)
      return number if iserr(number)
      return format if iserr(format)
      raise Exception.new("format #{format} not implemented") unless format.is_a?(Numeric)
      number.round(format).to_s
    end
    
    def round(number,decimal_places)
      return number if iserr(number)
      return decimal_places if iserr(decimal_places)
      return :na unless number.is_a?(Numeric)
      return :na unless decimal_places.is_a?(Numeric)
      number.round(decimal_places)
    end
    
    def roundup(number,decimal_places)
      return number if iserr(number)
      return decimal_places if iserr(decimal_places)
      return :na unless number.is_a?(Numeric)
      return :na unless decimal_places.is_a?(Numeric)
       (number * 10**decimal_places).ceil.to_f / 10**decimal_places
    end

    def rounddown(number,decimal_places)
      return number if iserr(number)
      return decimal_places if iserr(decimal_places)
      return :na unless number.is_a?(Numeric)
      return :na unless decimal_places.is_a?(Numeric)
       (number * 10**decimal_places).floor.to_f / 10**decimal_places
    end
    
    def mod(number,divisor)
      return number if iserr(number)
      return divisor if iserr(divisor)
      return :na unless number.is_a?(Numeric)
      return :na unless divisor.is_a?(Numeric)
      number % divisor
    end
    
    def pmt(rate,periods,principal)
      -principal*(rate*((1+rate)**periods))/(((1+rate)**periods)-1)
    end
    
    def npv(rate,*cashflows)
      discount_factor = 1
      flatten_and_inject(cashflows) do |pv,cashlfow|
        discount_factor = discount_factor * (1 + rate)
        cashlfow.is_a?(Numeric) ? pv + (cashlfow / discount_factor) : pv
      end
    end
            
    def sum(*args)
      flatten_and_inject(args) do |counter,arg|
        arg.is_a?(Numeric) ? counter + arg.to_f : counter
      end
    end
  
    def sumif(check_range,criteria,sum_range = check_range)
      sumifs(sum_range,check_range,criteria)
    end
  
    def sumifs(sum_range,*args)
      return :na if iserr(sum_range)
      return :na if args.any? { |c| iserr(c) }
      if sum_range.is_a?(ExcelRange)
        return :na unless sum_range.single_row_or_column?
        sum_range = sum_range.to_a
      else
        sum_range = [sum_range]
      end
      checks = Hash[*args].to_a.map do |check|
        check_range, check_value = check.first, check.last
        if check_range.is_a?(ExcelRange)
          return :na unless check_range.single_row_or_column?
          check_range = check_range.to_a
        else
          check_range = [check_range]
        end
        check_range = check_range.map { |c| c.to_s.downcase }
        check_value = check_value.to_s.downcase
        [check_range,check_value]
      end
      accumulator = 0
      sum_range.each_with_index do |potential_sum,i|
        next unless checks.all? do |c|
          if c.last =~ /^>([0-9.]+)$/
            c.first[i].to_f > $1.to_f
          elsif c.last =~ /^<([0-9.]+)$/
            c.first[i].to_f < $1.to_f
          else
            c.first[i] == c.last
          end
        end
        next unless potential_sum.is_a?(Numeric)
        accumulator = accumulator + potential_sum
      end
      accumulator
    end
  
    def sumproduct(*ranges)
      ranges.map do |range|
        return :na unless range.respond_to?(:to_a)
        range.to_a
      end.transpose.map do |values|
        values.inject(1) do |cell,total|
          total * cell
        end
      end.inject(0) do |product,total| 
        total + product
      end
    rescue IndexError
      return :value
    end
  
    def choose(choice,*choices)
      return choice if iserr(choice)
      choices[choice - 1]
    end
  
    def count(*args)
      flatten_and_inject(args) do |counter,arg|
        arg.is_a?(Numeric) && !arg.is_a?(Empty) ? counter + 1 : counter
      end
    end
    
    def countif(check_range,criteria,count_range = check_range)
      countifs(count_range,check_range,criteria)
    end
    
    def countifs(count_range,*args)
      return :na if iserr(count_range)
      return :na if args.any? { |c| iserr(c) }
      if count_range.is_a?(ExcelRange)
        return :na unless count_range.single_row_or_column?
        count_range = count_range.to_a
      else
        count_range = [count_range]
      end
      checks = Hash[*args].to_a.map do |check|
        check_range, check_value = check.first, check.last
        if check_range.is_a?(ExcelRange)
          return :na unless check_range.single_row_or_column?
          check_range = check_range.to_a
        else
          check_range = [check_range]
        end
        check_range = check_range.map { |c| c.to_s.downcase }
        check_value = check_value.to_s.downcase
        [check_range,check_value]
      end
      accumulator = 0
      count_range.each_with_index do |potential_count,i|
        next unless checks.all? do |c|
          if c.last =~ /^>([0-9.]+)$/
            c.first[i].to_f > $1.to_f
          elsif c.last =~ /^<([0-9.]+)$/
            c.first[i].to_f < $1.to_f
          else
            c.first[i] == c.last
          end
        end
        next unless potential_count.is_a?(Numeric) && !potential_count.is_a?(Empty)
        accumulator = accumulator + 1
      end
      accumulator
    end
  
    def counta(*args)
      flatten_and_inject(args) do |counter,arg|
        arg.is_a?(Empty) ? counter : counter + 1
      end
    end
  
    def max(*args)
      args = args.map { |arg| arg.respond_to?(:to_a) ? arg.to_a : arg }
      args.flatten!
      if (error = args.find { |arg| iserr(arg) }) 
        return error
      end
      args.delete_if { |arg| !arg.kind_of?(Numeric) }
      args.delete_if { |arg| arg.kind_of?(Empty) }
      args.max
    end
  
    def abs(value)
      return value if iserr(value)
      return :value unless value.respond_to?(:abs)
      value.abs
    end
  
    def min(*args)
      args = args.map { |arg| arg.respond_to?(:to_a) ? arg.to_a : arg }
      args.flatten!
      args.delete_if { |arg| !arg.kind_of?(Numeric) }
      args.delete_if { |arg| arg.kind_of?(Empty) }
      return error if error = args.find { |arg| iserr(arg) }
      args.min    
    end
  
    def flatten_and_inject(args,&block)
      args = args.map do |arg|
        return arg if iserr(arg)
        arg.respond_to?(:to_a) ? arg.to_a : arg
      end
      args.flatten!
      args.inject(0,&block)
    end
  
    def left(string,characters = 1)
      return string if iserr(string)
      return characters if iserr(characters)
      string.slice(0,characters)
    end
  
    def find(string_to_find,string_to_search,start_index = 1)
      return string_to_find if iserr(string_to_find)
      return string_to_search if iserr(string_to_search)
      return start_index if iserr(start_index)
      result = string_to_search.index(string_to_find,start_index - 1 )
      result ? result + 1 : :value
    end
  
    def average(*args)
      total = sum(*args)
      number = count(*args)
      return total if iserr(total)
      return number if iserr(number)
      total / number
    end
  
    def subtotal(type,*args)
      case type
      when 1.0, 101.0; average(*args)
      when 2.0, 102.0; count(*args)
      when 3.0, 103.0; counta(*args)
      when 9.0, 109.0; sum(*args)
      else 
        raise Exception.new("subtotal type #{type} on #{args} not implemented")
      end
    end
  
    def match(lookup_value,lookup_array,match_type = 1.0)
      formula_cache[[:match,lookup_value,lookup_array,match_type]] ||= calculate_match(lookup_value,lookup_array,match_type)
    end
  
    def calculate_match(lookup_value,lookup_array,match_type = 1.0)
      return lookup_value if iserr(lookup_value)
      return lookup_array if iserr(lookup_array)
      return match_type if iserr(match_type)
      lookup_value = lookup_value.downcase if lookup_value.respond_to?(:downcase)
      case match_type      
      when 0, 0.0, false
        lookup_array.each_with_index do |item,index|
          item = item.downcase if item.respond_to?(:downcase)
          return index+1 if lookup_value == item
        end
        return :na
      when 1, 1.0, true
        lookup_array.each_with_index do |item, index|
          next if lookup_value.is_a?(String) && !item.is_a?(String)
          next if lookup_value.is_a?(Numeric) && !item.is_a?(Numeric)
          item = item.downcase if item.respond_to?(:downcase)
          if item > lookup_value
            return :na if index == 0
            return index
          end
        end
        return lookup_array.to_a.size
      when -1, -1.0
        lookup_array.each_with_index do |item, index|
          next if lookup_value.is_a?(String) && !item.is_a?(String)
          next if lookup_value.is_a?(Numeric) && !item.is_a?(Numeric)
          item = item.downcase if item.respond_to?(:downcase)
          if item < lookup_value
            return :na if index == 0
            return index
          end
        end
        return lookup_array.to_a.size - 1
      end
      return :na
    end
  
    def index(area,row_number, column_number = nil)
      formula_cache[[:index,area,row_number,column_number]] ||= calculate_index_formula(area,row_number, column_number)
    end
  
    def calculate_index_formula(area,row_number, column_number = nil, method = :index)
      return area if iserr(area)
      return row_number if iserr(row_number)
      if area.single_row?
        return area.send(method,row_number,column_number) if column_number
        return area.send(method,1,row_number)
      elsif area.single_column?
        return area.send(method,row_number,column_number) if column_number
        return area.send(method,row_number,1)
      else
        return :ref unless row_number && column_number
        return area.column(column_number-1) if row_number == 0.0
        return area.row(row_number-1) if column_number == 0.0
        return area.send(method,row_number,column_number)
      end
    end
  
    def na(*arg)
      :na
    end
  
    def ref(*arg)
      :ref
    end
  
    def iserr(arg)
      return true if arg.is_a? Symbol
      return true if arg.respond_to?(:nan?) && arg.nan?
      return true if arg.respond_to?(:infinite?) && arg.infinite?
      false
    end
  
    def excel_if(condition,true_value,false_value = false)
      condition ? true_value : false_value
    end
    
    def excel_comparison(left,comparison,right)
      left = left.downcase if left.is_a?(String)
      right = right.downcase if right.is_a?(String)
      left.send(comparison,right)
    end
    
    def iferror(value,value_if_error)
      iserr(value) ? value_if_error : value
    rescue  ZeroDivisionError => e
      puts e
      return :div0
    end
  
    def excel_and(*args)
      args.all? {|a| a == true }
    end
  
    def excel_or(*args)
      args.any? { |a| a == true }
    end
  
    def a(start_cell,end_cell)
      result_cache[[:a,start_cell,end_cell]] ||= Area.new(self,start_cell,end_cell)
    end
  
    def r(start_row,end_row)
      result_cache[[:r,start_row,end_row]] ||= Rows.new(self,start_row,end_row)
    end
  
    def c(start_column_number,end_column_number)
      result_cache[[:r,start_column_number,end_column_number]] ||= Columns.new(self,start_column_number,end_column_number)
    end
  
    def s(full_name_of_worksheet)
      full_name_of_worksheet = $1 if full_name_of_worksheet.to_s =~ /^(\d+)\.0+$/
      self.send(@worksheet_names[full_name_of_worksheet])
    end
  
    def t(table_name)
      table = @workbook_tables[table_name]
      return :ref unless table
      return table if table.is_a?(Table)
      @workbook_tables[table_name] = eval(table)
    end
  
    def m(*potential_excel_matrices,&block)
      ExcelMatrixCollection.new(*potential_excel_matrices).matrix_map(&block)
    end
  
    def variable_name
      @variable_name ||= self.class.to_s.downcase
    end
  
    def indirect(reference_text,refering_cell = nil)
      parsed_reference = formula_cache[[:indirect,reference_text]] ||= Formula.parse(reference_text)
      reference = parsed_reference.visit(RuntimeFormulaBuilder.new(self,refering_cell && Reference.new(refering_cell)))
      formula_cache[[:indirect_result,reference]] ||= eval(reference)
    end
  
  end
end