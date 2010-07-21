module RubyFromExcel
  class ArrayFormulaBuilder < FormulaBuilder
  
    def sum_function(*args)
      "sum(#{args.map {|a| a.visit(self)}.join(',')})"
    end
  
    def sumif_function(check_range,criteria,sum_range = nil)
      if sum_range
        "m(#{criteria.visit(self)}) { |r1| sumif(#{check_range.visit(self)},r1,#{sum_range.visit(self)}) }"
      else
        "m(#{criteria.visit(self)}) { |r1| sumif(#{check_range.visit(self)},r1) }"
      end
    end
  
    def sumifs_function(sum_range,*args)
      checks = Hash[*args]
      "m(#{checks.values.map {|a| a.visit(self) }.join(',')}) { |#{checks.values.map.with_index {|a,i| "r#{i+1}"}.join(',')}| sumifs(#{sum_range.visit(self)},#{checks.map.with_index { |a,i| "#{a.first.visit(self)},r#{i+1}" }.join(',')}) }"
    end
  
    def standard_function(name_to_use_in_ruby,args)
      arg_names = (1..(args.size)).map { |i| "r#{i}" }
      "m(#{args.map {|a| a.visit(self) }.join(',')}) { |#{arg_names.join(',')}| #{name_to_use_in_ruby}(#{arg_names.join(',')}) }"
    end
  
    def index_function(index_range,*args)
      arg_names = (1..(args.size)).map { |i| "r#{i}" }
      "m(#{args.map {|a| a.visit(self) }.join(',')}) { |#{arg_names.join(',')}| index(#{index_range.visit(self)},#{arg_names.join(',')}) }"
    end
  
    def match_function(lookup_value,lookup_array,match_type = nil)
      if match_type
        "m(#{lookup_value.visit(self)},#{match_type.visit(self)}) { |r1,r2| match(r1,#{lookup_array.visit(self)},r2) }"
      else
        "m(#{lookup_value.visit(self)}) { |r1| match(r1,#{lookup_array.visit(self)}) }"
      end
    end
  
    def comparison(left,comparison,right)
      "m(#{left.visit(self)},#{right.visit(self)}) { |r1,r2| r1#{comparison.visit(self)}r2 }"
    end

    def arithmetic(*strings)
      matrix_strings = []
      matrix_string_references = []
      strings = strings.map do |s|
        if s.type == :operator
          s.visit(self)
        else
          matrix_strings << s.visit(self)
          matrix_string_references << "r#{matrix_string_references.size+1}"
          matrix_string_references.last
        end
      end
      "m(#{matrix_strings.join(',')}) { |#{matrix_string_references.join(',')}| #{strings.join} }"
    end
  end
end