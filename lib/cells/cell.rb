module RubyFromExcel
  class Cell
  
    attr_accessor :worksheet, :reference, :dependencies, :xml_value, :xml_type
  
    def initialize(worksheet,xml = nil)
      self.worksheet = worksheet
      parse_xml xml
      debug
    end
  
    def parse_xml(xml)
      return nil unless xml
      self.reference = Reference.new(xml['r'],worksheet)
      self.xml_value = xml.at_css("v").content
      self.xml_type = xml['t']
    end
  
    def alter_other_cells_if_required
      # nil
    end
  
    def work_out_dependencies
      self.dependencies = []
    end
    
    def to_ruby(r = RubyScriptWriter.new)
      r.put_simple_method reference.to_ruby, ruby_value
      r.to_s
    end

    def to_test(r = RubySpecWriter.new)
      r.put_spec "cell #{reference.to_ruby} should equal #{value}" do
        r.puts "#{worksheet}.#{reference.to_ruby}.should #{test}"
      end
      r.to_s
    end
  
    def test
      return "== #{value}" if xml_type
      "be_within(#{tolerance_for(value.to_f)}).of(#{value.to_f})"
    end
  
    def tolerance_for(value,tolerance = 0.1, maximum = 1e-8)
      tolerance = value.abs * tolerance
      tolerance = maximum if tolerance < maximum
      tolerance
    end
  
    def can_be_replaced_with_value?
      false
    end
  
    def value
      value = xml_value
      case xml_type
      when 'str'
        return value.inspect
      when 's'
        return SharedStrings.shared_string_for(value.to_i).inspect
      when 'e'
        return ":#{value.gsub(/[^a-zA-Z]/,'').downcase}"
      when 'b'
        return 'true' if value == '1'
        return 'false' if value == '0'
        return Formula.parse(value).visit(FormulaBuilder.new)
      else
        return Formula.parse(value).visit(FormulaBuilder.new)
      end
    end
  
    def value_for_including
      value = xml_value
      case xml_type
      when 'str'
        return value
      when 's'
        return SharedStrings.shared_string_for(value.to_i)
      when 'e'
        return ":#{value.gsub(/[^a-zA-Z]/,'').downcase}"
      when 'b'
        return 'true' if value == '1'
        return 'false' if value == '0'
        return Formula.parse(value).visit(FormulaBuilder.new)
      else
        return Formula.parse(value).visit(FormulaBuilder.new)
      end    
    end
  
    def to_s
      "#{worksheet}.#{reference}"
    end
  
    def inspect
      "(cell: #{to_s} with formula:#{respond_to?(:ast) ? ast.inspect : "na"} and value '#{value}' of type '#{xml_type}')"
    end
    
    def debug
      RubyFromExcel.debug(:cells,"#{worksheet.name}.#{reference} -> #{} -> #{xml_value} (#{xml_type})")
    end
    
  end
end
