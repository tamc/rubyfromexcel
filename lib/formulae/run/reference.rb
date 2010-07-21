module RubyFromExcel
  class Reference
    attr_accessor :row, :column, :worksheet
    attr_accessor :absolute_row, :absolute_column
  
    def self.ruby_for(column_number,row_number,worksheet = nil)
      worksheet ? "#{worksheet}.#{integer_to_column(column_number)}#{row_number.to_i}" : "#{integer_to_column(column_number)}#{row_number.to_i}"
    end
  
    def self.column_to_integer(string)
      @column_to_integer_cache ||= {}
      return @column_to_integer_cache[string] ||= string.downcase.each_byte.to_a.reverse.each.with_index.inject(0) do |memo,byte_with_index|
        memo = memo + ((byte_with_index.first - 96) * (26**byte_with_index.last))
      end
    end

    def self.integer_to_column(integer)
      @integer_to_column_cache ||= {}
      return @integer_to_column_cache[integer] ||= self.calculate_integer_to_column(integer)
    end
  
    def self.calculate_integer_to_column(integer)
      raise Exception.new("Column numbering starts at 1, received #{integer}") unless integer >= 1
      text = (integer-1).to_i.to_s(26)
      (text[0...-1].tr('1-9a-z','abcdefghijklmnopqrstuvwxyz')+text[-1,1].tr('0-9a-z','abcdefghijklmnopqrstuvwxyz')).gsub('a0','z').gsub(/([b-z])0/) { $1.tr('b-z','a-y')+"z" }
    end

    def initialize(reference_as_text, worksheet = nil)
      @worksheet = worksheet
      reference_as_text.downcase!
      @row = reference_as_text[/\d+/]
      @column = reference_as_text[/[a-z]+/]
      @absolute_row = (reference_as_text =~ /\$\d/)
      @absolute_column = (reference_as_text =~ /\$[a-z]/)
    end
  
    def shift(offset_array)
      self.dup.shift!(offset_array)
    end
  
    def shift!(offset_array)
      return self unless offset_array && offset_array.size == 2
      shift_row_by! offset_array.first
      shift_column_by! offset_array.last
      self
    end
  
    def to_ruby(include_worksheet = false)
       include_worksheet ? "#{worksheet}.#{column}#{row && row.to_i}".downcase  : "#{column}#{row && row.to_i}".downcase
    end
  
    alias to_s to_ruby
  
    def row_number
      row.to_i
    end
  
    def column_number
      Reference.column_to_integer(column)
    end

    def -(other_reference)
      [row_number - other_reference.row_number, column_number - other_reference.column_number]
    end

    private

    def shift_row_by!(offset)
      return if absolute_row
      self.row = (row_number + offset).to_s
    end

    def shift_column_by!(offset)
      return if absolute_column
      self.column = Reference.integer_to_column(column_number + offset)
    end

  end
end