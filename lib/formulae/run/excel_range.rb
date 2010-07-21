module RubyFromExcel

  class ExcelRange
  end

  class ExcelRangeCommon < ExcelRange
    include Enumerable
    attr_accessor :sheet
  
    def each &block
      block ? to_enum.each { |e| yield e } : to_enum
    end
  
    private
  
    def defined_cells(regexp = /^[a-z]{1,3}\d+$/io)
      sheet.methods.find_all {|m| m =~ regexp }
    end
  
    def value_at(reference)
      sheet.send(reference.to_s)
    rescue ZeroDivisionError => e
      return :div0
    end
  end

  class Area < ExcelRangeCommon
    attr_accessor :start_cell, :end_cell
  
    def initialize(sheet, start_cell,end_cell)
      self.sheet = sheet
      self.start_cell = Reference.new(start_cell)
      self.end_cell = Reference.new(end_cell)
    end
  
    def row(index) # index = 0 for top row
      row_number = row_number_for_index(index)
      Area.new(sheet,"#{start_cell.column}#{row_number}","#{end_cell.column}#{row_number}")
    end
  
    def rows(start_index, end_index)
      start_row_number = row_number_for_index(start_index)
      end_row_number = row_number_for_index(end_index)
      Area.new(sheet,"#{start_cell.column}#{start_row_number}","#{end_cell.column}#{end_row_number}")
    end
  
    def column(index) # index = 0 for first column
      column_number = start_cell.column_number + index
      Area.new(sheet,Reference.ruby_for(column_number,start_cell.row_number),Reference.ruby_for(column_number,end_cell.row_number))
    end
  
    def index(row_number,column_number) #origin at 1,1
      reference = index_reference(row_number,column_number)
      return reference if reference.is_a?(Symbol)
      value_at(reference.to_ruby)
    end
  
    def index_reference(row_number,column_number)
      ref = array_formula_index(row_number-1, column_number-1)
      return ref if ref.is_a?(Symbol)
      ref.worksheet = sheet
      ref
    end
  
    def array_formula_offset(row_index,column_index)
      value_at array_formula_index(row_index,column_index).to_ruby
    end
  
    def array_formula_index(row_index,column_index)
      column_index = 0 if single_column?
      row_index = 0 if single_row?
      reference = start_cell.shift([row_index,column_index])
      return :ref unless include?(reference)
      reference
    end
  
    def each(&block)
      return @array.each(&block) if @array
      start_cell.row_number.upto(end_cell.row_number) do |row|
        start_cell.column_number.upto(end_cell.column_number) do |column_number|
          yield value_at(Reference.ruby_for(column_number,row))
        end
      end    
    end
  
    # def to_enum
    #   to_reference_enum.map {|reference| value_at reference}
    # end
  
    def to_a
      @array ||= self.map {|v| v }
    end
  
    def dependencies
      @dependencies ||=
      (start_cell.row_number..end_cell.row_number).map do |row|
        (start_cell.column_number..end_cell.column_number).map do |column_number|
          Reference.ruby_for(column_number,row,sheet)
        end
      end.flatten
    end
  
    def to_reference_enum
      Enumerator.new do |yielder|
        start_cell.row_number.upto(end_cell.row_number) do |row|
          start_cell.column_number.upto(end_cell.column_number) do |column_number|
            yielder << Reference.ruby_for(column_number,row)
          end
        end
      end    
    end
  
    def to_excel_matrix
      ExcelMatrix.new(to_grid)
    end
  
    def to_grid
      (start_cell.row_number..end_cell.row_number).map do |row|
        (start_cell.column_number..end_cell.column_number).map do |column_number|
          value_at(Reference.ruby_for(column_number,row))
        end
      end
    end
  
    def single_row_or_column?
      single_row? || single_column?
    end
  
    def single_column?
      start_cell.column == end_cell.column
    end
  
    def single_row?
      start_cell.row_number == end_cell.row_number
    end
  
    def to_s
      "#{sheet}.a('#{start_cell.to_ruby}','#{end_cell.to_ruby}')"
    end
  
    def include_cell?(cell)
      return false unless cell.worksheet == sheet
      include?(cell.reference)
    end
  
    def row_number_for_index(index)
      return (end_cell.row_number + 1 + index) if index < 0
      row_number = start_cell.row_number + index
    end
  
    def include?(reference)
      return false unless reference.column_number >= start_cell.column_number
      return false unless reference.column_number <= end_cell.column_number
      return false unless reference.row_number >= start_cell.row_number
      return false unless reference.row_number <= end_cell.row_number
      true
    end
  
    def method_missing(method,*args,&block)
      super unless start_cell.to_s == end_cell.to_s
      super unless value_at(start_cell).respond_to?(method)
      value_at(start_cell).send(method,*args,&block)
    end
  
  end

  class Columns < ExcelRangeCommon
  
    attr_accessor :start_column_number, :end_column_number
  
    def initialize(sheet,start_column,end_column)
      self.sheet = sheet
      self.start_column_number = Reference.column_to_integer(start_column)
      self.end_column_number = Reference.column_to_integer(end_column)
    end

    def array_formula_index(row_index,column_index)
      column_index = 0 if single_column?
      reference = Reference.new(Reference.ruby_for(start_column_number+column_index,row_index+1))
      return :na unless include?(reference)
      reference.to_ruby
    end
  
    def to_enum
      Enumerator.new do |yielder|
        (start_column_number..end_column_number).each do |column_number|
          (1..maximum_row_in_column_number(column_number)).each do |row|
            yielder << value_at(Reference.ruby_for(column_number,row))
          end
        end
      end
    end
  
    private
  
    def include?(reference)
      return false unless reference.column_number >= start_column_number
      return false unless reference.column_number <= end_column_number
      true
    end
  
    def maximum_row_in_column_number(column_number)
      cells_in_column = defined_cells(/^#{Reference.integer_to_column(column_number)}\d+$/i)
      cells_in_column.empty? ? 0 : cells_in_column.last[/\d+/i].to_i
    end
  
    def single_column?
      start_column_number == end_column_number
    end
  end

  class Rows < ExcelRangeCommon
    attr_accessor :start_row_number, :end_row_number

    def initialize(sheet,start_row_number,end_row_number)
      self.sheet = sheet
      self.start_row_number = start_row_number.to_i
      self.end_row_number = end_row_number.to_i
    end

    def array_formula_index(row_index,column_index)
      row_index = 0 if single_row?
      reference = Reference.new(Reference.ruby_for(column_index+1,row_index+1))
      return :na unless include?(reference)
      reference.to_ruby
    end
  
    def to_enum
      Enumerator.new do |yielder|
        (start_row_number..end_row_number).each do |row|
          (1..maximum_column_number_in_row_number(row)).each do |column_number|
            yielder << value_at(Reference.ruby_for(column_number,row))
          end
        end
      end
    end
  
    private
  
    def maximum_column_number_in_row_number(row_number)
      cells_in_row = defined_cells(/^[a-z]{1,3}#{row_number}$/i)
      cells_in_row.empty? ? 0 : Reference.column_to_integer(cells_in_row.last[/[a-z]+/i])
    end
  
    def single_row?
      start_row_number == end_row_number
    end
  
    def include?(reference)
      return false unless reference.row_number >= start_row_number
      return false unless reference.row_number <= end_row_number
      true
    end
  
  end
end