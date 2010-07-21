module RubyFromExcel
  class ExcelMatrixCollection
    include Enumerable
  
    attr_accessor :matrices, :max_rows, :max_columns
  
    def initialize(*args)
      self.matrices = args.map do |a|
        a.respond_to?(:to_excel_matrix) ? a.to_excel_matrix : ExcelMatrix.new(a)
      end
      coerce_to_equal_sizes
    end
  
    def coerce_to_equal_sizes
      self.max_rows = matrices.max_by(&:rows).rows
      self.max_columns = matrices.max_by(&:columns).columns
      matrices.each do |m|
        m.add_columns!(max_columns - m.columns)
        m.add_rows!(max_rows - m.rows)
      end
    end
  
    def each &block
      block ? to_enum.each { |e| yield e } : to_enum
    end
  
    def to_enum
      Enumerator.new do |yielder|
        0.upto(max_rows-1) do |j|
          0.upto(max_columns-1) do |i|
            yielder << matrices.map { |m| m.values[j][i] }
          end
        end
      end
    end
  
    def matrix_map
      results = Array.new(max_rows) { Array.new(max_columns) }
      0.upto(max_rows-1) do |j|
        0.upto(max_columns-1) do |i|
          results[j][i] =  yield *matrices.map { |m| m.values[j][i] }
        end
      end
      em = ExcelMatrix.new(results)
      if em.rows == 1 && em.columns == 1
        return em.values.first.first
      else
        return em
      end
    end
  
  end


  class ExcelMatrix < ExcelRange
    include Enumerable
  
    attr_accessor :values
  
    def initialize(values)
      @values = values.is_a?(Array) ? (values.first.is_a?(Array) ? values : [values] ) : [[values]]
    end
  
    # origin 0,0
    def array_formula_offset(row_index,column_index)
      row_index = 0 if values.size == 1
      column_index = 0 if values[row_index].size == 1
      values[row_index][column_index] || :na
    end
  
    def rows
      return 0 unless values
      values.size
    end
  
    def columns
      return 0 unless values && values.first
      values.first.size
    end
  
    def add_rows!(number = 1)
      return if number < 1
      number.times { self.values << self.values.first }
    end
  
    def add_columns!(number = 1)
      return if number < 1
      number.times do 
        values.each do |row|
          row << row.last
        end
      end
    end
  
    def each &block
      block ? to_enum.each { |e| yield e } : to_enum
    end
  
    def to_enum
      Enumerator.new do |yielder|
        values.each do |row|
          row.each do |value|
            yielder << value
          end
        end
      end
    end
  
    def to_excel_matrix
      self
    end
  
  end
end