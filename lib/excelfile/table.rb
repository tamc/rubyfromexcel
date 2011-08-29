module RubyFromExcel
  
  class TableNotFoundException < Exception
  end

  class Table
  
    def self.reference_for(table_name,structured_reference,cell_making_the_reference = nil)
      # puts "'#{table_name.downcase}' not found in #{tables.keys}" unless tables.include?(table_name.downcase)
      return ":ref" unless tables.include?(table_name.downcase)
      tables[table_name.downcase].reference_for(structured_reference,cell_making_the_reference)
    end
  
    def self.reference_for_local_reference(cell,structured_reference)
      table = tables.find { |n,t| t.include?(cell) }
      return ":ref" unless table
      table.last.reference_for(structured_reference,cell.reference)
    end
  
    def self.tables
      @@tables ||= {}
    end
  
    def self.add(table)
      tables[table.name.downcase] = table
    end

    attr_reader :worksheet, :name, :reference, :column_name_array, :number_of_total_rows, :all, :data, :headers, :totals, :column_names  
  
    def self.from_xml(worksheet,xml)
      Table.new(worksheet,xml['displayName'],xml['ref'],xml.css('tableColumn').map {|c| c['name']}.to_a,xml['totalsRowCount'])
    end
  
    def initialize(worksheet,name,reference,column_name_array,number_of_total_rows)
      @worksheet = worksheet
      @name = name
      @reference = reference
      @column_name_array = column_name_array
      @number_of_total_rows = number_of_total_rows
      @all = Area.new(worksheet,*reference.split(':'))
      index_of_first_total_row = -number_of_total_rows.to_i
      @data = @all.rows(1,index_of_first_total_row - 1)
      @headers = @all.row(0)
      @totals = @all.rows(index_of_first_total_row,-1)
      @column_names = {}
      column_name_array.each_with_index do |name,index|
        @column_names[name.strip.downcase] = index # column['id'].to_i
      end
      Table.add(self)
      RubyFromExcel.debug(:tables,"#{worksheet.name}.#{name} -> Table #{reference.inspect},#{column_name_array.inspect},#{number_of_total_rows}")
    end
  
    def column(name)
      name = $1 if name.to_s =~ /^(\d+)\.0+$/
      return :na unless column = column_names[name.to_s.downcase]
      data.column(column)
    end
  
    def reference_for(structured_reference,cell_making_the_reference = nil)
      structured_reference.strip!
      return this_row(cell_making_the_reference) if structured_reference == '#This Row'
      if structured_reference =~ /\[(.*?)\],\[(.*?)\]:\[(.*?)\]/
        return this_row_area_intersection(cell_making_the_reference,area_for_simple_reference($2),area_for_simple_reference($3)) if $1 == '#This Row'
        # return all.column(column_names[$2]) if $1 == '#All'
        # return intersection(area_for_simple_reference($1),area_for_simple_reference($2))
      elsif structured_reference =~ /\[(.*?)\],\[(.*?)\]/
        return this_row_intersection(cell_making_the_reference,area_for_simple_reference($2)) if $1 == '#This Row'
        return all.column(column_names[$2.to_s.downcase]) if $1 == '#All'
        return intersection(area_for_simple_reference($1),area_for_simple_reference($2))
      end
      if cell_making_the_reference
        if cell_making_the_reference.worksheet == self.worksheet
          if data.include?(cell_making_the_reference)
            case structured_reference
            when /#totals/i, /#headers/i
              return this_column_intersection(cell_making_the_reference,area_for_simple_reference(structured_reference))
            when /^#/i
              return area_for_simple_reference(structured_reference) 
            else
              return this_row_intersection(cell_making_the_reference,area_for_simple_reference(structured_reference))
            end
          end
        end
      end
      return area_for_simple_reference(structured_reference) 
    end
  
    def inspect
      %Q{'Table.new(#{worksheet},#{name.inspect},#{reference.inspect},#{column_name_array.inspect},#{number_of_total_rows})'}
    end  
  
    def include?(cell)
      @all.include_cell?(cell)
    end
  
    private
  
    def this_row(cell_making_the_reference)
      data.row(cell_making_the_reference.row_number - data.start_cell.row_number)
    end
  
    def this_row_intersection(cell_making_the_reference,column_area)
      "#{worksheet}.#{Reference.ruby_for(column_area.start_cell.column_number,cell_making_the_reference.row_number)}"
    end
    
    def this_row_area_intersection(cell_making_the_reference,first_column_area,last_column_area)
      row = cell_making_the_reference.row_number
      first_column = Reference.ruby_for(first_column_area.start_cell.column_number,row)
      last_column = Reference.ruby_for(last_column_area.start_cell.column_number,row)
      "#{worksheet}.a('#{first_column}','#{last_column}')"
    end
  
    def this_column_intersection(cell_making_the_reference,row_area)
  "#{worksheet}.#{Reference.ruby_for(cell_making_the_reference.column_number,row_area.start_cell.row_number)}"
    end
  
    def intersection(row_area,colum_area)
      row_area, colum_area = colum_area, row_area unless row_area.single_row?
      "#{worksheet}.#{Reference.ruby_for(colum_area.start_cell.column_number,row_area.start_cell.row_number)}"
    end
  
    def area_for_simple_reference(structured_reference)
      structured_reference.strip!
      return self.send(structured_reference.downcase[1..-1]) if structured_reference.start_with?('#')
      column(structured_reference)
    end
  
  end
end