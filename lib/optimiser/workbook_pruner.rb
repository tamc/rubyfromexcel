module RubyFromExcel

  class WorkbookPruner
    attr_reader :workbook, :cells_to_keep, :output_sheet_names
  
    def initialize(workbook)
      @workbook = workbook
    end
  
    def prune_cells_not_needed_for_output_sheets(*output_sheet_names)
      @cells_to_keep = {}
      workbook.work_out_dependencies
      @output_sheet_names = output_sheet_names
      find_dependencies_of output_sheet_names
      delete_cells_that_we_dont_want_to_keep
      cells_to_keep = nil # So that we can garbage collect
      GC.start
    end

    def find_dependencies_of(output_sheet_names)
      output_sheet_names.each do |output_sheet_name|
        puts "Cascading dependencies for #{output_sheet_name}"
        output_sheet = workbook.worksheets[SheetNames.instance[output_sheet_name]] 
        raise Exception.new("#{output_sheet_name} not found #{SheetNames.instance.inspect} #{workbook.worksheets.keys.inspect}") unless output_sheet
        output_sheet.cells.each do |reference,cell|
          keep_dependencies_for(cell)
        end
      end
      puts "#{cells_to_keep.size} cells kept, #{workbook.total_cells - cells_to_keep.size} pruned"
    end
  
    def keep_dependencies_for(cell)
      return unless cell
      return if cells_to_keep.has_key?(cell.to_s)
      cells_to_keep[cell.to_s] = cell
      cell.dependencies.each do |reference|
        keep_dependencies_for workbook.cell(reference)
      end
    end
  
    def delete_cells_that_we_dont_want_to_keep
      workbook.worksheets.each do |name,sheet|
        sheet.cells.delete_if do |reference,cell|
          !cells_to_keep.has_key?(cell.to_s)
        end
      end
    end
  
    def convert_cells_to_values_when_independent_of_input_sheets(*input_sheet_names)
      input_sheet_names = input_sheet_names.map { |name| SheetNames.instance[name] }
      count = 0
      workbook.worksheets.each do |name,sheet|
        puts "Converting cells into values in #{name}"
        sheet.cells.each do |reference,cell|
          next if cell.is_a?(ValueCell)
          unless depends_on_input_sheets?(cell,input_sheet_names)
            # puts "converting #{cell}"
            count = count + 1
            sheet.replace_cell(reference,ValueCell.for(cell))
          end
        end
      end
      puts "#{count} formula cells replaced with their values"
      GC.start
      prune_cells_not_needed_for_output_sheets(*output_sheet_names)
    end
  
    def depends_on_input_sheets?(cell,input_sheet_names,stack_level = 0)
      return true if stack_level > 100
      return false unless cell
      return true if input_sheet_names.include?(cell.worksheet.variable_name)
      return false unless cell.dependencies
      return true if cell.dependencies.any? do |reference|
        depends_on_input_sheets?(workbook.cell(reference),input_sheet_names,stack_level + 1)
      end
      return false
    end
  
  end
end