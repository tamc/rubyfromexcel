# encoding: utf-8

module RubyFromExcel
  class Workbook
  
    attr_reader   :relationships
    attr_reader   :worksheets
    attr_reader   :worksheet_array
    attr_reader   :named_references
    attr_accessor :indirects_used
  
    def initialize(filename)
      @worksheets = {}
      @worksheet_array = []
      @named_references = {}
      @indirects_used = false
      @relationships = Relationships.for_file(filename)
      xml = File.open(filename) { |f| Nokogiri::XML(f) }.root
      load_shared_strings
      puts "\nLoaded shared strings"
      load_worksheets_from xml
      puts "Loaded #{worksheets.size} worksheets with #{total_cells} cells in total."
      work_out_named_references_from(xml)
      puts "Loaded named references"
      GC.start
    end
  
    def load_shared_strings
      return unless relationships.shared_strings
      SharedStrings.instance.load_strings_from_xml(File.open(relationships.shared_strings) { |f| Nokogiri::XML(f) }.root)    
    end
  
    def load_worksheets_from(xml)
      xml.css("sheet").each do |s|
        worksheet_filename = relationships[s['id']]
        worksheet = Worksheet.from_file(worksheet_filename,self)
        worksheets[worksheet.name] = worksheet
        SheetNames.instance[s['name']] = worksheet.name
        RubyFromExcel.debug(:worksheet_names,"#{worksheet.name}: #{s['name'].inspect}")
        worksheet_array << worksheet
        puts "Loaded #{worksheet.name} with #{worksheet.cells.size} cells"
      end
    end
  
    def work_out_named_references_from(xml)
      xml.css('definedName').each do |defined_name_xml|
        reference_name = defined_name_xml['name'].downcase # .gsub(/[^a-z0-9_]/,'_')
        reference_value = defined_name_xml.content
        if reference_value.start_with?('[')
          puts "Sorry, #{reference_name} (#{reference_value}) has a link to an external workbook. Skipping."
          next
        end
        reference = Formula.parse(reference_value).visit(FormulaBuilder.new)
        if defined_name_xml["localSheetId"]
          worksheet_array[defined_name_xml["localSheetId"].to_i].named_references[reference_name] = reference
        else
          named_references[reference_name] = reference
        end
        RubyFromExcel.debug(:named_references,"#{defined_name_xml['name'].inspect} (#{defined_name_xml["localSheetId"]}) -> #{reference_name.inspect} (#{defined_name_xml["localSheetId"] ? worksheet_array[defined_name_xml["localSheetId"].to_i].name.inspect : ""}) -> #{reference_value.inspect}")
      end
    end
  
    def work_out_dependencies
      puts "Working out dependencies..."
      worksheets.each do |name,worksheet|
        puts "Working out dependencies for #{name}"
        worksheet.work_out_dependencies
      end
      puts "Finished working out dependencies"
    end
  
    def workbook_pruner
      @workbook_pruner ||= WorkbookPruner.new(self)
    end
  
    def prune_cells_not_needed_for_output_sheets(*output_sheets)
      workbook_pruner.prune_cells_not_needed_for_output_sheets(*output_sheets)
    end
  
    def convert_cells_to_values_when_independent_of_input_sheets(*input_sheets)
      workbook_pruner.convert_cells_to_values_when_independent_of_input_sheets(*input_sheets)
    end
  
    def cell(reference)
      reference =~ /^(sheet\d+)\.([a-z]+\d+)$/
      worksheets[$1].cell($2)
    end
  
    def to_ruby(r = RubyScriptWriter.new)
      r.put_coding
      r.puts "require 'rubyfromexcel'"
      r.puts
      r.put_class 'Spreadsheet' do
        r.puts 'include RubyFromExcel::ExcelFunctions'
        r.puts
        if indirects_used
          r.put_method "initialize" do
            r.puts "@worksheet_names = #{Hash[SheetNames.instance.sort]}"
            r.puts "@workbook_tables = #{Hash[Table.tables.sort]}"
          end
          named_references.each do |name,reference|
            r.put_simple_method name.downcase.gsub(/[^\p{word}]/,'_'), reference
          end
        end
      end
      r.puts 'Dir[File.join(File.dirname(__FILE__),"sheets/","sheet*.rb")].each {|f| Spreadsheet.autoload(File.basename(f,".rb").capitalize,f)}'
      r.to_s
    end
  
    def total_cells
      worksheets.inject(0) { |total,a| a.last.cells.size + total }
    end
  
  end
end