module RubyFromExcel
  class Worksheet
  
    def self.from_file(filename)
      xml = File.open(filename) { |f| Nokogiri::XML(f).root }
      worksheet = Worksheet.new(xml)
      relationships = Relationships.for_file(filename)
      xml.css('tablePart').each do |table_reference_xml|
        table_xml = File.open(relationships[table_reference_xml['id']]) {|f| Nokogiri::XML(f).root }
        Table.from_xml(worksheet,table_xml)
      end
      worksheet
    end
  
    attr_accessor :cells
    attr_accessor :name
    attr_reader   :named_references
    attr_accessor :workbook
  
    def initialize(xml)
      self.cells = {}
      @named_references = {}
      load_cells_from xml
      GC.start
    end

    def load_cells_from(xml)
      xml.css("c").each do |cell_xml|
         new_cell = create_cell_from cell_xml
         next unless new_cell
         self.cells[new_cell.reference.to_s] = new_cell
      end
      xml = nil
      let_cells_alter_other_cells_if_required
    end
  
    def create_cell_from(xml)
      formula = xml.at_css("f")
      if formula
        formula_type = formula['t']
        ref = formula['ref']
        return SimpleFormulaCell.new(self,xml) unless formula_type
        return SharingFormulaCell.new(self,xml) if formula_type == 'shared' && ref
        return SharedFormulaCell.new(self,xml) if formula_type == 'shared'
        return ArrayingFormulaCell.new(self,xml) if formula_type == 'array' && ref =~ /:/
        return SingleCellArrayFormulaCell.new(self,xml) if formula_type == 'array'
      end
      return ValueCell.new(self,xml) if xml.at_css("v")
      nil
    end
  
    def let_cells_alter_other_cells_if_required
      cells.each do |reference,cell|
        cell.alter_other_cells_if_required
      end
    end
  
    def work_out_dependencies
      cells.each do |reference,cell|
        cell.work_out_dependencies
      end
    end
  
    def cell(reference)
      cells[reference]
    end
  
    def replace_cell(reference,new_cell)
      cells[reference] = new_cell
    end
  
    def to_ruby(r = RubyScriptWriter.new)
      r.put_coding
      r.comment SheetNames.instance.key(variable_name)
      r.put_class class_name, 'Spreadsheet' do
        cells.each do |reference,cell|
          begin
            cell.to_ruby(r)
          rescue Exception => e
            puts "Error in #{cell.inspect}"
            raise
          end
        end
        if workbook.indirects_used
          named_references.each do |name,reference|
            r.put_simple_method name, reference
          end
        end
      end
      r.to_s
    end
  
    def to_test(r = RubySpecWriter.new)
      r.put_coding
      r.puts "require_relative '../spreadsheet'"
      r.comment SheetNames.instance.key(variable_name)
      r.put_description "'#{class_name}'" do
        r.put_simple_method variable_name, "$spreadsheet ||= Spreadsheet.new; $spreadsheet.#{variable_name}"
        r.puts
        cells.each do |reference,cell|
          cell.to_test(r)
        end
      end
      r.to_s
    end
  
    def class_name
      name.gsub(/\/(.?)/) { "::#{$1.upcase}" }.gsub(/(?:^|_)(.)/) { $1.upcase }
    end
  
    def variable_name
      name.gsub(/([a-z])([A-Z])/,'\1_\2').downcase.gsub(/[^a-z0-9_]/,'_')
    end
  
    alias to_s variable_name
  
    def worksheet_file_name
      variable_name
    end
  
  end
end