module RubyFromExcel
  class Relationships < Hash
  
    attr_reader :shared_strings
  
    def self.for_file(filename)
      root_directory = File.dirname(filename)
      relationships_xml_file = File.join(root_directory,'_rels',"#{File.basename(filename)}.rels")
      return Relationships.new unless File.exist?(relationships_xml_file)
      xml = File.open(relationships_xml_file) { |f| Nokogiri::XML(f) }
      Relationships.new(xml,root_directory)
    end
  
    def initialize(xml = nil, root_directory = nil)
      return unless xml
      xml.css("Relationship").each do |relationship|
        filename = root_directory ? File.expand_path(File.join(root_directory,relationship['Target'])) : relationship['Target']
        self[relationship['Id']] = filename
        @shared_strings = filename if relationship['Type'] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
      end
    end
  
  end
end