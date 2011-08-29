# FIXME: This is obviously bad programming. What if we multithread?!
module RubyFromExcel
  class SharedStrings < Array
    include Singleton
  
    def self.shared_string_for(index)
      self.instance.shared_string_for(index)
    end
  
    def load_strings_from_xml(xml)
      xml.css("si").each do |si|
        push si.css("t").map(&:content).join
        RubyFromExcel.debug(:shared_strings,"#{self.size-1}: #{self.last.inspect}")
      end
    end
  
    def shared_string_for(index)
      at(index.to_i)
    end
  
  end
end