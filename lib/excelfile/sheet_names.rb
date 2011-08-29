# FIXME: This is obviously bad programming. What if we multithread?!
module RubyFromExcel
  class SheetNames < Hash
    include Singleton
    
    def marshal_dump
      map { |a,b| [a,b] }
    end

    def marshal_load array
      array.each do |sheet|
        self[sheet.first] = sheet.last
      end
    end
    
  end
end