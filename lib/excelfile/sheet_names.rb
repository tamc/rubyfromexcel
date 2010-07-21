# FIXME: This is obviously bad programming. What if we multithread?!
module RubyFromExcel
  class SheetNames < Hash
    include Singleton
  end
end