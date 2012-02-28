Gem::Specification.new do |s|
  s.name = "rubyfromexcel"
  s.add_dependency('nokogiri','>= 1.4.1')
  s.add_dependency('rubyscriptwriter','>= 0.0.1')
  s.add_dependency('rubypeg','>= 0.0.4')
  s.required_ruby_version = "~>1.9.1"
  s.version = '0.0.23'
  s.author = "Thomas Counsell, Green on Black Ltd"
  s.email = "ruby-from-excel@greenonblack.com"
  s.homepage = "http://github.com/tamc/rubyfromexcel"
  s.platform = Gem::Platform::RUBY
  s.summary = "Converts .xlxs files into pure ruby 1.9 code so that they can be executed without excel"
  s.files = ["LICENSE", "README", "{spec,lib,bin,doc,examples}/**/*"].map{|p| Dir[p]}.flatten
  s.executables = ["rubyfromexcel"]
  s.require_path = "lib"
  s.has_rdoc = false
end
