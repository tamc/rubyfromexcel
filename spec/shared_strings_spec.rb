require_relative 'spec_helper'

describe SharedStrings do
  
  it "should return a shared string when provided with an index" do
    xml = Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>This a second shared string</t><phoneticPr fontId="3" type="noConversion"/></si><si><r><t>This is, h</t></r><r><rPr><b/><sz val="10"/><rFont val="Verdana"/></rPr><t>opefully, the first</t></r><r><rPr><sz val="10"/><rFont val="Verdana"/></rPr><t xml:space="preserve"> shared string</t></r><phoneticPr fontId="3" type="noConversion"/></si></sst>')
    shared_strings = SharedStrings.instance
    shared_strings.clear
    shared_strings.load_strings_from_xml xml
    SharedStrings.shared_string_for(0).should == "This a second shared string"
    SharedStrings.shared_string_for(1).should == "This is, hopefully, the first shared string"
  end
end