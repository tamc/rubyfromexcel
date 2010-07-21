require_relative 'spec_helper'

describe Relationships do

relationship_xml =<<END
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
END

it "it should be created from xml be equivalent to a Hash where the keys are the relationship ids and the values are the target files" do
  wr = Relationships.new(Nokogiri::XML(relationship_xml).root)
  wr['rId3'].should == 'worksheets/sheet3.xml'
end

it "when given a root_directory will prepend that to the target files" do
  wr = Relationships.new(Nokogiri::XML(relationship_xml).root, root_directory = '/usr/local')
  wr['rId3'].should == '/usr/local/worksheets/sheet3.xml'
end

it "should have a Relationships.for_file(path_to_workkbook) that loads the xml automatically" do
  File.should_receive(:open).with('/usr/local/excel_dir/xl/_rels/workbook.xml.rels').and_yield(StringIO.new(relationship_xml))
  File.should_receive(:exist?).with('/usr/local/excel_dir/xl/_rels/workbook.xml.rels').and_return(true)
  
  wr = Relationships.for_file('/usr/local/excel_dir/xl/workbook.xml')
  wr['rId3'].should == '/usr/local/excel_dir/xl/worksheets/sheet3.xml'
  
  File.should_receive(:open).with('/usr/local/excel_dir/xl/worksheets/_rels/sheet1.xml.rels').and_yield(StringIO.new(relationship_xml))
    File.should_receive(:exist?).with('/usr/local/excel_dir/xl/worksheets/_rels/sheet1.xml.rels').and_return(true)
  wr = Relationships.for_file('/usr/local/excel_dir/xl/worksheets/sheet1.xml')
  wr['rId3'].should == '/usr/local/excel_dir/xl/worksheets/worksheets/sheet3.xml'
  
end

it "should return an empty Relationships if relationships xml file doesn't exist" do
  wr = Relationships.for_file('missing_file')
  wr.size.should == 0
end

it "should pick out the sharedStrings relationship and put in an accessor" do
  wr = Relationships.new(Nokogiri::XML(relationship_xml).root)
  wr.shared_strings.should == 'sharedStrings.xml'
end
end