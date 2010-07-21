require_relative 'spec_helper'

describe ValueCell do
  
  it "creates a ruby method that returns the cell's value as a float if it is a number" do
    cell = ValueCell.new(mock('worksheet',:ruby_name => 'sheet1'),Nokogiri::XML('<c r="C22"><v>24</v></c>').root)
    cell.to_ruby.should == "def c22; 24.0; end\n"    
  end

  it "creates a ruby method that returns the cell's value as a ruby boolean if it is a boolean" do
    cell = ValueCell.new(mock('worksheet',:ruby_name => 'sheet1'),Nokogiri::XML('<c r="C22" t="b"><v>TRUE</v></c>').root)
    cell.to_ruby.should == "def c22; true; end\n"
    cell = ValueCell.new(mock('worksheet',:ruby_name => 'sheet1'),Nokogiri::XML('<c r="C22" t="b"><v>1</v></c>').root)
    cell.to_ruby.should == "def c22; true; end\n"

  end
  
  it "creates a ruby method that returns the cell's value as a ruby string if it is a string" do
    cell = ValueCell.new(mock('worksheet',:ruby_name => 'sheet1'),Nokogiri::XML('<c r="B6" t="str"><v>hello</v></c>').root)
    cell.to_ruby.should == %Q{def b6; "hello"; end\n}
  end

  it "creates a ruby method that returns the cell's value as a ruby string if it is a shared string" do
    cell = ValueCell.new(mock('worksheet',:ruby_name => 'sheet1'),Nokogiri::XML('<c r="B6" t="s"><v>7</v></c>').root)
    SharedStrings.should_receive(:shared_string_for).with(7).and_return('hello')
    cell.to_ruby.should == %Q{def b6; "hello"; end\n}
  end  
  
  it "creates a ruby method that returns an appropriate symbol if it is an error" do
    cell = ValueCell.new(mock('worksheet',:ruby_name => 'sheet1'),Nokogiri::XML('<c r="B6" t="e"><v>#N/A</v></c>').root)
    cell.to_ruby.should == "def b6; :na; end\n"
  end
  
  it "has a method that confirms it can be replaced with its value" do
    cell = ValueCell.new(mock('worksheet',:ruby_name => 'sheet1'),Nokogiri::XML('<c r="B6" t="str"><v>hello</v></c>').root)
    cell.can_be_replaced_with_value?.should be_true
  end
  
end

describe ValueCell, "can list the cells upon which it depends" do
  
  it "in simple cases knows what it depends upon" do
    cell = ValueCell.new(mock('worksheet',:ruby_name => 'sheet1'),Nokogiri::XML('<c r="B6" t="s"><v>7</v></c>').root)
    cell.work_out_dependencies
    cell.dependencies.should == []
  end

end