# coding: utf-8
require_relative '../spreadsheet'
# Outputs
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell b1 should equal 121.0' do
    sheet1.b1.should be_close(121.0,12.1)
  end

  it 'cell b2 should equal 99.0' do
    sheet1.b2.should be_close(99.0,9.9)
  end

  it 'cell b3 should equal "In result"' do
    sheet1.b3.should == "In result"
  end

end

