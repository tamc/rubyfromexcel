# coding: utf-8
require_relative '../spreadsheet'
# Sheet1
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell d3 should equal 3.0' do
    sheet1.d3.should be_close(3.0,0.3)
  end

  it 'cell d4 should equal 7.0' do
    sheet1.d4.should be_close(7.0,0.7)
  end

  it 'cell d5 should equal 11.0' do
    sheet1.d5.should be_close(11.0,1.1)
  end

end

