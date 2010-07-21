# coding: utf-8
require_relative '../spreadsheet'
# Sheet1
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell a4 should equal 1.0' do
    sheet1.a4.should be_close(1.0,0.1)
  end

  it 'cell a5 should equal 2.0' do
    sheet1.a5.should be_close(2.0,0.2)
  end

end

