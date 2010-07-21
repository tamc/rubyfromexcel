# coding: utf-8
require_relative '../spreadsheet'
# Sheet1 (2)
describe 'Sheet2' do
  def sheet2; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet2; end

  it 'cell a2 should equal 5.0' do
    sheet2.a2.should be_close(5.0,0.5)
  end

  it 'cell a3 should equal 3.0' do
    sheet2.a3.should be_close(3.0,0.3)
  end

end

