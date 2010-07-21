# coding: utf-8
require_relative '../spreadsheet'
# Calcs
describe 'Sheet2' do
  def sheet2; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet2; end

  it 'cell a1 should equal 121.0' do
    sheet2.a1.should be_close(121.0,12.1)
  end

  it 'cell a4 should equal 99.0' do
    sheet2.a4.should be_close(99.0,9.9)
  end

  it 'cell c9 should equal "In result"' do
    sheet2.c9.should == "In result"
  end

end

