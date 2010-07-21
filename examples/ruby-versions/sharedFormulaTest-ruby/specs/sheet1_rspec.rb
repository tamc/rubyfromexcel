# coding: utf-8
require_relative '../spreadsheet'
# Sheet1
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell a2 should equal 2.0' do
    sheet1.a2.should be_close(2.0,0.2)
  end

  it 'cell a3 should equal 4.0' do
    sheet1.a3.should be_close(4.0,0.4)
  end

  it 'cell a4 should equal 8.0' do
    sheet1.a4.should be_close(8.0,0.8)
  end

  it 'cell a5 should equal 16.0' do
    sheet1.a5.should be_close(16.0,1.6)
  end

  it 'cell a6 should equal 32.0' do
    sheet1.a6.should be_close(32.0,3.2)
  end

  it 'cell a7 should equal 64.0' do
    sheet1.a7.should be_close(64.0,6.4)
  end

  it 'cell a8 should equal 128.0' do
    sheet1.a8.should be_close(128.0,12.8)
  end

  it 'cell a9 should equal 256.0' do
    sheet1.a9.should be_close(256.0,25.6)
  end

  it 'cell a10 should equal 512.0' do
    sheet1.a10.should be_close(512.0,51.2)
  end

end

