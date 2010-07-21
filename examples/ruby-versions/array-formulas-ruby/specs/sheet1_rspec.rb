# coding: utf-8
require_relative '../spreadsheet'
# Sheet1
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell b3 should equal 2.0' do
    sheet1.b3.should be_close(2.0,0.2)
  end

  it 'cell c3 should equal 3.0' do
    sheet1.c3.should be_close(3.0,0.3)
  end

  it 'cell d3 should equal 4.0' do
    sheet1.d3.should be_close(4.0,0.4)
  end

  it 'cell e3 should equal 5.0' do
    sheet1.e3.should be_close(5.0,0.5)
  end

  it 'cell b4 should equal 3.0' do
    sheet1.b4.should be_close(3.0,0.3)
  end

  it 'cell c4 should equal 4.0' do
    sheet1.c4.should be_close(4.0,0.4)
  end

  it 'cell d4 should equal 5.0' do
    sheet1.d4.should be_close(5.0,0.5)
  end

  it 'cell e4 should equal 6.0' do
    sheet1.e4.should be_close(6.0,0.6)
  end

  it 'cell b5 should equal 4.0' do
    sheet1.b5.should be_close(4.0,0.4)
  end

  it 'cell c5 should equal 5.0' do
    sheet1.c5.should be_close(5.0,0.5)
  end

  it 'cell d5 should equal 6.0' do
    sheet1.d5.should be_close(6.0,0.6)
  end

  it 'cell e5 should equal 7.0' do
    sheet1.e5.should be_close(7.0,0.7)
  end

  it 'cell b6 should equal 5.0' do
    sheet1.b6.should be_close(5.0,0.5)
  end

  it 'cell c6 should equal 6.0' do
    sheet1.c6.should be_close(6.0,0.6)
  end

  it 'cell d6 should equal 7.0' do
    sheet1.d6.should be_close(7.0,0.7)
  end

  it 'cell e6 should equal 8.0' do
    sheet1.e6.should be_close(8.0,0.8)
  end

  it 'cell b11 should equal 2.0' do
    sheet1.b11.should be_close(2.0,0.2)
  end

  it 'cell c11 should equal 2.0' do
    sheet1.c11.should be_close(2.0,0.2)
  end

  it 'cell d11 should equal 2.0' do
    sheet1.d11.should be_close(2.0,0.2)
  end

  it 'cell e11 should equal 2.0' do
    sheet1.e11.should be_close(2.0,0.2)
  end

  it 'cell b12 should equal 3.0' do
    sheet1.b12.should be_close(3.0,0.3)
  end

  it 'cell c12 should equal 3.0' do
    sheet1.c12.should be_close(3.0,0.3)
  end

  it 'cell d12 should equal 3.0' do
    sheet1.d12.should be_close(3.0,0.3)
  end

  it 'cell e12 should equal 3.0' do
    sheet1.e12.should be_close(3.0,0.3)
  end

  it 'cell b13 should equal 4.0' do
    sheet1.b13.should be_close(4.0,0.4)
  end

  it 'cell c13 should equal 4.0' do
    sheet1.c13.should be_close(4.0,0.4)
  end

  it 'cell d13 should equal 4.0' do
    sheet1.d13.should be_close(4.0,0.4)
  end

  it 'cell e13 should equal 4.0' do
    sheet1.e13.should be_close(4.0,0.4)
  end

  it 'cell b14 should equal 5.0' do
    sheet1.b14.should be_close(5.0,0.5)
  end

  it 'cell c14 should equal 5.0' do
    sheet1.c14.should be_close(5.0,0.5)
  end

  it 'cell d14 should equal 5.0' do
    sheet1.d14.should be_close(5.0,0.5)
  end

  it 'cell e14 should equal 5.0' do
    sheet1.e14.should be_close(5.0,0.5)
  end

  it 'cell c21 should equal 4.0' do
    sheet1.c21.should be_close(4.0,0.4)
  end

  it 'cell d21 should equal 8.0' do
    sheet1.d21.should be_close(8.0,0.8)
  end

  it 'cell e21 should equal 12.0' do
    sheet1.e21.should be_close(12.0,1.2)
  end

  it 'cell f21 should equal 16.0' do
    sheet1.f21.should be_close(16.0,1.6)
  end

  it 'cell g21 should equal :na' do
    sheet1.g21.should == :na
  end

end

