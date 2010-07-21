# coding: utf-8
# Sheet1
class Sheet1 < Spreadsheet
  def b2; "Column A"; end
  def c2; "Column B "; end
  def d2; "Column C"; end
  def b3; 1.0; end
  def c3; 2.0; end
  def d3; @d3 ||= sheet1.b3+sheet1.c3; end
  def b4; 3.0; end
  def c4; 4.0; end
  def d4; @d4 ||= sheet1.b4+sheet1.c4; end
  def b5; 5.0; end
  def c5; 6.0; end
  def d5; @d5 ||= sheet1.b5+sheet1.c5; end
end

