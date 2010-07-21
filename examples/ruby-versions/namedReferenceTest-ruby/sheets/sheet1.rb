# coding: utf-8
# Sheet1
class Sheet1 < Spreadsheet
  def a2; 1.0; end
  def a3; 2.0; end
  def a4; @a4 ||= sheet1.a2; end
  def a5; @a5 ||= sheet1.a3; end
end

