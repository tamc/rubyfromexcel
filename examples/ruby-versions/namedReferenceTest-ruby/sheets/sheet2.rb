# coding: utf-8
# Sheet1 (2)
class Sheet2 < Spreadsheet
  def a1; 5.0; end
  def a2; @a2 ||= sheet2.a1; end
  def a3; @a3 ||= sum(sheet1.a('a2','a3')); end
end

