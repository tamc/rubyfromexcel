# coding: utf-8
# Sheet1
class Sheet1 < Spreadsheet
  def a1; 1.0; end
  def a2; @a2 ||= a1*2.0; end
  def a3; @a3 ||= a2*2.0; end
  def a4; @a4 ||= a3*2.0; end
  def a5; @a5 ||= a4*2.0; end
  def a6; @a6 ||= a5*2.0; end
  def a7; @a7 ||= a6*2.0; end
  def a8; @a8 ||= a7*2.0; end
  def a9; @a9 ||= a8*2.0; end
  def a10; @a10 ||= a9*2.0; end
end

