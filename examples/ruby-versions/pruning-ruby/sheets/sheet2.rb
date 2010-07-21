# coding: utf-8
# Calcs
class Sheet2 < Spreadsheet
  def a1; @a1 ||= sum(a('a2','a7')); end
  def a2; 1.0; end
  def a3; 2.0; end
  def a4; @a4 ||= sheet3.a1; end
  def a5; 4.0; end
  def a6; 10.0; end
  def a7; 5.0; end
  def c8; "Inputs"; end
  def c9; @c9 ||= sheet3.a3; end
end

