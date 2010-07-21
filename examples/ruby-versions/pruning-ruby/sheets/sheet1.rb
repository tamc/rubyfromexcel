# coding: utf-8
# Outputs
class Sheet1 < Spreadsheet
  def a1; "Result"; end
  def b1; @b1 ||= sheet2.a1; end
  def a2; "Input"; end
  def b2; @b2 ||= sheet3.a1; end
  def b3; @b3 ||= sheet2.c9; end
  def b4; "Doesn't depend on an input"; end
end

