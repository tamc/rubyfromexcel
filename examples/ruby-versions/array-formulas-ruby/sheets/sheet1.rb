# coding: utf-8
# Sheet1
class Sheet1 < Spreadsheet
  def b2; 1.0; end
  def c2; 2.0; end
  def d2; 3.0; end
  def e2; 4.0; end
  def a3; 1.0; end
  def b3_array; @b3_array ||= m(a('b2','e2'),a('a3','a6')) { |r1,r2| r1+r2 }; end
  def b3; @b3 ||= b3_array.array_formula_offset(0,0); end
  def c3; @c3 ||= b3_array.array_formula_offset(0,1); end
  def d3; @d3 ||= b3_array.array_formula_offset(0,2); end
  def e3; @e3 ||= b3_array.array_formula_offset(0,3); end
  def a4; 2.0; end
  def b4; @b4 ||= b3_array.array_formula_offset(1,0); end
  def c4; @c4 ||= b3_array.array_formula_offset(1,1); end
  def d4; @d4 ||= b3_array.array_formula_offset(1,2); end
  def e4; @e4 ||= b3_array.array_formula_offset(1,3); end
  def a5; 3.0; end
  def b5; @b5 ||= b3_array.array_formula_offset(2,0); end
  def c5; @c5 ||= b3_array.array_formula_offset(2,1); end
  def d5; @d5 ||= b3_array.array_formula_offset(2,2); end
  def e5; @e5 ||= b3_array.array_formula_offset(2,3); end
  def a6; 4.0; end
  def b6; @b6 ||= b3_array.array_formula_offset(3,0); end
  def c6; @c6 ||= b3_array.array_formula_offset(3,1); end
  def d6; @d6 ||= b3_array.array_formula_offset(3,2); end
  def e6; @e6 ||= b3_array.array_formula_offset(3,3); end
  def a11; 1.0; end
  def b11_array; @b11_array ||= m(a('a11','a14'),b2) { |r1,r2| r1+r2 }; end
  def b11; @b11 ||= b11_array.array_formula_offset(0,0); end
  def c11; @c11 ||= b11_array.array_formula_offset(0,1); end
  def d11; @d11 ||= b11_array.array_formula_offset(0,2); end
  def e11; @e11 ||= b11_array.array_formula_offset(0,3); end
  def a12; 2.0; end
  def b12; @b12 ||= b11_array.array_formula_offset(1,0); end
  def c12; @c12 ||= b11_array.array_formula_offset(1,1); end
  def d12; @d12 ||= b11_array.array_formula_offset(1,2); end
  def e12; @e12 ||= b11_array.array_formula_offset(1,3); end
  def a13; 3.0; end
  def b13; @b13 ||= b11_array.array_formula_offset(2,0); end
  def c13; @c13 ||= b11_array.array_formula_offset(2,1); end
  def d13; @d13 ||= b11_array.array_formula_offset(2,2); end
  def e13; @e13 ||= b11_array.array_formula_offset(2,3); end
  def a14; 4.0; end
  def b14; @b14 ||= b11_array.array_formula_offset(3,0); end
  def c14; @c14 ||= b11_array.array_formula_offset(3,1); end
  def d14; @d14 ||= b11_array.array_formula_offset(3,2); end
  def e14; @e14 ||= b11_array.array_formula_offset(3,3); end
  def c21_array; @c21_array ||= m(2.0,sheet2.a('b15','e15')) { |r1,r2| r1*r2 }; end
  def c21; @c21 ||= c21_array.array_formula_offset(0,0); end
  def d21; @d21 ||= c21_array.array_formula_offset(0,1); end
  def e21; @e21 ||= c21_array.array_formula_offset(0,2); end
  def f21; @f21 ||= c21_array.array_formula_offset(0,3); end
  def g21; @g21 ||= c21_array.array_formula_offset(0,4); end
  def d24; "This is, hopefully, the first shared string"; end
  def d25; "This a second shared string"; end
end

