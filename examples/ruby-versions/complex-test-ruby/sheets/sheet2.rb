# coding: utf-8
# OLD UK
class Sheet2 < Spreadsheet
  def b2; "UK - Possible phase III auction revenues"; end
  def b4; "Emissions"; end
  def c4; 2013.0; end
  def d4; 2014.0; end
  def e4; 2015.0; end
  def f4; 2016.0; end
  def g4; 2017.0; end
  def h4; 2018.0; end
  def i4; 2019.0; end
  def j4; 2020.0; end
  def l4; "Change"; end
  def b5; "Power sector"; end
  def c5; 163.25; end
  def d5; @d5 ||= c5*(1.0-l5); end
  def e5; @e5 ||= d5*(1.0-l5); end
  def f5; @f5 ||= e5*(1.0-l5); end
  def g5; @g5 ||= f5*(1.0-l5); end
  def h5; @h5 ||= g5*(1.0-l5); end
  def i5; @i5 ||= h5*(1.0-l5); end
  def j5; @j5 ||= i5*(1.0-l5); end
  def k5; "mtCO2"; end
  def l5; 0.0174; end
  def b6; "Other"; end
  def c6; @c6 ||= c7-c5; end
  def d6; @d6 ||= d7-d5; end
  def e6; @e6 ||= e7-e5; end
  def f6; @f6 ||= f7-f5; end
  def g6; @g6 ||= g7-g5; end
  def h6; @h6 ||= h7-h5; end
  def i6; @i6 ||= i7-i5; end
  def j6; @j6 ||= j7-j5; end
  def k6; "mtCO3"; end
  def l6; @l6 ||= -((j6/c6)**(1.0/(j4-c4))-1.0); end
  def b7; "Total"; end
  def c7; @c7 ||= 278.35*(1.0+average(0.066,0.071)); end
  def d7; @d7 ||= c7*(1.0-l7); end
  def e7; @e7 ||= d7*(1.0-l7); end
  def f7; @f7 ||= e7*(1.0-l7); end
  def g7; @g7 ||= f7*(1.0-l7); end
  def h7; @h7 ||= g7*(1.0-l7); end
  def i7; @i7 ||= h7*(1.0-l7); end
  def j7; @j7 ||= i7*(1.0-l7); end
  def k7; "mtCO4"; end
  def l7; 0.0174; end
  def b9; "Proportion auctioned"; end
  def c9; @c9 ||= c4; end
  def d9; @d9 ||= d4; end
  def e9; @e9 ||= e4; end
  def f9; @f9 ||= f4; end
  def g9; @g9 ||= g4; end
  def h9; @h9 ||= h4; end
  def i9; @i9 ||= i4; end
  def j9; @j9 ||= j4; end
  def b10; "Power sector"; end
  def c10; 1.0; end
  def d10; @d10 ||= c10; end
  def e10; @e10 ||= d10; end
  def f10; @f10 ||= e10; end
  def g10; @g10 ||= f10; end
  def h10; @h10 ||= g10; end
  def i10; @i10 ||= h10; end
  def j10; @j10 ||= i10; end
  def k10; "%"; end
  def b11; "Other"; end
  def c11; 0.2; end
  def d11; @d11 ||= c11+l11; end
  def e11; @e11 ||= d11+l11; end
  def f11; @f11 ||= e11+l11; end
  def g11; @g11 ||= f11+l11; end
  def h11; @h11 ||= g11+l11; end
  def i11; @i11 ||= h11+l11; end
  def j11; 1.0; end
  def k11; "%"; end
  def l11; @l11 ||= (j11-c11)/(j9-c9); end
  def b12; "Total"; end
  def c12; @c12 ||= ((c10*c5)+(c11*c6))/c7; end
  def d12; @d12 ||= ((d10*d5)+(d11*d6))/d7; end
  def e12; @e12 ||= ((e10*e5)+(e11*e6))/e7; end
  def f12; @f12 ||= ((f10*f5)+(f11*f6))/f7; end
  def g12; @g12 ||= ((g10*g5)+(g11*g6))/g7; end
  def h12; @h12 ||= ((h10*h5)+(h11*h6))/h7; end
  def i12; @i12 ||= ((i10*i5)+(i11*i6))/i7; end
  def j12; @j12 ||= ((j10*j5)+(j11*j6))/j7; end
  def k12; "%"; end
  def c14; @c14 ||= c9; end
  def d14; @d14 ||= d9; end
  def e14; @e14 ||= e9; end
  def f14; @f14 ||= f9; end
  def g14; @g14 ||= g9; end
  def h14; @h14 ||= h9; end
  def i14; @i14 ||= i9; end
  def j14; @j14 ||= j9; end
  def b15; "CO2 price"; end
  def c15; 20.0; end
  def d15; @d15 ||= c15; end
  def e15; @e15 ||= d15; end
  def f15; @f15 ||= e15; end
  def g15; @g15 ||= f15; end
  def h15; @h15 ||= g15; end
  def i15; @i15 ||= h15; end
  def j15; @j15 ||= i15; end
  def k15; "€/tCO2"; end
  def b17; "Auction revenues"; end
  def c17; @c17 ||= c9; end
  def d17; @d17 ||= d9; end
  def e17; @e17 ||= e9; end
  def f17; @f17 ||= f9; end
  def g17; @g17 ||= g9; end
  def h17; @h17 ||= h9; end
  def i17; @i17 ||= i9; end
  def j17; @j17 ||= j9; end
  def b18; "Power sector"; end
  def c18; @c18 ||= c10*c5*c15/1000.0; end
  def d18; @d18 ||= d10*d5*d15/1000.0; end
  def e18; @e18 ||= e10*e5*e15/1000.0; end
  def f18; @f18 ||= f10*f5*f15/1000.0; end
  def g18; @g18 ||= g10*g5*g15/1000.0; end
  def h18; @h18 ||= h10*h5*h15/1000.0; end
  def i18; @i18 ||= i10*i5*i15/1000.0; end
  def j18; @j18 ||= j10*j5*j15/1000.0; end
  def k18; "€bn"; end
  def b19; "Other"; end
  def c19; @c19 ||= c11*c6*c15/1000.0; end
  def d19; @d19 ||= d11*d6*d15/1000.0; end
  def e19; @e19 ||= e11*e6*e15/1000.0; end
  def f19; @f19 ||= f11*f6*f15/1000.0; end
  def g19; @g19 ||= g11*g6*g15/1000.0; end
  def h19; @h19 ||= h11*h6*h15/1000.0; end
  def i19; @i19 ||= i11*i6*i15/1000.0; end
  def j19; @j19 ||= j11*j6*j15/1000.0; end
  def k19; "€bn"; end
  def b20; "Total"; end
  def c20; @c20 ||= c12*c7*c15/1000.0; end
  def d20; @d20 ||= d12*d7*d15/1000.0; end
  def e20; @e20 ||= e12*e7*e15/1000.0; end
  def f20; @f20 ||= f12*f7*f15/1000.0; end
  def g20; @g20 ||= g12*g7*g15/1000.0; end
  def h20; @h20 ||= h12*h7*h15/1000.0; end
  def i20; @i20 ||= i12*i7*i15/1000.0; end
  def j20; @j20 ||= j12*j7*j15/1000.0; end
  def k20; "€bn"; end
  def k22; " "; end
end

