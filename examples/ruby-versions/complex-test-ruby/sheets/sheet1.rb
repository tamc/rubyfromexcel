# coding: utf-8
# EU
class Sheet1 < Spreadsheet
  def b1; "Note, numbers not checked. Worry about power sector figure. Worry about the UK share of auction revenues."; end
  def d2; "Expansion"; end
  def d3; "Additions"; end
  def o3; "Annual"; end
  def b4; "EU-27 Emissions"; end
  def c4; "2005-6"; end
  def d4; "Phase II"; end
  def e4; "Phase III"; end
  def f4; 2013.0; end
  def g4; 2014.0; end
  def h4; 2015.0; end
  def i4; 2016.0; end
  def j4; 2017.0; end
  def k4; 2018.0; end
  def l4; 2019.0; end
  def m4; 2020.0; end
  def o4; "Change"; end
  def b5; "Power sector"; end
  def c5; 1150.0; end
  def d5; 50.0; end
  def e5; 50.0; end
  def f5; @f5 ||= +(c5+d5+e5)*0.9; end
  def g5; @g5 ||= f5*(1.0-o5); end
  def h5; @h5 ||= g5*(1.0-o5); end
  def i5; @i5 ||= h5*(1.0-o5); end
  def j5; @j5 ||= i5*(1.0-o5); end
  def k5; @k5 ||= j5*(1.0-o5); end
  def l5; @l5 ||= k5*(1.0-o5); end
  def m5; @m5 ||= l5*(1.0-o5); end
  def n5; "mtCO2"; end
  def o5; 0.02; end
  def b6; "Leakage sectors"; end
  def c6; 350.0; end
  def f6; @f6 ||= +(c6+d6+e6)*0.95; end
  def g6; @g6 ||= f6*(1.0-o5); end
  def h6; @h6 ||= g6*(1.0-o5); end
  def i6; @i6 ||= h6*(1.0-o5); end
  def j6; @j6 ||= i6*(1.0-o5); end
  def k6; @k6 ||= j6*(1.0-o5); end
  def l6; @l6 ||= k6*(1.0-o5); end
  def m6; @m6 ||= l6*(1.0-o5); end
  def n6; "mtCO2"; end
  def o6; 0.01; end
  def b7; "Other sectors"; end
  def c7; 550.0; end
  def d7; 100.0; end
  def e7; 100.0; end
  def f7; @f7 ||= +(c7+d7+e7)*0.95; end
  def g7; @g7 ||= f7*(1.0-o5); end
  def h7; @h7 ||= g7*(1.0-o5); end
  def i7; @i7 ||= h7*(1.0-o5); end
  def j7; @j7 ||= i7*(1.0-o5); end
  def k7; @k7 ||= j7*(1.0-o5); end
  def l7; @l7 ||= k7*(1.0-o5); end
  def m7; @m7 ||= l7*(1.0-o5); end
  def n7; "mtCO2"; end
  def o7; 0.01; end
  def b8; "Total"; end
  def f8; @f8 ||= +sum(a('f5','f7')); end
  def g8; @g8 ||= +sum(a('g5','g7')); end
  def h8; @h8 ||= +sum(a('h5','h7')); end
  def i8; @i8 ||= +sum(a('i5','i7')); end
  def j8; @j8 ||= +sum(a('j5','j7')); end
  def k8; @k8 ||= +sum(a('k5','k7')); end
  def l8; @l8 ||= +sum(a('l5','l7')); end
  def m8; @m8 ||= +sum(a('m5','m7')); end
  def n8; "mtCO2"; end
  def b10; "Proportion allocated for free"; end
  def f10; @f10 ||= f4; end
  def g10; @g10 ||= g4; end
  def h10; @h10 ||= h4; end
  def i10; @i10 ||= i4; end
  def j10; @j10 ||= j4; end
  def k10; @k10 ||= k4; end
  def l10; @l10 ||= l4; end
  def m10; @m10 ||= m4; end
  def b11; "Power sector"; end
  def f11; 0.0; end
  def g11; @g11 ||= f11; end
  def h11; @h11 ||= g11; end
  def i11; @i11 ||= h11; end
  def j11; @j11 ||= i11; end
  def k11; @k11 ||= j11; end
  def l11; @l11 ||= k11; end
  def m11; @m11 ||= l11; end
  def n11; "%"; end
  def b12; @b12 ||= +b6; end
  def f12; 0.9; end
  def g12; @g12 ||= +f12+o12; end
  def h12; @h12 ||= +g12+o12; end
  def i12; @i12 ||= +h12+o12; end
  def j12; @j12 ||= +i12+o12; end
  def k12; @k12 ||= +j12+o12; end
  def l12; @l12 ||= +k12+o12; end
  def m12; @m12 ||= +l12+o12; end
  def o12; -0.025; end
  def b13; "Other"; end
  def f13; 0.8; end
  def g13; @g13 ||= f13+o13; end
  def h13; @h13 ||= g13+o13; end
  def i13; @i13 ||= h13+o13; end
  def j13; @j13 ||= i13+o13; end
  def k13; @k13 ||= j13+o13; end
  def l13; @l13 ||= k13+o13; end
  def m13; 0.0; end
  def n13; "%"; end
  def o13; @o13 ||= (m13-f13)/(m10-f10); end
  def b14; "Average non-power sectors"; end
  def f14; @f14 ||= +(f6*f12+f7*f13)/(f6+f7); end
  def g14; @g14 ||= +(g6*g12+g7*g13)/(g6+g7); end
  def h14; @h14 ||= +(h6*h12+h7*h13)/(h6+h7); end
  def i14; @i14 ||= +(i6*i12+i7*i13)/(i6+i7); end
  def j14; @j14 ||= +(j6*j12+j7*j13)/(j6+j7); end
  def k14; @k14 ||= +(k6*k12+k7*k13)/(k6+k7); end
  def l14; @l14 ||= +(l6*l12+l7*l13)/(l6+l7); end
  def m14; @m14 ||= +(m6*m12+m7*m13)/(m6+m7); end
  def b15; "Total free allocation, % emissions"; end
  def f15; @f15 ||= ((f11*f5)+(f12*f6)+(f13*f7))/f8; end
  def g15; @g15 ||= ((g11*g5)+(g12*g6)+(g13*g7))/g8; end
  def h15; @h15 ||= ((h11*h5)+(h12*h6)+(h13*h7))/h8; end
  def i15; @i15 ||= ((i11*i5)+(i12*i6)+(i13*i7))/i8; end
  def j15; @j15 ||= ((j11*j5)+(j12*j6)+(j13*j7))/j8; end
  def k15; @k15 ||= ((k11*k5)+(k12*k6)+(k13*k7))/k8; end
  def l15; @l15 ||= ((l11*l5)+(l12*l6)+(l13*l7))/l8; end
  def m15; @m15 ||= ((m11*m5)+(m12*m6)+(m13*m7))/m8; end
  def n15; "%"; end
  def b17; "Proportion auctioned, other sectors"; end
  def f17; @f17 ||= 1.0-f14; end
  def g17; @g17 ||= 1.0-g14; end
  def h17; @h17 ||= 1.0-h14; end
  def i17; @i17 ||= 1.0-i14; end
  def j17; @j17 ||= 1.0-j14; end
  def k17; @k17 ||= 1.0-k14; end
  def l17; @l17 ||= 1.0-l14; end
  def m17; @m17 ||= 1.0-m14; end
  def b18; "Total allowances"; end
  def d18; 2010.0; end
  def f18; @f18 ||= f10; end
  def g18; @g18 ||= g10; end
  def h18; @h18 ||= h10; end
  def i18; @i18 ||= i10; end
  def j18; @j18 ||= j10; end
  def k18; @k18 ||= k10; end
  def l18; @l18 ||= l10; end
  def m18; @m18 ||= m10; end
  def b19; "Total allowances"; end
  def d19; @d19 ||= +(2100.0+100.0)*0.935+150.0; end
  def f19; @f19 ||= +d19*(1.0-3.0*o19); end
  def g19; @g19 ||= +f19*(1.0-o19); end
  def h19; @h19 ||= +g19*(1.0-o19); end
  def i19; @i19 ||= +h19*(1.0-o19); end
  def j19; @j19 ||= +i19*(1.0-o19); end
  def k19; @k19 ||= +j19*(1.0-o19); end
  def l19; @l19 ||= +k19*(1.0-o19); end
  def m19; @m19 ||= +l19*(1.0-o19); end
  def o19; 0.0174; end
  def p19; "Note: 1720 \"based on current scope\""; end
  def b20; "Total free allocation, MtCO2"; end
  def f20; @f20 ||= +f6*f12+f7*f13*f19/f8; end
  def g20; @g20 ||= +f6*g12+f7*g13*g19/g8; end
  def h20; @h20 ||= +f6*h12+f7*h13*h19/h8; end
  def i20; @i20 ||= +f6*i12+f7*i13*i19/i8; end
  def j20; @j20 ||= +f6*j12+f7*j13*j19/j8; end
  def k20; @k20 ||= +f6*k12+f7*k13*k19/k8; end
  def l20; @l20 ||= +f6*l12+f7*l13*l19/l8; end
  def m20; @m20 ||= +f6*m12+f7*m13*m19/m8; end
  def b21; "Volume available for auctioning"; end
  def f21; @f21 ||= +f19-f20; end
  def g21; @g21 ||= +g19-g20; end
  def h21; @h21 ||= +h19-h20; end
  def i21; @i21 ||= +i19-i20; end
  def j21; @j21 ||= +j19-j20; end
  def k21; @k21 ||= +k19-k20; end
  def l21; @l21 ||= +l19-l20; end
  def m21; @m21 ||= +m19-m20; end
  def b23; "Carbon Price"; end
  def f23; @f23 ||= f10; end
  def g23; @g23 ||= g10; end
  def h23; @h23 ||= h10; end
  def i23; @i23 ||= i10; end
  def j23; @j23 ||= j10; end
  def k23; @k23 ||= k10; end
  def l23; @l23 ||= l10; end
  def m23; @m23 ||= m10; end
  def b24; "Price per allowance"; end
  def f24; 25.0; end
  def g24; @g24 ||= +f24*(1.0+o24); end
  def h24; @h24 ||= +g24*(1.0+o24); end
  def i24; @i24 ||= +h24*(1.0+o24); end
  def j24; @j24 ||= +i24*(1.0+o24); end
  def k24; @k24 ||= +j24*(1.0+o24); end
  def l24; @l24 ||= +k24*(1.0+o24); end
  def m24; @m24 ||= +l24*(1.0+o24); end
  def n24; "€/tCO2"; end
  def o24; 0.05; end
  def b25; "Total revenue from auctions"; end
  def f25; @f25 ||= +f24*f21; end
  def g25; @g25 ||= +g24*g21; end
  def h25; @h25 ||= +h24*h21; end
  def i25; @i25 ||= +i24*i21; end
  def j25; @j25 ||= +j24*j21; end
  def k25; @k25 ||= +k24*k21; end
  def l25; @l25 ||= +l24*l21; end
  def m25; @m25 ||= +m24*m21; end
  def b27; "EU-27 Auction volumes- bought into proportion to net shortfall"; end
  def f27; @f27 ||= f10; end
  def g27; @g27 ||= g10; end
  def h27; @h27 ||= h10; end
  def i27; @i27 ||= i10; end
  def j27; @j27 ||= j10; end
  def k27; @k27 ||= k10; end
  def l27; @l27 ||= l10; end
  def m27; @m27 ||= m10; end
  def b28; "Power sector"; end
  def f28; @f28 ||= +f5*(1.0-f11)*(f19/f8); end
  def g28; @g28 ||= +g5*(1.0-g11)*(g19/g8); end
  def h28; @h28 ||= +h5*(1.0-h11)*(h19/h8); end
  def i28; @i28 ||= +i5*(1.0-i11)*(i19/i8); end
  def j28; @j28 ||= +j5*(1.0-j11)*(j19/j8); end
  def k28; @k28 ||= +k5*(1.0-k11)*(k19/k8); end
  def l28; @l28 ||= +l5*(1.0-l11)*(l19/l8); end
  def m28; @m28 ||= +m5*(1.0-m11)*(m19/m8); end
  def b29; @b29 ||= +b6; end
  def f29; @f29 ||= +f6*(1.0-f12)*(f19/f8); end
  def g29; @g29 ||= +g6*(1.0-g12)*(g19/g8); end
  def h29; @h29 ||= +h6*(1.0-h12)*(h19/h8); end
  def i29; @i29 ||= +i6*(1.0-i12)*(i19/i8); end
  def j29; @j29 ||= +j6*(1.0-j12)*(j19/j8); end
  def k29; @k29 ||= +k6*(1.0-k12)*(k19/k8); end
  def l29; @l29 ||= +l6*(1.0-l12)*(l19/l8); end
  def m29; @m29 ||= +m6*(1.0-m12)*(m19/m8); end
  def b30; "Other"; end
  def f30; @f30 ||= +f7*(1.0-f13)*(f19/f8); end
  def g30; @g30 ||= +g7*(1.0-g13)*(g19/g8); end
  def h30; @h30 ||= +h7*(1.0-h13)*(h19/h8); end
  def i30; @i30 ||= +i7*(1.0-i13)*(i19/i8); end
  def j30; @j30 ||= +j7*(1.0-j13)*(j19/j8); end
  def k30; @k30 ||= +k7*(1.0-k13)*(k19/k8); end
  def l30; @l30 ||= +l7*(1.0-l13)*(l19/l8); end
  def m30; @m30 ||= +m7*(1.0-m13)*(m19/m8); end
  def b31; "Total"; end
  def f31; @f31 ||= +sum(a('f28','f30')); end
  def g31; @g31 ||= +sum(a('g28','g30')); end
  def h31; @h31 ||= +sum(a('h28','h30')); end
  def i31; @i31 ||= +sum(a('i28','i30')); end
  def j31; @j31 ||= +sum(a('j28','j30')); end
  def k31; @k31 ||= +sum(a('k28','k30')); end
  def l31; @l31 ||= +sum(a('l28','l30')); end
  def m31; @m31 ||= +sum(a('m28','m30')); end
  def b33; "Revenues"; end
  def b34; "Power sector"; end
  def f34; @f34 ||= +f28*f24; end
  def g34; @g34 ||= +g28*g24; end
  def h34; @h34 ||= +h28*h24; end
  def i34; @i34 ||= +i28*i24; end
  def j34; @j34 ||= +j28*j24; end
  def k34; @k34 ||= +k28*k24; end
  def l34; @l34 ||= +l28*l24; end
  def m34; @m34 ||= +m28*m24; end
  def b35; "Other sectors"; end
  def f35; @f35 ||= +(f29+f30)*f24; end
  def g35; @g35 ||= +(g29+g30)*g24; end
  def h35; @h35 ||= +(h29+h30)*h24; end
  def i35; @i35 ||= +(i29+i30)*i24; end
  def j35; @j35 ||= +(j29+j30)*j24; end
  def k35; @k35 ||= +(k29+k30)*k24; end
  def l35; @l35 ||= +(l29+l30)*l24; end
  def m35; @m35 ||= +(m29+m30)*m24; end
  def b39; "2005 UK ETS Emissions"; end
  def f39; 242.0; end
  def g39; "mtCO2"; end
  def n39; " "; end
  def b40; "2005 EU ETS Emissions"; end
  def f40; 1785.0; end
  def g40; "mtCO2"; end
  def b41; "Basic UK share of auction revenues"; end
  def f41; @f41 ||= f39/f40; end
  def b42; "Amount of share auctioned in UK"; end
  def f42; 0.9; end
  def b43; "Actual UK share of auction revenues"; end
  def f43; @f43 ||= f42*f41; end
  def b45; "UK Auction revenues"; end
  def f45; @f45 ||= f4; end
  def g45; @g45 ||= g4; end
  def h45; @h45 ||= h4; end
  def i45; @i45 ||= i4; end
  def j45; @j45 ||= j4; end
  def k45; @k45 ||= k4; end
  def l45; @l45 ||= l4; end
  def m45; @m45 ||= m4; end
  def b46; "Total"; end
  def f46; @f46 ||= f43*f31; end
  def g46; @g46 ||= f43*g31; end
  def h46; @h46 ||= f43*h31; end
  def i46; @i46 ||= f43*i31; end
  def j46; @j46 ||= f43*j31; end
  def k46; @k46 ||= f43*k31; end
  def l46; @l46 ||= f43*l31; end
  def m46; @m46 ||= f43*m31; end
  def n46; "€bn"; end
end

