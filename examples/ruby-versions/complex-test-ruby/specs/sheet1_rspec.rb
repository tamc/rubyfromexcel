# coding: utf-8
require_relative '../spreadsheet'
# EU
describe 'Sheet1' do
  def sheet1; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet1; end

  it 'cell f5 should equal 1125.0' do
    sheet1.f5.should be_close(1125.0,112.5)
  end

  it 'cell g5 should equal 1102.5' do
    sheet1.g5.should be_close(1102.5,110.25)
  end

  it 'cell h5 should equal 1080.45' do
    sheet1.h5.should be_close(1080.45,108.045)
  end

  it 'cell i5 should equal 1058.841' do
    sheet1.i5.should be_close(1058.841,105.8841)
  end

  it 'cell j5 should equal 1037.66418' do
    sheet1.j5.should be_close(1037.66418,103.766418)
  end

  it 'cell k5 should equal 1016.9108964' do
    sheet1.k5.should be_close(1016.9108964,101.69108964)
  end

  it 'cell l5 should equal 996.572678472' do
    sheet1.l5.should be_close(996.572678472,99.6572678472)
  end

  it 'cell m5 should equal 976.64122490256' do
    sheet1.m5.should be_close(976.64122490256,97.664122490256)
  end

  it 'cell f6 should equal 332.5' do
    sheet1.f6.should be_close(332.5,33.25)
  end

  it 'cell g6 should equal 325.85' do
    sheet1.g6.should be_close(325.85,32.585)
  end

  it 'cell h6 should equal 319.333' do
    sheet1.h6.should be_close(319.333,31.9333)
  end

  it 'cell i6 should equal 312.94634' do
    sheet1.i6.should be_close(312.94634,31.294634)
  end

  it 'cell j6 should equal 306.6874132' do
    sheet1.j6.should be_close(306.6874132,30.66874132)
  end

  it 'cell k6 should equal 300.553664936' do
    sheet1.k6.should be_close(300.553664936,30.0553664936)
  end

  it 'cell l6 should equal 294.54259163728' do
    sheet1.l6.should be_close(294.54259163728,29.454259163728)
  end

  it 'cell m6 should equal 288.651739804534' do
    sheet1.m6.should be_close(288.651739804534,28.8651739804534)
  end

  it 'cell f7 should equal 712.5' do
    sheet1.f7.should be_close(712.5,71.25)
  end

  it 'cell g7 should equal 698.25' do
    sheet1.g7.should be_close(698.25,69.825)
  end

  it 'cell h7 should equal 684.285' do
    sheet1.h7.should be_close(684.285,68.4285)
  end

  it 'cell i7 should equal 670.5993' do
    sheet1.i7.should be_close(670.5993,67.05993)
  end

  it 'cell j7 should equal 657.187314' do
    sheet1.j7.should be_close(657.187314,65.7187314)
  end

  it 'cell k7 should equal 644.04356772' do
    sheet1.k7.should be_close(644.04356772,64.404356772)
  end

  it 'cell l7 should equal 631.1626963656' do
    sheet1.l7.should be_close(631.1626963656,63.11626963656)
  end

  it 'cell m7 should equal 618.539442438288' do
    sheet1.m7.should be_close(618.539442438288,61.8539442438288)
  end

  it 'cell f8 should equal 2170.0' do
    sheet1.f8.should be_close(2170.0,217.0)
  end

  it 'cell g8 should equal 2126.6' do
    sheet1.g8.should be_close(2126.6,212.66)
  end

  it 'cell h8 should equal 2084.068' do
    sheet1.h8.should be_close(2084.068,208.4068)
  end

  it 'cell i8 should equal 2042.38664' do
    sheet1.i8.should be_close(2042.38664,204.238664)
  end

  it 'cell j8 should equal 2001.5389072' do
    sheet1.j8.should be_close(2001.5389072,200.15389072)
  end

  it 'cell k8 should equal 1961.508129056' do
    sheet1.k8.should be_close(1961.508129056,196.1508129056)
  end

  it 'cell l8 should equal 1922.27796647488' do
    sheet1.l8.should be_close(1922.27796647488,192.227796647488)
  end

  it 'cell m8 should equal 1883.83240714538' do
    sheet1.m8.should be_close(1883.83240714538,188.383240714538)
  end

  it 'cell f10 should equal 2013.0' do
    sheet1.f10.should be_close(2013.0,201.3)
  end

  it 'cell g10 should equal 2014.0' do
    sheet1.g10.should be_close(2014.0,201.4)
  end

  it 'cell h10 should equal 2015.0' do
    sheet1.h10.should be_close(2015.0,201.5)
  end

  it 'cell i10 should equal 2016.0' do
    sheet1.i10.should be_close(2016.0,201.6)
  end

  it 'cell j10 should equal 2017.0' do
    sheet1.j10.should be_close(2017.0,201.7)
  end

  it 'cell k10 should equal 2018.0' do
    sheet1.k10.should be_close(2018.0,201.8)
  end

  it 'cell l10 should equal 2019.0' do
    sheet1.l10.should be_close(2019.0,201.9)
  end

  it 'cell m10 should equal 2020.0' do
    sheet1.m10.should be_close(2020.0,202.0)
  end

  it 'cell g11 should equal 0.0' do
    sheet1.g11.should be_close(0.0,1.0e-08)
  end

  it 'cell h11 should equal 0.0' do
    sheet1.h11.should be_close(0.0,1.0e-08)
  end

  it 'cell i11 should equal 0.0' do
    sheet1.i11.should be_close(0.0,1.0e-08)
  end

  it 'cell j11 should equal 0.0' do
    sheet1.j11.should be_close(0.0,1.0e-08)
  end

  it 'cell k11 should equal 0.0' do
    sheet1.k11.should be_close(0.0,1.0e-08)
  end

  it 'cell l11 should equal 0.0' do
    sheet1.l11.should be_close(0.0,1.0e-08)
  end

  it 'cell m11 should equal 0.0' do
    sheet1.m11.should be_close(0.0,1.0e-08)
  end

  it 'cell b12 should equal "Leakage sectors"' do
    sheet1.b12.should == "Leakage sectors"
  end

  it 'cell g12 should equal 0.875' do
    sheet1.g12.should be_close(0.875,0.0875)
  end

  it 'cell h12 should equal 0.85' do
    sheet1.h12.should be_close(0.85,0.085)
  end

  it 'cell i12 should equal 0.825' do
    sheet1.i12.should be_close(0.825,0.0825)
  end

  it 'cell j12 should equal 0.8' do
    sheet1.j12.should be_close(0.8,0.08)
  end

  it 'cell k12 should equal 0.775' do
    sheet1.k12.should be_close(0.775,0.0775)
  end

  it 'cell l12 should equal 0.75' do
    sheet1.l12.should be_close(0.75,0.075)
  end

  it 'cell m12 should equal 0.725' do
    sheet1.m12.should be_close(0.725,0.0725)
  end

  it 'cell g13 should equal 0.685714285714286' do
    sheet1.g13.should be_close(0.685714285714286,0.0685714285714286)
  end

  it 'cell h13 should equal 0.571428571428571' do
    sheet1.h13.should be_close(0.571428571428571,0.0571428571428571)
  end

  it 'cell i13 should equal 0.457142857142857' do
    sheet1.i13.should be_close(0.457142857142857,0.0457142857142857)
  end

  it 'cell j13 should equal 0.342857142857143' do
    sheet1.j13.should be_close(0.342857142857143,0.0342857142857143)
  end

  it 'cell k13 should equal 0.228571428571428' do
    sheet1.k13.should be_close(0.228571428571428,0.0228571428571428)
  end

  it 'cell l13 should equal 0.114285714285714' do
    sheet1.l13.should be_close(0.114285714285714,0.0114285714285714)
  end

  it 'cell o13 should equal -0.114285714285714' do
    sheet1.o13.should be_close(-0.114285714285714,0.0114285714285714)
  end

  it 'cell f14 should equal 0.831818181818182' do
    sheet1.f14.should be_close(0.831818181818182,0.0831818181818182)
  end

  it 'cell g14 should equal 0.745941558441559' do
    sheet1.g14.should be_close(0.745941558441559,0.0745941558441559)
  end

  it 'cell h14 should equal 0.660064935064935' do
    sheet1.h14.should be_close(0.660064935064935,0.0660064935064935)
  end

  it 'cell i14 should equal 0.574188311688312' do
    sheet1.i14.should be_close(0.574188311688312,0.0574188311688312)
  end

  it 'cell j14 should equal 0.488311688311688' do
    sheet1.j14.should be_close(0.488311688311688,0.0488311688311688)
  end

  it 'cell k14 should equal 0.402435064935065' do
    sheet1.k14.should be_close(0.402435064935065,0.0402435064935065)
  end

  it 'cell l14 should equal 0.316558441558441' do
    sheet1.l14.should be_close(0.316558441558441,0.0316558441558441)
  end

  it 'cell m14 should equal 0.230681818181818' do
    sheet1.m14.should be_close(0.230681818181818,0.0230681818181818)
  end

  it 'cell f15 should equal 0.400576036866359' do
    sheet1.f15.should be_close(0.400576036866359,0.0400576036866359)
  end

  it 'cell g15 should equal 0.359220704410797' do
    sheet1.g15.should be_close(0.359220704410797,0.0359220704410797)
  end

  it 'cell h15 should equal 0.317865371955234' do
    sheet1.h15.should be_close(0.317865371955234,0.0317865371955234)
  end

  it 'cell i15 should equal 0.276510039499671' do
    sheet1.i15.should be_close(0.276510039499671,0.0276510039499671)
  end

  it 'cell j15 should equal 0.235154707044108' do
    sheet1.j15.should be_close(0.235154707044108,0.0235154707044108)
  end

  it 'cell k15 should equal 0.193799374588545' do
    sheet1.k15.should be_close(0.193799374588545,0.0193799374588545)
  end

  it 'cell l15 should equal 0.152444042132982' do
    sheet1.l15.should be_close(0.152444042132982,0.0152444042132982)
  end

  it 'cell m15 should equal 0.111088709677419' do
    sheet1.m15.should be_close(0.111088709677419,0.0111088709677419)
  end

  it 'cell f17 should equal 0.168181818181818' do
    sheet1.f17.should be_close(0.168181818181818,0.0168181818181818)
  end

  it 'cell g17 should equal 0.254058441558441' do
    sheet1.g17.should be_close(0.254058441558441,0.0254058441558441)
  end

  it 'cell h17 should equal 0.339935064935065' do
    sheet1.h17.should be_close(0.339935064935065,0.0339935064935065)
  end

  it 'cell i17 should equal 0.425811688311688' do
    sheet1.i17.should be_close(0.425811688311688,0.0425811688311688)
  end

  it 'cell j17 should equal 0.511688311688312' do
    sheet1.j17.should be_close(0.511688311688312,0.0511688311688312)
  end

  it 'cell k17 should equal 0.597564935064935' do
    sheet1.k17.should be_close(0.597564935064935,0.0597564935064935)
  end

  it 'cell l17 should equal 0.683441558441559' do
    sheet1.l17.should be_close(0.683441558441559,0.0683441558441559)
  end

  it 'cell m17 should equal 0.769318181818182' do
    sheet1.m17.should be_close(0.769318181818182,0.0769318181818182)
  end

  it 'cell f18 should equal 2013.0' do
    sheet1.f18.should be_close(2013.0,201.3)
  end

  it 'cell g18 should equal 2014.0' do
    sheet1.g18.should be_close(2014.0,201.4)
  end

  it 'cell h18 should equal 2015.0' do
    sheet1.h18.should be_close(2015.0,201.5)
  end

  it 'cell i18 should equal 2016.0' do
    sheet1.i18.should be_close(2016.0,201.6)
  end

  it 'cell j18 should equal 2017.0' do
    sheet1.j18.should be_close(2017.0,201.7)
  end

  it 'cell k18 should equal 2018.0' do
    sheet1.k18.should be_close(2018.0,201.8)
  end

  it 'cell l18 should equal 2019.0' do
    sheet1.l18.should be_close(2019.0,201.9)
  end

  it 'cell m18 should equal 2020.0' do
    sheet1.m18.should be_close(2020.0,202.0)
  end

  it 'cell d19 should equal 2207.0' do
    sheet1.d19.should be_close(2207.0,220.7)
  end

  it 'cell f19 should equal 2091.7946' do
    sheet1.f19.should be_close(2091.7946,209.17946)
  end

  it 'cell g19 should equal 2055.39737396' do
    sheet1.g19.should be_close(2055.39737396,205.539737396)
  end

  it 'cell h19 should equal 2019.6334596531' do
    sheet1.h19.should be_close(2019.6334596531,201.96334596531)
  end

  it 'cell i19 should equal 1984.49183745513' do
    sheet1.i19.should be_close(1984.49183745513,198.449183745513)
  end

  it 'cell j19 should equal 1949.96167948341' do
    sheet1.j19.should be_close(1949.96167948341,194.996167948341)
  end

  it 'cell k19 should equal 1916.0323462604' do
    sheet1.k19.should be_close(1916.0323462604,191.60323462604)
  end

  it 'cell l19 should equal 1882.69338343547' do
    sheet1.l19.should be_close(1882.69338343547,188.269338343547)
  end

  it 'cell m19 should equal 1849.93451856369' do
    sheet1.m19.should be_close(1849.93451856369,184.993451856369)
  end

  it 'cell f20 should equal 848.707567741935' do
    sheet1.f20.should be_close(848.707567741935,84.8707567741935)
  end

  it 'cell g20 should equal 763.150624836641' do
    sheet1.g20.should be_close(763.150624836641,76.3150624836641)
  end

  it 'cell h20 should equal 677.1799459732' do
    sheet1.h20.should be_close(677.1799459732,67.71799459732)
  end

  it 'cell i20 should equal 590.793879521034' do
    sheet1.i20.should be_close(590.793879521034,59.0793879521034)
  end

  it 'cell j20 should equal 503.990767997985' do
    sheet1.j20.should be_close(503.990767997985,50.3990767997985)
  end

  it 'cell k20 should equal 416.768948050898' do
    sheet1.k20.should be_close(416.768948050898,41.6768948050898)
  end

  it 'cell l20 should equal 329.126750436129' do
    sheet1.l20.should be_close(329.126750436129,32.9126750436129)
  end

  it 'cell m20 should equal 241.0625' do
    sheet1.m20.should be_close(241.0625,24.10625)
  end

  it 'cell f21 should equal 1243.08703225806' do
    sheet1.f21.should be_close(1243.08703225806,124.308703225806)
  end

  it 'cell g21 should equal 1292.24674912336' do
    sheet1.g21.should be_close(1292.24674912336,129.224674912336)
  end

  it 'cell h21 should equal 1342.4535136799' do
    sheet1.h21.should be_close(1342.4535136799,134.24535136799)
  end

  it 'cell i21 should equal 1393.6979579341' do
    sheet1.i21.should be_close(1393.6979579341,139.36979579341)
  end

  it 'cell j21 should equal 1445.97091148543' do
    sheet1.j21.should be_close(1445.97091148543,144.597091148543)
  end

  it 'cell k21 should equal 1499.2633982095' do
    sheet1.k21.should be_close(1499.2633982095,149.92633982095)
  end

  it 'cell l21 should equal 1553.56663299934' do
    sheet1.l21.should be_close(1553.56663299934,155.356663299934)
  end

  it 'cell m21 should equal 1608.87201856369' do
    sheet1.m21.should be_close(1608.87201856369,160.887201856369)
  end

  it 'cell f23 should equal 2013.0' do
    sheet1.f23.should be_close(2013.0,201.3)
  end

  it 'cell g23 should equal 2014.0' do
    sheet1.g23.should be_close(2014.0,201.4)
  end

  it 'cell h23 should equal 2015.0' do
    sheet1.h23.should be_close(2015.0,201.5)
  end

  it 'cell i23 should equal 2016.0' do
    sheet1.i23.should be_close(2016.0,201.6)
  end

  it 'cell j23 should equal 2017.0' do
    sheet1.j23.should be_close(2017.0,201.7)
  end

  it 'cell k23 should equal 2018.0' do
    sheet1.k23.should be_close(2018.0,201.8)
  end

  it 'cell l23 should equal 2019.0' do
    sheet1.l23.should be_close(2019.0,201.9)
  end

  it 'cell m23 should equal 2020.0' do
    sheet1.m23.should be_close(2020.0,202.0)
  end

  it 'cell g24 should equal 26.25' do
    sheet1.g24.should be_close(26.25,2.625)
  end

  it 'cell h24 should equal 27.5625' do
    sheet1.h24.should be_close(27.5625,2.75625)
  end

  it 'cell i24 should equal 28.940625' do
    sheet1.i24.should be_close(28.940625,2.8940625)
  end

  it 'cell j24 should equal 30.38765625' do
    sheet1.j24.should be_close(30.38765625,3.038765625)
  end

  it 'cell k24 should equal 31.9070390625' do
    sheet1.k24.should be_close(31.9070390625,3.19070390625)
  end

  it 'cell l24 should equal 33.502391015625' do
    sheet1.l24.should be_close(33.502391015625,3.3502391015625)
  end

  it 'cell m24 should equal 35.1775105664063' do
    sheet1.m24.should be_close(35.1775105664063,3.51775105664063)
  end

  it 'cell f25 should equal 31077.1758064516' do
    sheet1.f25.should be_close(31077.1758064516,3107.71758064516)
  end

  it 'cell g25 should equal 33921.4771644882' do
    sheet1.g25.should be_close(33921.4771644882,3392.14771644882)
  end

  it 'cell h25 should equal 37001.3749708021' do
    sheet1.h25.should be_close(37001.3749708021,3700.13749708021)
  end

  it 'cell i25 should equal 40334.4899638365' do
    sheet1.i25.should be_close(40334.4899638365,4033.44899638365)
  end

  it 'cell j25 should equal 43939.6670057184' do
    sheet1.j25.should be_close(43939.6670057184,4393.96670057184)
  end

  it 'cell k25 should equal 47837.0558116471' do
    sheet1.k25.should be_close(47837.0558116471,4783.70558116471)
  end

  it 'cell l25 should equal 52048.196807572' do
    sheet1.l25.should be_close(52048.196807572,5204.8196807572)
  end

  it 'cell m25 should equal 56596.1124330197' do
    sheet1.m25.should be_close(56596.1124330197,5659.61124330197)
  end

  it 'cell f27 should equal 2013.0' do
    sheet1.f27.should be_close(2013.0,201.3)
  end

  it 'cell g27 should equal 2014.0' do
    sheet1.g27.should be_close(2014.0,201.4)
  end

  it 'cell h27 should equal 2015.0' do
    sheet1.h27.should be_close(2015.0,201.5)
  end

  it 'cell i27 should equal 2016.0' do
    sheet1.i27.should be_close(2016.0,201.6)
  end

  it 'cell j27 should equal 2017.0' do
    sheet1.j27.should be_close(2017.0,201.7)
  end

  it 'cell k27 should equal 2018.0' do
    sheet1.k27.should be_close(2018.0,201.8)
  end

  it 'cell l27 should equal 2019.0' do
    sheet1.l27.should be_close(2019.0,201.9)
  end

  it 'cell m27 should equal 2020.0' do
    sheet1.m27.should be_close(2020.0,202.0)
  end

  it 'cell f28 should equal 1084.45572580645' do
    sheet1.f28.should be_close(1084.45572580645,108.445572580645)
  end

  it 'cell g28 should equal 1065.58619617742' do
    sheet1.g28.should be_close(1065.58619617742,106.558619617742)
  end

  it 'cell h28 should equal 1047.04499636393' do
    sheet1.h28.should be_close(1047.04499636393,104.704499636393)
  end

  it 'cell i28 should equal 1028.8264134272' do
    sheet1.i28.should be_close(1028.8264134272,102.88264134272)
  end

  it 'cell j28 should equal 1010.92483383357' do
    sheet1.j28.should be_close(1010.92483383357,101.092483383357)
  end

  it 'cell k28 should equal 993.334741724863' do
    sheet1.k28.should be_close(993.334741724863,99.3334741724863)
  end

  it 'cell l28 should equal 976.05071721885' do
    sheet1.l28.should be_close(976.05071721885,97.605071721885)
  end

  it 'cell m28 should equal 959.067434739242' do
    sheet1.m28.should be_close(959.067434739242,95.9067434739242)
  end

  it 'cell b29 should equal "Leakage sectors"' do
    sheet1.b29.should == "Leakage sectors"
  end

  it 'cell f29 should equal 32.0516914516129' do
    sheet1.f29.should be_close(32.0516914516129,3.20516914516129)
  end

  it 'cell g29 should equal 39.3674900254435' do
    sheet1.g29.should be_close(39.3674900254435,3.93674900254435)
  end

  it 'cell h29 should equal 46.418994838801' do
    sheet1.h29.should be_close(46.418994838801,4.6418994838801)
  end

  it 'cell i29 should equal 53.2131883833735' do
    sheet1.i29.should be_close(53.2131883833735,5.32131883833735)
  end

  it 'cell j29 should equal 59.7568901777175' do
    sheet1.j29.should be_close(59.7568901777175,5.97568901777175)
  end

  it 'cell k29 should equal 66.0567603247034' do
    sheet1.k29.should be_close(66.0567603247034,6.60567603247034)
  end

  it 'cell l29 should equal 72.1193029945039' do
    sheet1.l29.should be_close(72.1193029945039,7.21193029945039)
  end

  it 'cell m29 should equal 77.9508698346395' do
    sheet1.m29.should be_close(77.9508698346395,7.79508698346395)
  end

  it 'cell f30 should equal 137.364391935484' do
    sheet1.f30.should be_close(137.364391935484,13.7364391935484)
  end

  it 'cell g30 should equal 212.102395239124' do
    sheet1.g30.should be_close(212.102395239124,21.2102395239124)
  end

  it 'cell h30 should equal 284.197927584496' do
    sheet1.h30.should be_close(284.197927584496,28.4197927584496)
  end

  it 'cell i30 should equal 353.720319283066' do
    sheet1.i30.should be_close(353.720319283066,35.3720319283066)
  end

  it 'cell j30 should equal 420.73728798597' do
    sheet1.j30.should be_close(420.73728798597,42.073728798597)
  end

  it 'cell k30 should equal 485.314973814147' do
    sheet1.k30.should be_close(485.314973814147,48.5314973814147)
  end

  it 'cell l30 should equal 547.517973754193' do
    sheet1.l30.should be_close(547.517973754193,54.7517973754193)
  end

  it 'cell m30 should equal 607.409375334853' do
    sheet1.m30.should be_close(607.409375334853,60.7409375334853)
  end

  it 'cell f31 should equal 1253.87180919355' do
    sheet1.f31.should be_close(1253.87180919355,125.387180919355)
  end

  it 'cell g31 should equal 1317.05608144199' do
    sheet1.g31.should be_close(1317.05608144199,131.705608144199)
  end

  it 'cell h31 should equal 1377.66191878723' do
    sheet1.h31.should be_close(1377.66191878723,137.766191878723)
  end

  it 'cell i31 should equal 1435.75992109364' do
    sheet1.i31.should be_close(1435.75992109364,143.575992109364)
  end

  it 'cell j31 should equal 1491.41901199725' do
    sheet1.j31.should be_close(1491.41901199725,149.141901199725)
  end

  it 'cell k31 should equal 1544.70647586371' do
    sheet1.k31.should be_close(1544.70647586371,154.470647586371)
  end

  it 'cell l31 should equal 1595.68799396755' do
    sheet1.l31.should be_close(1595.68799396755,159.568799396755)
  end

  it 'cell m31 should equal 1644.42767990873' do
    sheet1.m31.should be_close(1644.42767990873,164.442767990873)
  end

  it 'cell f34 should equal 27111.3931451613' do
    sheet1.f34.should be_close(27111.3931451613,2711.13931451613)
  end

  it 'cell g34 should equal 27971.6376496573' do
    sheet1.g34.should be_close(27971.6376496573,2797.16376496573)
  end

  it 'cell h34 should equal 28859.1777122809' do
    sheet1.h34.should be_close(28859.1777122809,2885.91777122809)
  end

  it 'cell i34 should equal 29774.8794210916' do
    sheet1.i34.should be_close(29774.8794210916,2977.48794210916)
  end

  it 'cell j34 should equal 30719.6363451228' do
    sheet1.j34.should be_close(30719.6363451228,3071.96363451228)
  end

  it 'cell k34 should equal 31694.3704063535' do
    sheet1.k34.should be_close(31694.3704063535,3169.43704063535)
  end

  it 'cell l34 should equal 32700.0327793471' do
    sheet1.l34.should be_close(32700.0327793471,3270.00327793471)
  end

  it 'cell m34 should equal 33737.6048194358' do
    sheet1.m34.should be_close(33737.6048194358,3373.76048194358)
  end

  it 'cell f35 should equal 4235.40208467742' do
    sheet1.f35.should be_close(4235.40208467742,423.540208467742)
  end

  it 'cell g35 should equal 6601.08448819491' do
    sheet1.g35.should be_close(6601.08448819491,660.108448819491)
  end

  it 'cell h35 should equal 9112.62892429212' do
    sheet1.h35.should be_close(9112.62892429212,911.262892429212)
  end

  it 'cell i35 should equal 11776.910045309' do
    sheet1.i35.should be_close(11776.910045309,1177.6910045309)
  end

  it 'cell j35 should equal 14601.0919161644' do
    sheet1.j35.should be_close(14601.0919161644,1460.10919161644)
  end

  it 'cell k35 should equal 17592.6394591267' do
    sheet1.k35.should be_close(17592.6394591267,1759.26394591267)
  end

  it 'cell l35 should equal 20759.3303334919' do
    sheet1.l35.should be_close(20759.3303334919,2075.93303334919)
  end

  it 'cell m35 should equal 24109.2672662446' do
    sheet1.m35.should be_close(24109.2672662446,2410.92672662446)
  end

  it 'cell f41 should equal 0.135574229691877' do
    sheet1.f41.should be_close(0.135574229691877,0.0135574229691877)
  end

  it 'cell f43 should equal 0.122016806722689' do
    sheet1.f43.should be_close(0.122016806722689,0.0122016806722689)
  end

  it 'cell f45 should equal 2013.0' do
    sheet1.f45.should be_close(2013.0,201.3)
  end

  it 'cell g45 should equal 2014.0' do
    sheet1.g45.should be_close(2014.0,201.4)
  end

  it 'cell h45 should equal 2015.0' do
    sheet1.h45.should be_close(2015.0,201.5)
  end

  it 'cell i45 should equal 2016.0' do
    sheet1.i45.should be_close(2016.0,201.6)
  end

  it 'cell j45 should equal 2017.0' do
    sheet1.j45.should be_close(2017.0,201.7)
  end

  it 'cell k45 should equal 2018.0' do
    sheet1.k45.should be_close(2018.0,201.8)
  end

  it 'cell l45 should equal 2019.0' do
    sheet1.l45.should be_close(2019.0,201.9)
  end

  it 'cell m45 should equal 2020.0' do
    sheet1.m45.should be_close(2020.0,202.0)
  end

  it 'cell f46 should equal 152.993434197398' do
    sheet1.f46.should be_close(152.993434197398,15.2993434197398)
  end

  it 'cell g46 should equal 160.702977332249' do
    sheet1.g46.should be_close(160.702977332249,16.0702977332249)
  end

  it 'cell h46 should equal 168.09790807387' do
    sheet1.h46.should be_close(168.09790807387,16.809790807387)
  end

  it 'cell i46 should equal 175.186840792266' do
    sheet1.i46.should be_close(175.186840792266,17.5186840792266)
  end

  it 'cell j46 should equal 181.978185329413' do
    sheet1.j46.should be_close(181.978185329413,18.1978185329413)
  end

  it 'cell k46 should equal 188.480151508749' do
    sheet1.k46.should be_close(188.480151508749,18.8480151508749)
  end

  it 'cell l46 should equal 194.700753549654' do
    sheet1.l46.should be_close(194.700753549654,19.4700753549654)
  end

  it 'cell m46 should equal 200.647814388864' do
    sheet1.m46.should be_close(200.647814388864,20.0647814388864)
  end

end

