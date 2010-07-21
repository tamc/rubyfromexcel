# coding: utf-8
require_relative '../spreadsheet'
# OLD UK
describe 'Sheet2' do
  def sheet2; $spreadsheet ||= Spreadsheet.new; $spreadsheet.sheet2; end

  it 'cell d5 should equal 160.40945' do
    sheet2.d5.should be_close(160.40945,16.040945)
  end

  it 'cell e5 should equal 157.61832557' do
    sheet2.e5.should be_close(157.61832557,15.761832557)
  end

  it 'cell f5 should equal 154.875766705082' do
    sheet2.f5.should be_close(154.875766705082,15.4875766705082)
  end

  it 'cell g5 should equal 152.180928364414' do
    sheet2.g5.should be_close(152.180928364414,15.2180928364414)
  end

  it 'cell h5 should equal 149.532980210873' do
    sheet2.h5.should be_close(149.532980210873,14.9532980210873)
  end

  it 'cell i5 should equal 146.931106355204' do
    sheet2.i5.should be_close(146.931106355204,14.6931106355204)
  end

  it 'cell j5 should equal 144.374505104623' do
    sheet2.j5.should be_close(144.374505104623,14.4374505104623)
  end

  it 'cell c6 should equal 134.166975' do
    sheet2.c6.should be_close(134.166975,13.4166975)
  end

  it 'cell d6 should equal 131.832469635' do
    sheet2.d6.should be_close(131.832469635,13.1832469635)
  end

  it 'cell e6 should equal 129.538584663351' do
    sheet2.e6.should be_close(129.538584663351,12.9538584663351)
  end

  it 'cell f6 should equal 127.284613290209' do
    sheet2.f6.should be_close(127.284613290209,12.7284613290209)
  end

  it 'cell g6 should equal 125.069861018959' do
    sheet2.g6.should be_close(125.069861018959,12.5069861018959)
  end

  it 'cell h6 should equal 122.893645437229' do
    sheet2.h6.should be_close(122.893645437229,12.2893645437229)
  end

  it 'cell i6 should equal 120.755296006621' do
    sheet2.i6.should be_close(120.755296006621,12.0755296006621)
  end

  it 'cell j6 should equal 118.654153856106' do
    sheet2.j6.should be_close(118.654153856106,11.8654153856106)
  end

  it 'cell l6 should equal 0.0174000000000001' do
    sheet2.l6.should be_close(0.0174000000000001,0.00174000000000001)
  end

  it 'cell c7 should equal 297.416975' do
    sheet2.c7.should be_close(297.416975,29.7416975)
  end

  it 'cell d7 should equal 292.241919635' do
    sheet2.d7.should be_close(292.241919635,29.2241919635)
  end

  it 'cell e7 should equal 287.156910233351' do
    sheet2.e7.should be_close(287.156910233351,28.7156910233351)
  end

  it 'cell f7 should equal 282.160379995291' do
    sheet2.f7.should be_close(282.160379995291,28.2160379995291)
  end

  it 'cell g7 should equal 277.250789383373' do
    sheet2.g7.should be_close(277.250789383373,27.7250789383373)
  end

  it 'cell h7 should equal 272.426625648102' do
    sheet2.h7.should be_close(272.426625648102,27.2426625648102)
  end

  it 'cell i7 should equal 267.686402361825' do
    sheet2.i7.should be_close(267.686402361825,26.7686402361825)
  end

  it 'cell j7 should equal 263.028658960729' do
    sheet2.j7.should be_close(263.028658960729,26.3028658960729)
  end

  it 'cell c9 should equal 2013.0' do
    sheet2.c9.should be_close(2013.0,201.3)
  end

  it 'cell d9 should equal 2014.0' do
    sheet2.d9.should be_close(2014.0,201.4)
  end

  it 'cell e9 should equal 2015.0' do
    sheet2.e9.should be_close(2015.0,201.5)
  end

  it 'cell f9 should equal 2016.0' do
    sheet2.f9.should be_close(2016.0,201.6)
  end

  it 'cell g9 should equal 2017.0' do
    sheet2.g9.should be_close(2017.0,201.7)
  end

  it 'cell h9 should equal 2018.0' do
    sheet2.h9.should be_close(2018.0,201.8)
  end

  it 'cell i9 should equal 2019.0' do
    sheet2.i9.should be_close(2019.0,201.9)
  end

  it 'cell j9 should equal 2020.0' do
    sheet2.j9.should be_close(2020.0,202.0)
  end

  it 'cell d10 should equal 1.0' do
    sheet2.d10.should be_close(1.0,0.1)
  end

  it 'cell e10 should equal 1.0' do
    sheet2.e10.should be_close(1.0,0.1)
  end

  it 'cell f10 should equal 1.0' do
    sheet2.f10.should be_close(1.0,0.1)
  end

  it 'cell g10 should equal 1.0' do
    sheet2.g10.should be_close(1.0,0.1)
  end

  it 'cell h10 should equal 1.0' do
    sheet2.h10.should be_close(1.0,0.1)
  end

  it 'cell i10 should equal 1.0' do
    sheet2.i10.should be_close(1.0,0.1)
  end

  it 'cell j10 should equal 1.0' do
    sheet2.j10.should be_close(1.0,0.1)
  end

  it 'cell d11 should equal 0.314285714285714' do
    sheet2.d11.should be_close(0.314285714285714,0.0314285714285714)
  end

  it 'cell e11 should equal 0.428571428571429' do
    sheet2.e11.should be_close(0.428571428571429,0.0428571428571429)
  end

  it 'cell f11 should equal 0.542857142857143' do
    sheet2.f11.should be_close(0.542857142857143,0.0542857142857143)
  end

  it 'cell g11 should equal 0.657142857142857' do
    sheet2.g11.should be_close(0.657142857142857,0.0657142857142857)
  end

  it 'cell h11 should equal 0.771428571428572' do
    sheet2.h11.should be_close(0.771428571428572,0.0771428571428572)
  end

  it 'cell i11 should equal 0.885714285714286' do
    sheet2.i11.should be_close(0.885714285714286,0.0885714285714286)
  end

  it 'cell l11 should equal 0.114285714285714' do
    sheet2.l11.should be_close(0.114285714285714,0.0114285714285714)
  end

  it 'cell c12 should equal 0.639114142694781' do
    sheet2.c12.should be_close(0.639114142694781,0.0639114142694781)
  end

  it 'cell d12 should equal 0.690669265166955' do
    sheet2.d12.should be_close(0.690669265166955,0.0690669265166955)
  end

  it 'cell e12 should equal 0.742224387639129' do
    sheet2.e12.should be_close(0.742224387639129,0.0742224387639129)
  end

  it 'cell f12 should equal 0.793779510111303' do
    sheet2.f12.should be_close(0.793779510111303,0.0793779510111303)
  end

  it 'cell g12 should equal 0.845334632583478' do
    sheet2.g12.should be_close(0.845334632583478,0.0845334632583478)
  end

  it 'cell h12 should equal 0.896889755055652' do
    sheet2.h12.should be_close(0.896889755055652,0.0896889755055652)
  end

  it 'cell i12 should equal 0.948444877527826' do
    sheet2.i12.should be_close(0.948444877527826,0.0948444877527826)
  end

  it 'cell j12 should equal 1.0' do
    sheet2.j12.should be_close(1.0,0.1)
  end

  it 'cell c14 should equal 2013.0' do
    sheet2.c14.should be_close(2013.0,201.3)
  end

  it 'cell d14 should equal 2014.0' do
    sheet2.d14.should be_close(2014.0,201.4)
  end

  it 'cell e14 should equal 2015.0' do
    sheet2.e14.should be_close(2015.0,201.5)
  end

  it 'cell f14 should equal 2016.0' do
    sheet2.f14.should be_close(2016.0,201.6)
  end

  it 'cell g14 should equal 2017.0' do
    sheet2.g14.should be_close(2017.0,201.7)
  end

  it 'cell h14 should equal 2018.0' do
    sheet2.h14.should be_close(2018.0,201.8)
  end

  it 'cell i14 should equal 2019.0' do
    sheet2.i14.should be_close(2019.0,201.9)
  end

  it 'cell j14 should equal 2020.0' do
    sheet2.j14.should be_close(2020.0,202.0)
  end

  it 'cell d15 should equal 20.0' do
    sheet2.d15.should be_close(20.0,2.0)
  end

  it 'cell e15 should equal 20.0' do
    sheet2.e15.should be_close(20.0,2.0)
  end

  it 'cell f15 should equal 20.0' do
    sheet2.f15.should be_close(20.0,2.0)
  end

  it 'cell g15 should equal 20.0' do
    sheet2.g15.should be_close(20.0,2.0)
  end

  it 'cell h15 should equal 20.0' do
    sheet2.h15.should be_close(20.0,2.0)
  end

  it 'cell i15 should equal 20.0' do
    sheet2.i15.should be_close(20.0,2.0)
  end

  it 'cell j15 should equal 20.0' do
    sheet2.j15.should be_close(20.0,2.0)
  end

  it 'cell c17 should equal 2013.0' do
    sheet2.c17.should be_close(2013.0,201.3)
  end

  it 'cell d17 should equal 2014.0' do
    sheet2.d17.should be_close(2014.0,201.4)
  end

  it 'cell e17 should equal 2015.0' do
    sheet2.e17.should be_close(2015.0,201.5)
  end

  it 'cell f17 should equal 2016.0' do
    sheet2.f17.should be_close(2016.0,201.6)
  end

  it 'cell g17 should equal 2017.0' do
    sheet2.g17.should be_close(2017.0,201.7)
  end

  it 'cell h17 should equal 2018.0' do
    sheet2.h17.should be_close(2018.0,201.8)
  end

  it 'cell i17 should equal 2019.0' do
    sheet2.i17.should be_close(2019.0,201.9)
  end

  it 'cell j17 should equal 2020.0' do
    sheet2.j17.should be_close(2020.0,202.0)
  end

  it 'cell c18 should equal 3.265' do
    sheet2.c18.should be_close(3.265,0.3265)
  end

  it 'cell d18 should equal 3.208189' do
    sheet2.d18.should be_close(3.208189,0.3208189)
  end

  it 'cell e18 should equal 3.1523665114' do
    sheet2.e18.should be_close(3.1523665114,0.31523665114)
  end

  it 'cell f18 should equal 3.09751533410164' do
    sheet2.f18.should be_close(3.09751533410164,0.309751533410164)
  end

  it 'cell g18 should equal 3.04361856728827' do
    sheet2.g18.should be_close(3.04361856728827,0.304361856728827)
  end

  it 'cell h18 should equal 2.99065960421746' do
    sheet2.h18.should be_close(2.99065960421746,0.299065960421746)
  end

  it 'cell i18 should equal 2.93862212710407' do
    sheet2.i18.should be_close(2.93862212710407,0.293862212710407)
  end

  it 'cell j18 should equal 2.88749010209246' do
    sheet2.j18.should be_close(2.88749010209246,0.288749010209246)
  end

  it 'cell c19 should equal 0.5366679' do
    sheet2.c19.should be_close(0.5366679,0.05366679)
  end

  it 'cell d19 should equal 0.828661237705715' do
    sheet2.d19.should be_close(0.828661237705715,0.0828661237705715)
  end

  it 'cell e19 should equal 1.11033072568587' do
    sheet2.e19.should be_close(1.11033072568587,0.111033072568587)
  end

  it 'cell f19 should equal 1.38194723000798' do
    sheet2.f19.should be_close(1.38194723000798,0.138194723000798)
  end

  it 'cell g19 should equal 1.64377531624918' do
    sheet2.g19.should be_close(1.64377531624918,0.164377531624918)
  end

  it 'cell h19 should equal 1.89607338674582' do
    sheet2.h19.should be_close(1.89607338674582,0.189607338674582)
  end

  it 'cell i19 should equal 2.13909381497444' do
    sheet2.i19.should be_close(2.13909381497444,0.213909381497444)
  end

  it 'cell j19 should equal 2.37308307712212' do
    sheet2.j19.should be_close(2.37308307712212,0.237308307712212)
  end

  it 'cell c20 should equal 3.8016679' do
    sheet2.c20.should be_close(3.8016679,0.38016679)
  end

  it 'cell d20 should equal 4.03685023770571' do
    sheet2.d20.should be_close(4.03685023770571,0.403685023770571)
  end

  it 'cell e20 should equal 4.26269723708587' do
    sheet2.e20.should be_close(4.26269723708587,0.426269723708587)
  end

  it 'cell f20 should equal 4.47946256410962' do
    sheet2.f20.should be_close(4.47946256410962,0.447946256410962)
  end

  it 'cell g20 should equal 4.68739388353745' do
    sheet2.g20.should be_close(4.68739388353745,0.468739388353745)
  end

  it 'cell h20 should equal 4.88673299096328' do
    sheet2.h20.should be_close(4.88673299096328,0.488673299096328)
  end

  it 'cell i20 should equal 5.07771594207851' do
    sheet2.i20.should be_close(5.07771594207851,0.507771594207851)
  end

  it 'cell j20 should equal 5.26057317921459' do
    sheet2.j20.should be_close(5.26057317921459,0.526057317921459)
  end

end

