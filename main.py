import docxtpl
from PIL import ImageFont, ImageDraw, Image
from sympy import *
import numpy
import math
from sympy.parsing.latex import parse_latex
from lxml import etree
import latex2mathml.converter
from docxtpl import DocxTemplate
from docxtpl import InlineImage
from docx.shared import Mm

econd = [100,220,330,470]
e24 = [100,110,120,130,150,160,180,200,220,240,270,300,330,360,390,430,470,510,560,620,680,750,820,910]

Om = {0: "Ом",3: "кОм"}
Far = {-3: "мФ",-6: "мкФ",-9: "нФ", -12: "пФ"}



def round_to_n(x,n):
    if  0 < abs(x) < 1:
        nn = abs(floor(math.log10(abs(x))))
        return round(x,nn-1+n)
    else:
        return round(x, n)


def bligg(arr, x,bigger = False):
    count = 0
    razn = []
    for j in arr:
        razn.append(j-x)
        count+=1
    count = 0
    min = 9999999999
    minind = 0
    for j in razn:
        j = j if bigger else abs(j)
        if (j < min) and (j>=0) :
            min = j
            minind = count
        count+=1
    return(minind)


def nomin(xc,bigger = False):
    count = 0
    xc = xc*1.1 if bigger else xc
    shag = 1 if xc < 100 else -1
    tmp = xc
    while not( 100 <= tmp < 1000 ):
        count += 1
        tmp = xc * pow(10,count*shag)
    #count-=1
    ind = bligg(econd, tmp, bigger) if bigger else bligg(e24, tmp, bigger)
    count = count if not bigger else (count if tmp < econd[ind] else count-1 )
    resul = econd[ind]* pow(10,-count*shag) if bigger else e24[ind]* pow(10,-count*shag)
    #print(tmp,-count*shag, count, shag, e24[ind]* pow(10,-count*shag))
    return(resul)



def numberss(ttm):
    gh = ttm
    if(gh>1000):
        cck = 1
        while( not (gh>=1 and gh<=1000)):
            gh=ttm/pow(10,cck*3)
            cck+=1
        cck-=1
        return(gh,cck*3)
    elif gh<1 :
        cck = -1
        while (not (gh >= 1 and gh <= 1000)):
            gh = ttm * pow(10,-cck*3)
            cck -= 1
        cck+=1
        if cck*3 < -12:
            return (gh * pow(10,-(cck*3+12)),-12)
        return (gh, cck * 3)
    return (ttm,0)


def latex_to_word(latex_input):
    mathml = latex2mathml.converter.convert(latex_input)
    tree = etree.fromstring(mathml)
    xslt = etree.parse('MML2OMML.XSL')
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    xml = str(new_dom)
    xml = xml.split("\n")[1]


    xml = xml.replace('<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mml="http://www.w3.org/1998/Math/MathML">', '<m:oMathPara><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">')
    xml = xml + r"</m:oMathPara>"
    xml = xml.replace("<m:t>","""<w:rPr><w:rFonts w:ascii="GOST type B" w:hAnsi="GOST type B"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><m:t>""")
    return xml


def late_func(text,func,znach,edinic = "",precis1 = 3, precis2 = 3, prived = False):
    eqq = text
    znach = list(znach.items())
    zqq = eqq
    mass = {}
    for j in znach:
        eqq = eqq.replace(j[0],numpy.format_float_positional(j[1], trim='-', precision=precis1,fractional=False))
    zqq = parse_latex(eqq)
    resul = zqq.evalf()
    if prived:
        rr = numberss(resul)
        resull = rr[0]
        if rr[1] == 0:
            vivod = numpy.format_float_positional(resull, trim='-', precision=precis2,fractional=True)
        else:
            vivod = numpy.format_float_positional(resull, trim='-', precision=precis2,fractional=True) + r"\cdot 10^{  "+str(rr[1])+"  }"
    else:
        vivod = numpy.format_float_positional(resul, trim='-', precision=precis2,fractional=True)
    latt = func + text + " = " + eqq + " = " + vivod + ((r"\mathrm{  " + edinic + "  }") if edinic != "" else "" )
    return ([latt, round_to_n(resul,precis2)])





def cordcrug(x,y,r):
    return [x-r,y-r,x+r,y+r]


class priamai:
    k = 0
    b = 0

    def __init__(self,coord):  #[x1,y1,x2,y2]
        self.k = (coord[1]-coord[3])/(coord[0]-coord[2])
        self.b = coord[3] - self.k * coord[2]
    def y(self,x):
        return (self.k * x +self.b)
    def x(self,y):
        return ((y-self.b)/self.k)

class lineika: #[x11,x12,x21,x22]
    x11 = 0
    x12 = 0
    x21 = 0
    x22 = 0
    def __init__(self,mass):
        self.x11 = mass[0]
        self.x12 = mass[1]
        self.x21 = mass[2]
        self.x22 = mass[3]
    def OneToTwo(self,xx):
        return (self.x21 + (self.x22-self.x21)*((xx-self.x11)/(self.x12-self.x11)))
    def TwoToOne(self,xx):
        return (self.x11 + (self.x12-self.x11)*((xx-self.x21)/(self.x22-self.x21)))



class tranz3102:
    pathim = "C:\\Users\\User\\PycharmProjects\\analog\\вах3102.jpg"
    img = 0

    Ibb = 0
    Ikk = 0
    delt = 0
    h21 = 0
    Ik0 = 0
    Uke0 = 0
    Ube0 = 0
    Ibe0 = 0
    Uvix = 0
    ind = 0
    E = 0
    Ikcord = ([633,1455,1561,1455],[613,1305,1467,1295],[607,1155,1325,1121],[431,1005,1181,935],[417,833,1007,723],[395,665,911,547],[325,525,691,413],[339,373,665,253],[369,265,581,143])
    Ikfunc = 0
    Ibcord = ([3089,97,3029,1307],[3383,93,3183,1345])
    Ibfunc = 0

    linIk = lineika([1605,65,0,50])
    linUke = lineika([238,1748,0,40])

    linIb = lineika([1605,75,0,100])
    linUbe = lineika([2250,3470,0,0.8])

    def __init__(self, typ, Ik0, Uke0, Uvix,E):
        self.img = Image.open(self.pathim)
        toki = {1:25,2:15,3:10}
        h21i = {1:200,2:300,3:400}
        self.delt = toki[typ]
        self.h21 = h21i[typ]
        self.Ibb = [I * self.delt for I in range(1,10)]
        self.Ikk = [I*5 for I in range(2,10)]
        ind = bligg(self.Ikk,Ik0)+1
        self.Ibe0 = self.Ibb[ind]
        self.ind = ind
        self.Uvix = Uvix
        self.E = E
        self.Ikfunc = list(map(priamai,self.Ikcord))
        self.Ibfunc = list(map(priamai,self.Ibcord))
        self.Ik0 =  round(self.linIk.OneToTwo(self.Ikfunc[ind].y(self.linUke.TwoToOne(Uke0))),3)
        self.Uke0 = Uke0
        self.Ube0 = self.linUbe.OneToTwo(self.Ibfunc[1].x(self.linIb.TwoToOne(self.Ibe0)))
    def yparam(self):
        dIb = self.delt
        lat_dIb = r"\Delta I_{b} = " + numpy.format_float_positional(dIb,precision=3,trim='-') + r"\mathrm{мкА}"
        dUbe = self.linUbe.OneToTwo(self.Ibfunc[1].x(self.linIb.TwoToOne(self.Ibe0+dIb))) - self.linUbe.OneToTwo(self.Ibfunc[1].x(self.linIb.TwoToOne(self.Ibe0)))
        lat_dUbe = r"\Delta U_{be} = " + numpy.format_float_positional(dUbe,precision=3,trim='-') + r"\mathrm{В}"
        dIk = self.linIk.TwoToOne(self.Ikfunc[self.ind].y(self.linUke.OneToTwo(30))) - self.linIk.TwoToOne(self.Ikfunc[self.ind].y(self.linUke.OneToTwo(5)))
        lat_dIk = r"\Delta I_{k} = " + numpy.format_float_positional(dIk,precision=3,trim='-') + r"\mathrm{мА}"
        dUke = 25
        lat_dUke = r"\Delta U_{ke} = " + numpy.format_float_positional(dUke,precision=3,trim='-') + r"\mathrm{В}"
        dIbos = self.linIb.OneToTwo(self.Ibfunc[0].y(self.linUbe.TwoToOne(self.Ube0))) - self.linIb.OneToTwo(self.Ibfunc[1].y(self.linUbe.TwoToOne(self.Ube0)))
        lat_dIbos = r"\Delta I_{bos} = " + numpy.format_float_positional(dIbos,precision=3,trim='-') + r"\mathrm{мкА}"
        dIks = self.linIk.OneToTwo(self.Ikfunc[self.ind+1].y(self.linUke.TwoToOne(5))) - self.linIk.OneToTwo(self.Ikfunc[self.ind].y(self.linUke.TwoToOne(5)))
        lat_dIks = r"\Delta I_{ks} = " + numpy.format_float_positional(dIks,precision=3,trim='-') + r"\mathrm{мА}"
        latt_y11 = late_func(r"\frac {\Delta I_{b}} {\Delta U_{be}}","Y_{11} = ",{r"\Delta U_{be}":dUbe, r"\Delta I_{b}":dIb/1000 },"мСм")
        latt_y12 = late_func(r"\frac {\Delta I_{bos}} {\Delta U_{ke}}","Y_{12} = ",{r"\Delta I_{bos}":dIbos/1000, r"\Delta U_{ke}":5 },"мСм")
        latt_y21 = late_func(r"\frac {\Delta I_{ks}} {\Delta U_{be}}", "Y_{21} = ",{r"\Delta I_{ks}": dIks, r"\Delta U_{be}":dUbe}, "мСм")
        latt_y22 = late_func(r"\frac {\Delta I_{k}} {\Delta U_{ke}}", "Y_{22} = ",{r"\Delta I_{k}": dIk, r"\Delta U_{ke}": dUke}, "мСм")
        return [((lat_dIb,dIb), (lat_dUbe,dUbe), (lat_dIk,dIk), (lat_dUke,dUke), (lat_dIks,dIks), (lat_dIbos,dIbos)),(latt_y11,latt_y12,latt_y21,latt_y22)]
    def drawWork(self,kpy):
        draw = ImageDraw.Draw(self.img)
        rab = [self.linUke.TwoToOne(self.Uke0), self.Ikfunc[self.ind].y(self.linUke.TwoToOne(self.Uke0))]
        post = priamai([rab[0],rab[1],self.linUke.TwoToOne(self.E),self.linIk.TwoToOne(0)])
        perem = priamai([rab[0],rab[1],self.linUke.TwoToOne(self.Uvix+self.Uke0),self.Ikfunc[0].y(self.Uvix+self.Uke0)])
        draw.ellipse(cordcrug(rab[0], rab[1], 20), (255, 100, 0)) # tochka
        draw.ellipse(cordcrug(self.linUbe.TwoToOne(self.Ube0), self.linIb.TwoToOne(self.Ibe0), 20), (255, 100, 0))  # tochka


        draw.line([self.linUke.TwoToOne(0) ,perem.y(self.linUke.TwoToOne(0)), perem.x(self.linIk.TwoToOne(0)),self.linIk.TwoToOne(0)], (255, 100, 0), 10) #perem
        draw.line([self.linUke.TwoToOne(0),post.y(self.linUke.TwoToOne(0)),self.linUke.TwoToOne(self.E),self.linIk.TwoToOne(0)],(255, 100, 0),10) #post


        draw.line([self.linUke.TwoToOne(self.Uvix+self.Uke0), self.linIk.TwoToOne(50), self.linUke.TwoToOne(self.Uvix+self.Uke0),self.linIk.TwoToOne(-5)], (207, 238, 0), 9) #ogranich bokovoe
        draw.line([self.linUke.TwoToOne(self.Uke0-self.Uvix), self.linIk.TwoToOne(50),self.linUke.TwoToOne(self.Uke0 - self.Uvix), self.linIk.TwoToOne(-5)], (207, 238, 0), 9) #ogranich bokovoe

        draw.line([self.linUke.TwoToOne(self.Uvix+self.Uke0) ,perem.y(self.linUke.TwoToOne(self.Uvix+self.Uke0)), self.linUke.TwoToOne(-5) ,perem.y(self.linUke.TwoToOne(self.Uvix+self.Uke0))], (207, 238, 0), 9) #ogranich goriz
        draw.line([self.linUke.TwoToOne(self.Uvix + self.Uke0), perem.y(self.linUke.TwoToOne(self.Uke0 - self.Uvix)), self.linUke.TwoToOne(-5),perem.y(self.linUke.TwoToOne(self.Uke0 - self.Uvix))], (207, 238, 0), 9) #ogranich goriz


        draw.line([rab[0],self.linIk.TwoToOne(-5),rab[0],rab[1]], (207, 238, 0), 5) # seredka vertical
        draw.line([self.linUke.TwoToOne(-5), rab[1], rab[0], rab[1]], (207, 238, 0), 5) # seredka gorizont

        draw.line([self.linUbe.TwoToOne(self.Ube0),self.linIb.TwoToOne(0), self.linUbe.TwoToOne(self.Ube0),self.linIb.TwoToOne(self.Ibe0)],(207, 238, 0), 9)
        draw.line([self.linUbe.TwoToOne(0), self.linIb.TwoToOne(self.Ibe0), self.linUbe.TwoToOne(self.Ube0),self.linIb.TwoToOne(self.Ibe0)], (207, 238, 0), 9)
        self.img.save("3102__"+ str(kpy) +".jpg")
        self.img = 0
        self.draw = 0





def raschet(E,Rn,Cn,Kooc,Mn,Mv,fn,fv,Uvix,kpy,typt):
    # E = 18
    # Rn = 9040
    # Cn  = 30000
    # Kooc = 10.56
    # Mn = 0.7
    # Mv = 0.7
    Mnp = sqrt(Mn)+0.02
    Mne = Mn/Mnp
    # fn = 100000
    # fv = 1000000
    # Uvix = 0.004233
    Uke0 = 3+1+Uvix


    Uvixl = "U_{вых} = " + str(Uvix) + r"\mathrm{В}"
    Ivixl = late_func(r" \frac {U_{вых} } {R_{н} } ", "I_{вых} = ", {"U_{вых}": Uvix,"R_{н}": Rn/1000 }, "мА")
    El = "E_{пит} = " + str(E) + r"\mathrm{В}"
    Pl = late_func(r"\frac {U_{вых}^{2} } {2 \cdot R_{н} }", "P_{вых} = ",{"U_{вых}": Uvix,"R_{н}": Rn/1000 }, "мВт")
    PPl = late_func(r"3 \cdot P_{вых}", "P_{pac} = ", {"P_{вых}":Pl[1]}, "мВт")
    Fgrl = late_func(r"3 \cdot f_{в} " , "f_{гр} = ", { "f_{в}": fv}, "Гц", prived=True)

    # print(Ivixl)
    # print(El)
    # print(Pl)
    # print(PPl)
    # print(Fgrl)

    tr = tranz3102(typt,Ivixl[1]*2,Uke0,Uvix,E)
    yy = tr.yparam()
    tr.drawWork(kpy)
    ypar = yy[1]
    y11 = ypar[0]
    y12 = ypar[1]
    y21 = ypar[2]
    y22 = ypar[3]
    deltparam = yy[0]
    Ik0 = tr.Ik0
    Uk0 = tr.Uke0
    Ib0 = tr.Ibe0
    Ube0 = tr.Ube0
    Ik0l = "I_{k0} = " + numpy.format_float_positional(Ik0,precision=3,trim='-') + r"\mathrm{мА}"
    Uk0l = "U_{ke0} = " + numpy.format_float_positional(Uk0,precision=3,trim='-') + r"\mathrm{В}"
    Ib0l = "I_{b0} = " + numpy.format_float_positional(Ib0,precision=3,trim='-') + r"\mathrm{мкА}"
    Ube0l = "U_{be0} = " + numpy.format_float_positional(Ube0,precision=3,trim='-') + r"\mathrm{В}"


    Re = late_func(r"\frac {0.2 \cdot E} {I_{k0} + I_{b0}}","R_{e} = ",{"E": E, "I_{k0}":Ik0/pow(10,3), "I_{b0}":Ib0/pow(10,6)},"Ом")
    Re[1] = nomin(Re[1])
    Rk = late_func(r"\frac {E - (I_{k0} + I_{b0}) \cdot R_{e} - U_{ke0}} {I_{k0}}","R_{k} = ",{"E":E,"U_{ke0}":Uk0,"I_{k0}":Ik0/pow(10,3),"R_{e}": Re[1], "I_{b0}":Ib0/pow(10,6)},"Ом")
    Rk[1] = nomin(Rk[1])
    Ku = late_func(r"\frac {Y_{21}} {Y_{22} + Y_{k} + Y_{н}}","K_{u} = ", {"Y_{21}": y21[1],"Y_{22}": y22[1], "Y_{k}": 1000/Rk[1], "Y_{н}": 1000/Rn })
    F = late_func(r"\frac {K_{u}}  {K_{ooc}}","F = ",{"K_{u}":Ku[1],"K_{ooc}": Kooc})
    Rooc = late_func(r"\frac {F - 1} {Y_{21}}", "R_{ooc} = ", {"F":F[1],"Y_{21}":y21[1]/1000}, "Ом")
    Rooc[1] = nomin(Rooc[1])
    Idel = late_func(r"3 \cdot I_{b0}", "I_{дел} = ", {"I_{b0}": Ib0},"мкА")
    R2 = late_func(r"\frac {U_{be0} + (R_{e} + R_{оос} ) \cdot ( I_{k0} + I_{b0}*10^{-3} )} {I_{дел}}", "R_{2} = ", {"U_{be0}": Ube0, "R_{e}": Re[1], "I_{k0}": Ik0/pow(10,3), "I_{b0}": Ib0/pow(10,3), "I_{дел}": Idel[1]/pow(10,3),"R_{оос}":Rooc[1]}, "кОм")
    R1 = late_func(r"\frac {E - ( U_{be0} + (R_{e} + R_{оос} ) \cdot ( I_{k0} + I_{b0}*10^{-3} ) )} { I_{дел} }", "R_{1} = ", {"E": E, "U_{be0}": Ube0, "R_{e}": Re[1], "I_{k0}": Ik0/pow(10,3), "I_{b0}": Ib0/pow(10,3), "I_{дел}": Idel[1]/pow(10,3),"R_{оос}":Rooc[1] }, "кОм" )
    R2[1] = nomin(R2[1])
    R1[1] = nomin(R1[1])
    Rvt = late_func(r"( 1 + \beta ) \cdot R_{ooc}", "R_{vt} = ", { r"\beta": tr.h21, "R_{ooc}": Rooc[1]/1000}, "кОм")
    Rdel = late_func(r"\frac {R_{1} \cdot R_{2}} {R_{1} + R_{2}}", "R_{дел} = ", {"R_{1}": R1[1],"R_{2}":R2[1]}, "кОм")
    Ri = late_func(r"\frac {1} { Y_{22}} ", "R_{i} = ",{"Y_{22}": y22[1]}, "кОм")
    an = late_func(r" \sqrt{(\frac {1} { M_{ср} })^2 - 1  } ", "a_{н} = ", {"M_{ср}":Mnp})
    Cp = late_func(r" \frac {1} {2 \cdot \pi \cdot f_{н} \cdot a_{н} \cdot (R_{н} + \frac {R_{k} \cdot R_{i}} { R_{k} + R_{i} }  )} ", "C_{p} = ", {"\pi": 3.1415 , "f_{н}":fn, "a_{н}": an[1],"R_{н}": Rn, "R_{k}": Rk[1], "R_{i}": Ri[1]*1000}, "Ф", precis1= 2, prived=True)


    Ce = late_func(r"\frac {1} {2 \cdot \pi \cdot f_{н} \cdot R_{e}} \cdot \sqrt{ \frac { ( M_{сэ} \cdot F )^2 - 1 } {1 - M_{сэ}} }", "C_{э} = ", { "M_{сэ}": Mne, "F": F[1], "R_{e}": Re[1], "\pi": 3.1415 , "f_{н}":fn }, "Ф", precis1= 3, prived=True)

    C0 = late_func(r"C_{транз} + C_{монтаж} + C_{н}", "C_{0} = ", {"C_{транз}": 5, "C_{монтаж}": 3,"C_{н}": Cn}, "пФ")
    Ye = late_func(r" Y_{22} + Y_{k} + Y_{н} +  ( 2 \cdot \pi \cdot f_{в} \cdot C_{н} ) ", "Y_{екв} = ", { "\pi": 3.1415 , "f_{в}":fv, "Y_{22}": y22[1]/pow(10,3),"Y_{k}": 1/Rk[1], "Y_{н}": 1/Rn, "C_{н}": Cn/pow(10,12)}, "См",precis2= 3, precis1=3 )
    av = late_func(r"\frac {2 \cdot \pi \cdot f_{в} \cdot C_{0}} { Y_{екв} }" , "a_{в} = " , { "\pi": 3.1415 , "f_{в}":fv, "C_{0}": C0[1]/pow(10,12), "Y_{екв}": Ye[1]})
    Mv = late_func(r"\frac {1} {\sqrt{ a_{в}^2 + 1 }}", "M_{в}", { "a_{в}": av[1]},precis2=3 )
    Uvx = late_func(r"\frac {U_{вых}} {K_{ooc}}", "U_{вх} = ", {"U_{вых}": Uvix, "K_{ooc}": Kooc},"В", precis2=3)
    Rvx = late_func(r" \frac {R_{дел} \cdot R_{vt}} {R_{дел} + R_{vt}} ", "R_{вх} = ", {"R_{дел}":Rdel[1], "R_{vt}":Rvt[1]},"кОм",precis2=2)
    contl = {}
    formul = (Uvixl,Ivixl,El,Pl,PPl,Fgrl,Ik0l,Ib0l,Uk0l,Ube0l,yy[0][0][0],yy[0][1][0],yy[0][2][0],yy[0][3][0],yy[0][4][0],yy[0][5][0], y11[0],y12[0],y21[0],y22[0],Re[0],Rk[0],Ku[0],F[0],Rooc[0],Idel[0],R1[0],R2[0],Rdel[0],Rvt[0],Ri[0],Cp[0],Ce[0],C0[0],Ye[0],av[0],Mv[0],Rvx[0],Uvx[0],an[0])
    count = 1
    print(formul)



    contl["Re"] = numpy.format_float_positional(Re[1], precision=3, trim='-') + " " + Om[0]
    contl["Rk"] = numpy.format_float_positional(Rk[1], precision=3, trim='-') + " " + Om[0]
    contl["R1"] = numpy.format_float_positional(R1[1], precision=3, trim='-') + " " + Om[3]
    contl["R2"] = numpy.format_float_positional(R2[1], precision=3, trim='-') + " " + Om[3]
    Cee = numberss(nomin(Ce[1],True))
    Cpp = numberss(nomin(Cp[1],True))
    print(Cpp,Cp,nomin(Cp[1],True))
    contl["Ce"] = numpy.format_float_positional(Cee[0], precision=4, trim='-') + Far[Cee[1]]
    contl["Cp"] = numpy.format_float_positional(Cpp[0], precision=4, trim='-') + Far[Cpp[1]]
    contl["de"] = numpy.format_float_positional(Mne, precision=3, trim='-')
    contl["dp"] = numpy.format_float_positional(Mnp, precision=3, trim='-')
    contl["Roc"] = numpy.format_float_positional(Rooc[1], precision=3, trim='-') + " " + Om[0]
    contl["Kooc"] = str(Kooc)
    contl["ind1"] = str(2 + 2*(kpy-1))
    contl["ind2"] = str(3 + 2 * (kpy - 1))
    contl["ind3"] = str(kpy)

    for g in formul:
        if str(type(g)) == r"<class 'tuple'>" or str(type(g)) == r"<class 'list'>":
            # print(latex_to_word(g[0]))
            contl["form" + str(count)] = latex_to_word(g[0])
        else:
            # print(latex_to_word(g))
            contl["form" + str(count)] = latex_to_word(g)
        count+=1
    doc = DocxTemplate("каскад.docx")
    pathim = "C:\\Users\\User\\Documents\\analogafan\\kpy.png"
    img = Image.open(pathim)
    draw = ImageDraw.Draw(img)
    font1 = ImageFont.truetype("GOST_B.TTF", 20)
    draw.text((175,284), contl["R1"], font=font1, fill=('#1C0606'))
    draw.text((175, 425), contl["R2"], font=font1, fill=('#1C0606'))
    draw.text((248, 566), contl["Re"], font=font1, fill=('#1C0606'))
    draw.text((358, 558), contl["Ce"], font=font1, fill=('#1C0606'))
    draw.text((296, 213), contl["Rk"], font=font1, fill=('#1C0606'))
    draw.text((387, 304), contl["Cp"], font=font1, fill=('#1C0606'))
    draw.text((295, 454), contl["Roc"], font=font1, fill=('#1C0606'))
    img.save("kpy_" + str(kpy) + ".png")
    img = 0
    draw = 0
    contl["img2"] = docxtpl.InlineImage(doc,"kpy_" + str(kpy) + ".png",Mm(100),Mm(108))
    contl["img"] = docxtpl.InlineImage(doc,"3102__"+ str(kpy) +".jpg",Mm(170),Mm(120))
    doc.render(contl)
    doc.save("каскадпосчит__"+ str(kpy) +".docx")
    tr = 0
    return (round(Rvx[1]*1000,3),round(Uvx[1],4))





# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    kpu = 3
    E = 15
    Rn = 300
    Cn  = 0
    Kooc = 12.7
    Mn = 0.891
    Mv = 0.891
    fn = 480
    fv = 50000
    Uvix = 4
    vv = raschet(E,Rn,Cn,Kooc,Mn,Mv,fn,fv,Uvix,1,2)
    Cn = 0
    E = E*0.8
    for i in range(2,kpu+1):
        vv = raschet(E,vv[0],Cn,Kooc,Mn,Mv,fn,fv,vv[1],i,2)


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
