# -*- coding: utf-8 -*-
"""Elektrikçi saha cihazı montaj talimatı PDF üretici (Cofle Manage).
Metal Enjeksiyon — OPTOKUPLÖRLÜ (PC817) sürüm: 300T, 550T (yeni kurulum) + 400T (retrofit).
Cihaz başına TEK sayfa: özet uyarı + bölümlü bağlantı tablosu + blok şema + kontrol listesi.

Çalıştırma (repo kökünden):  python -X utf8 talimat_uret.py
Çıktı:  belgeler/Montaj_Talimati_Metal_Opto.pdf

Yeni makine eklemek: aşağıdaki 'makine_sayfasi(...)' çağrılarına satır ekleyin,
ör. story.append(PageBreak()); story += makine_sayfasi('M13', 'Montaj 13')
Çalışan makineyi opto'ya çevirmek için retrofit=True verin.
Kablo standardı: röle NO -> opto pin2 (400T'de beyaz tel), röle COM -> GND (sarı tel)."""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, Line, String, PolyLine, Circle

FW = r'C:\Windows\Fonts'
pdfmetrics.registerFont(TTFont('TR',   os.path.join(FW, 'arial.ttf')))
pdfmetrics.registerFont(TTFont('TR-B', os.path.join(FW, 'arialbd.ttf')))
pdfmetrics.registerFontFamily('TR', normal='TR', bold='TR-B')

NAVY  = colors.HexColor('#1e293b'); BLUE = colors.HexColor('#2563eb')
GREEN = colors.HexColor('#16a34a'); RED  = colors.HexColor('#dc2626')
AMBER = colors.HexColor('#b45309'); GREY = colors.HexColor('#64748b')
PURP  = colors.HexColor('#7c3aed')
LIGHT = colors.HexColor('#f1f5f9'); LINE = colors.HexColor('#cbd5e1')

def st(name, **kw):
    base = dict(fontName='TR', fontSize=10, leading=14, textColor=NAVY); base.update(kw)
    return ParagraphStyle(name, **base)

S_TITLE = st('t', fontName='TR-B', fontSize=18, leading=22)
S_SUB   = st('s', fontSize=10.5, textColor=GREY)
S_H     = st('h', fontName='TR-B', fontSize=12, leading=14, textColor=BLUE, spaceBefore=6, spaceAfter=3)
S_BODY  = st('b')
S_NOTE  = st('n', fontSize=8.4, textColor=GREY, leading=11)
S_CELL  = st('c', fontSize=8.4, leading=10.5)
S_CELLB = st('cb', fontName='TR-B', fontSize=8.4, leading=10.5)
S_CELLM = st('cm', fontName='TR-B', fontSize=9.5, leading=12, textColor=BLUE)

def kutu(txt, renk, zemin, w=170):
    p = Paragraph(txt, st('kx', fontSize=9.0, leading=12.2, textColor=renk))
    t = Table([[p]], colWidths=[w*mm])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),zemin),('BOX',(0,0),(-1,-1),0.8,renk),
        ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
        ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
    return t

def baglanti_tablosu(sections):
    """Bölüm başlıklı + numaralı bağlantı tablosu."""
    data = [[Paragraph('#', S_CELLB), Paragraph('NEREDEN', S_CELLB),
             Paragraph('NEREYE', S_CELLB), Paragraph('KOMPONENT / NOT', S_CELLB)]]
    spans = []; sec_rows = []; n = 0
    r = 1
    for baslik, satirlar in sections:
        data.append([Paragraph(baslik, st('sec', fontName='TR-B', fontSize=8.4, textColor=colors.white)), '', '', ''])
        spans.append(('SPAN',(0,r),(-1,r))); sec_rows.append(r); r += 1
        for a, b, c in satirlar:
            n += 1
            data.append([Paragraph(str(n), S_CELLB), Paragraph(a, S_CELL),
                         Paragraph(b, S_CELL), Paragraph(c, S_CELL)]); r += 1
    t = Table(data, colWidths=[8*mm, 58*mm, 46*mm, 58*mm], repeatRows=1)
    sty = [('FONTNAME',(0,0),(-1,-1),'TR'),
           ('BACKGROUND',(0,0),(-1,0), NAVY), ('TEXTCOLOR',(0,0),(-1,0), colors.white),
           ('GRID',(0,0),(-1,-1),0.5,LINE), ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
           ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
           ('TOPPADDING',(0,0),(-1,-1),2.4),('BOTTOMPADDING',(0,0),(-1,-1),2.4)]
    for sr in sec_rows:
        sty.append(('BACKGROUND',(0,sr),(-1,sr), BLUE))
    sty += spans
    veri = [i for i in range(1, len(data)) if i not in sec_rows]
    for k, i in enumerate(veri):
        if k % 2 == 1: sty.append(('BACKGROUND',(0,i),(-1,i), LIGHT))
    t.setStyle(TableStyle(sty))
    return t

def kontrol_listesi(items):
    rows = []
    for it in items:
        box = Table([['']], colWidths=[4.3*mm], rowHeights=[4.3*mm])
        box.setStyle(TableStyle([('BOX',(0,0),(-1,-1),1.0,NAVY)]))
        rows.append([box, Paragraph(it, S_CELL)])
    t = Table(rows, colWidths=[8*mm, 162*mm])
    t.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),2.2),('BOTTOMPADDING',(0,0),(-1,-1),2.2),('LEFTPADDING',(0,0),(0,-1),2)]))
    return t

def sema(cihaz_id):
    """Optokuplörlü (PC817) sinyal yolu blok şeması. Güç zinciri: bkz. Tablo C."""
    d = Drawing(420, 132)
    def box(x,y,w,h,title,sub,fill,stroke):
        d.add(Rect(x,y,w,h, fillColor=fill, strokeColor=stroke, strokeWidth=1.1))
        if sub:
            d.add(String(x+w/2, y+h/2+1, title, fontName='TR-B', fontSize=7.6, fillColor=NAVY, textAnchor='middle'))
            d.add(String(x+w/2, y+h/2-8, sub, fontName='TR', fontSize=6.1, fillColor=GREY, textAnchor='middle'))
        else:
            d.add(String(x+w/2, y+h/2-3, title, fontName='TR-B', fontSize=7.6, fillColor=NAVY, textAnchor='middle'))
    def lbl(x,y,t,col=NAVY,anchor='start',fs=6.1,b=True):
        d.add(String(x,y,t, fontName='TR-B' if b else 'TR', fontSize=fs, fillColor=col, textAnchor=anchor))
    def wire(pts,col):
        d.add(PolyLine(pts, strokeColor=col, strokeWidth=1.4))
    def dot(x,y,col=GREY):
        d.add(Circle(x,y,2.0, fillColor=col, strokeColor=col))
    # kutular
    box(30,106,58,20,'5V','',colors.HexColor('#fef2f2'),RED)
    box(104,106,60,20,'2.2 kΩ','',LIGHT,GREY)
    box(176,60,92,62,'PC817','optokuplör',colors.HexColor('#f5f3ff'),PURP)
    box(30,64,86,32,'RÖLE','kuru kontak',colors.HexColor('#fff7ed'),AMBER)
    box(296,54,96,72,'ESP32',cihaz_id,colors.HexColor('#eff6ff'),BLUE)
    # PC817 bacak etiketleri
    lbl(180,108,'1 anot'); lbl(180,70,'2 katot'); lbl(264,108,'4 koll.',NAVY,'end'); lbl(264,70,'3 emit.',NAVY,'end')
    # röle uçları
    lbl(108,88,'NO'); lbl(104,68,'COM')
    # esp32 pin + güç notu
    lbl(300,108,'GPIO25'); lbl(300,70,'GND'); lbl(300,60,'5V/VIN ← LM2596 (Tablo C)',GREY,'start',5.6,False)
    # kablolar
    wire([88,116, 104,116], RED)                        # 5V -> direnç
    wire([164,116, 170,116, 170,110, 176,110], RED)     # direnç -> pin1 (anot)
    wire([176,72, 146,72, 146,86, 116,86], AMBER)       # pin2 (katot) -> röle NO
    wire([116,70, 116,16], GREY)                         # röle COM -> ortak GND
    wire([268,110, 296,110], BLUE)                       # pin4 (koll.) -> GPIO25
    wire([268,72, 268,16], GREY)                         # pin3 (emit.) -> ortak GND
    wire([296,72, 284,72, 284,16], GREY)                 # ESP32 GND -> ortak GND
    wire([100,16, 300,16], GREY)                         # ortak GND hattı
    dot(116,16); dot(268,16); dot(284,16)
    lbl(200,8,'ortak GND',GREY,'middle',5.9,False)
    return d

def makine_sayfasi(cihaz_id, makine_ad, retrofit=False):
    el = []
    hdr = Table([[Paragraph(f'CİHAZ: {cihaz_id}', st('hx', fontName='TR-B', fontSize=14, textColor=colors.white)),
                  Paragraph(f'Makine: {makine_ad} &nbsp;|&nbsp; Bölüm: Metal Enjeksiyon &nbsp;|&nbsp; Optokuplörlü',
                            st('hy', fontSize=9.5, textColor=colors.white, alignment=TA_LEFT))]],
                 colWidths=[50*mm, 120*mm])
    hdr.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),BLUE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('LEFTPADDING',(0,0),(-1,-1),10),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
    el += [hdr, Spacer(1, 5)]
    if retrofit:
        el.append(kutu(f'<b>RETROFIT — {cihaz_id} ŞU AN ÇALIŞIYOR:</b> <b>1)</b> Makineyi/pano enerjisini '
            '<b>durdur</b>. <b>2)</b> Karta giden <b>mevcut röle→GPIO25 doğrudan telini SÖK</b> (artık araya opto giriyor). '
            '<b>3)</b> Aşağıdaki tabloyu uygula. Röle uçları yerinde (beyaz=NO, sarı=COM); <b>sadece kart tarafı değişir</b>.',
            AMBER, colors.HexColor('#fffbeb')))
        el.append(Spacer(1, 4))
    el.append(kutu('<b>ÖNEMLİ:</b> <b>1)</b> <b>PC817 yönü:</b> köşedeki nokta/çentik = <b>pin 1 (anot)</b>, ters takma. '
        '<b>2)</b> <b>LM2596 çıkışını ESP32 bağlamadan ÖNCE 5.0V</b>’a ayarla. <b>3)</b> Opto çıkışına (GPIO25) <b>harici direnç KOYMA</b> '
        '(dahili pull-up var). <b>4)</b> Anteni metal kutunun <b>DIŞINA</b> tak.', RED, colors.HexColor('#fef2f2')))
    el.append(Spacer(1, 4))
    el.append(Paragraph('Bağlantılar', S_H))
    el.append(baglanti_tablosu([
        ('A) GİRİŞ — optokuplör LED tarafı (makine sinyali)', [
            ('ESP32 5V/VIN', 'PC817 pin 1 (anot)', 'arada <b>2.2kΩ</b> seri direnç'),
            ('PC817 pin 2 (katot)', 'Röle NO ucu', 'üretim sinyali (400T’de <b>beyaz</b> tel)'),
            ('Röle COM ucu', 'ESP32 GND', '400T’de <b>sarı</b> tel')]),
        ('B) ÇIKIŞ — optokuplör transistör tarafı (karta sinyal)', [
            ('PC817 pin 4 (kollektör)', 'ESP32 GPIO25', '<b>harici direnç YOK</b> (dahili pull-up)'),
            ('PC817 pin 3 (emitter)', 'ESP32 GND', '—')]),
        ('C) GÜÇ — pano 24V → 5V (LM2596)', [
            ('Pano 24V (+)', 'LM2596 IN+', '—'),
            ('Pano 0V (−)', 'LM2596 IN−', '—'),
            ('LM2596 OUT+', 'ESP32 5V/VIN', '<b>önce 5.0V’a ayarla</b>'),
            ('LM2596 OUT−', 'ESP32 GND', 'tüm GND ortak')]),
        ('D) ANTEN', [
            ('IPEX pigtail', 'ESP32 U.FL soketi', '“tık” oturt'),
            ('SMA çubuk anten', 'metal kutu DIŞINA', 'metal WiFi’yi keser')]),
    ]))
    el.append(Paragraph('Şema  (sinyal yolu — güç zinciri Tablo C’de)', S_H))
    el.append(sema(cihaz_id))
    el.append(Paragraph('Kontrol Listesi', S_H))
    el.append(kontrol_listesi([
        'Pano enerjisi kapalıyken bağlandı',
        'PC817 yönü doğru — pin 1 (nokta/çentikli köşe) anot tarafında',
        'Giriş: 5V → 2.2kΩ → pin 1,  pin 2 → röle NO,  röle COM → GND',
        'Çıkış: pin 4 → GPIO25,  pin 3 → GND  (harici pull-up YOK)',
        'LM2596 çıkışı multimetreyle 5.0V doğrulandı (ESP32 bağlı değilken)',
        'Anten metal kutunun DIŞINDA',
        'ESP32 enerjilendi (kart LED’i yanıyor)',
        'Yazılım: cihaz online + makine cycle’ında sayım merkeze ulaştı']))
    return el

os.makedirs('belgeler', exist_ok=True)
OUT = os.path.join('belgeler', 'Montaj_Talimati_Metal_Opto.pdf')
doc = SimpleDocTemplate(OUT, pagesize=A4, leftMargin=20*mm, rightMargin=20*mm,
                        topMargin=15*mm, bottomMargin=15*mm,
                        title='Cofle Manage - Montaj Talimati (Metal, Optokuplorlu)')
story = []
story.append(Paragraph('COFLE MANAGE', S_TITLE))
story.append(Paragraph('Saha Cihazı Montaj Talimatı &mdash; Metal Enjeksiyon (Optokuplörlü)', S_SUB))
story.append(Spacer(1, 4))
story.append(Table([['']], colWidths=[170*mm], rowHeights=[2], style=[('LINEBELOW',(0,0),(-1,-1),1.2,BLUE)]))
story.append(Spacer(1, 9))
story.append(kutu('<b>Kapsam:</b> 3 metal enjeksiyon makinesine <b>optokuplörlü (PC817)</b> ESP32 sayaç kartı. '
                  '<b>300T &amp; 550T</b> = yeni kurulum. <b>400T</b> = retrofit (şu an çalışıyor; önce durdurulur, '
                  'mevcut doğrudan tel sökülür). Makinelerdeki sinyal röleleri <b>kuru kontak</b> ve hazırdır; '
                  'kart tarafından başlanır. Her makine için ayrı sayfadaki adımları izleyin.', NAVY, LIGHT))
story.append(Paragraph('Neden optokuplör?', S_H))
story.append(Paragraph('Optokuplör, sinyali <b>akımla sürülen bir iç LED</b> üzerinden geçirir — bu, hat parazitinin '
    'GPIO’yu tetikleyip <b>hayalet/eksik sayıma</b> yol açmasını engeller (özellikle 400T sorunu). Devreye eklenen '
    'tek yeni parça PC817 (4 bacaklı entegre) + 2.2kΩ dirençtir; <b>harici/görünür LED yoktur</b> — LED entegrenin içindedir.',
    S_BODY))
story.append(Paragraph('Her Cihaz İçin Kit İçeriği', S_H))
kit = [['1×','ESP32-WROOM-32U geliştirme kartı'],['1×','PC817 optokuplör (4 bacaklı entegre)'],
       ['1×','2.2kΩ direnç (opto LED akım sınırlayıcı)'],
       ['1×','LM2596 ayarlanabilir güç regülatörü (24V→5V)'],
       ['1×','IPEX anten seti (pigtail + SMA çubuk anten + somun)'],
       ['ops.','100nF seramik kondansatör (ek filtre — opto varken çoğunlukla gerekmez)'],
       ['','Bağlantı kablosu / klemens']]
kt = Table([[Paragraph(a,S_CELLM),Paragraph(b,S_CELL)] for a,b in kit], colWidths=[14*mm,156*mm])
kt.setStyle(TableStyle([('ROWBACKGROUNDS',(0,0),(-1,-1),[colors.white,LIGHT]),('GRID',(0,0),(-1,-1),0.4,LINE),
    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),3.4),('BOTTOMPADDING',(0,0),(-1,-1),3.4),('LEFTPADDING',(0,0),(-1,-1),6)]))
story.append(kt)
story.append(Paragraph('ESP32 Pin Haritası', S_H))
story.append(Paragraph('<b>GPIO25</b> = sinyal girişi (opto pin 4’ten) &nbsp;&bull;&nbsp; <b>GND</b> = ortak şase/0V '
    '&nbsp;&bull;&nbsp; <b>5V/VIN</b> = besleme (LM2596’dan; opto LED’i de buradan beslenir) &nbsp;&bull;&nbsp; '
    '<b>U.FL</b> = anten soketi. <i>(3V3 bu sürümde harici kullanılmaz.)</i>', S_BODY))
story.append(Paragraph('PC817 Bacak Düzeni', S_H))
story.append(Paragraph('Köşedeki <b>nokta = pin 1</b>. Saat yönünde: <b>1 anot</b> ve <b>2 katot</b> = giriş (LED) tarafı; '
    '<b>3 emitter</b> ve <b>4 kollektör</b> = çıkış (transistör) tarafı. Giriş ve çıkış <b>karşılıklı</b> kenarlardadır.', S_BODY))
story.append(Paragraph('Genel Güvenlik ve Sıra', S_H))
story.append(kutu('<b>1.</b> <b>PC817’yi doğru yönde tak</b> — köşedeki nokta = pin 1 (anot). Ters takılırsa sinyal gelmez.'
    '<br/><b>2.</b> LM2596 çıkışını her zaman <b>ESP32 bağlı DEĞİLKEN</b> multimetreyle <b>5.0V</b>’a ayarla.'
    '<br/><b>3.</b> Opto çıkışına (GPIO25) <b>ikinci bir direnç koyma</b> — ESP32’nin dahili pull-up’ı yeterli.'
    '<br/><b>4.</b> Anteni metal kutunun <b>DIŞINA</b> monte et; bağlantıları pano enerjisi <b>kapalıyken</b> yap.',
    RED, colors.HexColor('#fef2f2')))
story.append(Spacer(1, 5))
story.append(Paragraph('Not: Üç makinede de kart aynı yazılımla çalışır; sinyal mantığı her iki yöntemde aynıdır '
    '(üretim anı = GPIO25 düşük). Bu sürüm metal makinelerini tek standarda — optokuplörlü — alır.', S_NOTE))
story.append(PageBreak()); story += makine_sayfasi('300T', '300 Ton')
story.append(PageBreak()); story += makine_sayfasi('550T', '550 Ton')
story.append(PageBreak()); story += makine_sayfasi('400T', '400 Ton', retrofit=True)

def alt(c, d):
    c.saveState(); c.setFont('TR', 7.5); c.setFillColor(GREY)
    c.drawString(20*mm, 9*mm, 'Cofle Manage — Saha Montaj Talimatı (Metal, Optokuplörlü: 300T · 550T · 400T)')
    c.drawRightString(190*mm, 9*mm, 'Sayfa %d' % d.page)
    c.setStrokeColor(LINE); c.setLineWidth(0.5); c.line(20*mm, 12*mm, 190*mm, 12*mm); c.restoreState()

doc.build(story, onFirstPage=alt, onLaterPages=alt)
print('OK:', os.path.abspath(OUT), '|', os.path.getsize(OUT), 'byte')
