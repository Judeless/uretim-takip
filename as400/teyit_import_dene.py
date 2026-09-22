# -*- coding: utf-8 -*-
"""
teyit_import_dene.py — BPRCF0I teyit aktarımının ELLE denenmesi (Simone Rota testi).

Tek satır yazar, ne yazdığını EKRANA DÖKER ve programın cevabını bekler. Yazdığımız
her satır J0STE2 = 'T'+yymmddHHMMSS+sayaç ile damgalıdır; gerekirse ERP tarafında
bu anahtarla bulunup elle düzeltilebilir.

KULLANIM
  python as400/teyit_import_dene.py --liste                     # acik emirleri goster
  python as400/teyit_import_dene.py --liste --article 10.300.2933W
  python as400/teyit_import_dene.py --son 10                    # tablodaki son satirlar
  python as400/teyit_import_dene.py --emir 26-200385 --adet 2   # DENEME TEYIDI (A)
  python as400/teyit_import_dene.py --emir 26-200385 --adet 2 --flsa S --onayla
  python as400/teyit_import_dene.py --emir 26-200385 --adet 1 --causale 001 --depo 001
  python as400/teyit_import_dene.py --rrn 123                   # satirin durumunu oku

GUVENLIK
  · Varsayilan J0FLSA='A' (ara teyit). 'S' KAPANIS teyididir; --onayla sart.
  · Adet ust siniri 5 (deneme); daha fazlasi icin --onayla.
  · Yazmadan once emir XPRO90'da aranir, kalan adet gosterilir.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import as400_teyit_import as TI

# Windows konsolu cp1254: ok/tire gibi karakterler cokerdi (cikti kaybolur,
# yazim hic yapilmaz). Ciktiyi UTF-8'e cevir, cevrilemeyeni '?' yap.
for _akim in (sys.stdout, sys.stderr):
    try:
        _akim.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


def yaz(b, *sat):
    print('\n' + b)
    for s in sat:
        print('   ' + s)


def main():
    ap = argparse.ArgumentParser(description='BPRCF0I teyit aktarimi denemesi')
    ap.add_argument('--emir', help="uretim emri: '26-200385' ya da '20-26-200385'")
    ap.add_argument('--adet', type=float, help='teyit edilecek adet (J0QTRI)')
    ap.add_argument('--flsa', default='A', choices=['A', 'S'],
                    help="A = ara teyit (varsayilan), S = KAPANIS teyidi")
    ap.add_argument('--bekle', type=int, default=90, help='cevap icin bekleme (sn)')
    ap.add_argument('--liste', action='store_true', help='acik uretim emirlerini listele')
    ap.add_argument('--article', help='--liste icin referans kodu suzgeci')
    ap.add_argument('--son', type=int, help='BPRCF0I tablosundaki son N satiri goster')
    ap.add_argument('--bizim', action='store_true', help="--son icin yalniz bizim satirlarimiz")
    ap.add_argument('--rrn', type=int, help='tek satirin durumunu oku')
    ap.add_argument('--kod', help='J0ARTI olarak yazilacak referans kodu '
                                 '(verilmezse ERP kaydindaki article yazilir)')
    ap.add_argument('--kodsuz', action='store_true',
                    help='J0ARTI hic yazma (ilk denemelerdeki sade hali)')
    ap.add_argument('--causale', help='J0CRCD rientro neden kodu (5250 ekranindaki deger)')
    ap.add_argument('--depo', help='J0MGPR ana depo kodu')
    ap.add_argument('--commessa', help='J0COMM is emri referansi')
    ap.add_argument('--onayla', action='store_true', help='S bayragi / buyuk adet icin onay')
    a = ap.parse_args()

    cn = TI.baglan(timeout=30)
    try:
        if a.liste:
            satir = TI.acik_emirler(a.article, n=25, cn=cn)
            yaz('ACIK URETIM EMIRLERI (XPRO90, durum 40/45/50)',
                '%-12s %-18s %-5s %9s %9s %9s' % ('EMIR', 'ARTICLE', 'DRM', 'ADET', 'TEYIT', 'KALAN'))
            for s in satir:
                print('   %-12s %-18s %-5s %9.0f %9.0f %9.0f'
                      % (s['emir'], s['article'][:18], s['durum'], s['adet'], s['teyit'], s['kalan']))
            if not satir:
                print('   (kayit yok)')
            return

        if a.son:
            satir = TI.son_satirlar(a.son, cn=cn, yalniz_bizim=a.bizim)
            yaz('BPRCF0I SON SATIRLAR',
                '%-6s %-12s %8s %-4s %-16s %-4s %-12s %s'
                % ('RRN', 'EMIR', 'ADET', 'A/S', 'ANAHTAR', 'DRM', 'HAREKET', 'NOT'))
            for s in satir:
                print('   %-6s %-12s %8s %-4s %-16s %-4s %-12s %s'
                      % (s['rrn'], s.get('emir', ''), s.get('j0qtri', ''), s.get('j0flsa', ''),
                         s.get('j0ste2', ''), s.get('j0stat', '') or '-',
                         s.get('hareket_no', ''), (s.get('not') or '')[:60]))
            if not satir:
                print('   (tablo bos)')
            return

        if a.rrn:
            d = TI.teyit_durum(a.rrn, cn=cn, bekleme_sn=0)
            yaz('SATIR DURUMU', *[f'{k}: {v}' for k, v in d.items()])
            return

        if not a.emir or not a.adet:
            ap.error('--emir ve --adet gerekli (ya da --liste / --son / --rrn)')

        parcali = TI.emir_parcala(a.emir)
        bilgi = TI.emir_bilgi(parcali, cn=cn)
        yaz('YAZILACAK TEYIT',
            f'emir        : {a.emir}  →  J0RED1={parcali[0]}, J0RED2={parcali[1]}, J0RENU={parcali[2]}',
            f'adet (J0QTRI): {a.adet:g}',
            f"bayrak (J0FLSA): {a.flsa}  ({'ARA teyit' if a.flsa == 'A' else 'KAPANIS teyidi — emri kapatir'})")
        ekler = {}
        if not a.kodsuz:
            kod = a.kod or (bilgi or {}).get('article') or ''
            if kod:
                ekler['J0ARTI'] = kod
        for ad, deger in (('J0CRCD', a.causale), ('J0MGPR', a.depo), ('J0COMM', a.commessa)):
            if deger:
                ekler[ad] = deger
        if ekler:
            print('   ek alanlar  : ' + ' · '.join(f'{x}={y}' for x, y in ekler.items()))
        if bilgi:
            print('   ERP kaydi   : %s · durum %s · adet %.0f · teyit %.0f · KALAN %.0f'
                  % (bilgi['article'], bilgi['durum'], bilgi['adet'], bilgi['teyit'], bilgi['kalan']))
            if a.adet > bilgi['kalan'] and not a.onayla:
                print(f"   DURDURULDU: adet kalan adetten buyuk ({a.adet:g} > {bilgi['kalan']:.0f}). "
                      f"--onayla ile zorlayabilirsin.")
                return
        else:
            print('   ERP kaydi   : BULUNAMADI (XPRO90 acik emirlerinde yok) — 01E bekleyebilirsin')

        if a.flsa == 'S' and not a.onayla:
            print('   DURDURULDU: S bayragi emri KAPATIR. Eminsen --onayla ekle.')
            return
        if a.adet > 5 and not a.onayla:
            print('   DURDURULDU: deneme icin adet siniri 5. Daha fazlasi icin --onayla ekle.')
            return

        sonuc = TI.teyit_yaz(a.emir, a.adet, flsa=a.flsa, cn=cn, bekleme_sn=a.bekle,
                             ekler=ekler, kalan_kontrol=not a.onayla)
        yaz('SONUC',
            f"durum      : {sonuc.get('durum')}  (ok={sonuc.get('ok')})",
            f"anahtar    : {sonuc.get('anahtar', '')}   ← ERP'de J0STE2 bu deger",
            f"rrn        : {sonuc.get('rrn', '')}",
            f"teyit sira : {sonuc.get('teyit_sira', '')}  (J0CNPR)",
            f"hareket no : {sonuc.get('hareket_no', '')}  (J0MGSS-AA-NU-PG)",
            f"J0STAT     : {sonuc.get('j0stat', '')}   1=OK 2=HATA 3=UYARI",
            f"not        : {sonuc.get('not', '')}")
        if sonuc.get('sql'):
            print('   sql        : ' + sonuc['sql'])
            print('   alanlar    : ' + str(sonuc.get('alanlar')))
    finally:
        cn.close()


if __name__ == '__main__':
    main()
