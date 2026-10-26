# -*- coding: utf-8 -*-
"""Süre onarımı (2026-10-26) — Excel otomatik doldurma (Flash Fill) hasarını geri alır.

Ana Veri v6'da G (çevrim) sütununun BOŞ hücreleri D sütunundan türetilmiş değerlerle
dolmuştu (TK-2 → 924 / 120, TK-1 → 913 / 110, Satınalma → '1Satınalma0'…). Yükleme bu
değerleri TK2 referanslarına ve — süre değişince geçmiş kayıtlar da güncellendiği için —
Nisan'dan beri GEÇMİŞ üretim kayıtlarına işledi. Onarım verisi (JSON) 21.10 11:54
sunucu kopyasıyla karşılaştırılarak üretildi.

Kullanım (sunucuda, C:\\cofle\\uretim_takip içinde):
    python sunucu\\sure_onar.py sunucu\\sure_onarim_20261026.json             yalnız rapor
    python sunucu\\sure_onar.py sunucu\\sure_onarim_20261026.json --uygula    yedek + onarım
    --grup 924,120,test   yalnız bu gruplar (varsayılan: hepsi)
    --yalniz-referans     geçmiş üretim kayıtlarına DOKUNMA, yalnız referans süreleri
                          (kullanıcı 2026-10-26: "geçmişe dönük güncellemeye gerek yok")

Her satır YALNIZ değeri hâlâ hasarlı değere eşitse geri alınır — arada panelden
düzeltilmiş bir süreye dokunulmaz."""
import argparse
import io
import json
import os
import sqlite3
import sys
from datetime import datetime

for _a in (sys.stdout, sys.stderr):
    try:
        _a.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB = os.path.join(KOK, 'uretim.db')


def esit(a, b):
    return abs(float(a or 0) - float(b or 0)) < 0.001


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('veri')
    ap.add_argument('--uygula', action='store_true')
    ap.add_argument('--grup', default='')
    ap.add_argument('--yalniz-referans', action='store_true')
    arg = ap.parse_args()
    veri = json.load(io.open(arg.veri, encoding='utf-8'))
    if arg.yalniz_referans:
        veri['kayitlar'] = []                 # geçmiş kayıtlar olduğu gibi kalır
    gruplar = {g.strip() for g in arg.grup.split(',') if g.strip()}
    sec = (lambda x: not gruplar or x['grup'] in gruplar)
    if not os.path.exists(DB):
        sys.exit(f'[HATA] veritabanı yok: {DB}')

    cn = sqlite3.connect(DB, timeout=30)
    sayac = {}

    def say(anahtar):
        sayac[anahtar] = sayac.get(anahtar, 0) + 1

    try:
        cn.execute('BEGIN IMMEDIATE')
        for r in filter(sec, veri['referanslar']):
            m = cn.execute("SELECT COALESCE(hedef_cycle_time_sn,0), COALESCE(kaynak_suresi_sn,0), "
                           "COALESCE(soktak_suresi_sn,0) FROM referans_listesi WHERE id=? AND bolum=?",
                           (r['id'], r['bolum'])).fetchone()
            if not m:
                say((r['grup'], 'referans', 'bulunamadı'))
                continue
            if not (esit(m[0], r['simdi']) and esit(m[1], r['simdi_kaynak']) and esit(m[2], r['simdi_soktak'])):
                say((r['grup'], 'referans', 'değişmiş, dokunulmadı'))
                continue
            cn.execute("UPDATE referans_listesi SET hedef_cycle_time_sn=?, kaynak_suresi_sn=?, soktak_suresi_sn=? "
                       "WHERE id=?", (r['eski'], r['eski_kaynak'], r['eski_soktak'], r['id']))
            say((r['grup'], 'referans', 'geri alındı'))
        for k in filter(sec, veri['kayitlar']):
            m = cn.execute("SELECT COALESCE(cycle_time_sn,0) FROM uretim_kayitlari WHERE id=?", (k['id'],)).fetchone()
            if not m:
                say((k['grup'], 'kayıt', 'bulunamadı'))
                continue
            if not esit(m[0], k['simdi']):
                say((k['grup'], 'kayıt', 'değişmiş, dokunulmadı'))
                continue
            cn.execute("UPDATE uretim_kayitlari SET cycle_time_sn=? WHERE id=?", (k['eski'], k['id']))
            say((k['grup'], 'kayıt', 'geri alındı'))

        print(veri.get('aciklama', ''))
        print()
        for (grup, tur, durum), n in sorted(sayac.items()):
            print(f'  {grup:>5} · {tur:<9} · {durum:<22} {n}')
        if not arg.uygula:
            cn.rollback()
            print('\nRAPOR — hiçbir şey değişmedi. Uygulamak için --uygula.')
            return
        # Yedek: aynı bağlantının işlemi açıkken alınamaz → önce geri al, yedekle, yeniden uygula
        cn.rollback()
        hedef_klasor = r'C:\cofle\yedek' if os.path.isdir(r'C:\cofle\yedek') else KOK
        yedek = os.path.join(hedef_klasor, 'uretim_ONCE_sure_onarim_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.db')
        with sqlite3.connect(yedek) as y:
            cn.backup(y)
        print(f'\n[OK] yedek: {yedek}')
        sayac.clear()
        cn.execute('BEGIN IMMEDIATE')
        for r in filter(sec, veri['referanslar']):
            cn.execute("UPDATE referans_listesi SET hedef_cycle_time_sn=?, kaynak_suresi_sn=?, soktak_suresi_sn=? "
                       "WHERE id=? AND bolum=? AND ABS(COALESCE(hedef_cycle_time_sn,0)-?)<0.001 "
                       "AND ABS(COALESCE(kaynak_suresi_sn,0)-?)<0.001 AND ABS(COALESCE(soktak_suresi_sn,0)-?)<0.001",
                       (r['eski'], r['eski_kaynak'], r['eski_soktak'], r['id'], r['bolum'],
                        r['simdi'], r['simdi_kaynak'], r['simdi_soktak']))
        for k in filter(sec, veri['kayitlar']):
            cn.execute("UPDATE uretim_kayitlari SET cycle_time_sn=? WHERE id=? "
                       "AND ABS(COALESCE(cycle_time_sn,0)-?)<0.001", (k['eski'], k['id'], k['simdi']))
        # Onarım verisi üretildikten SONRA açılan kayıtlar referansın hasarlı süresiyle açıldı →
        # referansın eski süresine çekilir (aynı eşleşme: import_excel'in geri yazımı)
        yeni = 0
        for r in ([] if arg.yalniz_referans else filter(sec, veri['referanslar'])):
            yeni += cn.execute(
                "UPDATE uretim_kayitlari SET cycle_time_sn=? WHERE id>? "
                "AND UPPER(REPLACE(referans_kodu,' ',''))=UPPER(REPLACE(?,' ','')) "
                "AND ABS(COALESCE(cycle_time_sn,0)-?)<0.001 AND vardiya_id IN (SELECT id FROM vardiyalar "
                "WHERE COALESCE(lokasyon,'TK2')='TK2' AND COALESCE(bolum,'kaynak')=?)",
                (r['eski'], veri.get('son_kayit_id') or 0, r['kod'], r['simdi'], r['bolum'])).rowcount
        if yeni:
            print(f'[OK] onarım verisinden sonra açılmış {yeni} kayıt da düzeltildi.')
        cn.commit()
        print('[OK] onarım uygulandı.')
    finally:
        cn.close()


if __name__ == '__main__':
    main()
