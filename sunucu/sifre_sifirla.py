# -*- coding: utf-8 -*-
"""PANEL ŞİFRE SIFIRLAMA (sunucuda, 2026-10-27) — yönetici şifresi unutulunca.

Panelden şifre sıfırlamak için giriş yapmış bir yönetici gerekir; TEK yönetici şifresini
unuttuysa (Aspanda: 'admin') bu betik sunucunun kendisinde çalıştırılır. Yeni şifre
RASTGELE üretilir, yalnız bu ekrana BİR KEZ yazılır ve geçici işaretlenir — ilk girişte
yeni şifre belirlemeniz istenir. Şifre hiçbir dosyaya yazılmaz.

    python sunucu\\sifre_sifirla.py C:\\aspanda            (kullanıcı adı sorulur, varsayılan admin)
    python sunucu\\sifre_sifirla.py C:\\aspanda admin
"""
import os
import secrets
import sqlite3
import sys

from werkzeug.security import generate_password_hash

for _a in (sys.stdout, sys.stderr):
    try:
        _a.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


def main():
    klasor = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    db = os.path.join(klasor, 'uretim.db')
    if not os.path.exists(db):
        sys.exit(f'[HATA] veritabanı yok: {db}')
    cn = sqlite3.connect(db, timeout=30)
    try:
        kullanicilar = cn.execute(
            "SELECT id, kullanici_adi, COALESCE(ad_soyad,''), rol, aktif FROM panel_kullanicilari "
            "ORDER BY rol<>'admin', kullanici_adi").fetchall()
        print(f'\n{db}\nPanel kullanıcıları:')
        for _id, ka, ad, rol, aktif in kullanicilar:
            print(f'   {ka:<24} {rol:<10} {ad}' + ('' if aktif else '   (PASİF)'))
        hedef = sys.argv[2] if len(sys.argv) > 2 else (
            input('\nŞifresi sıfırlanacak kullanıcı adı [admin]: ').strip() or 'admin')
        row = next((k for k in kullanicilar if k[1].lower() == hedef.lower()), None)
        if not row:
            sys.exit(f'[HATA] "{hedef}" adında panel kullanıcısı yok')
        sifre = secrets.token_urlsafe(9)
        cn.execute("UPDATE panel_kullanicilari SET sifre_hash=?, sifre_gecici=1 WHERE id=?",
                   (generate_password_hash(sifre), row[0]))
        cn.commit()
    finally:
        cn.close()
    print('\n' + '=' * 52)
    print(f'  Kullanıcı     : {row[1]}')
    print(f'  GEÇİCİ ŞİFRE  : {sifre}')
    print('=' * 52)
    print('  İlk girişte yeni şifre belirlemeniz istenecek.')
    print('  Bu şifre hiçbir yere kaydedilmedi — şimdi not alın.')
    if not row[4]:
        print('  DİKKAT: bu kullanıcı PASİF — giriş yapabilmesi için panelden etkinleştirilmeli.')
    print('  "Çok fazla hatalı deneme" kilidi 5 dakikada kendiliğinden açılır.')


if __name__ == '__main__':
    main()
