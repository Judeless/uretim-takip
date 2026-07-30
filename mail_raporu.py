# -*- coding: utf-8 -*-
"""
mail_raporu.py — Günlük üretim raporunu Excel olarak üretir ve e-posta ile gönderir.

Akış (scheduler her gün gönderim saatinde çağırır):
  1. O günün tüm üretim kayıtlarını çeker (tarih, tesis, bölüm, operatör, referans, adet)
  2. Bir .xlsx dosyası oluşturur (data/gunluk_rapor/YYYY-MM-DD_UretimRaporu.xlsx)
  3. mail_alicilari tablosundaki aktif alıcılara SMTP ile e-posta atar (Excel ekli)

SMTP bilgileri mail_config.json'da tutulur (gitignore — git'e girmez).
Config yoksa/etkin değilse özellik sessizce devre dışıdır (cofle_test paterni).
"""
import os
import json
import smtplib
import sqlite3
import ssl
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH     = os.path.join(PROJECT_DIR, 'uretim.db')
CONFIG_YOL  = os.path.join(PROJECT_DIR, 'mail_config.json')
RAPOR_DIR   = os.path.join(PROJECT_DIR, 'data', 'gunluk_rapor')
RAPOR_GUN_LIMIT = 60  # data/gunluk_rapor'da en fazla bu kadar günlük dosya tutulur

# Bölüm kod → görünen ad (app.py BOLUM_AD ile aynı — rapor okunur olsun)
BOLUM_AD = {
    'kaynak': 'Robot Kaynak',
    'montaj': 'Montaj',
    'metal':  'Metal Enjeksiyon',
    'isleme': 'İşleme',
    'lazer':  'Lazer Kesim',
    'pres':   'Pres Abkant',
}


# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────
def config_yukle():
    """mail_config.json'u okur. Yoksa/bozuksa None döner.
    encoding='utf-8-sig': Notepad (özellikle Server 2019) UTF-8 dosyaya BOM ekler;
    utf-8-sig BOM'u otomatik yutar (BOM'suz dosyayı da sorunsuz okur)."""
    if not os.path.exists(CONFIG_YOL):
        return None
    try:
        with open(CONFIG_YOL, encoding='utf-8-sig') as f:
            return json.load(f)
    except Exception as e:
        print(f'[MAIL] Config okunamadı: {e}')
        return None


def config_hatasi():
    """Dosya VAR ama JSON'u bozuksa hata metni, sorun yoksa ''.

    2026-07-30: config elle duzenlenirken bir satirin sonunda virgul unutuldu;
    dosya okunamaz oldu ve panel "SMTP tanimli degil" dedi. Mesaj yaniltiyordu
    (dosya oradaydi) ve o gun rapor maili hic gitmeyecekti. Gercek sebebi ve
    hatanin satirini panele tasiyoruz."""
    if not os.path.exists(CONFIG_YOL):
        return ''
    try:
        with open(CONFIG_YOL, encoding='utf-8-sig') as f:
            json.load(f)
        return ''
    except Exception as e:
        return str(e)


def yontem_al(c=None):
    """Gönderim yöntemi: 'outlook' (açık Outlook masaüstü — şifre gerekmez) | 'smtp'."""
    c = c or config_yukle() or {}
    return str(c.get('yontem', 'smtp')).lower()


def host_uygun(c=None):
    """config'te 'sadece_host' varsa, sadece o makinede gönderim yapılır.

    2026-07-29: sistem hem laptopta hem canlı sunucuda (ProManage) ayakta olduğu
    için alıcılara İKİ mail gitti. Laptopun veritabanı bayat olduğundan onunki
    "üretim bulunmamaktadır" diyordu — yani yanlış makinenin gönderdiği mail
    doğrusuyla çelişiyor. Kopya kurulumda config de kopyalandığı için 'etkin'
    bayrağı tek başına yetmiyor; makine kimliği gerekiyor.
    Alan yoksa davranış değişmez (geriye uyumlu)."""
    c = c if c is not None else (config_yukle() or {})
    beklenen = (c.get('sadece_host') or '').strip()
    if not beklenen:
        return True
    try:
        import socket
        return socket.gethostname().strip().upper() == beklenen.upper()
    except Exception:
        return True          # host okunamadı → engelleme (güvenli yön: mevcut davranış)


def etkin():
    """Mail entegrasyonu kullanıma hazır mı? (config var + etkin + doğru makine
    + yönteme göre zorunlu alanlar)"""
    c = config_yukle()
    if not c or not c.get('etkin'):
        return False
    if not host_uygun(c):
        return False
    if yontem_al(c) == 'outlook':
        return True  # Outlook masaüstü — ek bilgi gerekmez (açık ve girişli olmalı)
    return all(c.get(k) for k in ('smtp_host', 'smtp_port', 'gonderen'))


def gonderim_saati():
    """(saat, dakika) — config'ten, yoksa 17:00. scheduler bunu okur."""
    c = config_yukle() or {}
    raw = str(c.get('gonderim_saati', '17:00'))
    try:
        sa, dk = raw.split(':')
        return int(sa), int(dk)
    except Exception:
        return 17, 0


def durum_ozeti():
    """Dashboard göstergesi için maskelenmiş durum (şifre ASLA dönmez)."""
    c = config_yukle()
    if not c:
        h = config_hatasi()
        # dosya var ama bozuk -> "tanimli degil" demek yaniltici olur
        return {'config_var': bool(h), 'etkin': False, 'config_hatasi': h}
    import socket
    try:
        bu_host = socket.gethostname()
    except Exception:
        bu_host = ''
    return {
        'config_var': True,
        'etkin': bool(etkin()),
        # Host kilidi gönderimi SESSIZCE durdurabilir (config'te sadece_host yanlış
        # yazılırsa). Panelde nedeni görünsün diye dönüyoruz.
        'sadece_host': (c.get('sadece_host') or ''),
        'bu_host': bu_host,
        'host_engeli': bool(c.get('etkin')) and not host_uygun(c),
        'yontem': yontem_al(c),
        'smtp_host': c.get('smtp_host', ''),
        'smtp_port': c.get('smtp_port', ''),
        'gonderen': c.get('gonderen', ''),
        'gonderim_saati': c.get('gonderim_saati', '17:00'),
        'kapsam_lokasyon': c.get('kapsam_lokasyon', 'HEPSI'),
    }


# ─────────────────────────────────────────────────────────────
# ALICILAR (mail_alicilari tablosu)
# ─────────────────────────────────────────────────────────────
def _db(conn=None):
    if conn is not None:
        return conn, False
    c = sqlite3.connect(DB_PATH, timeout=20.0)
    c.row_factory = sqlite3.Row
    return c, True


def aktif_alicilar(conn=None):
    """aktif=1 olan alıcı e-postalarının listesi."""
    c, kapat = _db(conn)
    try:
        rows = c.execute(
            "SELECT email FROM mail_alicilari WHERE COALESCE(aktif,1)=1 ORDER BY email"
        ).fetchall()
        return [r['email'] for r in rows]
    finally:
        if kapat:
            c.close()


# ─────────────────────────────────────────────────────────────
# VERİ + EXCEL
# ─────────────────────────────────────────────────────────────
def gunluk_veri(tarih=None, lokasyon=None):
    """O günün üretim satırları: (tesis, bolum, operator, referans, adet).
    lokasyon verilirse yalnız o tesis; verilmezse (None/HEPSI) tüm tesisler.
    """
    if tarih is None:
        tarih = date.today().isoformat()
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.row_factory = sqlite3.Row
    try:
        sql = """
            SELECT
                COALESCE(v.lokasyon, 'TK2')      AS tesis,
                COALESCE(v.bolum, 'kaynak')      AS bolum,
                v.operator_adi                   AS operator,
                u.referans_kodu                  AS referans,
                SUM(u.ok_adet)                   AS adet
            FROM uretim_kayitlari u
            JOIN vardiyalar v ON v.id = u.vardiya_id
            WHERE v.tarih = ?
              AND u.referans_kodu IS NOT NULL AND u.referans_kodu != ''
        """
        params = [tarih]
        if lokasyon and lokasyon.upper() != 'HEPSI':
            sql += " AND COALESCE(v.lokasyon,'TK2') = ?"
            params.append(lokasyon.upper())
        sql += """
            GROUP BY tesis, bolum, v.operator_adi, u.referans_kodu
            HAVING SUM(u.ok_adet) > 0
            ORDER BY tesis, bolum, v.operator_adi, u.referans_kodu
        """
        return conn.execute(sql, params).fetchall()
    finally:
        conn.close()


def excel_olustur(tarih=None, lokasyon=None):
    """Günlük üretim Excel'i üretir. (dosya_yolu, satir_sayisi, toplam_adet) döner."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    if tarih is None:
        tarih = date.today().isoformat()
    rows = gunluk_veri(tarih, lokasyon)

    os.makedirs(RAPOR_DIR, exist_ok=True)
    dosya = os.path.join(RAPOR_DIR, f'{tarih}_UretimRaporu.xlsx')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Günlük Üretim'

    basliklar = ['Tarih', 'Tesis', 'Bölüm', 'Operatör', 'Referans Kodu', 'Adet']
    ws.append(basliklar)

    # Başlık stili (koyu lacivert — Cofle kimliği)
    baslik_fill = PatternFill('solid', fgColor='1E293B')
    baslik_font = Font(bold=True, color='FFFFFF', size=11)
    ince = Side(style='thin', color='D0D7E2')
    kenar = Border(left=ince, right=ince, top=ince, bottom=ince)
    for i in range(1, len(basliklar) + 1):
        c = ws.cell(row=1, column=i)
        c.fill = baslik_fill
        c.font = baslik_font
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = kenar

    tarih_tr = _tarih_tr(tarih)
    toplam_adet = 0
    for r in rows:
        toplam_adet += int(r['adet'] or 0)
        ws.append([
            tarih_tr,
            r['tesis'],
            BOLUM_AD.get(r['bolum'], r['bolum']),
            r['operator'] or '',
            r['referans'],
            int(r['adet'] or 0),
        ])

    # Gövde hücre kenarları + Adet sağa yasla
    for satir in range(2, ws.max_row + 1):
        for sutun in range(1, len(basliklar) + 1):
            hc = ws.cell(row=satir, column=sutun)
            hc.border = kenar
            if sutun == 6:
                hc.alignment = Alignment(horizontal='right')

    # Toplam satırı
    if rows:
        tsatir = ws.max_row + 1
        ws.cell(row=tsatir, column=5, value='TOPLAM').font = Font(bold=True)
        ws.cell(row=tsatir, column=5).alignment = Alignment(horizontal='right')
        tc = ws.cell(row=tsatir, column=6, value=toplam_adet)
        tc.font = Font(bold=True)
        tc.alignment = Alignment(horizontal='right')

    # Kolon genişlikleri
    for kol, gen in zip('ABCDEF', (12, 8, 18, 22, 20, 10)):
        ws.column_dimensions[kol].width = gen
    ws.freeze_panes = 'A2'

    wb.save(dosya)
    _eski_raporlari_temizle()
    return dosya, len(rows), toplam_adet


def _eski_raporlari_temizle():
    if not os.path.isdir(RAPOR_DIR):
        return
    dosyalar = sorted(
        (f for f in os.listdir(RAPOR_DIR) if f.endswith('.xlsx')),
        reverse=True
    )
    for eski in dosyalar[RAPOR_GUN_LIMIT:]:
        try:
            os.remove(os.path.join(RAPOR_DIR, eski))
        except Exception:
            pass


def _tarih_tr(iso):
    """'2026-07-10' → '10.07.2026'."""
    try:
        p = iso.split('-')
        return f'{p[2]}.{p[1]}.{p[0]}'
    except Exception:
        return iso


# ─────────────────────────────────────────────────────────────
# GÖNDERİM
# ─────────────────────────────────────────────────────────────
def _outlook_gonder(alicilar, konu, govde, ek_yol):
    """Açık Outlook masaüstü uygulaması üzerinden gönderir (win32com COM) — ŞİFRE GEREKMEZ.
    Şart: Outlook aynı Windows oturumunda çalışıyor ve giriş yapılmış olmalı.
    Not: COM çağrısı arka planda çalıştığından pythoncom.CoInitialize() zorunlu."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        ol = win32com.client.Dispatch('Outlook.Application')
        mail = ol.CreateItem(0)  # 0 = olMailItem
        mail.To = '; '.join(alicilar)
        mail.Subject = konu
        mail.Body = govde
        if ek_yol and os.path.exists(ek_yol):
            mail.Attachments.Add(os.path.abspath(ek_yol))
        mail.Send()
    finally:
        pythoncom.CoUninitialize()


def _smtp_gonder(cfg, alicilar, konu, govde, ek_yol):
    """Tek SMTP oturumunda maili kurar ve gönderir. Hata fırlatır (çağıran yakalar)."""
    msg = MIMEMultipart()
    gonderen = cfg['gonderen']
    msg['From'] = formataddr((cfg.get('gonderen_ad', 'Cofle Forge'), gonderen))
    msg['To'] = ', '.join(alicilar)
    msg['Subject'] = konu
    msg.attach(MIMEText(govde, 'plain', 'utf-8'))

    if ek_yol and os.path.exists(ek_yol):
        with open(ek_yol, 'rb') as f:
            ek = MIMEApplication(f.read(), _subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            ek.add_header('Content-Disposition', 'attachment', filename=os.path.basename(ek_yol))
            msg.attach(ek)

    host = cfg['smtp_host']
    port = int(cfg['smtp_port'])
    kullanici = cfg.get('kullanici') or gonderen
    sifre = cfg.get('sifre', '')
    mod = str(cfg.get('guvenlik', 'starttls')).lower()  # starttls | ssl | none

    if mod == 'ssl':
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, timeout=30, context=ctx) as srv:
            if sifre:
                srv.login(kullanici, sifre)
            srv.sendmail(gonderen, alicilar, msg.as_string())
    else:
        with smtplib.SMTP(host, port, timeout=30) as srv:
            srv.ehlo()
            if mod == 'starttls':
                srv.starttls(context=ssl.create_default_context())
                srv.ehlo()
            if sifre:
                srv.login(kullanici, sifre)
            srv.sendmail(gonderen, alicilar, msg.as_string())


def gunluk_mail_gonder(tarih=None, zorla_alicilar=None):
    """Günlük raporu üretip alıcılara gönderir. scheduler ve 'şimdi gönder' bunu çağırır.
    zorla_alicilar verilirse (test) DB yerine o listeye gönderir.
    Döner: {basarili, mesaj, ...}
    """
    cfg = config_yukle()
    if not cfg or not cfg.get('etkin'):
        return {'basarili': False, 'mesaj': 'Mail yapılandırılmamış (mail_config.json yok/etkin değil).', 'atlandi': True}
    yontem = yontem_al(cfg)
    if yontem == 'smtp' and not all(cfg.get(k) for k in ('smtp_host', 'smtp_port', 'gonderen')):
        return {'basarili': False, 'mesaj': 'SMTP config eksik: smtp_host / smtp_port / gonderen zorunlu.'}

    if tarih is None:
        tarih = date.today().isoformat()
    lokasyon = cfg.get('kapsam_lokasyon', 'HEPSI')

    alicilar = zorla_alicilar if zorla_alicilar else aktif_alicilar()
    if not alicilar:
        return {'basarili': False, 'mesaj': 'Aktif alıcı yok — dashboard\'dan e-posta ekleyin.', 'atlandi': True}

    try:
        dosya, satir, toplam = excel_olustur(tarih, lokasyon)
    except Exception as e:
        return {'basarili': False, 'mesaj': f'Excel oluşturulamadı: {e}'}

    tarih_tr = _tarih_tr(tarih)
    konu = f'Cofle Forge — Günlük Üretim Raporu ({tarih_tr})'
    if satir:
        govde = (
            f'Merhaba,\n\n'
            f'{tarih_tr} tarihli günlük üretim raporu ektedir.\n\n'
            f'Özet:\n'
            f'  • Kayıt satırı : {satir}\n'
            f'  • Toplam üretim: {toplam:,} adet\n\n'
            f'Detaylar ekteki Excel dosyasındadır.\n\n'
            f'Bu e-posta Cofle Forge tarafından otomatik gönderilmiştir.'
        ).replace(',', '.')
    else:
        govde = (
            f'Merhaba,\n\n'
            f'{tarih_tr} tarihinde sisteme kayıtlı üretim bulunmamaktadır.\n\n'
            f'Bu e-posta Cofle Forge tarafından otomatik gönderilmiştir.'
        )

    try:
        ek = dosya if satir else None
        if yontem == 'outlook':
            _outlook_gonder(alicilar, konu, govde, ek)
        else:
            _smtp_gonder(cfg, alicilar, konu, govde, ek)
    except Exception as e:
        return {'basarili': False, 'mesaj': f'Gönderim hatası: {e}', 'alici_sayisi': len(alicilar)}

    print(f'[MAIL] {tarih} raporu gonderildi -> {len(alicilar)} alici, {satir} satir, {toplam} adet')
    return {
        'basarili': True,
        'mesaj': f'{len(alicilar)} alıcıya gönderildi ({satir} satır, {toplam} adet).',
        'alici_sayisi': len(alicilar),
        'satir': satir,
        'toplam_adet': toplam,
        'dosya': os.path.basename(dosya),
    }


def test_gonder(email, tarih=None):
    """Tek adrese test maili — kurulumu doğrulamak için."""
    return gunluk_mail_gonder(tarih=tarih, zorla_alicilar=[email])


if __name__ == '__main__':
    # Elle test: python mail_raporu.py [YYYY-MM-DD]
    import sys
    t = sys.argv[1] if len(sys.argv) > 1 else None
    print('Config etkin mi:', etkin())
    if t or True:
        d, n, top = excel_olustur(t)
        print(f'Excel: {d}  ({n} satır, {top} adet)')
