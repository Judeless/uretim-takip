# Handoff: Cofle Forge UI v3 — Endüstriyel Açık Tema

## Overview
Yönetici dashboard'u ve operatör (mobile_v2) ekranlarının açık temasının sadeleştirilmiş,
endüstriyel yeniden tasarımı. Hedef repo: **Judeless/uretim-takip** (branch: main).
Kapsam: `[data-theme="light"]` yüzeyleri + bileşen şablonları. **Koyu tema (:root) ve
andon ekranları DEĞİŞMEZ.**

## About the Design Files
`Cofle Forge UI Redesign v3 Endustriyel.dc.html` (support.js ile birlikte) **HTML tasarım
referansıdır** — üretim kodu değildir. Görev: bu görünümü mevcut kod tabanında
(dashboard.html, mobile_v2.html — vanilla HTML/CSS/JS + Chart.js) mevcut sınıf/değişken
altyapısını KULLANARAK yeniden üretmek. Yapı, id'ler ve JS akışları korunur; değişen şey
CSS değişkenleri, bileşen stilleri ve küçük markup rötuşlarıdır.

## Fidelity
**High-fidelity.** Renk, tipografi, ölçüler ve durum renklendirmesi kesindir.

## Design Tokens — [data-theme="light"] bloğu bu değerlerle güncellenir
```
--bg: #e9ebef;            /* nötr gri (lavanta değil) */
--bg-1: #f4f5f7;  --bg-2: #eceef2;
--sidebar-bg: #fbfbfc;    /* düz, saydamlık yok */
--card: #ffffff;          /* düz beyaz, blur yok */
--card-border: #d7dae2;   --border: #d7dae2;  --border-subtle: #eceef2;
--text: #1e1b3a;  --muted: #6b7080;  --dim: #9aa0ad;
--primary: #4f46e5;  --primary-hover: #4338ca;
--success: #16a34a (metin #15803d);  --warning: #d97706 (metin #b45309);
--danger: #dc2626 (metin #b91c1c);  --info: #0e7490;
--ink: #1e1b3a;           /* YENİ: aktif chrome dolgusu */
--shadow: none;           /* gölge kullanılmaz */
--radius: 8px;            /* 16px değil */
Fontlar aynı: Inter (--font-main), JetBrains Mono (--font-mono)
```

## Bileşen kuralları (light temada tüm sayfalara aynı)
1. **Kart:** düz #FFF + 1px #D7DAE2, radius 8, gölgesiz. `backdrop-filter` ve
   `linear-gradient(160deg, rgba(255,255,255,...))` yüzeyleri light temada kapatılır.
2. **Chrome:** aktif sekme / birincil buton = #1E1B3A dolgu + beyaz metin
   (grad-violet pill DEĞİL). İkincil buton: beyaz + 1px #C2C6D1. Segmented kontroller
   köşeli (radius 7), 1px iç ayraç.
3. **Sidebar:** #FBFBFC; aktif öğe #EEF0F6 zemin + solda 3px #4F46E5 çubuk.
   "Tezgâhlar" yerine açılır **Bölümler** grubu: Robot Kaynak / Montaj / Metal
   Enjeksiyon alt öğeleri (sayı rozetli). Topbar'daki `bolum-toggle` KALDIRILIR;
   bölüm filtresi bu alt sekmelere taşınır (mevcut `bolumSec()` mantığı yeniden bağlanır).
   Genel Bakış HER ZAMAN fabrika geneli; başlıkta `TÜM BÖLÜMLER` mono etiketi.
4. **KPI kartı:** üst gradyan şerit ve ::before kaldırılır. Etiket 9.5px/800/+12%
   uppercase #6B7080; değer JetBrains Mono 800. OEE kartında cetvelli gauge:
   12px bar (iz #E7E9EF, dolgu #4F46E5, radius 2) + %85 hedef çentiği (2px #1E1B3A)
   + altında tik cetveli `repeating-linear-gradient(90deg,#C2C6D1 0 1px,transparent 1px 10%)`.
5. **oee-bar (tablo/robot listesi):** 12px, radius 2, DÜZ dolgu — norm #4F46E5,
   %60-84 arası duruma göre #D97706, kritik #DC2626 — + hedef çentiği. Gradyan yok.
6. **Durum ikonları:** emoji kaldırılır (📊📦⏸📡). Yerine 7px kare durum noktası
   (yeşil/amber blink 1.6s/kırmızı blink 1.2s) ve köşeli mono etiketler:
   `SAYAÇ` (yeşil sınır) · `ELLE` (gri) · `ARIZA` (kırmızı, blink) · `● AÇIK` · `KAPANDI`.
   Tag reçetesi: 2px 8px padding, 1px durum-rengi sınır, radius 3, JetBrains Mono 9px/700.
7. **Tablolar:** başlık altında 2px #1E1B3A çizgi; satırlar 1px #ECEEF2 hairline;
   sayısal sütunlar sağa dayalı mono; zebra yok; arıza satırı zemini #FDEFEF;
   hedef/OEE sütununda 6px mini bar.
8. **Grafikler:** saatlik nabız — bar dolgusu #A5B0F1, mola saati #D7DAE2, aktif saat
   #4F46E5 + 2px ink outline, hedef çizgisi 1.5px dashed #9AA0AD. Duruş donutu conic:
   kırmızı/amber/mavi #2563EB/cyan #0E7490. Chart.js paleti bu değerlerden beslenir.
9. **Hareket:** TEK animasyon `cfBlink` (opacity 1→0.3) — arıza/canlı noktalar ve ARIZA
   tag'i. Hover transform/glow animasyonları light temada kaldırılır.
10. **Operatör (mobile_v2):** sayaç halkası conic #4F46E5 / iz #E7E9EF, iç beyaz;
    OK değeri mono 46px. Referans kartı ve NOK/TAMİR stepper'ları beyaz kart reçetesi;
    dokunma hedefleri ≥44px. CTA'lar: VARDİYAYI BAŞLAT = #1E1B3A; DURUŞ BİLDİR = #D97706
    dolu; VARDİYAYI KAPAT = beyaz + 2px #B91C1C sınır. Üst yeşil operatör barı düz
    #EEF8F1 + 1px #B7DCC3.

## Sayfa → repo eşlemesi
| Tasarım sayfası | Repo dosyası / bölge |
| --- | --- |
| 01 Genel Bakış | dashboard.html — sidebar, topbar, kpi-grid, tezgâh OEE listesi, saatlik nabız |
| 02 Üretim | dashboard.html — inner-tab-bar, filter-bar, detaylı üretim kayıtları tablosu |
| 03 Duruşlar | dashboard.html — duruş sebebi kırılımı, duruş tipi özeti, robot bazlı duruş |
| 04 Vardiyalar | dashboard.html — vardiya kayıtları + özet kartlar |
| 05 Vardiya Aç | mobile_v2.html — masa seçimi + PIN |
| 06 Aktif Vardiya | mobile_v2.html — Ekran C (sayaç, referans, NOK/tamir, duruş) |
| 07 Stil Rehberi | token + bileşen reçeteleri (bu README ile aynı) |

## Uygulama sırası (önerilen commit'ler)
1. `[data-theme="light"]` değişken bloğunu güncelle (yukarıdaki tokens) — tek commit.
2. Light temaya özel override'lar: kart/kpi/filter/tab/tablo sınıflarında blur-gradyan-gölge
   kapatma (`[data-theme="light"] .kpi-card { background:#fff; box-shadow:none; ... }`).
3. Bar bileşenleri: .oee-bar yükseklik/renk/çentik + cetvel tiki; hedef çentiği için
   `.oee-bar-target` elemanı ekle (JS'te width set eden yerlere dokunma).
4. Emoji → durum noktası + mono tag değişimi (innerSecme başlıkları, badge render eden JS
   şablonları: `sayacBadge`, durum badge'leri).
5. Sidebar "Bölümler" grubu + bolum-toggle kaldırma; `bolumSec()` çağrılarını nav'a bağla;
   Özet sekmesini bölüm filtresiz (tümü) çalıştır.
6. mobile_v2 light rötuşları (CTA renkleri, kart reçetesi, stepper).
7. Chart.js renk güncellemesi (light temada).
Her adımda koyu tema regresyonu kontrol edilir (değişiklikler `[data-theme="light"]`
seçicisi altında kalmalı; ortak bileşen değişimlerinde iki temada da test).

## Files
- `Cofle Forge UI Redesign v3 Endustriyel.dc.html` + `support.js` — tarayıcıda açılır referans
- Kök projede: v2 (renkli cam) ve koyu ilk deneme dosyaları — uygulanmayacak, yalnız tarihçe
