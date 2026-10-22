/* BÖLÜM KAPASİTESİ → PowerPoint sunumu (kullanıcı 2026-10-22: "kapasite çıktısına göre üst
 * yönetime sunum yapmak için pp olarak sunum indirebilelim — grafikli görselli").
 *
 * Panelin o anki hesabından (kişi / gün / verim düzeltmeleri, devir, 4-8 hafta görünümü) üretilir:
 * veri dashboard'daki bkSunumVerisi()'nden gelir, burada YALNIZ çizilir. Tarayıcıda PptxGenJS
 * (CDN) ile, laptopta node + pptxgenjs ile aynı fonksiyon çalışır (görsel kontrol için).
 *   bkSunumOlustur(PptxGenJS, veri, { logo: 'data:image/png;base64,…' }) → pptx nesnesi
 */
(function (kok) {
    'use strict';

    var R = {
        koyu: '17152E', mor: '6D28D9', morAcik: 'C4B5FD', morZemin: 'F4F1FD',
        metin: '14122B', ikinci: '4E4C63', soluk: '6B6982', cizgi: 'E3E1EC', beyaz: 'FFFFFF',
        mevcut: '94A3B8',
        ok: '157A3C', okZ: 'E4F3EA', uyari: '8A5A05', uyariZ: 'FBF1DC', asim: 'B4231C', asimZ: 'FBE4E2',
        yokZ: 'F1F1F5'
    };
    var YAZI = 'Calibri';

    function sayi(x, b) {
        if (x == null || !isFinite(x)) return '—';
        return Number(x).toLocaleString('tr-TR', { minimumFractionDigits: b == null ? 1 : b, maximumFractionDigits: b == null ? 1 : b });
    }
    function durum(oran) {
        if (oran == null || !isFinite(oran)) return { ad: 'yok', renk: R.soluk, zemin: R.yokZ, isaret: '' };
        if (oran > 100) return { ad: 'asim', renk: R.asim, zemin: R.asimZ, isaret: '▲' };
        if (oran > 85) return { ad: 'uyari', renk: R.uyari, zemin: R.uyariZ, isaret: '!' };
        return { ad: 'ok', renk: R.ok, zemin: R.okZ, isaret: '✓' };
    }
    function yuzde(oran) { return oran == null || !isFinite(oran) ? '—' : '%' + Math.round(oran); }
    function bolumAdi(b) { return (b.tesis === 'TK2' ? '' : b.tesis + ' ') + b.ad; }

    // Başlık + alt bilgi (her içerik sayfası)
    function sayfa(p, v, baslik, alt) {
        var s = p.addSlide();
        s.background = { color: R.beyaz };
        s.addText(baslik, { x: 0.6, y: 0.35, w: 12.1, h: 0.7, fontFace: YAZI, fontSize: 30, bold: true, color: R.metin, margin: 0, isTextBox: true });
        if (alt) s.addText(alt, { x: 0.6, y: 1.02, w: 12.1, h: 0.4, fontFace: YAZI, fontSize: 14, color: R.ikinci, margin: 0, isTextBox: true });
        s.addText('Cofle Forge · Bölüm Kapasitesi · ' + v.tarih + ' · ' + v.donem, {
            x: 0.6, y: 7.05, w: 9, h: 0.3, fontFace: YAZI, fontSize: 10, color: R.soluk, margin: 0, isTextBox: true });
        s.slideNumber = { x: 12.3, y: 7.05, w: 0.5, h: 0.3, fontFace: YAZI, fontSize: 10, color: R.soluk };
        return s;
    }

    function kpiKart(p, s, x, y, w, h, deger, etiket, renk, not) {
        s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: R.morZemin }, line: { color: R.morZemin }, rectRadius: 0.12 });
        s.addText(deger, { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.85, fontFace: YAZI, fontSize: 38, bold: true, color: renk || R.metin, margin: 0, isTextBox: true });
        s.addText(etiket, { x: x + 0.25, y: y + 1.05, w: w - 0.5, h: 0.4, fontFace: YAZI, fontSize: 13, bold: true, color: R.ikinci, margin: 0, isTextBox: true });
        if (not) s.addText(not, { x: x + 0.25, y: y + 1.42, w: w - 0.5, h: 0.4, fontFace: YAZI, fontSize: 11, color: R.soluk, margin: 0, isTextBox: true });
    }

    function bkSunumOlustur(PptxGen, v, sec) {
        sec = sec || {};
        var p = new PptxGen();
        p.layout = 'LAYOUT_WIDE';                    // 13,33 × 7,5 in
        p.author = 'Cofle Forge';
        p.title = 'Bölüm Kapasite Analizi ' + v.donem;

        // Sunuma girenler: kişisi ya da yükü olan bölümler; fason (dış üretici) özet toplamlara girmez
        var bol = v.bolumler.filter(function (b) { return b.kisi > 0 || (b.ort || 0) >= 0.05; });
        var ic = bol.filter(function (b) { return b.tesis === 'TK2' || b.tesis === 'TK1'; });
        var olculen = ic.filter(function (b) { return b.kisi > 0 && b.ort != null && b.ort >= 0.05; });
        function ek(b) { return Math.max(0, (b.ort || 0) - (b.kisi || 0)); }
        var siraliEk = olculen.slice().sort(function (a, b) { return ek(b) - ek(a) || (b.oran || 0) - (a.oran || 0); });
        var nH = v.haftalar.length;

        // ── 1) KAPAK ─────────────────────────────────────────────────────────────
        var s = p.addSlide();
        s.background = { color: R.koyu };
        if (sec.logo) s.addImage({ data: sec.logo, x: 0.6, y: 0.55, w: 3.5, h: 1.05 });
        s.addText('Bölüm Kapasite Analizi', { x: 0.6, y: 2.55, w: 12, h: 1.0, fontFace: YAZI, fontSize: 46, bold: true, color: R.beyaz, margin: 0, isTextBox: true });
        s.addText(v.donem + ' · ' + nH + ' haftalık görünüm · ' + v.haftalar.reduce(function (t, h) { return t + h.gun; }, 0) + ' çalışma günü', {
            x: 0.6, y: 3.6, w: 12, h: 0.6, fontFace: YAZI, fontSize: 24, color: R.morAcik, margin: 0, isTextBox: true });
        s.addText('Hazırlanma: ' + v.tarih, { x: 0.6, y: 4.3, w: 12, h: 0.5, fontFace: YAZI, fontSize: 16, color: 'D9D6EC', margin: 0, isTextBox: true });
        s.addText('Veri: AS400 açık iş emirleri (son çekim ' + v.talep_ts + ') · Birim süre ve bölüm: Ana Veri (' + v.ana_veri + ')', {
            x: 0.6, y: 6.7, w: 12.1, h: 0.4, fontFace: YAZI, fontSize: 12, color: 'A5A3BD', margin: 0, isTextBox: true });

        // ── 2) YÖNETİCİ ÖZETİ ──────────────────────────────────────────────────────
        s = sayfa(p, v, 'Yönetici özeti', v.donem + ' dönem ortalaması · mevcut kişi sayılarıyla');
        var tk2 = olculen.filter(function (b) { return b.tesis === 'TK2'; });
        var gerekTk2 = tk2.reduce(function (t, b) { return t + b.ort; }, 0), mevcutTk2 = tk2.reduce(function (t, b) { return t + b.kisi; }, 0);
        var genelOran = mevcutTk2 > 0 ? gerekTk2 / mevcutTk2 * 100 : null;
        var asan = olculen.filter(function (b) { return b.oran > 100; });
        var bakiye = ic.reduce(function (t, b) { return t + (b.bakiye_kh || 0); }, 0);
        var ekKisi = olculen.reduce(function (t, b) { return t + Math.max(0, b.ort - b.kisi); }, 0);
        var ekSaat = olculen.reduce(function (t, b) { return t + Math.max(0, b.ort - b.kisi) * b.haftalik_saat; }, 0);
        kpiKart(p, s, 0.6, 1.7, 2.9, 1.95, yuzde(genelOran), 'Ortalama doluluk · TK2', durum(genelOran).renk,
                sayi(gerekTk2) + ' kişi gerekli / ' + sayi(mevcutTk2, 0) + ' mevcut');
        kpiKart(p, s, 3.7, 1.7, 2.9, 1.95, asan.length + ' / ' + olculen.length, 'Kapasiteyi aşan bölüm', asan.length ? R.asim : R.ok,
                asan.length ? asan.map(function (b) { return b.ad; }).slice(0, 3).join(', ') : 'tüm bölümler kapasite içinde');
        kpiKart(p, s, 6.8, 1.7, 2.9, 1.95, sayi(bakiye), 'Bakiye · kişi-hafta', bakiye >= 1 ? R.asim : R.metin, 'bitiş tarihi geçmiş açık iş');
        kpiKart(p, s, 9.9, 1.7, 2.85, 1.95, '+' + sayi(ekKisi), 'Ek kişi ihtiyacı', ekKisi > 0.05 ? R.asim : R.ok,
                '≈ ' + sayi(ekSaat, 0) + ' saat/hafta fazla mesai');
        var maddeler = [];
        siraliEk.filter(function (b) { return ek(b) > 0.05; }).slice(0, 3).forEach(function (b) {
            maddeler.push(bolumAdi(b) + ': dönemde ortalama ' + sayi(b.ort) + ' kişi gerekiyor, mevcut ' + sayi(b.kisi, b.kisi % 1 ? 1 : 0) +
                          ' kişi (' + yuzde(b.oran) + ') → +' + sayi(ek(b)) + ' kişi.');
        });
        if (v.devirli) {
            olculen.filter(function (b) { return (b.bakiye_kh || 0) >= 0.5; })
                   .sort(function (a, b) { return (b.donem_sonu_kh || 0) - (a.donem_sonu_kh || 0); }).slice(0, 2).forEach(function (b) {
                maddeler.push(bolumAdi(b) + ': mevcut kişiyle bakiye ' + (b.donem_sonu_kh > 0.05
                    ? 'dönem içinde erimiyor — sonunda ' + sayi(b.donem_sonu_kh) + ' kişi-hafta kalıyor.'
                    : (b.erime_haftasi ? b.erime_haftasi + ' haftasında eriyor.' : 'dönem içinde eriyor.')));
            });
        }
        ic.filter(function (b) { return b.kapsama != null && b.kapsama < 80 && b.ort >= 0.05; })
          .sort(function (a, b) { return a.kapsama - b.kapsama; }).slice(0, 1).forEach(function (b) {
            maddeler.push(bolumAdi(b) + ': süre kapsaması %' + Math.round(b.kapsama) + ' — süresi Ana Veri\'de olmayan ürünlerin yükü tahmin.');
        });
        s.addText('Öne çıkanlar', { x: 0.6, y: 3.95, w: 12, h: 0.45, fontFace: YAZI, fontSize: 18, bold: true, color: R.metin, margin: 0, isTextBox: true });
        s.addText(maddeler.map(function (m, i) { return { text: m, options: { bullet: true, breakLine: i < maddeler.length - 1 } }; }), {
            x: 0.6, y: 4.45, w: 12.1, h: 2.45, fontFace: YAZI, fontSize: 15, color: R.metin, paraSpaceAfter: 6, valign: 'top', margin: 0, isTextBox: true });

        // ── 3) GEREKEN vs MEVCUT KİŞİ ────────────────────────────────────────────────
        s = sayfa(p, v, 'Gereken ve mevcut kişi', 'Dönem ortalaması (' + v.donem + ')' + (v.devirli ? ' · bakiye dahil' : ''));
        var g3 = olculen.slice().concat(ic.filter(function (b) { return !(b.kisi > 0) && b.ort >= 0.05; }));
        s.addChart(p.charts.BAR, [
            { name: 'Gereken kişi (ort.)', labels: g3.map(bolumAdi), values: g3.map(function (b) { return Math.round(b.ort * 10) / 10; }) },
            { name: 'Mevcut kişi', labels: g3.map(bolumAdi), values: g3.map(function (b) { return b.kisi; }) }
        ], {
            x: 0.6, y: 1.55, w: 8.4, h: 5.35, barDir: 'bar', barGrouping: 'clustered', barGapWidthPct: 55,
            chartColors: [R.mor, R.mevcut], catAxisOrientation: 'maxMin',
            showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', dataLabelFontSize: 10, dataLabelColor: R.metin,
            catAxisLabelColor: R.metin, catAxisLabelFontSize: 12, catAxisLabelFontFace: YAZI,
            valAxisLabelColor: R.soluk, valAxisLabelFontSize: 10, valGridLine: { color: R.cizgi, size: 0.5 }, catGridLine: { style: 'none' },
            showLegend: true, legendPos: 'b', legendFontSize: 12, legendFontFace: YAZI
        });
        var t3 = [[
            { text: 'Bölüm', options: { bold: true, color: R.ikinci, fill: { color: R.morZemin } } },
            { text: 'Doluluk', options: { bold: true, color: R.ikinci, fill: { color: R.morZemin }, align: 'center' } }
        ]];
        g3.forEach(function (b) {
            var d = durum(b.oran);
            t3.push([{ text: bolumAdi(b) }, { text: (d.isaret ? d.isaret + ' ' : '') + (b.kisi > 0 ? yuzde(b.oran) : 'kişi yok'),
                options: { bold: true, color: d.renk, fill: { color: d.zemin }, align: 'center' } }]);
        });
        s.addTable(t3, { x: 9.3, y: 1.6, w: 3.45, colW: [2.15, 1.3], fontFace: YAZI, fontSize: 12, color: R.metin,
                         border: { type: 'solid', pt: 0.5, color: R.cizgi }, rowH: 0.36, valign: 'middle' });
        s.addText('Doluluk = gereken ÷ mevcut kişi · ✓ ≤%85 · ! %85–100 · ▲ %100 üstü', {
            x: 9.3, y: 6.55, w: 3.45, h: 0.35, fontFace: YAZI, fontSize: 9, color: R.soluk, margin: 0, isTextBox: true });

        // ── 4) HAFTALIK TABLO ─────────────────────────────────────────────────────────
        s = sayfa(p, v, 'Haftalık gereken kişi', 'Her hücre: o hafta gereken kişi ve mevcut kişiye göre doluluk' +
                  (v.devirli ? ' · ilk hafta bakiyeyi taşır, bitmeyen iş sonraki haftaya devreder' : ''));
        var bas4 = [{ text: 'Bölüm' }, { text: 'Mevcut' }, { text: 'Bakiye\n(kişi-hf)' }].concat(v.haftalar.map(function (h) {
            return { text: h.no + '\n' + h.tarih + ' · ' + sayi(h.gun, h.gun % 1 ? 1 : 0) + ' gün' + (h.kismi ? ' (kalan)' : '') };
        })).concat([{ text: 'Dönem\nort.' }]);
        bas4.forEach(function (c) { c.options = { bold: true, color: R.beyaz, fill: { color: R.koyu }, align: 'center' }; });
        var t4 = [bas4];
        var bol4 = ic.filter(function (b) { return (b.ort || 0) >= 0.05; });
        bol4.forEach(function (b) {
            var satir = [{ text: bolumAdi(b), options: { bold: true } },
                         { text: sayi(b.kisi, b.kisi % 1 ? 1 : 0), options: { align: 'center' } },
                         { text: (b.bakiye_kh || 0) >= 0.05 ? sayi(b.bakiye_kh) : '—', options: { align: 'center', color: R.ikinci } }];
            b.hafta.forEach(function (h) {
                var d = durum(b.kisi > 0 ? h.oran : null);
                var m = h.gereken == null ? 'çalışma yok' : h.gereken < 0.05 ? (h.adet > 0 ? '<0,1' : '—')
                      : sayi(h.gereken) + (b.kisi > 0 ? '\n' + d.isaret + ' ' + yuzde(h.oran) : '');
                satir.push({ text: m, options: { align: 'center', color: h.gereken >= 0.05 ? d.renk : R.soluk,
                                                 fill: { color: h.gereken >= 0.05 ? d.zemin : R.beyaz }, bold: h.gereken >= 0.05 } });
            });
            var dO = durum(b.kisi > 0 ? b.oran : null);
            satir.push({ text: b.ort >= 0.05 ? sayi(b.ort) + (b.kisi > 0 ? '\n' + dO.isaret + ' ' + yuzde(b.oran) : '') : '—',
                         options: { align: 'center', bold: true, color: b.ort >= 0.05 ? dO.renk : R.soluk, fill: { color: b.ort >= 0.05 ? dO.zemin : R.beyaz } } });
            t4.push(satir);
        });
        var genis = 12.15, sabit = [2.1, 0.85, 0.95], son = 1.05, hw = (genis - sabit[0] - sabit[1] - sabit[2] - son) / nH;
        var satH = Math.min(0.55, 5.2 / t4.length);
        s.addTable(t4, { x: 0.6, y: 1.5, w: genis, colW: sabit.concat(v.haftalar.map(function () { return hw; })).concat([son]),
                         fontFace: YAZI, fontSize: nH > 4 ? 10 : 11, color: R.metin, border: { type: 'solid', pt: 0.5, color: R.cizgi },
                         rowH: satH, valign: 'middle', margin: 0.04 });

        // ── 5) BÖLÜM SAYFALARI ─────────────────────────────────────────────────────────
        ic.filter(function (b) { return b.ort >= 0.05; }).forEach(function (b) {
            var d = durum(b.kisi > 0 ? b.oran : null);
            s = sayfa(p, v, bolumAdi(b) + ' — haftalık iş yükü', 'Gereken kişi (haftanın çalışma gününe göre)' +
                      (v.devirli ? ' · önceki haftadan devreden iş dahil' : '') + ' · kesik çizgi: mevcut kişi');
            var et = v.haftalar.map(function (h) { return h.no + ' (' + sayi(h.gun, h.gun % 1 ? 1 : 0) + 'g)'; });
            var cubuk = [{ name: 'Haftanın işi', labels: et, values: b.hafta.map(function (h) { return Math.round((h.kendi_kisi || 0) * 10) / 10; }) }];
            if (v.devirli) cubuk.push({ name: 'Önceki haftadan devreden', labels: et,
                                         values: b.hafta.map(function (h) { return Math.round((h.devir_kisi || 0) * 10) / 10; }) });
            s.addChart([
                { type: p.charts.BAR, data: cubuk, options: { barGrouping: 'stacked', chartColors: [R.mor, R.morAcik], barGapWidthPct: 45 } },
                { type: p.charts.LINE, data: [{ name: 'Mevcut kişi', labels: et, values: b.hafta.map(function () { return b.kisi; }) }],
                  options: { chartColors: [R.metin], lineSize: 2, lineDash: 'dash', lineDataSymbol: 'none' } }
            ], {
                x: 0.6, y: 1.55, w: 8.3, h: 5.35, catAxisLabelColor: R.ikinci, catAxisLabelFontSize: 11, catAxisLabelFontFace: YAZI,
                valAxisLabelColor: R.soluk, valAxisLabelFontSize: 10, valGridLine: { color: R.cizgi, size: 0.5 }, catGridLine: { style: 'none' },
                showValAxisTitle: true, valAxisTitle: 'kişi', valAxisTitleColor: R.soluk, valAxisTitleFontSize: 11,
                showLegend: true, legendPos: 'b', legendFontSize: 11, legendFontFace: YAZI
            });
            // Sağ panel: anahtar rakamlar
            s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 9.2, y: 1.6, w: 3.55, h: 5.25, fill: { color: R.morZemin }, line: { color: R.morZemin }, rectRadius: 0.1 });
            s.addText(b.kisi > 0 ? yuzde(b.oran) : 'kişi yok', { x: 9.45, y: 1.8, w: 3.1, h: 0.85, fontFace: YAZI, fontSize: 40, bold: true, color: d.renk, margin: 0, isTextBox: true });
            s.addText('dönem ortalaması doluluk', { x: 9.45, y: 2.62, w: 3.1, h: 0.35, fontFace: YAZI, fontSize: 12, color: R.ikinci, margin: 0, isTextBox: true });
            var satirlar = [
                ['Mevcut kişi', sayi(b.kisi, b.kisi % 1 ? 1 : 0)],
                ['Gereken kişi (ort.)', sayi(b.ort)],
                ['Bakiye', sayi(b.bakiye_kh) + ' kişi-hafta'],
                ['Dönem sonu', v.devirli ? (b.donem_sonu_kh > 0.05 ? sayi(b.donem_sonu_kh) + ' kişi-hafta kalır' : (b.erime_haftasi ? b.erime_haftasi + ' haftasında erir' : 'dönem içinde erir')) : '—'],
                ['Süre kapsaması', b.kapsama == null ? '—' : '%' + Math.round(b.kapsama)],
                ['Günlük saat · verim', sayi(b.gunluk_saat) + ' sa · %' + Math.round(b.verim)]
            ];
            s.addTable(satirlar.map(function (r) {
                return [{ text: r[0], options: { color: R.ikinci } }, { text: r[1], options: { bold: true, align: 'right' } }];
            }), { x: 9.4, y: 3.1, w: 3.15, colW: [1.6, 1.55], fontFace: YAZI, fontSize: 12, color: R.metin, rowH: 0.42, valign: 'middle',
                  border: { type: 'solid', pt: 0.5, color: R.cizgi } });
            var ek = Math.max(0, b.ort - b.kisi);
            var yorum = !(b.kisi > 0) ? 'Mevcut kişi girilmemiş; ' + sayi(b.ort) + ' kişilik iş var.'
                : ek > 0.05 ? 'Dönemi karşılamak için +' + sayi(ek) + ' kişi ya da ≈ ' + sayi(ek * b.haftalik_saat, 0) + ' saat/hafta fazla mesai gerekiyor.'
                : 'Mevcut kişi dönemi karşılıyor; ' + sayi(b.kisi - b.ort) + ' kişilik pay var.';
            s.addText(yorum, { x: 9.45, y: 5.75, w: 3.1, h: 1.0, fontFace: YAZI, fontSize: 12, color: R.metin, valign: 'top', margin: 0, isTextBox: true });
        });

        // ── 6) ÖNERİLER ──────────────────────────────────────────────────────────────
        s = sayfa(p, v, 'Öneriler', 'Doluluğu %85 üstündeki bölümler · dönem ortalamasına göre');
        var riskli = siraliEk.filter(function (b) { return b.oran > 85; });
        if (!riskli.length) {
            s.addText('Tüm bölümler mevcut kişiyle dönem ortalamasında %85 doluluğun altında.', {
                x: 0.6, y: 2.0, w: 12, h: 0.6, fontFace: YAZI, fontSize: 20, color: R.ok, bold: true, margin: 0, isTextBox: true });
        } else {
            var t6 = [['Bölüm', 'Doluluk', 'Gereken (ort.)', 'Mevcut', 'Ek kişi', 'veya fazla mesai', 'Bakiye'].map(function (x) {
                return { text: x, options: { bold: true, color: R.beyaz, fill: { color: R.koyu }, align: 'center' } };
            })];
            riskli.forEach(function (b) {
                var d = durum(b.oran), ek = Math.max(0, b.ort - b.kisi);
                t6.push([
                    { text: bolumAdi(b), options: { bold: true } },
                    { text: d.isaret + ' ' + yuzde(b.oran), options: { bold: true, color: d.renk, fill: { color: d.zemin }, align: 'center' } },
                    { text: sayi(b.ort), options: { align: 'center' } },
                    { text: sayi(b.kisi, b.kisi % 1 ? 1 : 0), options: { align: 'center' } },
                    { text: ek > 0.05 ? '+' + sayi(ek) + ' kişi' : '—', options: { align: 'center', bold: true } },
                    { text: ek > 0.05 ? '≈ ' + sayi(ek * b.haftalik_saat, 0) + ' saat/hafta' : '—', options: { align: 'center' } },
                    { text: (b.bakiye_kh || 0) >= 0.05 ? sayi(b.bakiye_kh) + ' kişi-hf' : '—', options: { align: 'center' } }
                ]);
            });
            s.addTable(t6, { x: 0.6, y: 1.6, w: 12.15, colW: [2.45, 1.45, 1.6, 1.25, 1.6, 2.2, 1.6], fontFace: YAZI, fontSize: 14,
                             color: R.metin, rowH: 0.5, valign: 'middle', border: { type: 'solid', pt: 0.5, color: R.cizgi } });
            s.addText('Ek kişi: dönem ortalamasını karşılamak için gereken fark. Fazla mesai: aynı farkı mevcut kişiyle karşılamak için haftalık toplam ek saat (verim dahil).', {
                x: 0.6, y: 6.45, w: 12.1, h: 0.45, fontFace: YAZI, fontSize: 11, color: R.soluk, margin: 0, isTextBox: true });
        }

        // ── 7) YÖNTEM VE VARSAYIMLAR ─────────────────────────────────────────────────────
        s = sayfa(p, v, 'Yöntem ve varsayımlar', null);
        var y = [
            'İhtiyaç: AS400 açık iş emirleri (OPR + launch), kalan adet; emrin bitiş haftasına yazılır, Pazartesi bitişli emir önceki haftaya.',
            'Bölüm ve birim süre: Ana Veri (kaynakta söktak + kaynak süresi; süre yoksa vardiya başı kapasite).',
            'Süresi olmayan ürünler bölümün ortanca süresiyle ' + (v.tahmin ? 'tahmin edildi.' : 'hesaba katılmadı.'),
            'Gereken kişi = kişi-saat ÷ (haftanın çalışma günü × günlük saat); kişi-saat = standart süre ÷ verim.',
            v.devirli ? 'Devir: bakiye ilk haftaya, mevcut kişiyle bitmeyen iş sonraki haftaya geçer.'
                      : (v.yay ? 'Bakiye ilk 4 haftaya çalışma günü oranında dağıtıldı.' : 'Bakiye haftalık doluluğa dahil edilmedi.'),
            'Çalışma günü: hafta içi − resmî tatil − arife; içinde bulunulan haftada yalnız kalan günler.',
            'Ana Veri\'de bölümü olmayan üretim kodları (' + (v.atanmamis.kod || 0) + ' kod · ' +
                Number(v.atanmamis.adet || 0).toLocaleString('tr-TR') + ' adet) hesaba girmedi.'
        ];
        s.addText(y.map(function (m, i) { return { text: m, options: { bullet: true, breakLine: i < y.length - 1 } }; }), {
            x: 0.6, y: 1.3, w: 7.3, h: 5.6, fontFace: YAZI, fontSize: 14, color: R.metin, paraSpaceAfter: 8, valign: 'top', margin: 0, isTextBox: true });
        var t7 = [['Hafta', 'Tarih', 'Gün', 'Not'].map(function (x) {
            return { text: x, options: { bold: true, color: R.beyaz, fill: { color: R.koyu }, align: 'center' } };
        })];
        v.haftalar.forEach(function (h) {
            var not = [];
            if (h.kismi) not.push('kalan günler');
            if (h.elle) not.push('elle düzeltildi');
            h.tatil.forEach(function (t) { not.push(t); });
            t7.push([{ text: h.no, options: { bold: true, align: 'center' } }, { text: h.tarih, options: { align: 'center' } },
                     { text: sayi(h.gun, h.gun % 1 ? 1 : 0), options: { align: 'center', bold: true } },
                     { text: not.join(' · ') || '—', options: { color: R.ikinci } }]);
        });
        s.addTable(t7, { x: 8.2, y: 1.35, w: 4.55, colW: [0.75, 0.85, 0.6, 2.35], fontFace: YAZI, fontSize: 11, color: R.metin,
                         rowH: 0.38, valign: 'middle', border: { type: 'solid', pt: 0.5, color: R.cizgi } });
        return p;
    }

    kok.bkSunumOlustur = bkSunumOlustur;
    if (typeof module !== 'undefined' && module.exports) module.exports = { bkSunumOlustur: bkSunumOlustur };
})(typeof window !== 'undefined' ? window : globalThis);
