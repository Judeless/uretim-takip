
        // ═══════════════════════════════════════════════════════
        // STATE
        // ═══════════════════════════════════════════════════════
        let chartDurus = null;
        let chartRobot = null;
        let aktifSekme = 'vardiyalar';
        let ozetData = null;
        // Aktif bolum: 'kaynak' | 'montaj' | 'metal'. localStorage'dan okunur, default 'kaynak'.
        let aktifBolum = (function () {
            const v = localStorage.getItem('dashboard_aktif_bolum');
            return (v === 'montaj' || v === 'metal') ? v : 'kaynak';
        })();

        // Bolum bilgileri: etiketler ve sabit makine listeleri
        const BOLUM_BILGI = {
            kaynak: { ad: 'Robot Kaynak', tekil_etiket: 'Robot' },
            montaj: { ad: 'Montaj',        tekil_etiket: 'Hat' },
            metal:  { ad: 'Metal Enjeksiyon', tekil_etiket: 'Makine' }
        };

        // URL'lere bolum parametresi eklemek icin helper. Url '?' icermiyorsa otomatik ekler.
        function bolumQS(prefix) {
            const sep = (prefix === undefined || prefix === '?') ? '?' : '&';
            return sep + 'bolum=' + encodeURIComponent(aktifBolum);
        }

        // Bolume gore robot dropdown'unu ve etiketi yenile
        async function yukleRobotListesi() {
            const sel = document.getElementById('f-robot');
            const lbl = document.getElementById('f-robot-label');
            if (lbl) lbl.textContent = (BOLUM_BILGI[aktifBolum] || {}).tekil_etiket || 'Robot';
            if (!sel) return;
            // Mevcut secimi koru
            const onceki = sel.value;
            sel.innerHTML = '<option value="">Tümü</option>';
            try {
                const res = await fetch('/api/robotlar?bolum=' + encodeURIComponent(aktifBolum));
                const robotlar = await res.json();
                robotlar.forEach(r => {
                    const o = document.createElement('option');
                    o.value = r; o.textContent = r;
                    sel.appendChild(o);
                });
                // Eski secim hala listede varsa geri yukle
                if (onceki && Array.from(sel.options).some(op => op.value === onceki)) {
                    sel.value = onceki;
                }
            } catch (e) { console.warn('Robot listesi yüklenemedi:', e); }
        }

        // Sekmeleri ve butonlari bolume gore goster/gizle
        function bolumGorunurlugu() {
            // data-bolum-only="kaynak" olan inner-tab butonlarini gizle/goster
            document.querySelectorAll('[data-bolum-only]').forEach(el => {
                const izinli = el.getAttribute('data-bolum-only');
                el.style.display = (izinli === aktifBolum) ? '' : 'none';
            });
            // Aktif inner sekme gizlendiyse 'ozet'e dus
            const aktifBtn = document.querySelector('#inner-tab-bar .inner-tab-btn.active');
            if (aktifBtn && aktifBtn.style.display === 'none') {
                const ozetBtn = document.querySelector('#inner-tab-bar .inner-tab-btn[onclick*="\'ozet\'"]');
                if (ozetBtn) innerSecme('ozet', ozetBtn);
            }
            // Topbar buton aktif sinifi
            document.querySelectorAll('.bolum-btn').forEach(b => {
                b.classList.toggle('active', b.dataset.bolum === aktifBolum);
            });
            // Subtitle bolum adini yansitsin
            const sub = document.getElementById('subtitle-text');
            if (sub) {
                const bilgi = BOLUM_BILGI[aktifBolum] || BOLUM_BILGI.kaynak;
                sub.textContent = bilgi.ad;
            }
        }

        // Bolum degistir: state guncelle, kalici yap, UI'yi tazele
        async function bolumDegistir(yeni) {
            if (yeni !== 'kaynak' && yeni !== 'montaj' && yeni !== 'metal') return;
            if (yeni === aktifBolum) return;
            aktifBolum = yeni;
            localStorage.setItem('dashboard_aktif_bolum', yeni);
            // Robot filtresini sifirla (eski bolumun degerleri yeni bolumde anlamsiz)
            const sel = document.getElementById('f-robot');
            if (sel) sel.value = '';
            bolumGorunurlugu();
            await yukleRobotListesi();
            yukle();
            try { yukleReferanslar(); } catch (e) {}
        }

        // ─── EXPORT MODAL ───────────────────────────────────────
        function exportModalAc() {
            // Modal yoksa oluştur
            if (!document.getElementById('export-modal')) {
                const modal = document.createElement('div');
                modal.id = 'export-modal';
                modal.style.cssText = `
                    position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:500;
                    display:flex;align-items:center;justify-content:center;
                `;
                const bugun = new Date().toISOString().split('T')[0];
                const birAyOnce = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
                modal.innerHTML = `
                    <div style="background:#fff;border-radius:16px;padding:28px;width:420px;box-shadow:0 8px 40px rgba(0,0,0,.25)">
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px">
                            <h3 style="font-size:1rem;font-weight:800">📥 Excel'e Aktar</h3>
                            <button onclick="exportModalKapat()" style="background:#f1f5f9;border:none;border-radius:8px;padding:6px 12px;cursor:pointer;font-size:1rem">✕</button>
                        </div>
                        <div style="font-size:.82rem;color:#64748b;margin-bottom:16px;background:#f8fafc;padding:12px;border-radius:8px;border-left:3px solid #2563eb">
                            Veriler <b>Masaüstü/UretimTakipArsiv.xlsx</b> dosyasına kaydedilir.<br>
                            Dosya her kayıtta güncellenir (üzerine yazar).
                        </div>
                        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px">
                            <div>
                                <label style="font-size:.78rem;font-weight:700;color:#64748b;display:block;margin-bottom:4px">Başlangıç Tarihi</label>
                                <input type="date" id="exp-bas" value="${birAyOnce}" style="width:100%;padding:10px 12px;border:2px solid #e2e8f0;border-radius:8px;font-size:.88rem">
                            </div>
                            <div>
                                <label style="font-size:.78rem;font-weight:700;color:#64748b;display:block;margin-bottom:4px">Bitiş Tarihi</label>
                                <input type="date" id="exp-bit" value="${bugun}" style="width:100%;padding:10px 12px;border:2px solid #e2e8f0;border-radius:8px;font-size:.88rem">
                            </div>
                        </div>
                        <div style="display:flex;gap:8px">
                            <button onclick="exportModalKapat()" style="flex:1;padding:12px;border:2px solid #e2e8f0;border-radius:10px;background:#f8fafc;font-weight:700;cursor:pointer">İptal</button>
                            <button id="exp-btn" onclick="exportArsiv()" style="flex:2;padding:12px;background:#16a34a;color:#fff;border:none;border-radius:10px;font-weight:800;font-size:.92rem;cursor:pointer">
                                📥 Excel'e Kaydet
                            </button>
                        </div>
                        <div id="exp-sonuc" style="margin-top:14px;display:none"></div>
                    </div>
                `;
                document.body.appendChild(modal);
                modal.addEventListener('click', e => { if (e.target === modal) exportModalKapat(); });
            } else {
                document.getElementById('export-modal').style.display = 'flex';
            }
            document.getElementById('exp-sonuc').style.display = 'none';
        }

        function exportModalKapat() {
            const m = document.getElementById('export-modal');
            if (m) m.style.display = 'none';
        }

        async function exportArsiv() {
            const btn = document.getElementById('exp-btn');
            const sonucEl = document.getElementById('exp-sonuc');
            const bas = document.getElementById('exp-bas').value;
            const bit = document.getElementById('exp-bit').value;

            btn.textContent = '⏳ Kaydediliyor...';
            btn.disabled = true;
            sonucEl.style.display = 'none';

            try {
                const res = await fetch('/api/export_arsiv', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ tarih_bas: bas || null, tarih_bit: bit || null })
                });
                const d = await res.json();
                if (d.basarili) {
                    sonucEl.style.display = 'block';
                    sonucEl.style.color = '#16a34a';
                    sonucEl.innerHTML = `
                        ✅ <b>Kaydedildi!</b> — ${d.zaman}<br>
                        <span style="font-size:.78rem;color:#64748b">
                            ${d.vardiya_sayisi} vardiya · ${d.uretim_kayit} üretim · ${d.durus_kayit} duruş kaydı aktarıldı
                        </span>
                    `;
                    toast('Excel dosyası kaydedildi!', 'ok');
                } else {
                    sonucEl.style.display = 'block';
                    sonucEl.style.color = '#dc2626';
                    sonucEl.textContent = '❌ Hata: ' + (d.hata || 'Bilinmeyen hata');
                }
            } catch (e) {
                sonucEl.style.display = 'block';
                sonucEl.style.color = '#dc2626';
                sonucEl.textContent = '❌ Bağlantı hatası: ' + e.message;
            }
            btn.textContent = '📥 Excel\'e Kaydet';
            btn.disabled = false;
        }



        // ═════════════════════════════════════════════════════
        // REFERANS TAKİP PANEL
        // ═════════════════════════════════════════════════════
        const RT_DURUMLAR = [
            { key: 'launch_alinacak', label: 'Launch Alınacak', bg: '#1e293b', color: '#94a3b8' },
            { key: 'launch_alindi',   label: 'Launch Alındı',   bg: '#172554', color: '#93c5fd' },
            { key: 'launch_hazir',    label: 'Launch Hazır',    bg: '#14532d', color: '#86efac' },
            { key: 'uretimde',        label: 'Üretimde',        bg: '#451a03', color: '#fb923c' },
            { key: 'uretim_tamamlandi', label: 'Tamamlandı',   bg: '#365314', color: '#a3e635' }
        ];

        async function yukleRefTakipPanel() {
            try {
                const res = await fetch('/api/referans_takip');
                const list = await res.json();
                const tbody = document.getElementById('ref-takip-tbody');
                const count = document.getElementById('ref-takip-count');
                if (!list.length) {
                    tbody.innerHTML = '<tr><td colspan="7"><div class="empty"><div class="empty-icon">🚀</div><p>Kayıt yok</p></div></td></tr>';
                    count.textContent = '';
                    return;
                }
                count.textContent = list.length + ' kayıt';
                tbody.innerHTML = list.map(r => {
                    const aktifD = RT_DURUMLAR.find(d => d.key === r.durum) || RT_DURUMLAR[0];
                    const durumBtns = RT_DURUMLAR.map(d => `
                        <button onclick="refDurumGuncelle(${r.id},'${d.key}')"
                            style="padding:3px 8px;border-radius:6px;font-size:.68rem;font-weight:700;cursor:pointer;
                                   border:1px solid ${d.key === r.durum ? d.color : 'transparent'};
                                   background:${d.key === r.durum ? d.bg : 'transparent'};
                                   color:${d.key === r.durum ? d.color : 'var(--muted)'};
                                   transition:.2s">
                            ${d.label}
                        </button>`).join('');
                    const robotLbl = r.robot_no ? `<b>${r.robot_no}</b>${r.istasyon > 0 ? ' · İst.' + r.istasyon : ''}` : '—';
                    return `<tr>
                        <td style="font-weight:700">${r.referans_kodu}</td>
                        <td style="text-align:center">${r.hedef_adet > 0 ? r.hedef_adet + ' adet' : '—'}</td>
                        <td style="font-size:.8rem">${robotLbl}</td>
                        <td style="font-size:.8rem;color:var(--muted)">${r.aciklama || '—'}</td>
                        <td><div style="display:flex;gap:3px;flex-wrap:wrap">${durumBtns}</div></td>
                        <td style="font-size:.72rem;color:var(--muted)">${(r.guncelleme_tarihi||'').slice(0,16)}</td>
                        <td style="display:flex;gap:4px">
                            <button onclick="refTakipDuzenleAc(${r.id},'${r.referans_kodu.replace(/'/g, `\\'`).replace(/"/g, `&quot;`)}',${r.hedef_adet},'${r.robot_no||''}',${r.istasyon||0})"
                                style="background:#eff6ff;color:#2563eb;border:1px solid #bfdbfe;border-radius:6px;padding:3px 10px;font-size:.72rem;cursor:pointer;font-weight:700">
                                ✏️
                            </button>
                            <button onclick="refTakipSil(${r.id})"
                                style="background:#fef2f2;color:#dc2626;border:1px solid #fca5a5;border-radius:6px;padding:3px 10px;font-size:.72rem;cursor:pointer;font-weight:700">
                                🗑️
                            </button>
                        </td>
                    </tr>`;
                }).join('');
            } catch(e) {}
        }

        async function refTakipEkle() {
            const ref = document.getElementById('rt-ref').value.trim();
            if (!ref) { toast('Referans kodu zorunludur', 'error'); return; }
            try {
                await fetch('/api/referans_takip', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({
                        referans_kodu: ref,
                        hedef_adet: parseInt(document.getElementById('rt-adet').value) || 0,
                        robot_no: document.getElementById('rt-robot').value,
                        istasyon: parseInt(document.getElementById('rt-istasyon').value) || 0,
                        aciklama: document.getElementById('rt-aciklama').value.trim()
                    })
                });
                document.getElementById('rt-ref').value = '';
                document.getElementById('rt-adet').value = '';
                document.getElementById('rt-robot').value = '';
                document.getElementById('rt-istasyon').value = '0';
                document.getElementById('rt-aciklama').value = '';
                await yukleRefTakipPanel();
                toast('Referans eklendi', 'ok');
            } catch(e) { toast('Hata: ' + e.message, 'error'); }
        }

        // Inline düzenle
        let _rtDuzenleId = null;
        function refTakipDuzenleAc(id, ref, adet, robot, istasyon) {
            _rtDuzenleId = id;
            const yeniRef = prompt('Referans Kodu:', ref);
            if (yeniRef === null) return;
            const yeniAdet = prompt('Hedef Adet:', adet);
            if (yeniAdet === null) return;
            const yeniRobot = prompt('Robot (boş bırakabilirsiniz):', robot);
            if (yeniRobot === null) return;
            const yeniIst = prompt('İstasyon (0=belirtme, 1, 2):', istasyon);
            if (yeniIst === null) return;
            refTakipDuzenleKaydet(id, yeniRef.trim(), parseInt(yeniAdet)||0, yeniRobot.trim(), parseInt(yeniIst)||0);
        }

        async function refTakipDuzenleKaydet(id, ref, adet, robot, istasyon) {
            if (!ref) { toast('Referans kodu boş olamaz', 'error'); return; }
            try {
                await fetch('/api/referans_takip/' + id, {
                    method: 'PATCH',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ referans_kodu: ref, hedef_adet: adet, robot_no: robot, istasyon })
                });
                await yukleRefTakipPanel();
                toast('Kayıt güncellendi', 'ok');
            } catch(e) { toast('Hata', 'error'); }
        }

        async function refDurumGuncelle(id, durum) {
            try {
                await fetch('/api/referans_takip/' + id, {
                    method: 'PATCH',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ durum })
                });
                await yukleRefTakipPanel();
            } catch(e) {}
        }

        async function refTakipSil(id) {
            if (!confirm('Bu kaydı silmek istiyor musunuz?')) return;
            try {
                await fetch('/api/referans_takip/' + id, { method: 'DELETE' });
                await yukleRefTakipPanel();
                toast('Kayıt silindi', 'ok');
            } catch(e) {}
        }

        // ═══════════════════════════════════════════════════════
        // INIT
        // ═══════════════════════════════════════════════════════
        document.addEventListener('DOMContentLoaded', async () => {
            const bugun = new Date().toISOString().split('T')[0];
            document.getElementById('f-bas').value = bugun;
            document.getElementById('f-bit').value = bugun;

            // Bolum durumunu UI'ye yansit (aktifBolum localStorage'dan zaten okundu)
            bolumGorunurlugu();
            // Bolume gore robot listesi
            await yukleRobotListesi();

            yukle();
            yukleRefTakipPanel();
            initDragDrop();

            // Kaydedilmiş iç sekme tercihini yükle
            const kaydedilmisSekme = localStorage.getItem('dashboard_inner_tab') || 'ozet';
            const ilkBtn = document.querySelector(`#inner-tab-bar .inner-tab-btn[onclick*="'${kaydedilmisSekme}'"]`);
            if (ilkBtn) innerSecme(kaydedilmisSekme, ilkBtn);
            tabPinGuncelle(kaydedilmisSekme);
        });

        // ═══════════════════════════════════════════════════════
        // İÇ SEKME
        // ═══════════════════════════════════════════════════════
        let aktifInnerTab = 'ozet';

        function innerSecme(sekme, btn) {
            aktifInnerTab = sekme;
            // Pane geçişi
            document.querySelectorAll('.inner-pane').forEach(p => p.classList.remove('active'));
            const pane = document.getElementById('inner-pane-' + sekme);
            if (pane) pane.classList.add('active');
            // Buton aktif
            document.querySelectorAll('#inner-tab-bar .inner-tab-btn').forEach(b => b.classList.remove('active'));
            if (btn) btn.classList.add('active');
            tabPinGuncelle(sekme);
        }

        function tabPinGuncelle(sekme) {
            const kayitli = localStorage.getItem('dashboard_inner_tab') || 'ozet';
            const btn = document.getElementById('tab-pin-btn');
            if (!btn) return;
            if (kayitli === sekme) {
                btn.textContent = '📍 Başlangıç';
                btn.classList.add('pinned');
            } else {
                btn.textContent = '📍 Başlangıç yap';
                btn.classList.remove('pinned');
            }
        }

        function tabTercihiKaydet() {
            localStorage.setItem('dashboard_inner_tab', aktifInnerTab);
            tabPinGuncelle(aktifInnerTab);
            toast('Başlangıç sekmesi kaydedildi!', 'ok');
        }

        // ═══════════════════════════════════════════════════════
        // SEKME DEĞİŞTİR (Yan sidebar)
        // ═══════════════════════════════════════════════════════
        function sekmeDegistir(sekme) {
            const sekmeler = ['vardiyalar', 'referanslar', 'andon-ayarlar'];
            sekmeler.forEach(s => {
                const el = document.getElementById('sekme-' + s);
                if (el) el.style.display = (s === sekme) ? '' : 'none';
            });
            // sidebar nav aktif sınıfı
            document.querySelectorAll('.nav-item').forEach((a, i) => {
                a.classList.toggle('active', i === sekmeler.indexOf(sekme));
            });
            // Başlık güncelle
            const basliklar = {
                'vardiyalar': 'Dashboard',
                'referanslar': 'Referans Listesi',
                'andon-ayarlar': 'Andon Ekran Ayarları'
            };
            document.getElementById('page-title').textContent = basliklar[sekme] || sekme;
            document.getElementById('subtitle-text').textContent = '';
            if (sekme === 'andon-ayarlar') yukleAndonRobotlari();
        }

        // ═══════════════════════════════════════════════════════
        // ANDON AYARLARI
        // ═══════════════════════════════════════════════════════
        let andonRobotVerisi = [];

        async function yukleAndonRobotlari() {
            try {
                const res = await fetch('/api/andon_robot_ayarlari');
                andonRobotVerisi = await res.json();
                renderAndonRobotListe();
            } catch(e) {}
        }

        function renderAndonRobotListe() {
            const container = document.getElementById('andon-robot-liste');
            if (!container) return;
            container.innerHTML = andonRobotVerisi.map((r, i) => `
                <div style="background:var(--bg,#f8fafc);border:2px solid ${r.goster ? '#2563eb' : '#e2e8f0'};
                    border-radius:12px;padding:14px 16px;display:flex;align-items:center;gap:12px;
                    transition:border-color .2s;cursor:pointer" onclick="andonRobotToggle(${i})">
                    <div style="width:20px;height:20px;border-radius:6px;border:2px solid ${r.goster ? '#2563eb' : '#cbd5e1'};
                        background:${r.goster ? '#2563eb' : 'transparent'};display:flex;align-items:center;justify-content:center;
                        flex-shrink:0;transition:.2s">
                        ${r.goster ? '<span style="color:#fff;font-size:.7rem;font-weight:900">✓</span>' : ''}
                    </div>
                    <div style="flex:1">
                        <div style="font-weight:700;font-size:.88rem">${r.robot_no}</div>
                        <div style="font-size:.72rem;color:var(--muted,#64748b)">${r.goster ? '✅ Andonda görünür' : '⛔ Gizlendi'}</div>
                    </div>
                    <div style="font-size:.75rem;color:var(--muted,#64748b)">Sıra: ${r.sira + 1}</div>
                </div>
            `).join('');
        }

        function andonRobotToggle(i) {
            andonRobotVerisi[i].goster = andonRobotVerisi[i].goster ? 0 : 1;
            renderAndonRobotListe();
        }

        async function andonAyarlariKaydet() {
            try {
                await fetch('/api/andon_robot_ayarlari', {
                    method: 'PATCH',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(andonRobotVerisi)
                });
                toast('Andon ayarları kaydedildi!', 'ok');
            } catch(e) {
                toast('Hata: ' + e.message, 'error');
            }
        }


        // ═══════════════════════════════════════════════════════
        // DRAG & DROP PANELLER
        // ═══════════════════════════════════════════════════════
        function initDragDrop() {
            const container = document.getElementById('sekme-vardiyalar');
            const blocks = Array.from(container.querySelectorAll('.panel-block'));
            if (!blocks.length) return;

            // Kayitli siralamayi uygula
            const savedOrder = JSON.parse(localStorage.getItem('dashboard_panel_order') || 'null');
            if (savedOrder && savedOrder.length) {
                savedOrder.forEach(id => {
                    const el = container.querySelector(`.panel-block[data-block-id="${id}"]`);
                    if (el) container.appendChild(el);
                });
            }

            let draggingEl = null;

            blocks.forEach(block => {
                const handle = block.querySelector('.drag-handle');
                if (handle) {
                    handle.addEventListener('mousedown', () => block.setAttribute('draggable', 'true'));
                    handle.addEventListener('mouseup', () => block.setAttribute('draggable', 'false'));
                    handle.addEventListener('mouseleave', () => block.setAttribute('draggable', 'false'));
                }

                block.addEventListener('dragstart', (e) => {
                    draggingEl = block;
                    setTimeout(() => block.classList.add('dragging'), 0);
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', block.dataset.blockId);
                });

                block.addEventListener('dragend', () => {
                    block.classList.remove('dragging');
                    block.setAttribute('draggable', 'false');
                    draggingEl = null;
                    container.querySelectorAll('.panel-block').forEach(b => b.classList.remove('drag-over'));
                });
            });

            container.addEventListener('dragover', (e) => {
                e.preventDefault(); // Drop tetiklenmesi icin sart
                e.dataTransfer.dropEffect = 'move';

                const targetBlock = e.target.closest('.panel-block');
                if (targetBlock && targetBlock !== draggingEl) {
                    container.querySelectorAll('.panel-block').forEach(b => b.classList.remove('drag-over'));
                    targetBlock.classList.add('drag-over');
                }
            });

            container.addEventListener('drop', (e) => {
                e.preventDefault();
                container.querySelectorAll('.panel-block').forEach(b => b.classList.remove('drag-over'));

                if (!draggingEl) return;
                const targetBlock = e.target.closest('.panel-block');
                if (!targetBlock || targetBlock === draggingEl) return;

                // İmlecin konumuna gore uste veya alta yerlestir
                const rect = targetBlock.getBoundingClientRect();
                const midY = rect.top + rect.height / 2;
                if (e.clientY < midY) {
                    container.insertBefore(draggingEl, targetBlock);
                } else {
                    container.insertBefore(draggingEl, targetBlock.nextSibling);
                }

                // Siralamayi kaydet
                const newOrder = Array.from(container.querySelectorAll('.panel-block')).map(b => b.dataset.blockId);
                localStorage.setItem('dashboard_panel_order', JSON.stringify(newOrder));
            });
        }


        // ═══════════════════════════════════════════════════════
        // SEKME DEĞİŞTİR
        // ═══════════════════════════════════════════════════════
        function sekmeDegistir(s) {
            aktifSekme = s;
            document.getElementById('sekme-vardiyalar').style.display = s === 'vardiyalar' ? '' : 'none';
            document.getElementById('sekme-referanslar').style.display = s === 'referanslar' ? '' : 'none';
            document.querySelectorAll('.nav-item').forEach((n, i) => {
                n.classList.toggle('active', (i === 0 && s === 'vardiyalar') || (i === 1 && s === 'referanslar'));
            });
            document.getElementById('page-title').textContent = s === 'vardiyalar' ? 'Dashboard' : 'Referanslar';
            if (s === 'referanslar') yukleReferanslar();
        }

        // ═══════════════════════════════════════════════════════
        // ANA YÜKLEME
        // ═══════════════════════════════════════════════════════
        async function yukle() {
            const bas = document.getElementById('f-bas').value;
            const bit = document.getElementById('f-bit').value;
            const robot = document.getElementById('f-robot').value;
            const vtp = document.getElementById('f-vardiya').value;

            document.getElementById('subtitle-text').textContent = `${bas} – ${bit}`;
            document.getElementById('tbl-body').innerHTML = '<tr><td colspan="9"><div class="loading"><span class="spinner"></span>Yükleniyor...</div></td></tr>';

            try {
                // Özet verisi
                let url = `/api/ozet?tarih_bas=${bas}&tarih_bit=${bit}&bolum=${encodeURIComponent(aktifBolum)}`;
                if (robot) url += `&robot=${robot}`;
                const res = await fetch(url);
                ozetData = await res.json();

                // KPI güncelle
                guncelleKPI(ozetData);
                // Grafikler
                guncelleGrafik_durus(ozetData.durus_dagilim || []);
                guncelleGrafik_robot(ozetData.robot_uretim || []);
                // Referans tablosu
                guncelleRefUretim(ozetData.referans_uretim || []);
                // Duruş kırılımı
                guncelleDurusKirilim(ozetData.durus_dagilim || [], ozetData.durus_tipi_ozet || []);
                // Robot bazlı kırılımlar
                guncelleRobotRefUretim(ozetData.robot_referans_uretim || []);
                guncelleRobotDurus(ozetData.robot_durus_kirilim || []);

                // Tablo (vardiya filtresi client-side)
                let varList = ozetData.vardiyalar || [];
                if (vtp) varList = varList.filter(v => v.vardiya === vtp);
                guncelleTablo(varList);

            } catch (e) {
                console.error(e);
                toast('Veri yüklenemedi: ' + e.message, 'err');
            }

            // Fikstür ve Robot panellerini yükle
            yukleFiksturPanel();
            yukleRobotPrgPanel();
        }

        // ═════════════════════════════════════════════════════
        // FİKSTÜR ADRESLERİ PANEL
        // ═════════════════════════════════════════════════════
        async function yukleFiksturPanel() {
            try {
                const res = await fetch('/api/fikstur');
                const list = await res.json();
                const tbody = document.getElementById('fikstur-table-body');
                const count = document.getElementById('fikstur-count');
                count.textContent = list.length + ' kayıt';
                if (!list.length) {
                    tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;color:var(--muted);padding:16px">Henüz fikstür raf adresi girilmedi.</td></tr>';
                    return;
                }
                // Raf bazında grupla ve sırala
                const raflar = {};
                list.forEach(f => {
                    const raf = f.raf_adresi || 'Belirtilmedi';
                    if (!raflar[raf]) raflar[raf] = [];
                    raflar[raf].push(f);
                });
                tbody.innerHTML = '';
                Object.keys(raflar).sort().forEach(raf => {
                    const rafRow = document.createElement('tr');
                    rafRow.innerHTML = `<td colspan="3" style="background:#f1f5f9;font-weight:800;font-size:.78rem;color:var(--primary);padding:6px 12px">🗄️ Raf: ${raf}</td>`;
                    tbody.appendChild(rafRow);
                    raflar[raf].forEach(f => {
                        const tr = document.createElement('tr');
                        tr.innerHTML = `<td style="padding-left:24px"><b>${f.referans_kodu}</b></td><td>${f.raf_adresi || '—'}</td><td style="font-size:.8rem;color:var(--muted)">${f.notlar || ''}</td>`;
                        tbody.appendChild(tr);
                    });
                });
            } catch(e) {}
        }

        // ═════════════════════════════════════════════════════
        // ROBOT PROGRAMLARI PANEL (Nested header tablo)
        // ═════════════════════════════════════════════════════
        async function yukleRobotPrgPanel() {
            const ROBOTLAR = ['ABB1','ABB2','ABB3','ABB4','ABB5','ABB6','ABB7','ABB8','ABB9'];
            try {
                const res = await fetch('/api/robot_programlari');
                const list = await res.json();
                const thead = document.getElementById('robot-prg-thead');
                const tbody = document.getElementById('robot-prg-tbody');
                const count = document.getElementById('robot-prg-count');
                
                // Üst satır: boş köşe + her robot colspan=2
                let tr1 = '<tr><th rowspan="2" style="vertical-align:middle">—</th>';
                ROBOTLAR.forEach(r => { tr1 += `<th colspan="2" style="text-align:center">${r}</th>`; });
                tr1 += '</tr>';
                // Alt satır: her robot için İst. 1 ve İst. 2
                let tr2 = '<tr>';
                ROBOTLAR.forEach(() => { tr2 += '<th style="font-size:.7rem;font-weight:600">İst.1</th><th style="font-size:.7rem;font-weight:600">İst.2</th>'; });
                tr2 += '</tr>';
                thead.innerHTML = tr1 + tr2;

                tbody.innerHTML = '';
                let toplam = 0;

                // Tek satır: her robot için İst.1 ve İst.2 bitişik hücre
                const tr = document.createElement('tr');
                const corner = document.createElement('td');
                corner.style.cssText = 'font-weight:800;font-size:.78rem;background:#f1f5f9';
                corner.textContent = 'Kayıtlı Kodlar';
                tr.appendChild(corner);

                ROBOTLAR.forEach(rName => {
                    [1, 2].forEach(ist => {
                        const kayitlar = list.filter(x => x.robot_no === rName && x.istasyon === ist);
                        toplam += kayitlar.length;
                        const td = document.createElement('td');
                        td.style.verticalAlign = 'top';
                        if (kayitlar.length === 0) {
                            td.innerHTML = '<span style="color:var(--muted);font-size:.75rem">—</span>';
                        } else {
                            td.innerHTML = kayitlar.map(k =>
                                `<div style="background:#eff6ff;border-radius:4px;padding:2px 6px;margin-bottom:2px;font-size:.75rem;font-weight:600">${k.referans_kodu}</div>`
                            ).join('');
                        }
                        tr.appendChild(td);
                    });
                });
                tbody.appendChild(tr);
                count.textContent = toplam + ' kayıt';
            } catch(e) {}
        }


        // ═══════════════════════════════════════════════════════
        // KPI
        // ═══════════════════════════════════════════════════════
        function guncelleKPI(data) {
            const varList = data.vardiyalar || [];
            if (!varList.length) {
                ['oee', 'avail', 'perf', 'qual'].forEach(k => document.getElementById('kpi-' + k).textContent = '—');
                document.getElementById('kpi-oee-sub').textContent = 'Kayıt yok';
                return;
            }
            const ort = (key) => (varList.reduce((s, v) => s + (v[key] || 0), 0) / varList.length).toFixed(1);
            document.getElementById('kpi-oee').textContent = `%${ort('oee')}`;
            document.getElementById('kpi-avail').textContent = `%${ort('availability')}`;
            document.getElementById('kpi-perf').textContent = `%${ort('performance')}`;
            document.getElementById('kpi-qual').textContent = `%${ort('quality')}`;
            document.getElementById('kpi-oee-sub').textContent = `${varList.length} vardiya`;
        }

        // ═══════════════════════════════════════════════════════
        // GRAFİKLER
        // ═══════════════════════════════════════════════════════
        function guncelleGrafik_durus(data) {
            const ctx = document.getElementById('chart-durus').getContext('2d');
            if (chartDurus) chartDurus.destroy();
            if (!data.length) { ctx.clearRect(0, 0, 400, 220); return; }
            chartDurus = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: data.map(d => d.durus_sebebi),
                    datasets: [{
                        data: data.map(d => d.toplam_sure),
                        backgroundColor: ['#2563eb', '#16a34a', '#d97706', '#dc2626', '#7c3aed', '#0891b2', '#64748b'],
                        borderWidth: 2, borderColor: '#fff',
                    }]
                },
                options: {
                    responsive: true, maintainAspectRatio: false,
                    plugins: {
                        legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12 } },
                        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.raw} dk` } }
                    }
                }
            });
        }

        function guncelleGrafik_robot(data) {
            const ctx = document.getElementById('chart-robot').getContext('2d');
            if (chartRobot) chartRobot.destroy();
            if (!data.length) { return; }
            chartRobot = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: data.map(d => d.robot_no),
                    datasets: [
                        { label: 'OK', data: data.map(d => d.toplam_ok || 0), backgroundColor: '#16a34a', borderRadius: 6 },
                        { label: 'NOK', data: data.map(d => d.toplam_nok || 0), backgroundColor: '#dc2626', borderRadius: 6 },
                        { label: 'Hedef', data: data.map(d => d.toplam_hedef || 0), backgroundColor: 'rgba(37,99,235,.2)', borderRadius: 6, borderWidth: 2, borderColor: '#2563eb' },
                    ]
                },
                options: {
                    responsive: true, maintainAspectRatio: false,
                    plugins: { legend: { labels: { font: { size: 11 }, boxWidth: 12 } } },
                    scales: {
                        x: { grid: { display: false } },
                        y: { beginAtZero: true, grid: { color: '#f1f5f9' } }
                    }
                }
            });
        }

        // ═══════════════════════════════════════════════════════
        // TABLO
        // ═══════════════════════════════════════════════════════
        function guncelleTablo(varList) {
            document.getElementById('toplam-kayit').textContent = `${varList.length} kayıt`;
            const tbody = document.getElementById('tbl-body');
            if (!varList.length) {
                tbody.innerHTML = `<tr><td colspan="9"><div class="empty"><div class="empty-icon">📭</div><p>Bu tarih aralığında kayıt yok</p></div></td></tr>`;
                return;
            }
            tbody.innerHTML = varList.map(v => {
                const oeeColor = v.oee >= 80 ? '#16a34a' : v.oee >= 60 ? '#d97706' : '#dc2626';
                const vBadge = v.vardiya === 'Sabah' ? 'badge-blue' : v.vardiya === 'Öğle' ? 'badge-orange' : 'badge-purple';
                return `<tr style="cursor:pointer" onclick="panelAc(${getVardiyaId(v)})">
      <td><b>${v.tarih}</b></td>
      <td><span class="badge ${vBadge}">${v.vardiya}</span></td>
      <td>${v.operator}</td>
      <td><span class="badge badge-blue">${v.robot}</span></td>
      <td style="color:var(--success);font-weight:700">${v.ok}</td>
      <td style="color:${v.nok > 0 ? 'var(--danger)' : 'var(--muted)'}">${v.nok}</td>
      <td>${v.durus_dk} dk</td>
      <td>
        <div style="font-weight:800;color:${oeeColor}">${v.oee}%</div>
        <div class="oee-bar-bg"><div class="oee-bar" style="width:${v.oee}%;background:${oeeColor}"></div></div>
      </td>
      <td><button class="btn btn-danger" onclick="vardiyaSil(event,${getVardiyaId(v)})">🗑</button></td>
    </tr>`;
            }).join('');
        }

        // Ozetdaki vardiyaların id'sine şu an erişemiyoruz, bu yüzden api'dan çekeceğiz
        // vardiya_id alanını ozet'e ekledik
        function getVardiyaId(v) {
            return v.vardiya_id || 0;
        }

        // ═══════════════════════════════════════════════════════
        // REFERANS BAZLI ÜRETİM TABLOSU
        // ═══════════════════════════════════════════════════════
        function guncelleRefUretim(data) {
            const tbody = document.getElementById('ref-uretim-body');
            const count = document.getElementById('ref-uretim-count');
            if (!data.length) {
                tbody.innerHTML = '<tr><td colspan="6"><div class="empty"><div class="empty-icon">📭</div><p>Kayıt yok</p></div></td></tr>';
                count.textContent = '';
                return;
            }
            count.textContent = data.length + ' referans';
            tbody.innerHTML = data.map(r => {
                const kalite = r.toplam_uretim > 0 ? ((r.toplam_ok / r.toplam_uretim) * 100).toFixed(1) : '—';
                const kaliteColor = r.toplam_uretim > 0 ? (r.toplam_ok / r.toplam_uretim >= 0.95 ? '#16a34a' : r.toplam_ok / r.toplam_uretim >= 0.80 ? '#d97706' : '#dc2626') : '#64748b';
                const ct = r.ort_cycle_time_sn > 0 ? Math.round(r.ort_cycle_time_sn) + ' sn' : '—';
                const nokStyle = r.toplam_nok > 0 ? 'color:var(--danger);font-weight:700' : 'color:var(--muted)';
                return `<tr>
                    <td><b>${r.referans_kodu}</b></td>
                    <td style="text-align:right;font-weight:700">${r.toplam_uretim}</td>
                    <td style="text-align:right;color:var(--success);font-weight:700">${r.toplam_ok}</td>
                    <td style="text-align:right;${nokStyle}">${r.toplam_nok}</td>
                    <td style="text-align:right;color:${kaliteColor};font-weight:700">${kalite}%</td>
                    <td style="text-align:right;color:var(--muted)">${ct}</td>
                </tr>`;
            }).join('');
        }

        // ═══════════════════════════════════════════════════════
        // DURUŞ KIRILIM TABLOSU
        // ═══════════════════════════════════════════════════════
        function guncelleDurusKirilim(dagilim, tipiOzet) {
            // Sebep kırılımı
            const tbody = document.getElementById('durus-kirilim-body');
            const count = document.getElementById('durus-count');
            if (!dagilim.length) {
                tbody.innerHTML = '<tr><td colspan="4"><div class="empty"><div class="empty-icon">✅</div><p>Duruş yok</p></div></td></tr>';
                count.textContent = '';
            } else {
                count.textContent = dagilim.length + ' çeşit duruş';
                tbody.innerHTML = dagilim.map(d => {
                    const planli = d.durus_tipi === 'planli';
                    const badge = planli
                        ? '<span class="badge badge-green" style="font-size:.65rem">Planlı</span>'
                        : '<span class="badge badge-red" style="font-size:.65rem">Plansız</span>';
                    return `<tr>
                        <td><b>${d.durus_sebebi}</b></td>
                        <td>${badge}</td>
                        <td style="text-align:right;font-weight:700">${d.toplam_sure} dk</td>
                        <td style="text-align:right;color:var(--muted)">${d.adet}x</td>
                    </tr>`;
                }).join('');
            }

            // Tip özeti (sağ kart)
            const ozetEl = document.getElementById('durus-tipi-ozet');
            if (!tipiOzet.length) {
                ozetEl.innerHTML = '<div style="color:var(--muted);font-size:.85rem;text-align:center;padding:20px">Duruş yok</div>';
                return;
            }
            let toplam = tipiOzet.reduce((s, t) => s + t.toplam_sure, 0);
            ozetEl.innerHTML = tipiOzet.map(t => {
                const pct = toplam > 0 ? ((t.toplam_sure / toplam) * 100).toFixed(1) : 0;
                const isPlanlı = t.durus_tipi === 'planli';
                const renk = isPlanlı ? '#16a34a' : '#dc2626';
                const ikon = isPlanlı ? '✅' : '⚠️';
                return `<div style="margin-bottom:16px">
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
                        <span style="font-size:.82rem;font-weight:700">${ikon} ${isPlanlı ? 'Planlı Duruşlar' : 'Plansız Duruşlar'}</span>
                        <span style="font-size:.82rem;font-weight:800;color:${renk}">${t.toplam_sure} dk</span>
                    </div>
                    <div style="background:#f1f5f9;border-radius:6px;height:8px;overflow:hidden">
                        <div style="background:${renk};height:100%;width:${pct}%;border-radius:6px;transition:width .4s"></div>
                    </div>
                    <div style="display:flex;justify-content:space-between;margin-top:4px">
                        <span style="font-size:.7rem;color:var(--muted)">${t.adet} kayıt</span>
                        <span style="font-size:.7rem;color:var(--muted)">%${pct}</span>
                    </div>
                </div>`;
            }).join('');
        }

        // ═══════════════════════════════════════════════════════
        // ROBOT BAZLI REFERANS ÜRETİM KIRILIMi
        // ═══════════════════════════════════════════════════════
        function guncelleRobotRefUretim(data) {
            const tbody = document.getElementById('robot-ref-body');
            const count = document.getElementById('robot-ref-count');
            if (!data.length) {
                tbody.innerHTML = '<tr><td colspan="5"><div class="empty"><div class="empty-icon">📭</div><p>Kayıt yok</p></div></td></tr>';
                count.textContent = ''; return;
            }
            count.textContent = data.length + ' satır';
            let prevRobot = null;
            tbody.innerHTML = data.map(r => {
                const isNew = r.robot_no !== prevRobot;
                prevRobot = r.robot_no;
                const toplam = r.toplam_uretim || 0;
                const ok = r.toplam_ok || 0;
                const kalite = toplam > 0 ? ((ok / toplam) * 100).toFixed(1) : '—';
                const kaliteColor = toplam > 0
                    ? (ok / toplam >= 0.95 ? '#16a34a' : ok / toplam >= 0.80 ? '#d97706' : '#dc2626')
                    : '#64748b';
                return `<tr style="${isNew ? 'border-top:2px solid #e2e8f0' : ''}">
                    <td style="font-weight:${isNew ? '800' : '400'};color:${isNew ? 'var(--primary)' : 'transparent'}">
                        ${isNew ? '🤖 ' + r.robot_no : '↳'}
                    </td>
                    <td><b>${r.referans_kodu}</b></td>
                    <td style="text-align:right;font-weight:700">${toplam}</td>
                    <td style="text-align:right;color:var(--success);font-weight:700">${ok}</td>
                    <td style="text-align:right;color:${kaliteColor};font-weight:700">${kalite}%</td>
                </tr>`;
            }).join('');
        }

        // ═══════════════════════════════════════════════════════
        // ROBOT BAZLI DURUŞ KIRILIMi
        // ═══════════════════════════════════════════════════════
        function guncelleRobotDurus(data) {
            const tbody = document.getElementById('robot-durus-body');
            const count = document.getElementById('robot-durus-count');
            if (!data.length) {
                tbody.innerHTML = '<tr><td colspan="5"><div class="empty"><div class="empty-icon">✅</div><p>Duruş yok</p></div></td></tr>';
                count.textContent = ''; return;
            }
            count.textContent = data.length + ' satır';
            let prevRobot = null;
            tbody.innerHTML = data.map(d => {
                const isNew = d.robot_no !== prevRobot;
                prevRobot = d.robot_no;
                const planli = d.durus_tipi === 'planli';
                const badge = planli
                    ? '<span class="badge badge-green" style="font-size:.62rem">Planlı</span>'
                    : '<span class="badge badge-red" style="font-size:.62rem">Plansız</span>';
                return `<tr style="${isNew ? 'border-top:2px solid #e2e8f0' : ''}">
                    <td style="font-weight:${isNew ? '800' : '400'};color:${isNew ? 'var(--primary)' : 'transparent'}">
                        ${isNew ? '🤖 ' + d.robot_no : '↳'}
                    </td>
                    <td>${d.durus_sebebi}</td>
                    <td>${badge}</td>
                    <td style="text-align:right;font-weight:700">${d.toplam_sure} dk</td>
                    <td style="text-align:right;color:var(--muted)">${d.adet}x</td>
                </tr>`;
            }).join('');
        }

        // ═══════════════════════════════════════════════════════
        // DETAIL PANEL
        // ═══════════════════════════════════════════════════════
        async function panelAc(vid) {
            if (!vid) return;
            document.getElementById('detail-panel').classList.add('open');
            document.getElementById('detail-content').innerHTML = '<div class="loading"><span class="spinner"></span>Yükleniyor</div>';
            try {
                const [detayRes, oeeRes] = await Promise.all([
                    fetch(`/api/vardiya/${vid}`),
                    fetch(`/api/oee/${vid}`)
                ]);
                const detay = await detayRes.json();
                const oee = await oeeRes.json();
                const v = detay.vardiya;

                const oeeColor = oee.oee >= 80 ? '#16a34a' : oee.oee >= 60 ? '#d97706' : '#dc2626';

                let html = `
      <div class="oee-meter">
        <div class="meter-item"><div class="meter-val" style="color:${oeeColor}">${oee.oee}%</div><div class="meter-lbl">OEE</div></div>
        <div class="meter-item"><div class="meter-val" style="color:#16a34a">${oee.availability}%</div><div class="meter-lbl">Lık</div></div>
        <div class="meter-item"><div class="meter-val" style="color:#d97706">${oee.performance}%</div><div class="meter-lbl">Perf.</div></div>
        <div class="meter-item"><div class="meter-val" style="color:#0891b2">${oee.quality}%</div><div class="meter-lbl">Kalite</div></div>
      </div>
      <div class="detail-section">
        <h4>Vardiya Bilgileri</h4>
        <div class="detail-row"><span>Operatör</span><b>${v.operator_adi}</b></div>
        <div class="detail-row"><span>Robot</span><span class="badge badge-blue">${v.robot_no}</span></div>
        <div class="detail-row"><span>Tarih</span><b>${v.tarih}</b></div>
        <div class="detail-row"><span>Vardiya</span><b>${v.vardiya_turu}</b></div>
        <div class="detail-row"><span>Saat</span><b>${v.baslangic_saati} – ${v.bitis_saati}</b></div>
        <div class="detail-row"><span>Planlı Süre</span><b>${v.toplam_sure_dk} dk</b></div>
        <div class="detail-row"><span>Toplam Duruş</span><b style="color:var(--danger)">${oee.toplam_durus_dk} dk</b></div>
      </div>`;

                if (detay.uretim.length) {
                    html += `<div class="detail-section"><h4>Üretim Kayıtları</h4>`;
                    detay.uretim.forEach(u => {
                        html += `<div class="detail-row"><span>📦 ${u.referans_kodu}</span><span><b style="color:var(--success)">${u.ok_adet}</b> OK / <span style="color:var(--danger)">${u.nok_adet}</span> NOK (Hdf: ${u.hedef_adet})</span></div>`;
                    });
                    html += `</div>`;
                }

                if (detay.duruslar.length) {
                    html += `<div class="detail-section"><h4>Duruş Kayıtları</h4>`;
                    detay.duruslar.forEach(d => {
                        html += `<div class="detail-row"><span>⏸️ ${d.durus_sebebi}${d.baslangic_saati ? ' (' + d.baslangic_saati + ')' : ''}</span><b>${d.sure_dk} dk</b></div>`;
                        if (d.aciklama) html += `<div style="font-size:.76rem;color:var(--muted);padding:2px 0 6px 22px">${d.aciklama}</div>`;
                    });
                    html += `</div>`;
                }

                if (v.notlar) {
                    html += `<div class="detail-section"><h4>Notlar</h4><p style="font-size:.85rem;color:var(--muted)">${v.notlar}</p></div>`;
                }

                document.getElementById('detail-content').innerHTML = html;
            } catch (e) {
                document.getElementById('detail-content').innerHTML = '<p style="color:var(--danger)">Yüklenemedi</p>';
            }
        }

        function panelKapat() { document.getElementById('detail-panel').classList.remove('open'); }

        // ═══════════════════════════════════════════════════════
        // VARDİYA SİL
        // ═══════════════════════════════════════════════════════
        async function vardiyaSil(e, vid) {
            e.stopPropagation();
            if (!confirm('Bu vardiya kaydını silmek istiyor musunuz?')) return;
            try {
                await fetch(`/api/vardiya/${vid}`, { method: 'DELETE' });
                toast('Vardiya silindi', 'ok');
                yukle();
            } catch (err) {
                toast('Silinemedi', 'err');
            }
        }

        // ═══════════════════════════════════════════════════════
        // REFERANSLAR
        // ═══════════════════════════════════════════════════════
        let tumReferanslar = [];

        async function yukleReferanslar() {
            const tbody = document.getElementById('ref-tbl-body');
            tbody.innerHTML = '<tr><td colspan="4"><div class="loading"><span class="spinner"></span></div></td></tr>';
            try {
                const res = await fetch('/api/referanslar?q=&bolum=' + encodeURIComponent(aktifBolum));
                tumReferanslar = await res.json();
                renderReferanslar(tumReferanslar);
                yukleEksikReferanslar();
            } catch (e) {
                toast('Referanslar yüklenemedi', 'err');
            }
        }

        function renderReferanslar(data) {
            const tbody = document.getElementById('ref-tbl-body');
            if (!data.length) {
                tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:30px">Referans bulunamadı</td></tr>';
                return;
            }
            tbody.innerHTML = data.map(r => `
                <tr style="cursor:pointer" onclick="refDuzenle('${r.referans_kodu}', '${(r.aciklama || '').replace(/'/g, "\\'")}', ${r.hedef_cycle_time_sn || 0})">
                    <td><b>${r.referans_kodu}</b></td>
                    <td style="font-size:.82rem">${r.aciklama || '—'}</td>
                    <td style="text-align:right;font-weight:700">${r.hedef_cycle_time_sn || '0'} sn</td>
                    <td style="display:flex;gap:4px;justify-content:center" onclick="event.stopPropagation()">
                        <button class="btn" style="padding:4px 8px;font-size:.7rem;background:#eff6ff;color:#2563eb;border:1px solid #bfdbfe" 
                                onclick="refDuzenle('${r.referans_kodu}', '${(r.aciklama || '').replace(/'/g, "\\'")}', ${r.hedef_cycle_time_sn || 0})">✏️</button>
                        <button class="btn" style="padding:4px 8px;font-size:.7rem;background:#fef2f2;color:#dc2626;border:1px solid #fca5a5" 
                                onclick="referansSil('${r.referans_kodu}')">🗑️</button>
                    </td>
                </tr>
            `).join('');
        }

        function refAra() {
            const q = document.getElementById('ref-ara-input').value.toLowerCase();
            const filtered = tumReferanslar.filter(r => 
                r.referans_kodu.toLowerCase().includes(q) || 
                (r.aciklama && r.aciklama.toLowerCase().includes(q))
            );
            renderReferanslar(filtered);
        }

        async function yukleEksikReferanslar() {
            try {
                const res = await fetch('/api/referanslar/eksik');
                if (!res.ok) throw new Error('API hatası');
                const list = await res.json();
                const panel = document.getElementById('eksik-referanslar-panel');
                const listEl = document.getElementById('eksik-ref-list');
                
                if (list && list.length > 0) {
                    panel.style.display = 'block';
                    listEl.innerHTML = list.map(ref => `
                        <button onclick="refDuzenle('${ref}', '', 0)" 
                                style="background:#fff;border:1px solid #fca5a5;border-radius:10px;padding:6px 12px;font-size:.78rem;font-weight:700;color:#e11d48;cursor:pointer;box-shadow:0 2px 4px rgba(0,0,0,.05);transition:.2s">
                            ${ref} +
                        </button>
                    `).join('');
                } else {
                    panel.style.display = 'none';
                }
            } catch (e) {
                console.error('Eksik referanslar yüklenemedi:', e);
                document.getElementById('eksik-referanslar-panel').style.display = 'none';
            }
        }

        function refModalAc() { 
            document.getElementById('ref-modal-title').textContent = 'Yeni Referans Ekle';
            document.getElementById('ref-kod').disabled = false;
            document.getElementById('ref-modal').classList.add('open'); 
        }

        function refDuzenle(kod, desc, ct) {
            document.getElementById('ref-modal-title').textContent = 'Referans Düzenle';
            document.getElementById('ref-kod').value = kod;
            document.getElementById('ref-kod').disabled = true; // Kodu değiştirmeye izin verme (ID gibi)
            document.getElementById('ref-aciklama').value = desc;
            document.getElementById('ref-cycle').value = ct;
            document.getElementById('ref-modal').classList.add('open');
        }

        function refModalKapat(e) {
            if (!e || e.target === document.getElementById('ref-modal')) {
                document.getElementById('ref-modal').classList.remove('open');
                ['ref-kod', 'ref-aciklama', 'ref-cycle'].forEach(id => document.getElementById(id).value = '');
            }
        }

        async function refKaydet() {
            const kod = document.getElementById('ref-kod').value.trim();
            if (!kod) { toast('Referans kodu zorunlu', 'err'); return; }
            try {
                const res = await fetch('/api/referanslar', {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        referans_kodu: kod, 
                        aciklama: document.getElementById('ref-aciklama').value, 
                        hedef_cycle_time_sn: document.getElementById('ref-cycle').value || 0 
                    })
                });
                if (res.ok) {
                    toast('Referans kaydedildi', 'ok');
                    refModalKapat();
                    yukleReferanslar();
                } else {
                    const d = await res.json();
                    toast(d.hata || 'Kaydedilemedi', 'err');
                }
            } catch (e) { toast('Bağlantı hatası', 'err'); }
        }

        async function referansSil(kod) {
            if (!confirm(`'${kod}' referansını listeden silmek istiyor musunuz?`)) return;
            try {
                // Not: Henüz DELETE endpoint'i yoksa POST/DELETE mantığı eklenebilir. 
                // Standart olarak app.route('/api/referanslar/<int:id>', methods=['DELETE']) bekliyoruz.
                // kodu kullanarak silelim.
                const res = await fetch(`/api/referanslar/${kod}`, { method: 'DELETE' });
                if (res.ok) {
                    toast('Referans silindi', 'ok');
                    yukleReferanslar();
                } else {
                    toast('Silebilmek için Backend desteği gerekiyor.', 'err');
                }
            } catch (e) { toast('Silebilmek için Backend desteği gerekiyor.', 'err'); }
        }

        async function referansExcelGuncelle() {
            if (!confirm('Veritabanındaki tüm referans ve süreleri Masaüstü/Kaynakhane.xlsx dosyasına yazmak istiyor musunuz?')) return;
            toast('Excel güncelleniyor...', 'info');
            try {
                const res = await fetch('/api/referanslar/export_excel', { method: 'POST' });
                const d = await res.json();
                if (d.basarili) {
                    toast('✅ Excel başarıyla güncellendi!', 'ok');
                } else {
                    toast('❌ Hata: ' + d.hata, 'err');
                }
            } catch (e) { toast('Bağlantı hatası', 'err'); }
        }

        // ═══════════════════════════════════════════════════════
        // EXCEL VERİ AKTARMA
        // ═══════════════════════════════════════════════════════
        async function excelVeriAktar() {
            if (!confirm('Masaüstündeki Kaynakhane.xlsx dosyasından referans ve operatör bilgileri aktarılacak. Onaylıyor musunuz?')) return;

            toast('Veriler aktarılıyor, lütfen bekleyin...', 'info');
            try {
                const res = await fetch('/api/import_excel', { method: 'POST' });
                const data = await res.json();

                if (res.ok && data.basarili) {
                    toast(`✅ Başarılı! Ekl.: ${data.referanslar_eklenen} ref, Günc.: ${data.referanslar_guncellenen} ref, Ekl. Op.: ${data.operatorler_eklenen}`, 'ok');
                    if (aktifSekme === 'referanslar') {
                        yuklreferanslar();
                    }
                } else {
                    throw new Error(data.hata || 'Bilinmeyen hata');
                }
            } catch (err) {
                toast(`❌ Aktarım hatası: ${err.message}`, 'err');
            }
        }

        // ═══════════════════════════════════════════════════════
        // TOAST
        // ═══════════════════════════════════════════════════════
        let tTimer;
        function toast(msg, type = '') {
            const el = document.getElementById('toast');
            el.textContent = msg; el.className = 'toast show ' + type;
            if (tTimer) clearTimeout(tTimer);
            tTimer = setTimeout(() => el.classList.remove('show'), 3000);
        }
    
