const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);

const DB = {
    logs: path.join(DATA_DIR, 'logs.json'),
    users: path.join(DATA_DIR, 'users.json'),
    products: path.join(DATA_DIR, 'products.json'),
    structure: path.join(DATA_DIR, 'structure.json'),
    notes: path.join(DATA_DIR, 'notes.json')
};

app.use(bodyParser.json());

const readDB = (f) => {
    if (!fs.existsSync(f)) {
        if (f === DB.users) return [{ user: 'hakkı', pass: '2125', role: 'admin' }];
        if (f === DB.products) return [{ name: 'Su' }, { name: 'Kola' }];
        if (f === DB.structure) return [];
        return f === DB.notes ? {} : [];
    }
    return JSON.parse(fs.readFileSync(f));
};

const writeDB = (f, d) => fs.writeFileSync(f, JSON.stringify(d, null, 2));

// API Servisleri
app.post('/api/login', (req, res) => {
    const users = readDB(DB.users);
    const u = users.find(u => u.user === req.body.user && u.pass === req.body.pass);
    if (u) res.json({ success: true, role: u.role, user: u.user });
    else res.json({ success: false });
});

app.get('/api/logs', (req, res) => res.json(readDB(DB.logs)));
app.post('/api/logs', (req, res) => {
    const logs = readDB(DB.logs);
    logs.unshift({ ...req.body, date: new Date().toLocaleDateString('tr-TR'), endTime: new Date().toLocaleTimeString('tr-TR') });
    writeDB(DB.logs, logs); res.json({ success: true });
});

app.get('/api/notes', (req, res) => res.json(readDB(DB.notes)));
app.post('/api/notes', (req, res) => { writeDB(DB.notes, req.body); res.json({ success: true }); });
app.get('/api/structure', (req, res) => res.json(readDB(DB.structure)));
app.post('/api/structure', (req, res) => { writeDB(DB.structure, req.body); res.json({ success: true }); });
app.get('/api/products', (req, res) => res.json(readDB(DB.products)));
app.post('/api/products', (req, res) => { writeDB(DB.products, req.body); res.json({ success: true }); });
app.get('/api/users', (req, res) => res.json(readDB(DB.users)));
app.post('/api/users', (req, res) => {
    const u = readDB(DB.users); u.push({ ...req.body, role: 'staff' });
    writeDB(DB.users, u); res.json({ success: true });
});
app.delete('/api/users/:name', (req, res) => {
    let u = readDB(DB.users).filter(x => x.user !== req.params.name);
    writeDB(DB.users, u); res.json({ success: true });
});
app.post('/api/end-day', (req, res) => { writeDB(DB.logs, []); writeDB(DB.notes, {}); res.json({ success: true }); });

app.get('/api/export', (req, res) => {
    const logs = readDB(DB.logs);
    const ws = XLSX.utils.json_to_sheet(logs);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rapor");
    const buf = XLSX.write(wb, {bookType:'xlsx', type:'buffer'});
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=minibar_rapor.xlsx');
    res.send(buf);
});

// HTML Arayüzü
app.get('/', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Smart Minibar v17.7</title>
    <style>
        :root { --p: #2c3e50; --g: #2ecc71; --y: #f1c40f; --r: #e74c3c; --b: #3498db; --gr: #95a5a6; --note: #9b59b6; }
        body { font-family: sans-serif; background: #f4f7f6; margin: 0; padding: 10px; }
        .page { display: none; max-width: 1100px; margin: auto; background: white; padding: 15px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .active { display: block !important; }
        button { padding: 12px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
        input { width: 100%; padding: 12px; margin: 5px 0; border: 1px solid #ddd; border-radius: 8px; box-sizing: border-box; }
        .btn-p { background: var(--p); color: white; width: 100%; margin: 5px 0; }
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); gap: 10px; margin: 15px 0; }
        .btn-room { background: #fff; border: 1px solid #ddd; height: 55px; border-radius: 8px; display:flex; align-items:center; justify-content:center; cursor: pointer; font-weight: bold; position: relative; }
        
        .status-Müsait { background-color: var(--g) !important; color: white !important; }
        .status-Sonra { background-color: var(--y) !important; color: black !important; }
        .status-DND { background-color: var(--r) !important; color: white !important; }
        .has-note { border: 2px solid var(--note) !important; }
        .note-badge { position: absolute; top: 2px; right: 2px; font-size: 10px; }

        .admin-tabs { display: flex; gap: 5px; margin-bottom: 15px; background: #eee; padding: 5px; border-radius: 8px; }
        .a-tab { flex: 1; padding: 10px; font-size: 11px; background: none; color: #666; border: none; cursor: pointer; }
        .a-tab.active { background: white !important; color: var(--p) !important; box-shadow: 0 2px 5px rgba(0,0,0,0.1); border-radius: 6px; }
        .tab-content { display: none; }
        .tab-content.active { display: block !important; }
        
        table { width: 100%; border-collapse: collapse; font-size: 12px; }
        th, td { border: 1px solid #eee; padding: 10px; text-align: left; }
        .setup-box { background: #f8f9fa; padding: 15px; border-radius: 10px; margin-bottom: 20px; }
        #modal { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.75); z-index:9999; align-items:center; justify-content:center; }
        .modal-box { background:white; width:90%; max-width:400px; padding:20px; border-radius:15px; }
    </style>
</head>
<body>
    <div id="modal">
        <div class="modal-box"><h3 id="mTitle"></h3><div id="mBody"></div><button class="btn-p" onclick="document.getElementById('modal').style.display='none'">Kapat</button></div>
    </div>

    <!-- GİRİŞ SAYFASI -->
    <div id="loginPage" class="page active">
        <h2 style="text-align:center">Smart Minibar</h2>
        <input type="text" id="lUser" placeholder="Kullanıcı Adı">
        <input type="password" id="lPass" placeholder="Şifre">
        <button class="btn-p" onclick="login()">GİRİŞ YAP</button>
    </div>

    <!-- PERSONEL SAYFASI -->
    <div id="staffPage" class="page">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px">
            <b id="sn"></b> <button onclick="logout()">Çıkış</button>
        </div>
        <div id="staffContent">
            <div id="blockTabs" class="grid"></div>
            <div id="view_floors" class="grid" style="display:none"></div>
            <div id="view_rooms" class="grid" style="display:none"></div>
        </div>
        <div id="statusScreen" style="display:none">
            <div id="staffNoteBox" style="background:#f3e5f5; padding:10px; border-radius:8px; margin-bottom:10px; display:none; border-left:5px solid var(--note)"></div>
            <h2 id="targetRoomTitle" style="text-align:center"></h2>
            <button class="btn-p" style="background:var(--g); height:80px" onclick="openProductMenu()">MÜSAİT</button>
            <button class="btn-p" style="background:var(--y); color:#000; height:70px" onclick="submitLog('Sonra', '-')">SONRA</button>
            <button class="btn-p" style="background:var(--r); height:70px" onclick="submitLog('DND', '-')">DND</button>
            <button class="btn-p" onclick="goBackToRooms()">⬅ GERİ</button>
        </div>
        <div id="productMenu" style="display:none">
            <div id="pGrid" class="grid"></div>
            <button class="btn-p" style="background:var(--g); height:80px" onclick="processAndSubmit()">KAYDET</button>
        </div>
    </div>

    <!-- YÖNETİCİ SAYFASI -->
    <div id="adminPage" class="page">
        <div class="admin-tabs">
            <button class="a-tab active" onclick="switchAdminTab('t_live', this)">👁️ Canlı</button>
            <button class="a-tab" onclick="switchAdminTab('t_matrix', this)">🏢 Harita</button>
            <button class="a-tab" onclick="switchAdminTab('t_setup', this)">⚙️ Ayar</button>
            <button class="a-tab" onclick="switchAdminTab('t_end', this)">🧹 Gün Sonu</button>
            <button onclick="logout()" style="background:#ddd; border-radius:5px; padding:5px">X</button>
        </div>

        <div id="t_live" class="tab-content active">
            <table><thead><tr><th>Oda</th><th>Durum</th><th>Saat</th></tr></thead><tbody id="liveBody"></tbody></table>
        </div>
        
        <div id="t_matrix" class="tab-content">
            <button id="btnNoteMode" class="btn-p" onclick="toggleNoteMode()" style="background:var(--note)">📝 NOT MODU: KAPALI</button>
            <div id="matrixArea"></div>
        </div>

        <div id="t_setup" class="tab-content">
            <div class="setup-box">
                <b>Personel Ekle</b>
                <input type="text" id="inUN" placeholder="Ad"> <input type="password" id="inUP" placeholder="Şifre">
                <button class="btn-p" onclick="addStaff()">Ekle</button>
                <div id="uList"></div>
            </div>
            <div class="setup-box">
                <b>Oda Yapısı</b>
                <input type="text" id="inB" placeholder="Blok"> <input type="text" id="inF" placeholder="Kat"> <input type="text" id="inR" placeholder="Odalar (101,102)">
                <button class="btn-p" onclick="addStruct()">Kaydet</button>
            </div>
            <div class="setup-box">
                <b>Ürün Ekle</b>
                <input type="text" id="inP" placeholder="Ürün Adı">
                <button class="btn-p" onclick="addProd()">Ekle</button>
                <div id="pList"></div>
            </div>
        </div>

        <div id="t_end" class="tab-content" style="text-align:center">
            <button class="btn-p" onclick="window.open('/api/export')" style="background:var(--g)">📥 EXCEL RAPORU</button>
            <button class="btn-p" onclick="endDay()" style="background:var(--r); margin-top:20px">🧹 GÜNÜ SIFIRLA</button>
        </div>
    </div>

    <script>
        let currentUser = null, hotelData = [], products = [], logs = [], roomNotes = {}, counts = {}, selRoom = "", selBlock = "", selFloor = "";
        let isNoteMode = false;

        async function login() {
            const user = document.getElementById('lUser').value, pass = document.getElementById('lPass').value;
            const res = await fetch('/api/login', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user,pass}) });
            const data = await res.json();
            if(data.success) { currentUser = data; launchApp(); } else alert("Hata!");
        }

        function logout() { location.reload(); }

        function launchApp() {
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            if(currentUser.role === 'admin') { document.getElementById('adminPage').classList.add('active'); initAdmin(); }
            else { document.getElementById('staffPage').classList.add('active'); document.getElementById('sn').innerText = currentUser.user; initStaff(); }
        }

        function switchAdminTab(tabId, btn) {
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            document.querySelectorAll('.a-tab').forEach(b => b.classList.remove('active'));
            document.getElementById(tabId).classList.add('active');
            btn.classList.add('active');
            if(tabId === 't_matrix') refreshMatrix();
            if(tabId === 't_live') refreshLiveLogs();
        }

        // --- Ortak Veri Çekme ---
        async function loadAllData() {
            const [s, p, l, n, u] = await Promise.all([
                fetch('/api/structure').then(r=>r.json()),
                fetch('/api/products').then(r=>r.json()),
                fetch('/api/logs').then(r=>r.json()),
                fetch('/api/notes').then(r=>r.json()),
                fetch('/api/users').then(r=>r.json())
            ]);
            hotelData = s; products = p; logs = l; roomNotes = n;
            return {u};
        }

        // --- Admin İşlemleri ---
        async function initAdmin() {
            const data = await loadAllData();
            document.getElementById('uList').innerHTML = data.u.filter(x=>x.role!=='admin').map(x=>\`<div>\${x.user} <button onclick="delUser('\${x.user}')" style="color:red;background:none">Sil</button></div>\`).join('');
            document.getElementById('pList').innerHTML = products.map(x=>\`<div>\${x.name}</div>\`).join('');
            refreshLiveLogs();
        }

        function refreshLiveLogs() { document.getElementById('liveBody').innerHTML = logs.map(l => \`<tr><td>\${l.room}</td><td>\${l.status}</td><td>\${l.endTime}</td></tr>\`).join(''); }

        function refreshMatrix() {
            let h = ""; hotelData.forEach(b => {
                h += \`<h4>Blok \${b.name}</h4>\`; b.floors.forEach(f => {
                    h += '<div class="grid">'; f.rooms.forEach(r => {
                        const log = logs.find(l => String(l.room) === String(r));
                        const note = roomNotes[r];
                        h += \`<button class="btn-room \${log?'status-'+log.status:''} \${note?'has-note':''}" onclick="handleRoomClick('\${r}')">\${note?'<span class="note-badge">📝</span>':''}\${r}</button>\`;
                    }); h += '</div>';
                });
            }); document.getElementById('matrixArea').innerHTML = h;
        }

        function handleRoomClick(r) {
            if(isNoteMode) {
                const n = prompt("Not:", roomNotes[r]||"");
                if(n !== null) { if(n==="") delete roomNotes[r]; else roomNotes[r]=n; saveNotes(); }
            } else {
                const l = logs.filter(x => String(x.room) === String(r));
                document.getElementById('mTitle').innerText = "Oda " + r;
                document.getElementById('mBody').innerHTML = l.map(x=>\`<div>\${x.endTime} - \${x.status}</div>\`).join('') || "Kayıt yok.";
                document.getElementById('modal').style.display='flex';
            }
        }

        async function saveNotes() { await fetch('/api/notes', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(roomNotes) }); refreshMatrix(); }
        function toggleNoteMode() { isNoteMode = !isNoteMode; document.getElementById('btnNoteMode').innerText = isNoteMode ? "NOT MODU: AÇIK" : "NOT MODU: KAPALI"; }

        // --- Personel İşlemleri ---
        async function initStaff() { await loadAllData(); renderBlocks(); }
        function renderBlocks() {
            document.getElementById('staffContent').style.display='block'; document.getElementById('statusScreen').style.display='none'; document.getElementById('view_floors').style.display='none'; document.getElementById('view_rooms').style.display='none'; document.getElementById('blockTabs').style.display='grid';
            document.getElementById('blockTabs').innerHTML = hotelData.map(b => \`<button class="btn-room" onclick="selectBlock('\${b.name}')">\${b.name}</button>\`).join('');
        }
        function selectBlock(n) {
            selBlock = n; const b = hotelData.find(x => x.name === n);
            document.getElementById('blockTabs').style.display='none'; document.getElementById('view_floors').style.display='grid';
            document.getElementById('view_floors').innerHTML = b.floors.map(f => \`<button class="btn-room" onclick="selectFloor('\${f.name}')">Kat \${f.name}</button>\`).join('') + '<button class="btn-p" onclick="renderBlocks()">Geri</button>';
        }
        function selectFloor(n) {
            selFloor = n; const f = hotelData.find(x => x.name === selBlock).floors.find(x => x.name === n);
            document.getElementById('view_floors').style.display='none'; document.getElementById('view_rooms').style.display='grid'; document.getElementById('staffContent').style.display='block';
            document.getElementById('view_rooms').innerHTML = f.rooms.map(r => {
                const log = logs.find(l => String(l.room) === String(r));
                return \`<button class="btn-room \${log?'status-'+log.status:''} \${roomNotes[r]?'has-note':''}" onclick="openStatusMenu('\${r}')">\${r}</button>\`;
            }).join('') + \`<button class="btn-p" onclick="selectBlock('\${selBlock}')">Geri</button>\`;
        }
        function openStatusMenu(r) {
            selRoom = r; document.getElementById('staffContent').style.display='none'; document.getElementById('statusScreen').style.display='block'; document.getElementById('targetRoomTitle').innerText=r;
            const nb = document.getElementById('staffNoteBox');
            if(roomNotes[r]) { nb.innerText = roomNotes[r]; nb.style.display='block'; } else nb.style.display='none';
        }
        function goBackToRooms() { selectFloor(selFloor); }

        async function submitLog(status, details) {
            await fetch('/api/logs', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({room:selRoom, status, details, staff:currentUser.user}) });
            initStaff();
        }

        function openProductMenu() {
            document.getElementById('statusScreen').style.display='none'; document.getElementById('productMenu').style.display='block';
            counts = {}; products.forEach(p => counts[p.name] = 0);
            document.getElementById('pGrid').innerHTML = products.map((p,i) => \`<div style="border:1px solid #ddd; padding:10px;" onclick="counts['\${p.name}']++; this.querySelector('b').innerText=counts['\${p.name}']">\${p.name}<br><b>0</b></div>\`).join('');
        }
        function processAndSubmit() { let items = Object.entries(counts).filter(e => e[1] > 0).map(e => e[0] + " x" + e[1]); submitLog('Müsait', items.join(", ")); }

        // --- Ayarlar ---
        async function addStaff() { await fetch('/api/users', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user:document.getElementById('inUN').value, pass:document.getElementById('inUP').value}) }); initAdmin(); }
        async function delUser(n) { await fetch('/api/users/'+n, { method:'DELETE' }); initAdmin(); }
        async function addProd() { await fetch('/api/products', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify([...products, {name:document.getElementById('inP').value}]) }); initAdmin(); }
        async function addStruct() {
            let bN = document.getElementById('inB').value, fN = document.getElementById('inF').value, rN = document.getElementById('inR').value;
            let block = hotelData.find(x => x.name === bN) || {name:bN, floors:[]}; if(!hotelData.includes(block)) hotelData.push(block);
            block.floors.push({name:fN, rooms:rN.split(',').map(x=>x.trim())});
            await fetch('/api/structure', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(hotelData) }); initAdmin();
        }
        async function endDay() { if(confirm("Sıfırla?")) { await fetch('/api/end-day', {method:'POST'}); initAdmin(); } }
    </script>
</body>
</html>
    `);
});

app.listen(PORT, '0.0.0.0', () => console.log(`Smart Minibar v17.7 Aktif: ${PORT}`));
