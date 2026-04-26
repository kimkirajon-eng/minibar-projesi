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
    structure: path.join(DATA_DIR, 'structure.json')
};

app.use(bodyParser.json());

const readDB = (f) => {
    if (!fs.existsSync(f)) {
        if (f === DB.users) return [{ user: 'hakkı', pass: '2125', role: 'admin', sid: '' }];
        if (f === DB.products) return [{ name: 'Su' }, { name: 'Kola' }];
        return [];
    }
    return JSON.parse(fs.readFileSync(f));
};

const writeDB = (f, d) => fs.writeFileSync(f, JSON.stringify(d, null, 2));

app.get('/manifest.json', (req, res) => {
    res.json({
        "name": "Smart Minibar v16.6",
        "short_name": "Minibar",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#2c3e50",
        "theme_color": "#2c3e50"
    });
});

app.get('/sw.js', (req, res) => {
    res.setHeader('Content-Type', 'application/javascript');
    res.send(`
        const CACHE_NAME = 'minibar-v1';
        self.addEventListener('install', e => e.waitUntil(caches.open(CACHE_NAME)));
        self.addEventListener('fetch', e => e.respondWith(fetch(e.request).catch(() => caches.match(e.request))));
    `);
});

app.post('/api/login', (req, res) => {
    const users = readDB(DB.users);
    const idx = users.findIndex(u => u.user === req.body.user && u.pass === req.body.pass);
    if (idx !== -1) {
        const sid = Math.random().toString(36).substring(7);
        users[idx].sid = sid;
        writeDB(DB.users, users);
        res.json({ success: true, role: users[idx].role, user: users[idx].user, sid: sid });
    } else res.json({ success: false });
});

app.get('/api/logs', (req, res) => res.json(readDB(DB.logs)));

app.post('/api/logs', (req, res) => {
    const logs = readDB(DB.logs);
    if(logs.some(l => l.id === req.body.id)) return res.json({ success: true });
    logs.unshift({ ...req.body, date: new Date().toLocaleDateString('tr-TR'), endTime: new Date().toLocaleTimeString('tr-TR') });
    writeDB(DB.logs, logs); res.json({ success: true });
});

app.get('/api/export', (req, res) => {
    try {
        const logs = readDB(DB.logs);
        const ws = XLSX.utils.json_to_sheet(logs);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Rapor");
        const buf = XLSX.write(wb, {bookType:'xlsx', type:'buffer'});
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=rapor.xlsx');
        res.end(buf);
    } catch (e) { res.status(500).send("Hata"); }
});

app.get('/api/structure', (req, res) => res.json(readDB(DB.structure)));
app.post('/api/structure', (req, res) => { writeDB(DB.structure, req.body); res.json({ success: true }); });
app.get('/api/products', (req, res) => res.json(readDB(DB.products)));
app.post('/api/products', (req, res) => { writeDB(DB.products, req.body); res.json({ success: true }); });
app.get('/api/users', (req, res) => res.json(readDB(DB.users)));
app.post('/api/users', (req, res) => {
    const u = readDB(DB.users); u.push({ ...req.body, role: 'staff', sid: '' });
    writeDB(DB.users, u); res.json({ success: true });
});
app.delete('/api/users/:name', (req, res) => {
    let u = readDB(DB.users).filter(x => x.user !== req.params.name);
    writeDB(DB.users, u); res.json({ success: true });
});
app.post('/api/end-day', (req, res) => { writeDB(DB.logs, []); res.json({ success: true }); });

app.get('/', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Smart Minibar v16.6</title>
    <style>
        :root { --p: #2c3e50; --g: #2ecc71; --y: #f1c40f; --r: #e74c3c; --b: #3498db; --gr: #95a5a6; }
        body { font-family: sans-serif; background: #f4f7f6; margin: 0; padding: 10px; }
        .page { display: none; max-width: 1100px; margin: auto; background: white; padding: 15px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .active { display: block; }
        button { padding: 12px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
        input { width: 100%; padding: 10px; margin: 5px 0; border: 1px solid #ddd; border-radius: 8px; box-sizing: border-box; }
        .btn-p { background: var(--p); color: white; width: 100%; }
        .admin-tabs, .sub-tabs { display: flex; gap: 4px; background: #eee; padding: 5px; border-radius: 10px; margin-bottom: 15px; overflow-x: auto; }
        .a-tab, .s-tab { flex: 1; padding: 10px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 11px; white-space: nowrap; }
        .a-tab.active, .s-tab.active { background: white; color: var(--p); }
        .tab-content, .sub-content { display: none; }
        .tab-content.active, .sub-content.active { display: block; }
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(75px, 1fr)); gap: 8px; margin: 10px 0; }
        .btn-room { background: white; border: 1px solid #ddd; height: 45px; font-size: 11px; border-radius: 5px; display:flex; align-items:center; justify-content:center; position: relative; overflow: hidden; cursor: pointer; }
        .room-done-Müsait { background: var(--g) !important; color: white; }
        .room-done-Sonra { background: var(--y) !important; color: black; }
        .room-done-DND { background: var(--r) !important; color: white; }
        #modal { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); z-index:20000; align-items:center; justify-content:center; }
        .modal-box { background:white; width:90%; max-width:500px; border-radius:12px; padding:20px; position:relative; max-height:80vh; overflow-y:auto; }
        .close-modal { position:absolute; right:15px; top:10px; font-size:24px; cursor:pointer; }
        .history-bars { position: absolute; left: 0; top: 0; bottom: 0; display: flex; width: 8px; }
        .h-bar { width: 3px; height: 100%; }
        .h-r { background: var(--r); } .h-y { background: var(--y); } .h-g { background: var(--g); }
        .sub-item { display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px dashed #eee; font-size: 13px; }
        table { width: 100%; border-collapse: collapse; font-size: 11px; }
        th, td { border: 1px solid #eee; padding: 8px; text-align: left; }
    </style>
</head>
<body>
    <div id="modal"><div class="modal-box"><span class="close-modal" onclick="document.getElementById('modal').style.display='none'">&times;</span><h3 id="modalTitle"></h3><div id="modalBody"></div></div></div>
    <div id="loginPage" class="page active">
        <h2 style="text-align:center">Smart Minibar</h2>
        <input type="text" id="lUser" placeholder="Kullanıcı Adı">
        <input type="password" id="lPass" placeholder="Şifre">
        <button class="btn-p" onclick="login()">GİRİŞ YAP</button>
    </div>
    <div id="staffPage" class="page">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px">
            <b id="sn"></b> <button onclick="location.reload()" style="background:#eee">Çıkış</button>
        </div>
        <div id="staffContent"><div id="blockTabs" class="grid"></div><div id="view_floors" class="grid"></div><div id="view_rooms" class="grid"></div></div>
        <div id="statusScreen" style="display:none; text-align:center;">
            <h2 id="targetRoomTitle"></h2>
            <button class="btn-p" style="background:var(--g); height:70px; margin-bottom:10px" onclick="openProductMenu()">MÜSAİT</button>
            <button class="btn-p" style="background:var(--y); color:black; height:70px; margin-bottom:10px" onclick="submitLog('Sonra', '-')">SONRA</button>
            <button class="btn-p" style="background:var(--r); height:70px; margin-bottom:10px" onclick="submitLog('DND', '-')">DND</button>
            <button class="btn-p" onclick="backToRooms()" style="background:var(--gr); height:50px; margin-top:10px">⬅ GERİ</button>
        </div>
        <div id="productMenu" style="display:none">
            <h3 style="text-align:center">Harcama Girişi</h3>
            <div id="pGrid" class="grid"></div>
            <button class="btn-p" onclick="processAndSubmit()" style="background:var(--g); height:65px; font-size:18px;">KAYDI TAMAMLA</button>
            <button class="btn-p" onclick="openStatusMenu(selRoom)" style="background:var(--gr); height:50px; margin-top:10px">⬅ GERİ</button>
        </div>
    </div>
    <div id="adminPage" class="page">
        <div class="admin-tabs">
            <button class="a-tab active" onclick="switchAdminTab('t_live', this)">👁️ Canlı</button>
            <button class="a-tab" onclick="switchAdminTab('t_matrix', this)">🏢 Matris</button>
            <button class="a-tab" onclick="switchAdminTab('t_setup', this)">⚙️ Yapılandır</button>
            <button class="a-tab" onclick="switchAdminTab('t_end', this)">🧹 Gün Sonu</button>
            <button onclick="location.reload()" style="background:#ddd">Çıkış</button>
        </div>
        <div id="t_live" class="tab-content active"><table id="liveLogTable"><thead><tr><th>Zaman</th><th>Oda</th><th>Pers.</th><th>Durum</th><th>Harcamalar</th></tr></thead><tbody id="liveBody"></tbody></table></div>
        <div id="t_matrix" class="tab-content"><div id="matrixArea"></div></div>
        <div id="t_setup" class="tab-content">
            <div class="sub-tabs"><button class="s-tab active" onclick="switchSubTab('s_hotel', this)">Otel Yapısı</button><button class="s-tab" onclick="switchSubTab('s_prods', this)">Ürünler</button><button class="s-tab" onclick="switchSubTab('s_staff', this)">Personel</button></div>
            <div id="s_hotel" class="sub-content active">
                <input type="text" id="inB" placeholder="Blok"> <input type="text" id="inF" placeholder="Kat"> <input type="text" id="inR" placeholder="Odalar (101,102)">
                <button class="btn-p" onclick="addStruct()">YAPIYI EKLE</button>
                <div id="fullStructureList"></div>
            </div>
            <div id="s_prods" class="sub-content"><input type="text" id="inP" placeholder="Ürün Adı"><button class="btn-p" onclick="addProd()">EKLE</button><div id="pList"></div></div>
            <div id="s_staff" class="sub-content"><input type="text" id="inUN" placeholder="Ad"><input type="text" id="inUP" placeholder="Şifre"><button class="btn-p" onclick="addStaff()">EKLE</button><div id="uList"></div></div>
        </div>
        <div id="t_end" class="tab-content" style="text-align:center"><button class="btn-p" onclick="window.open('/api/export')" style="background:var(--g); margin-bottom:10px">📥 EXCEL İNDİR</button><button class="btn-p" onclick="endDay()" style="background:var(--r)">GÜNÜ SIFIRLA</button></div>
    </div>
    <script>
        let currentUser = "", currentSid = "", hotelData = [], logs = [], products = [], selBlock = "", selFloor = "", selRoom = "", startTime = "", counts = {};

        async function login() {
            const user = document.getElementById('lUser').value, pass = document.getElementById('lPass').value;
            const res = await fetch('/api/login', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user,pass}) });
            const data = await res.json();
            if(data.success) {
                currentUser = user; currentSid = data.sid;
                document.getElementById('loginPage').classList.remove('active');
                if(user === 'hakkı') { document.getElementById('adminPage').classList.add('active'); initAdmin(); }
                else { document.getElementById('staffPage').classList.add('active'); document.getElementById('sn').innerText="Personel: "+user; initStaff(); }
            } else alert("Hata!");
        }

        async function initStaff() {
            hotelData = await (await fetch('/api/structure')).json();
            products = await (await fetch('/api/products')).json();
            logs = await (await fetch('/api/logs')).json();
            document.getElementById('blockTabs').innerHTML = hotelData.map(b => '<button onclick="selectBlock(\\''+b.name+'\\')">'+b.name+'</button>').join('');
        }

        function selectBlock(n) { selBlock = n; const b = hotelData.find(x => x.name === n); document.getElementById('view_rooms').style.display = 'none'; document.getElementById('view_floors').style.display = 'grid'; document.getElementById('view_floors').innerHTML = b.floors.map(f => '<button style="background:var(--b); color:white" onclick="selectFloor(\\''+f.name+'\\')">'+f.name+'</button>').join(''); }
        function selectFloor(n) { selFloor = n; const f = hotelData.find(x => x.name === selBlock).floors.find(x => x.name === n); document.getElementById('view_floors').style.display = 'none'; document.getElementById('view_rooms').style.display = 'grid'; document.getElementById('view_rooms').innerHTML = f.rooms.map(r => { const today = new Date().toLocaleDateString('tr-TR'); const last = logs.find(l => String(l.room) === String(r) && l.date === today); return '<button class="btn-room room-done-'+(last?last.status:'')+'" onclick="openStatusMenu(\\''+r+'\\')">'+r+'</button>'; }).join(''); }
        function openStatusMenu(r) { selRoom = r; startTime = new Date().toLocaleTimeString(); document.getElementById('staffContent').style.display='none'; document.getElementById('productMenu').style.display='none'; document.getElementById('statusScreen').style.display='block'; document.getElementById('targetRoomTitle').innerText="Oda "+r; }
        function openProductMenu() { document.getElementById('statusScreen').style.display='none'; document.getElementById('productMenu').style.display='block'; counts = {}; products.forEach(p => counts[p.name] = 0); document.getElementById('pGrid').innerHTML = products.map((p,i) => '<div style="border:1px solid #ddd; padding:10px; text-align:center; border-radius:8px;" onclick="counts[\\''+p.name+'\\']++; document.getElementById(\\'c'+i+'\\').innerText=counts[\\''+p.name+'\\']">'+p.name+'<br><b id="c'+i+'" style="color:var(--b); font-size:18px;">0</b></div>').join(''); }
        function processAndSubmit() { let items = Object.entries(counts).filter(e => e[1] > 0).map(e => e[0] + " x" + e[1]); submitLog('Müsait', items.length > 0 ? items.join(", ") : "Kontrol Edildi"); }
        async function submitLog(status, details) { const entry = { id: Date.now(), room: selRoom, status, details, staff: currentUser, startTime: startTime }; await fetch('/api/logs', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(entry) }); alert("Başarılı"); location.reload(); }
        function backToRooms() { location.reload(); }

        async function initAdmin() { hotelData = await (await fetch('/api/structure')).json(); products = await (await fetch('/api/products')).json(); refreshAdminData(); setInterval(refreshAdminData, 5000); }
        function switchAdminTab(id, btn) { document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active')); document.querySelectorAll('.a-tab').forEach(b => b.classList.remove('active')); document.getElementById(id).classList.add('active'); btn.classList.add('active'); }
        function switchSubTab(id, btn) { document.querySelectorAll('.sub-content').forEach(c => c.classList.remove('active')); document.querySelectorAll('.s-tab').forEach(b => b.classList.remove('active')); document.getElementById(id).classList.add('active'); btn.classList.add('active'); }
        async function refreshAdminData() { const res = await fetch('/api/logs'); logs = await res.json(); document.getElementById('liveBody').innerHTML = logs.slice(0,25).map(l => '<tr><td>'+l.startTime+'</td><td>'+l.room+'</td><td>'+l.staff+'</td><td>'+l.status+'</td><td>'+l.details+'</td></tr>').join(''); if(document.getElementById('t_matrix').classList.contains('active')) renderMatrix(); renderSetupLists(); }
        function renderMatrix() { let h = ""; hotelData.forEach(b => { h += '<h4>'+b.name+'</h4><div class="grid">'; b.floors.forEach(f => { f.rooms.forEach(r => { h += '<div class="btn-room">'+r+'</div>'; }); }); h += '</div>'; }); document.getElementById('matrixArea').innerHTML = h; }
        async function renderSetupLists() { document.getElementById('pList').innerHTML = products.map((p,i) => '<div class="sub-item">'+p.name+'</div>').join(''); }
        async function addStruct() { const bV = document.getElementById('inB').value, fV = document.getElementById('inF').value, rV = document.getElementById('inR').value; let b = hotelData.find(x => x.name === bV) || {name:bV, floors:[]}; if(!hotelData.find(x=>x.name===bV)) hotelData.push(b); let f = b.floors.find(x => x.name === fV) || {name:fV, rooms:[]}; if(!b.floors.find(x=>x.name===fV)) b.floors.push(f); f.rooms = [...new Set([...f.rooms, ...rV.split(',').map(s=>s.trim())])]; await fetch('/api/structure', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(hotelData) }); location.reload(); }
        async function addProd() { products.push({name: document.getElementById('inP').value}); await fetch('/api/products', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(products) }); location.reload(); }
        async function addStaff() { await fetch('/api/users', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user:document.getElementById('inUN').value, pass:document.getElementById('inUP').value}) }); location.reload(); }
        async function endDay() { if(confirm("Sıfırla?")) { await fetch('/api/end-day',{method:'POST'}); location.reload(); } }
    </script>
</body>
</html>
    `);
});

app.listen(PORT, '0.0.0.0', () => console.log(`Smart Minibar v16.6 Aktif: ${PORT}`));
