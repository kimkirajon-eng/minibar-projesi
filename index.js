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
        if (f === DB.users) return [{ user: 'hakkı', pass: '2125', role: 'admin' }];
        if (f === DB.products) return [{ name: 'Su' }, { name: 'Kola' }];
        return [];
    }
    return JSON.parse(fs.readFileSync(f));
};

const writeDB = (f, d) => fs.writeFileSync(f, JSON.stringify(d, null, 2));

// API
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

app.get('/api/export', (req, res) => {
    const logs = readDB(DB.logs);
    const ws = XLSX.utils.json_to_sheet(logs);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rapor");
    const buf = XLSX.write(wb, {bookType:'xlsx', type:'buffer'});
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=rapor.xlsx');
    res.send(buf);
});

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
app.post('/api/end-day', (req, res) => { writeDB(DB.logs, []); res.json({ success: true }); });

app.get('/', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Smart Minibar v17.4</title>
    <style>
        :root { --p: #2c3e50; --g: #2ecc71; --y: #f1c40f; --r: #e74c3c; --b: #3498db; --gr: #95a5a6; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f3f5; margin: 0; padding: 10px; color: #333; }
        .page { display: none; max-width: 1100px; margin: auto; background: white; padding: 20px; border-radius: 16px; box-shadow: 0 10px 30px rgba(0,0,0,0.05); }
        .active { display: block; }
        
        /* Buton Genel Stil */
        button { padding: 12px; border: none; border-radius: 10px; cursor: pointer; font-weight: 600; transition: all 0.2s; }
        input { width: 100%; padding: 12px; margin: 8px 0; border: 1px solid #e0e0e0; border-radius: 10px; box-sizing: border-box; background: #fafafa; }
        .btn-p { background: var(--p); color: white; width: 100%; margin: 5px 0; }
        .btn-p:active { transform: scale(0.98); }

        /* Matris Oda Butonları - Profesyonel Görünüm */
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(85px, 1fr)); gap: 12px; margin: 15px 0; }
        .btn-room { 
            background: #ffffff; 
            border: 1px solid #edf2f7; 
            height: 55px; 
            border-radius: 12px; 
            display: flex; 
            align-items: center; 
            justify-content: center; 
            cursor: pointer; 
            font-size: 14px; 
            position: relative; 
            overflow: hidden; 
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
            color: #4a5568;
        }

        /* Sol Kenar Geçmiş Şeridi */
        .history-strip { 
            position: absolute; 
            left: 0; 
            top: 0; 
            bottom: 0; 
            width: 8px; 
            display: flex; 
            flex-direction: column; 
        }
        .h-segment { flex: 1; width: 100%; }
        .h-Müsait { background: var(--g); }
        .h-Sonra { background: var(--y); }
        .h-DND { background: var(--r); }

        /* Sağ Üst Durum Noktası */
        .status-dot {
            position: absolute;
            top: 8px;
            right: 8px;
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #cbd5e0;
        }
        .dot-Müsait { background: var(--g); box-shadow: 0 0 8px var(--g); }
        .dot-Sonra { background: var(--y); box-shadow: 0 0 8px var(--y); }
        .dot-DND { background: var(--r); box-shadow: 0 0 8px var(--r); }

        /* Aktif Oda Stil (Admin Panel Renklendirme) */
        .room-active { background: #f8fafc; border-color: #e2e8f0; font-weight: bold; color: #2d3748; }

        .admin-tabs { display: flex; gap: 8px; margin-bottom: 20px; background: #f7fafc; padding: 6px; border-radius: 12px; }
        .a-tab { flex: 1; padding: 10px; font-size: 12px; background: transparent; color: #718096; }
        .a-tab.active { background: white; color: var(--p); box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
        
        table { width: 100%; border-collapse: collapse; font-size: 12px; border-radius: 10px; overflow: hidden; }
        th { background: #f8fafc; color: #4a5568; padding: 12px; text-align: left; border-bottom: 2px solid #edf2f7; }
        td { padding: 12px; border-bottom: 1px solid #edf2f7; }
        
        tr.row-Müsait { border-left: 5px solid var(--g); }
        tr.row-Sonra { border-left: 5px solid var(--y); }
        tr.row-DND { border-left: 5px solid var(--r); }

        #modal { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); backdrop-filter: blur(4px); z-index:9999; align-items:center; justify-content:center; }
        .modal-box { background:white; width:90%; max-width:400px; padding:25px; border-radius:20px; box-shadow: 0 20px 40px rgba(0,0,0,0.2); }
    </style>
</head>
<body>
    <div id="modal">
        <div class="modal-box"><h3 id="mTitle" style="margin-top:0; color:var(--p)"></h3><div id="mBody" style="margin-bottom:20px"></div><button class="btn-p" onclick="document.getElementById('modal').style.display='none'">Kapat</button></div>
    </div>

    <div id="loginPage" class="page active">
        <h2 style="text-align:center; color:var(--p); margin-bottom:30px">Smart Minibar</h2>
        <input type="text" id="lUser" placeholder="Kullanıcı Adı">
        <input type="password" id="lPass" placeholder="Şifre">
        <button class="btn-p" onclick="login()">GİRİŞ YAP</button>
    </div>

    <div id="staffPage" class="page">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px">
            <b id="sn" style="color:var(--p)"></b> <button onclick="logout()" style="background:#f7fafc; color:var(--r); font-size:12px">Güvenli Çıkış</button>
        </div>
        <div id="staffContent">
            <div id="blockTabs" class="grid"></div>
            <div id="view_floors" class="grid" style="display:none"></div>
            <div id="view_rooms" class="grid" style="display:none"></div>
        </div>
        <div id="statusScreen" style="display:none; text-align:center;">
            <h2 id="targetRoomTitle" style="color:var(--p)"></h2>
            <button class="btn-p" style="background:var(--g); height:90px; font-size:18px; margin-bottom:12px" onclick="openProductMenu()">MÜSAİT / GİRİLDİ</button>
            <button class="btn-p" style="background:var(--y); color:#000; height:70px; margin-bottom:12px" onclick="submitLog('Sonra', '-')">SONRA BAKILACAK</button>
            <button class="btn-p" style="background:var(--r); height:70px" onclick="submitLog('DND', '-')">DND / GİRİLEMEZ</button>
            <button class="btn-p" style="background:#edf2f7; color:#4a5568; margin-top:20px" onclick="goBackToRooms()">⬅ VAZGEÇ / GERİ</button>
        </div>
        <div id="productMenu" style="display:none">
            <h3 style="text-align:center; color:var(--p)">Harcama Girişi</h3>
            <div id="pGrid" class="grid"></div>
            <button class="btn-p" style="background:var(--g); height:80px; font-size:20px; margin-top:15px" onclick="processAndSubmit()">KAYDI TAMAMLA</button>
            <button class="btn-p" style="background:#edf2f7; color:#4a5568; margin-top:10px" onclick="openStatusMenu(selRoom)">⬅ GERİ</button>
        </div>
    </div>

    <div id="adminPage" class="page">
        <div class="admin-tabs">
            <button class="a-tab active" onclick="switchAdminTab('t_live', this)">👁️ Canlı</button>
            <button class="a-tab" onclick="switchAdminTab('t_matrix', this)">🏢 Harita</button>
            <button class="a-tab" onclick="switchAdminTab('t_setup', this)">⚙️ Ayar</button>
            <button class="a-tab" onclick="switchAdminTab('t_end', this)">🧹 Gün Sonu</button>
            <button onclick="logout()" style="background:transparent; color:var(--r); font-size:18px">×</button>
        </div>
        <div id="t_live" class="tab-content active"><div style="overflow-x:auto"><table><thead><tr><th>Oda</th><th>Personel</th><th>Durum</th><th>Saat</th></tr></thead><tbody id="liveBody"></tbody></table></div></div>
        <div id="t_matrix" class="tab-content"><div id="matrixArea"></div></div>
        <div id="t_setup" class="tab-content">
            <div style="background:#f8fafc; padding:15px; border-radius:12px; margin-bottom:15px"><h4>Personel Yönetimi</h4><input type="text" id="inUN" placeholder="Ad"> <input type="text" id="inUP" placeholder="Şifre"><button class="btn-p" onclick="addStaff()" style="background:var(--b)">EKLE</button><div id="uList" style="margin-top:10px"></div></div>
            <div style="background:#f8fafc; padding:15px; border-radius:12px; margin-bottom:15px"><h4>Otel Yapısı</h4><input type="text" id="inB" placeholder="Blok"> <input type="text" id="inF" placeholder="Kat"> <input type="text" id="inR" placeholder="Odalar (101,102)"><button class="btn-p" onclick="addStruct()">KAYDET</button></div>
            <div style="background:#f8fafc; padding:15px; border-radius:12px;"><h4>Ürün Listesi</h4><input type="text" id="inP" placeholder="Ürün Adı"><button class="btn-p" onclick="addProd()">EKLE</button><div id="pList" style="margin-top:10px"></div></div>
        </div>
        <div id="t_end" class="tab-content" style="text-align:center"><button class="btn-p" onclick="window.open('/api/export')" style="background:var(--g); height:65px; font-size:18px">📥 RAPORU İNDİR (EXCEL)</button><button class="btn-p" onclick="endDay()" style="background:var(--r); margin-top:30px; opacity:0.8">🧹 TÜM VERİLERİ SIFIRLA</button></div>
    </div>

    <script>
        let currentUser = null, hotelData = [], products = [], logs = [], counts = {}, selRoom = "", selBlock = "", selFloor = "";
        let adminRefreshInterval = null;

        window.onload = () => {
            const savedUser = localStorage.getItem('minibar_user');
            if(savedUser) { currentUser = JSON.parse(savedUser); launchApp(); }
        };

        async function login() {
            const user = document.getElementById('lUser').value, pass = document.getElementById('lPass').value;
            const res = await fetch('/api/login', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user,pass}) });
            const data = await res.json();
            if(data.success) { currentUser = data; localStorage.setItem('minibar_user', JSON.stringify(data)); launchApp(); }
            else alert("Giriş başarısız!");
        }

        function logout() { localStorage.removeItem('minibar_user'); if(adminRefreshInterval) clearInterval(adminRefreshInterval); location.reload(); }

        function launchApp() {
            document.getElementById('loginPage').classList.remove('active');
            if(currentUser.role === 'admin') { document.getElementById('adminPage').classList.add('active'); initAdmin(); adminRefreshInterval = setInterval(autoUpdate, 5000); }
            else { document.getElementById('staffPage').classList.add('active'); document.getElementById('sn').innerText = "Hoş geldin, " + currentUser.user; initStaff(); }
        }

        async function autoUpdate() { 
            const res = await fetch('/api/logs'); logs = await res.json(); 
            if(document.getElementById('t_live').classList.contains('active')) refreshLiveLogs();
            if(document.getElementById('t_matrix').classList.contains('active')) refreshMatrix();
        }

        // --- STAFF ---
        async function initStaff() {
            const [s, p, l] = await Promise.all([fetch('/api/structure').then(r=>r.json()), fetch('/api/products').then(r=>r.json()), fetch('/api/logs').then(r=>r.json())]);
            hotelData = s; products = p; logs = l; renderBlocks();
        }

        function renderBlocks() {
            selBlock = ""; selFloor = "";
            document.getElementById('staffContent').style.display='block'; document.getElementById('statusScreen').style.display='none'; document.getElementById('productMenu').style.display='none';
            document.getElementById('view_floors').style.display='none'; document.getElementById('view_rooms').style.display='none'; document.getElementById('blockTabs').style.display='grid';
            document.getElementById('blockTabs').innerHTML = hotelData.map(b => '<button class="btn-room" onclick="selectBlock(\\''+b.name+'\\')">Blok '+b.name+'</button>').join('');
        }

        function selectBlock(n) {
            selBlock = n; const b = hotelData.find(x => x.name === n);
            document.getElementById('blockTabs').style.display='none'; document.getElementById('view_floors').style.display='grid'; document.getElementById('view_rooms').style.display='none';
            document.getElementById('view_floors').innerHTML = b.floors.map(f => '<button class="btn-room" style="background:var(--b); color:white" onclick="selectFloor(\\''+f.name+'\\')">Kat '+f.name+'</button>').join('') + '<button class="btn-p" style="background:#e2e8f0; color:#4a5568" onclick="renderBlocks()">⬅ GERİ</button>';
        }

        function selectFloor(n) {
            selFloor = n; const f = hotelData.find(x => x.name === selBlock).floors.find(x => x.name === n);
            document.getElementById('view_floors').style.display='none'; document.getElementById('view_rooms').style.display='grid'; document.getElementById('statusScreen').style.display='none'; document.getElementById('productMenu').style.display='none'; document.getElementById('staffContent').style.display='block';
            document.getElementById('view_rooms').innerHTML = f.rooms.map(r => { 
                const log = logs.find(l => String(l.room) === String(r)); 
                return '<button class="btn-room '+(log?'room-active':'')+'" onclick="openStatusMenu(\\''+r+'\\')">'+ (log? '<div class="status-dot dot-'+log.status+'"></div>' : '') + r + '</button>'; 
            }).join('') + '<button class="btn-p" style="background:#e2e8f0; color:#4a5568" onclick="selectBlock(\\''+selBlock+'\\')">⬅ GERİ</button>';
        }

        function goBackToRooms() { selectFloor(selFloor); }
        function openStatusMenu(r) { selRoom = r; document.getElementById('staffContent').style.display='none'; document.getElementById('statusScreen').style.display='block'; document.getElementById('targetRoomTitle').innerText="Oda "+r; }
        function openProductMenu() { 
            document.getElementById('statusScreen').style.display='none'; document.getElementById('productMenu').style.display='block'; counts = {}; products.forEach(p => counts[p.name] = 0);
            document.getElementById('pGrid').innerHTML = products.map((p,i) => '<div style="border:1px solid #edf2f7; padding:12px; text-align:center; border-radius:12px; background:white" onclick="counts[\\''+p.name+'\\']++; document.getElementById(\\'c'+i+'\\').innerText=counts[\\''+p.name+'\\']">'+p.name+'<br><b id="c'+i+'" style="color:var(--b); font-size:22px;">0</b></div>').join('');
        }
        async function submitLog(status, details) {
            await fetch('/api/logs', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({room:selRoom, status, details, staff:currentUser.user}) });
            const lRes = await fetch('/api/logs'); logs = await lRes.json();
            selectFloor(selFloor); 
        }
        function processAndSubmit() { let items = Object.entries(counts).filter(e => e[1] > 0).map(e => e[0] + " x" + e[1]); submitLog('Müsait', items.length > 0 ? items.join(", ") : "Kontrol Edildi"); }

        // --- ADMIN ---
        async function initAdmin() {
            const [p, s, u, l] = await Promise.all([fetch('/api/products').then(r=>r.json()), fetch('/api/structure').then(r=>r.json()), fetch('/api/users').then(r=>r.json()), fetch('/api/logs').then(r=>r.json())]);
            products = p; hotelData = s; logs = l;
            document.getElementById('pList').innerHTML = products.map(p => '<div style="font-size:13px; padding:5px; border-bottom:1px solid #eee">'+p.name+'</div>').join('');
            document.getElementById('uList').innerHTML = u.filter(u => u.role !== 'admin').map(u => '<div style="display:flex; justify-content:space-between; padding:5px; border-bottom:1px solid #eee">'+u.user+' <button onclick="delUser(\\''+u.user+'\\')" style="background:var(--r); color:white; padding:2px 6px; font-size:10px">SİL</button></div>').join('');
            refreshLiveLogs();
        }
        function refreshLiveLogs() { document.getElementById('liveBody').innerHTML = logs.slice(0,50).map(l => '<tr class="row-'+l.status+'"><td><b>'+l.room+'</b></td><td>'+l.staff+'</td><td>'+l.status+'</td><td>'+l.endTime+'</td></tr>').join(''); }
        
        function refreshMatrix() {
            let h = ""; hotelData.forEach(b => {
                h += '<h4 style="margin:20px 0 10px 0; color:var(--p)">Blok '+b.name+'</h4>'; 
                b.floors.forEach(f => {
                    h += '<div style="margin-bottom:15px"><small style="color:#718096">Kat '+f.name+'</small><div class="grid">';
                    f.rooms.forEach(r => { 
                        const roomLogs = logs.filter(l => String(l.room) === String(r)).reverse();
                        const lastLog = roomLogs[roomLogs.length - 1];
                        
                        let historyHtml = '<div class="history-strip">';
                        roomLogs.forEach(rl => { historyHtml += '<div class="h-segment h-'+rl.status+'"></div>'; });
                        historyHtml += '</div>';

                        const dotClass = lastLog ? 'dot-' + lastLog.status : '';
                        h += '<div class="btn-room '+(lastLog?'room-active':'')+'" onclick="showRoomDetail(\\''+r+'\\')">' + historyHtml + '<div class="status-dot '+dotClass+'"></div>' + r + '</div>'; 
                    });
                    h += '</div></div>';
                });
            }); 
            document.getElementById('matrixArea').innerHTML = h;
        }

        function showRoomDetail(r) {
            const roomLogs = logs.filter(l => String(l.room) === String(r));
            document.getElementById('mTitle').innerText = "Oda " + r;
            document.getElementById('mBody').innerHTML = roomLogs.length > 0 ? 
                roomLogs.map(l => '<div style="padding:10px; border-bottom:1px solid #edf2f7"><b style="color:var(--p)">'+l.endTime+'</b> - '+l.status+'<br><small style="color:#718096">'+l.staff+' | '+l.details+'</small></div>').join('') : 
                '<p style="color:#a0aec0">İşlem kaydı bulunamadı.</p>';
            document.getElementById('modal').style.display='flex';
        }

        function switchAdminTab(id, btn) {
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            document.querySelectorAll('.a-tab').forEach(b => b.classList.remove('active'));
            document.getElementById(id).classList.add('active'); btn.classList.add('active');
            if(id === 't_matrix') refreshMatrix(); else if(id === 't_live') refreshLiveLogs();
        }
        async function addStaff() { await fetch('/api/users', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user:document.getElementById('inUN').value, pass:document.getElementById('inUP').value}) }); initAdmin(); }
        async function delUser(n) { if(confirm('Silinsin mi?')) { await fetch('/api/users/'+n, { method:'DELETE' }); initAdmin(); } }
        async function addStruct() {
            let current = hotelData; let block = current.find(x => x.name === document.getElementById('inB').value);
            if(!block) { block = {name: document.getElementById('inB').value, floors: []}; current.push(block); }
            block.floors.push({name: document.getElementById('inF').value, rooms: document.getElementById('inR').value.split(',').map(x => x.trim())});
            await fetch('/api/structure', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(current) }); initAdmin();
        }
        async function addProd() { products.push({name: document.getElementById('inP').value}); await fetch('/api/products', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(products) }); initAdmin(); }
        async function endDay() { if(confirm("DİKKAT: Günü kapatmak tüm verileri silecektir! Onaylıyor musunuz?")) { await fetch('/api/end-day', { method:'POST' }); initAdmin(); } }
    </script>
</body>
</html>
    `);
});

app.listen(PORT, '0.0.0.0', () => console.log(`v17.4 Profesyonel Görünüm Aktif: ${PORT}`));
