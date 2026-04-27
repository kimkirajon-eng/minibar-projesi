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
    notes: path.join(DATA_DIR, 'notes.json'),
    stokLogs: path.join(DATA_DIR, 'stokLogs.json')
};

app.use(bodyParser.json());

const readDB = (f) => {
    if (!fs.existsSync(f)) {
        if (f === DB.users) return [{ user: 'hakkı', pass: '2125', role: 'admin' }];
        if (f === DB.products) return [{ name: 'Su', stok: 100 }, { name: 'Kola', stok: 50 }];
        if (f === DB.stokLogs) return [];
        if (f === DB.structure) return [];
        return f === DB.notes ? {} : [];
    }
    try { return JSON.parse(fs.readFileSync(f)); } catch (e) { return f === DB.notes ? {} : []; }
};

const writeDB = (f, d) => fs.writeFileSync(f, JSON.stringify(d, null, 2));

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
    writeDB(DB.logs, logs); 
    
    if (req.body.items && Array.isArray(req.body.items)) {
        let products = readDB(DB.products);
        req.body.items.forEach(item => {
            let p = products.find(x => x.name === item.name);
            if (p) p.stok = (p.stok || 0) - item.count;
        });
        writeDB(DB.products, products);
    }
    let notes = readDB(DB.notes);
    delete notes[req.body.room];
    writeDB(DB.notes, notes);
    res.json({ success: true });
});

app.post('/api/stok-islem', (req, res) => {
    let products = readDB(DB.products);
    let stokLogs = readDB(DB.stokLogs);
    const { productName, amount, source, type } = req.body;
    let p = products.find(x => x.name === productName);
    if(p) {
        const islemMiktari = parseInt(amount);
        if(type === 'giris') p.stok = (p.stok || 0) + islemMiktari;
        else p.stok = (p.stok || 0) - islemMiktari;
        
        stokLogs.unshift({
            productName,
            amount: type === 'giris' ? "+" + amount : "-" + amount,
            source,
            date: new Date().toLocaleDateString('tr-TR'),
            time: new Date().toLocaleTimeString('tr-TR')
        });
        writeDB(DB.products, products);
        writeDB(DB.stokLogs, stokLogs);
        res.json({ success: true });
    } else res.json({ success: false });
});

app.get('/api/stok-logs', (req, res) => res.json(readDB(DB.stokLogs)));
app.get('/api/products', (req, res) => res.json(readDB(DB.products)));
app.post('/api/products', (req, res) => { writeDB(DB.products, req.body); res.json({ success: true }); });
app.get('/api/structure', (req, res) => res.json(readDB(DB.structure)));
app.post('/api/structure', (req, res) => { writeDB(DB.structure, req.body); res.json({ success: true }); });
app.get('/api/notes', (req, res) => res.json(readDB(DB.notes)));
app.post('/api/notes', (req, res) => {
    const notes = readDB(DB.notes);
    notes[req.body.room] = req.body.note;
    writeDB(DB.notes, notes);
    res.json({ success: true });
});
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
    res.setHeader('Content-Disposition', 'attachment; filename=rapor.xlsx');
    res.send(buf);
});

app.get('/', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Smart Minibar v17.4</title>
    <style>
        :root { --p: #2c3e50; --g: #2ecc71; --y: #f1c40f; --r: #e74c3c; --b: #3498db; --gr: #95a5a6; --note: #9b59b6; }
        body { font-family: sans-serif; background: #f4f7f6; margin: 0; padding: 10px; }
        .page { display: none; max-width: 1100px; margin: auto; background: white; padding: 15px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .active { display: block; }
        button { padding: 12px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
        input, select { width: 100%; padding: 12px; margin: 8px 0; border: 1px solid #ddd; border-radius: 8px; box-sizing: border-box; }
        .btn-p { background: var(--p); color: white; width: 100%; margin: 5px 0; }
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); gap: 8px; margin: 15px 0; }
        .btn-room { background: #fff; border: 1px solid #ddd; height: 50px; border-radius: 8px; display:flex; align-items:center; justify-content:center; cursor: pointer; font-weight: bold; font-size: 12px; position: relative; overflow: hidden; }
        .has-note::after { content: "💬"; position: absolute; top: -5px; right: -2px; font-size: 14px; z-index: 10; }
        .note-bg { background-color: #fce4ec !important; border: 2px solid var(--note) !important; }
        .status-Müsait { background-color: var(--g) !important; color: white !important; }
        .status-Sonra { background-color: var(--y) !important; color: black !important; }
        .status-DND { background-color: var(--r) !important; color: white !important; }
        .history-container { position: absolute; left: 0; top: 0; bottom: 0; display: flex; flex-direction: row; gap: 0px; padding: 0px; }
        .h-bar { width: 6px; height: 100%; }
        .h-Müsait { background: var(--g); } .h-Sonra { background: var(--y); } .h-DND { background: var(--r); }
        .admin-tabs { display: flex; gap: 5px; margin-bottom: 15px; background: #eee; padding: 5px; border-radius: 8px; overflow-x: auto; }
        .a-tab { flex: 1; padding: 10px; font-size: 11px; white-space: nowrap; background: none; }
        .a-tab.active { background: white; color: var(--p); box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        table { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 10px; }
        th, td { border: 1px solid #eee; padding: 10px; text-align: left; }
        th { background: #34495e; color: white; }
        #modal { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.8); z-index:9999; align-items:center; justify-content:center; }
        .modal-box { background:white; width:90%; max-width:400px; padding:20px; border-radius:15px; overflow-y: auto; max-height: 80vh; }
        .staff-note-box { background: #fff3cd; color: #856404; padding: 10px; border-radius: 8px; margin-bottom: 10px; border: 1px solid #ffeeba; font-weight: bold; }
        .stok-form { background: #f9f9f9; padding: 15px; border-radius: 10px; border: 1px solid #eee; margin-bottom: 20px; }
        .stok-kritik { color: var(--r); font-weight: bold; }
    </style>
</head>
<body>
    <div id="modal">
        <div class="modal-box"><h3 id="mTitle"></h3><div id="mBody"></div><button class="btn-p" onclick="document.getElementById('modal').style.display='none'">Kapat</button></div>
    </div>

    <div id="loginPage" class="page active">
        <h2 style="text-align:center;">Smart Minibar</h2>
        <input type="text" id="lUser" placeholder="Kullanıcı Adı">
        <input type="password" id="lPass" placeholder="Şifre">
        <button class="btn-p" onclick="login()">GİRİŞ YAP</button>
    </div>

    <div id="staffPage" class="page">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px">
            <b id="sn"></b> <button onclick="logout()">Çıkış</button>
        </div>
        <div id="staffContent">
            <div id="blockTabs" class="grid"></div>
            <div id="view_floors" class="grid" style="display:none"></div>
            <div id="view_rooms" class="grid" style="display:none"></div>
        </div>
        <div id="statusScreen" style="display:none; text-align:center;">
            <div id="staffNoteView" class="staff-note-box" style="display:none"></div>
            <h2 id="targetRoomTitle"></h2>
            <button class="btn-p" style="background:var(--g); height:80px;" onclick="openProductMenu()">MÜSAİT / GİRİLDİ</button>
            <button class="btn-p" style="background:var(--y); color:#000; height:60px" onclick="submitLog('Sonra', '-')">SONRA BAKILACAK</button>
            <button class="btn-p" style="background:var(--r); height:60px" onclick="submitLog('DND', '-')">DND / GİRİLEMEZ</button>
            <button class="btn-p" style="background:var(--gr); margin-top:20px" onclick="goBackToRooms()">⬅ GERİ</button>
        </div>
        <div id="productMenu" style="display:none">
            <div id="pGrid" class="grid"></div>
            <button class="btn-p" style="background:var(--g); height:70px;" onclick="processAndSubmit()">KAYDI TAMAMLA</button>
            <button class="btn-p" style="background:var(--gr);" onclick="openStatusMenu(selRoom)">⬅ GERİ</button>
        </div>
    </div>

    <div id="adminPage" class="page">
        <div class="admin-tabs">
            <button class="a-tab active" onclick="switchAdminTab('t_live', this)">👁️ Canlı</button>
            <button class="a-tab" onclick="switchAdminTab('t_matrix', this)">🏢 Harita</button>
            <button class="a-tab" onclick="switchAdminTab('t_stok', this)">📦 Stok İşlem</button>
            <button class="a-tab" onclick="switchAdminTab('t_setup', this)">⚙️ Ayarlar</button>
            <button class="a-tab" onclick="switchAdminTab('t_end', this)">🧹 Gün Sonu</button>
            <button onclick="logout()">Çıkış</button>
        </div>
        <div id="t_live" class="tab-content active"><table><thead><tr><th>Oda</th><th>P.</th><th>Durum</th><th>Ürünler</th><th>Saat</th></tr></thead><tbody id="liveBody"></tbody></table></div>
        <div id="t_matrix" class="tab-content">
            <button id="noteModeBtn" class="btn-p" onclick="toggleNoteMode()">📝 NOT EKLEME MODU: KAPALI</button>
            <div id="matrixArea"></div>
        </div>
        <div id="t_stok" class="tab-content">
            <div class="stok-form">
                <h4>➕ Ürün Giriş / Çıkış Yap</h4>
                <select id="stokProdSelect"></select>
                <input id="stokAmount" type="number" placeholder="Adet">
                <select id="stokSource">
                    <option value="Ana Depo">Ana Depo</option>
                    <option value="Satın Alma">Satın Alma</option>
                    <option value="Fire / Zayi">Fire / Zayi</option>
                    <option value="Transfer">Transfer</option>
                    <option value="Düzeltme">Düzeltme</option>
                </select>
                <div style="display:flex; gap:10px;">
                    <button class="btn-p" style="background:var(--g); flex:1" onclick="submitStokIslem('giris')">Stoğa Ekle (+)</button>
                    <button class="btn-p" style="background:var(--r); flex:1" onclick="submitStokIslem('cikis')">Stoktan Çıkar (-)</button>
                </div>
            </div>
            <h4>Güncel Stok Listesi</h4>
            <table><thead><tr><th>Ürün</th><th>Mevcut Stok</th></tr></thead><tbody id="stokListBody"></tbody></table>
            <h4 style="margin-top:20px;">📜 Son Stok Hareketleri</h4>
            <table><thead><tr><th>Ürün</th><th>Miktar</th><th>Kaynak/Neden</th><th>Tarih/Saat</th></tr></thead><tbody id="stokHistoryBody"></tbody></table>
        </div>
        <div id="t_setup" class="tab-content">
            <h4>Personel</h4><div id="uList"></div><input id="inUN" placeholder="Ad"><input id="inUP" placeholder="Şifre"><button class="btn-p" onclick="addStaff()">Ekle</button>
            <h4>Yapı</h4><input id="inB" placeholder="Blok"><input id="inF" placeholder="Kat"><input id="inR" placeholder="Odalar (101,102)"><button class="btn-p" onclick="addStruct()">Kaydet</button>
            <h4>Ürün</h4><div id="pList"></div><input id="inP" placeholder="Ürün Adı"><input id="inPS" type="number" placeholder="Başlangıç Stoğu"><button class="btn-p" onclick="addProd()">Ekle</button>
        </div>
        <div id="t_end" class="tab-content" style="text-align:center"><button class="btn-p" onclick="window.open('/api/export')" style="background:var(--g)">EXCEL İNDİR</button><button class="btn-p" onclick="endDay()" style="background:var(--r)">SIFIRLA</button></div>
    </div>

    <script>
        let currentUser = null, hotelData = [], products = [], logs = [], notes = {}, counts = {}, selRoom = "", selBlock = "", selFloor = "", noteMode = false;

        async function login() {
            const userVal = document.getElementById('lUser').value;
            const passVal = document.getElementById('lPass').value;
            if(!userVal || !passVal) return alert("Bilgileri girin!");
            const res = await fetch('/api/login', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user:userVal, pass:passVal}) });
            const data = await res.json();
            if(data.success) { currentUser = data; launchApp(); } else alert("Hatalı giriş!");
        }

        function logout() { location.reload(); }

        function launchApp() {
            document.getElementById('loginPage').classList.remove('active');
            setInterval(autoUpdate, 5000);
            if(currentUser.role === 'admin') { 
                document.getElementById('adminPage').classList.add('active'); 
                initAdmin(); 
            } else { 
                document.getElementById('staffPage').classList.add('active'); 
                document.getElementById('sn').innerText = "P: " + currentUser.user; 
                initStaff(); 
            }
        }

        async function autoUpdate() {
            const [l, n, p] = await Promise.all([fetch('/api/logs').then(r=>r.json()), fetch('/api/notes').then(r=>r.json()), fetch('/api/products').then(r=>r.json())]);
            logs = l; notes = n; products = p;
            if(document.getElementById('staffPage').classList.contains('active') && document.getElementById('view_rooms').style.display === 'grid') selectFloor(selFloor);
            if(document.getElementById('adminPage').classList.contains('active')) {
                if(document.getElementById('t_live').classList.contains('active')) refreshLiveLogs();
                if(document.getElementById('t_matrix').classList.contains('active')) refreshMatrix();
                if(document.getElementById('t_stok').classList.contains('active')) refreshStokTab();
            }
        }

        async function initStaff() {
            const [s, p, l, n] = await Promise.all([fetch('/api/structure').then(r=>r.json()), fetch('/api/products').then(r=>r.json()), fetch('/api/logs').then(r=>r.json()), fetch('/api/notes').then(r=>r.json())]);
            hotelData = s; products = p; logs = l; notes = n; renderBlocks();
        }

        function renderBlocks() {
            const bt = document.getElementById('blockTabs'); bt.style.display='grid';
            document.getElementById('view_floors').style.display='none'; document.getElementById('view_rooms').style.display='none';
            bt.innerHTML = hotelData.map(b => '<button class="btn-room" onclick="selectBlock(\''+b.name+'\')">'+b.name+'</button>').join('');
        }

        function selectBlock(n) {
            selBlock = n; const b = hotelData.find(x => x.name === n);
            document.getElementById('blockTabs').style.display='none'; document.getElementById('view_floors').style.display='grid';
            document.getElementById('view_floors').innerHTML = b.floors.map(f => '<button class="btn-room" style="background:var(--b);color:#fff" onclick="selectFloor(\''+f.name+'\')">'+f.name+'. Kat</button>').join('') + '<button class="btn-p" onclick="renderBlocks()">Geri</button>';
        }

        function selectFloor(n) {
            selFloor = n; const f = hotelData.find(x => x.name === selBlock).floors.find(x => x.name === n);
            document.getElementById('view_floors').style.display='none'; document.getElementById('view_rooms').style.display='grid';
            document.getElementById('view_rooms').innerHTML = f.rooms.map(r => {
                const log = logs.find(l => String(l.room) === String(r));
                const noteCls = notes[r] ? 'note-bg has-note' : '';
                return '<button class="btn-room '+(log?'status-'+log.status:'')+' '+noteCls+'" onclick="openStatusMenu(\''+r+'\')">'+r+'</button>';
            }).join('') + '<button class="btn-p" onclick="selectBlock(\''+selBlock+'\')">Geri</button>';
        }

        function openStatusMenu(r) {
            selRoom = r; document.getElementById('staffContent').style.display='none'; document.getElementById('statusScreen').style.display='block';
            document.getElementById('targetRoomTitle').innerText = "Oda " + r;
            const snv = document.getElementById('staffNoteView');
            if(notes[r]) { snv.style.display='block'; snv.innerText = "NOT: " + notes[r]; } else snv.style.display='none';
        }

        function goBackToRooms() { document.getElementById('staffContent').style.display='block'; document.getElementById('statusScreen').style.display='none'; document.getElementById('productMenu').style.display='none'; selectFloor(selFloor); }

        function openProductMenu() {
            document.getElementById('statusScreen').style.display='none'; document.getElementById('productMenu').style.display='block'; counts = {}; products.forEach(p => counts[p.name] = 0);
            document.getElementById('pGrid').innerHTML = products.map((p,i) => '<div style="border:1px solid #ddd;padding:5px;text-align:center;border-radius:8px" onclick="counts[\''+p.name+'\']++; document.getElementById(\'c'+i+'\').innerText=counts[\''+p.name+'\']">'+p.name+'<br><small>Stok: '+(p.stok||0)+'</small><br><b id="c'+i+'" style="color:var(--b);font-size:20px">0</b></div>').join('');
        }

        async function submitLog(status, details, items = []) {
            await fetch('/api/logs', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({room:selRoom, status, details, items, staff:currentUser.user}) });
            goBackToRooms();
        }

        function processAndSubmit() {
            let itmArr = Object.entries(counts).filter(e => e[1] > 0).map(e => ({ name: e[0], count: e[1] }));
            let details = itmArr.length > 0 ? itmArr.map(i => i.name + " x" + i.count).join(", ") : "Kontrol Edildi";
            submitLog('Müsait', details, itmArr);
        }

        async function initAdmin() {
            const [p, s, u, l, n] = await Promise.all([fetch('/api/products').then(r=>r.json()), fetch('/api/structure').then(r=>r.json()), fetch('/api/users').then(r=>r.json()), fetch('/api/logs').then(r=>r.json()), fetch('/api/notes').then(r=>r.json())]);
            products = p; hotelData = s; logs = l; notes = n;
            document.getElementById('uList').innerHTML = u.filter(x => x.role !== 'admin').map(x => '<div>'+x.user+' <button onclick="delUser(\''+x.user+'\')">Sil</button></div>').join('');
            document.getElementById('pList').innerHTML = products.map(x => x.name + " (" + (x.stok||0) + ")").join(', ');
            refreshStokTab();
        }

        async function refreshStokTab() {
            document.getElementById('stokProdSelect').innerHTML = products.map(p => '<option value="'+p.name+'">'+p.name+'</option>').join('');
            document.getElementById('stokListBody').innerHTML = products.map(p => '<tr><td>'+p.name+'</td><td class="'+(p.stok < 10 ? 'stok-kritik' : '')+'">'+(p.stok || 0)+'</td></tr>').join('');
            const sh = await fetch('/api/stok-logs').then(r=>r.json());
            document.getElementById('stokHistoryBody').innerHTML = sh.map(h => '<tr><td>'+h.productName+'</td><td style="color:'+(h.amount.startsWith('+') ? 'green' : 'red')+'">'+h.amount+'</td><td>'+h.source+'</td><td>'+h.date+' '+h.time+'</td></tr>').join('');
        }

        async function submitStokIslem(type) {
            const body = { productName: document.getElementById('stokProdSelect').value, amount: document.getElementById('stokAmount').value, source: document.getElementById('stokSource').value, type: type };
            if(!body.amount || body.amount <= 0) return alert("Adet girin!");
            await fetch('/api/stok-islem', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body) });
            document.getElementById('stokAmount').value = ""; autoUpdate(); alert("Kaydedildi.");
        }

        function refreshMatrix() {
            let h = "";
            hotelData.forEach(b => {
                h += '<h4>'+b.name+'</h4>';
                b.floors.forEach(f => {
                    h += '<div class="grid">';
                    f.rooms.forEach(r => {
                        const roomLogs = logs.filter(l => String(l.room) === String(r)).reverse();
                        const lastLog = roomLogs[roomLogs.length-1];
                        const noteCls = notes[r] ? 'note-bg has-note' : '';
                        let hist = '<div class="history-container">' + roomLogs.map(rl => '<div class="h-bar h-'+rl.status+'"></div>').join('') + '</div>';
                        h += '<div class="btn-room '+(lastLog?'status-'+lastLog.status:'')+' '+noteCls+'" onclick="handleMatrixClick(\''+r+'\')">'+hist+r+'</div>';
                    });
                    h += '</div>';
                });
            });
            document.getElementById('matrixArea').innerHTML = h;
        }

        function handleMatrixClick(r) {
            if(noteMode) {
                const n = prompt(r + " için not:");
                if(n !== null) fetch('/api/notes', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({room:r, note:n}) }).then(() => autoUpdate());
            } else showRoomDetail(r);
        }

        function showRoomDetail(r) {
            const roomLogs = logs.filter(l => String(l.room) === String(r));
            document.getElementById('mTitle').innerText = "Oda " + r + " Tarihçesi";
            document.getElementById('mBody').innerHTML = (notes[r] ? '<div class="staff-note-box">NOT: '+notes[r]+'</div>' : '') + roomLogs.map(l => '<div><b>'+l.endTime+'</b>: '+l.status+' ('+l.staff+')<br><small>'+l.details+'</small></div>').join('<hr>');
            document.getElementById('modal').style.display='flex';
        }

        function switchAdminTab(id, btn) {
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            document.querySelectorAll('.a-tab').forEach(b => b.classList.remove('active'));
            document.getElementById(id).classList.add('active'); btn.classList.add('active');
            if(id === 't_matrix') refreshMatrix();
            if(id === 't_stok') refreshStokTab();
        }

        function toggleNoteMode() {
            noteMode = !noteMode;
            const btn = document.getElementById('noteModeBtn');
            btn.innerText = noteMode ? "📝 NOT EKLEME MODU: AÇIK" : "📝 NOT EKLEME MODU: KAPALI";
            btn.style.background = noteMode ? "var(--r)" : "var(--note)";
        }

        async function addStaff() { await fetch('/api/users', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user:document.getElementById('inUN').value, pass:document.getElementById('inUP').value}) }); initAdmin(); document.getElementById('inUN').value=''; document.getElementById('inUP').value=''; }
        async function delUser(n) { if(confirm('Sil?')) { await fetch('/api/users/'+n, { method:'DELETE' }); initAdmin(); } }
        async function addStruct() {
            let s = hotelData; let b = s.find(x => x.name === document.getElementById('inB').value);
            if(!b) { b = {name: document.getElementById('inB').value, floors: []}; s.push(b); }
            b.floors.push({name: document.getElementById('inF').value, rooms: document.getElementById('inR').value.split(',').map(x => x.trim())});
            await fetch('/api/structure', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(s) }); initAdmin();
        }
        async function addProd() {
            let p = products; p.push({name:document.getElementById('inP').value, stok: parseInt(document.getElementById('inPS').value)||0});
            await fetch('/api/products', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p) }); initAdmin();
        }
        async function endDay() { if(confirm("Sıfırla?")) { await fetch('/api/end-day', { method:'POST' }); location.reload(); } }
        function refreshLiveLogs() { 
            document.getElementById('liveBody').innerHTML = logs.map(l => '<tr class="status-'+l.status+'"><td><b>'+l.room+'</b></td><td>'+l.staff+'</td><td>'+l.status+'</td><td>'+(l.details || '-')+'</td><td>'+l.endTime+'</td></tr>').join(''); 
        }
    </script>
</body>
</html>
    `);
});

app.listen(PORT, '0.0.0.0', () => console.log(`v17.4 Smart Minibar Hazır: http://localhost:${PORT}`));

