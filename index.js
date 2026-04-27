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

app.get('/', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Smart Minibar v17.5</title>
    <style>
        :root { --p: #2c3e50; --g: #2ecc71; --y: #f1c40f; --r: #e74c3c; --b: #3498db; --gr: #95a5a6; --note: #9b59b6; }
        body { font-family: sans-serif; background: #f0f3f5; margin: 0; padding: 10px; }
        .page { display: none; max-width: 1100px; margin: auto; background: white; padding: 15px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .active { display: block; }
        button { padding: 12px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
        input, textarea { width: 100%; padding: 12px; margin: 8px 0; border: 1px solid #ddd; border-radius: 8px; box-sizing: border-box; }
        .btn-p { background: var(--p); color: white; width: 100%; margin: 5px 0; }
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); gap: 10px; margin: 15px 0; }
        .btn-room { background: #fff; border: 1px solid #ddd; height: 55px; border-radius: 8px; display:flex; align-items:center; justify-content:center; cursor: pointer; font-weight: bold; position: relative; overflow: hidden; }
        
        /* Renkler */
        .status-Müsait { background-color: var(--g) !important; color: white !important; }
        .status-Sonra { background-color: var(--y) !important; color: black !important; }
        .status-DND { background-color: var(--r) !important; color: white !important; }
        .has-note { border: 2px solid var(--note) !important; background-color: #f3e5f5 !important; }
        
        .note-icon { position: absolute; top: 2px; right: 2px; font-size: 12px; background: white; border-radius: 50%; width: 16px; height: 16px; display: flex; align-items: center; justify-content: center; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
        
        .admin-tabs { display: flex; gap: 5px; margin-bottom: 15px; background: #eee; padding: 5px; border-radius: 8px; }
        .a-tab { flex: 1; padding: 10px; font-size: 11px; background: none; }
        .a-tab.active { background: white; color: var(--p); }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        table { width: 100%; border-collapse: collapse; font-size: 11px; }
        th, td { border: 1px solid #eee; padding: 8px; text-align: left; }
        #modal { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); z-index:9999; align-items:center; justify-content:center; }
        .modal-box { background:white; width:90%; max-width:400px; padding:20px; border-radius:12px; }
    </style>
</head>
<body>
    <div id="modal">
        <div class="modal-box"><h3 id="mTitle"></h3><div id="mBody"></div><button class="btn-p" onclick="closeModal()">Kapat</button></div>
    </div>

    <div id="loginPage" class="page active">
        <h2 style="text-align:center">Smart Minibar</h2>
        <input type="text" id="lUser" placeholder="Kullanıcı Adı">
        <input type="password" id="lPass" placeholder="Şifre">
        <button class="btn-p" onclick="login()">GİRİŞ YAP</button>
    </div>

    <div id="staffPage" class="page">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px">
            <b id="sn"></b> <button onclick="logout()" style="background:#eee">Çıkış</button>
        </div>
        <div id="staffContent">
            <div id="blockTabs" class="grid"></div>
            <div id="view_floors" class="grid" style="display:none"></div>
            <div id="view_rooms" class="grid" style="display:none"></div>
        </div>
        <div id="statusScreen" style="display:none">
            <div id="staffNoteBox" style="background:#fff3cd; padding:10px; border-radius:8px; margin-bottom:10px; display:none; border-left:5px solid var(--note)"></div>
            <h2 id="targetRoomTitle" style="text-align:center"></h2>
            <button class="btn-p" style="background:var(--g); height:80px" onclick="openProductMenu()">MÜSAİT / GİRİLDİ</button>
            <button class="btn-p" style="background:var(--y); color:#000; height:70px" onclick="submitLog('Sonra', '-')">SONRA BAKILACAK</button>
            <button class="btn-p" style="background:var(--r); height:70px" onclick="submitLog('DND', '-')">DND / GİRİLEMEZ</button>
            <button class="btn-p" style="background:var(--gr); margin-top:20px" onclick="goBackToRooms()">⬅ GERİ</button>
        </div>
        <div id="productMenu" style="display:none">
            <div id="pGrid" class="grid"></div>
            <button class="btn-p" style="background:var(--g); height:80px" onclick="processAndSubmit()">KAYDI TAMAMLA</button>
        </div>
    </div>

    <div id="adminPage" class="page">
        <div class="admin-tabs">
            <button class="a-tab active" onclick="switchAdminTab('t_live', this)">👁️ Canlı</button>
            <button class="a-tab" onclick="switchAdminTab('t_matrix', this)">🏢 Harita</button>
            <button class="a-tab" onclick="switchAdminTab('t_setup', this)">⚙️ Ayar</button>
            <button onclick="logout()" style="background:#ddd">Çıkış</button>
        </div>
        <div id="t_live" class="tab-content active"><table><thead><tr><th>Oda</th><th>Durum</th><th>Saat</th></tr></thead><tbody id="liveBody"></tbody></table></div>
        <div id="t_matrix" class="tab-content">
            <button id="btnNoteMode" class="btn-p" onclick="toggleNoteMode()" style="background:var(--note); margin-bottom:10px">📝 NOT EKLEME MODU: KAPALI</button>
            <div id="matrixArea"></div>
        </div>
        <div id="t_setup" class="tab-content">
            <input type="text" id="inUN" placeholder="Personel Adı"> <input type="text" id="inUP" placeholder="Şifre">
            <button class="btn-p" onclick="addStaff()">PERSONEL EKLE</button>
            <button class="btn-p" style="background:var(--r); margin-top:20px" onclick="endDay()">🧹 GÜNÜ SIFIRLA</button>
        </div>
    </div>

    <script>
        let currentUser = null, hotelData = [], products = [], logs = [], roomNotes = {}, selRoom = "", selBlock = "", selFloor = "";
        let isNoteMode = false;

        window.onload = () => { const s = localStorage.getItem('minibar_user'); if(s) { currentUser = JSON.parse(s); launchApp(); } };

        async function login() {
            const user = document.getElementById('lUser').value, pass = document.getElementById('lPass').value;
            const res = await fetch('/api/login', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user,pass}) });
            const data = await res.json();
            if(data.success) { currentUser = data; localStorage.setItem('minibar_user', JSON.stringify(data)); launchApp(); } else alert("Hata!");
        }

        function logout() { localStorage.removeItem('minibar_user'); location.reload(); }

        function launchApp() {
            document.getElementById('loginPage').classList.remove('active');
            if(currentUser.role === 'admin') { document.getElementById('adminPage').classList.add('active'); initAdmin(); }
            else { document.getElementById('staffPage').classList.add('active'); document.getElementById('sn').innerText = "Personel: " + currentUser.user; initStaff(); }
        }

        async function initData() {
            const [s, p, l, n] = await Promise.all([
                fetch('/api/structure').then(r=>r.json()),
                fetch('/api/products').then(r=>r.json()),
                fetch('/api/logs').then(r=>r.json()),
                fetch('/api/notes').then(r=>r.json())
            ]);
            hotelData = s; products = p; logs = l; roomNotes = n;
        }

        // --- STAFF ---
        async function initStaff() { await initData(); renderBlocks(); }
        function renderBlocks() {
            document.getElementById('staffContent').style.display='block'; document.getElementById('statusScreen').style.display='none';
            document.getElementById('view_floors').style.display='none'; document.getElementById('view_rooms').style.display='none'; document.getElementById('blockTabs').style.display='grid';
            document.getElementById('blockTabs').innerHTML = hotelData.map(b => '<button class="btn-room" onclick="selectBlock(\\''+b.name+'\\')">Blok '+b.name+'</button>').join('');
        }
        function selectBlock(n) {
            selBlock = n; const b = hotelData.find(x => x.name === n);
            document.getElementById('blockTabs').style.display='none'; document.getElementById('view_floors').style.display='grid';
            document.getElementById('view_floors').innerHTML = b.floors.map(f => '<button class="btn-room" style="background:var(--b); color:white" onclick="selectFloor(\\''+f.name+'\\')">Kat '+f.name+'</button>').join('') + '<button class="btn-p" onclick="renderBlocks()">⬅ GERİ</button>';
        }
        function selectFloor(n) {
            selFloor = n; const f = hotelData.find(x => x.name === selBlock).floors.find(x => x.name === n);
            document.getElementById('view_floors').style.display='none'; document.getElementById('view_rooms').style.display='grid'; document.getElementById('statusScreen').style.display='none'; document.getElementById('staffContent').style.display='block';
            document.getElementById('view_rooms').innerHTML = f.rooms.map(r => {
                const log = logs.find(l => String(l.room) === String(r));
                const note = roomNotes[r];
                return \`<button class="btn-room \${log?'status-'+log.status:''} \${note?'has-note':''}" onclick="openStatusMenu('\${r}')">
                    \${note?'<span class="note-icon">📝</span>':''} \${r}
                </button>\`;
            }).join('') + '<button class="btn-p" onclick="selectBlock(\\''+selBlock+'\\')">⬅ GERİ</button>';
        }
        function openStatusMenu(r) {
            selRoom = r; document.getElementById('staffContent').style.display='none'; document.getElementById('statusScreen').style.display='block'; document.getElementById('targetRoomTitle').innerText="Oda "+r;
            const nb = document.getElementById('staffNoteBox');
            if(roomNotes[r]) { nb.innerHTML = "<b>YÖNETİCİ NOTU:</b><br>" + roomNotes[r]; nb.style.display='block'; } else nb.style.display='none';
        }
        function goBackToRooms() { selectFloor(selFloor); }

        async function submitLog(status, details) {
            await fetch('/api/logs', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({room:selRoom, status, details, staff:currentUser.user}) });
            initStaff();
        }

        function openProductMenu() {
            document.getElementById('statusScreen').style.display='none'; document.getElementById('productMenu').style.display='block';
            let counts = {}; products.forEach(p => counts[p.name] = 0);
            document.getElementById('pGrid').innerHTML = products.map((p,i) => \`<div style="border:1px solid #ddd; padding:10px; text-align:center; border-radius:8px;" onclick="this.querySelector('b').innerText = ++counts['\${p.name}']">\${p.name}<br><b>0</b></div>\`).join('');
            window.processAndSubmit = () => {
                let items = Object.entries(counts).filter(e => e[1] > 0).map(e => e[0] + " x" + e[1]);
                submitLog('Müsait', items.length > 0 ? items.join(", ") : "Kontrol Edildi");
            };
        }

        // --- ADMIN ---
        async function initAdmin() { await initData(); refreshLiveLogs(); refreshMatrix(); }
        function toggleNoteMode() { isNoteMode = !isNoteMode; document.getElementById('btnNoteMode').innerText = "📝 NOT EKLEME MODU: " + (isNoteMode ? "AÇIK" : "KAPALI"); }
        
        function refreshMatrix() {
            let h = ""; hotelData.forEach(b => {
                h += '<h4>Blok '+b.name+'</h4>'; b.floors.forEach(f => {
                    h += '<div class="grid">'; f.rooms.forEach(r => {
                        const log = logs.find(l => String(l.room) === String(r));
                        const note = roomNotes[r];
                        h += \`<button class="btn-room \${log?'status-'+log.status:''} \${note?'has-note':''}" onclick="handleRoomClick('\${r}')">
                            \${note?'<span class="note-icon">📝</span>':''} \${r}
                        </button>\`;
                    }); h += '</div>';
                });
            }); document.getElementById('matrixArea').innerHTML = h;
        }

        function handleRoomClick(r) {
            if(isNoteMode) {
                const n = prompt("Oda " + r + " için not yazın (Boş bırakırsanız silinir):", roomNotes[r] || "");
                if(n !== null) { if(n === "") delete roomNotes[r]; else roomNotes[r] = n; saveNotes(); }
            } else showRoomDetail(r);
        }

        async function saveNotes() { 
            await fetch('/api/notes', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(roomNotes) }); 
            initAdmin(); 
        }

        function showRoomDetail(r) {
            const l = logs.filter(x => String(x.room) === String(r));
            document.getElementById('mTitle').innerText = "Oda " + r;
            document.getElementById('mBody').innerHTML = (roomNotes[r] ? "<p style='color:var(--note)'><b>NOT:</b> "+roomNotes[r]+"</p><hr>" : "") + 
                (l.length ? l.map(x => "<div>"+x.endTime+" - "+x.status+" ("+x.staff+")</div>").join('') : "Kayıt yok.");
            document.getElementById('modal').style.display='flex';
        }

        function closeModal() { document.getElementById('modal').style.display='none'; }
        function switchAdminTab(id, btn) { 
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active')); 
            document.querySelectorAll('.a-tab').forEach(b => b.classList.remove('active'));
            document.getElementById(id).classList.add('active'); btn.classList.add('active'); 
        }
        function refreshLiveLogs() { document.getElementById('liveBody').innerHTML = logs.slice(0,20).map(l => "<tr><td>"+l.room+"</td><td>"+l.status+"</td><td>"+l.endTime+"</td></tr>").join(''); }
        async function addStaff() { await fetch('/api/users', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user:document.getElementById('inUN').value, pass:document.getElementById('inUP').value}) }); initAdmin(); }
        async function endDay() { if(confirm("Sıfırla?")) { await fetch('/api/end-day', { method:'POST' }); initAdmin(); } }
    </script>
</body>
</html>
    `);
});

app.listen(PORT, '0.0.0.0', () => console.log(`v17.5 Aktif: ${PORT}`));
