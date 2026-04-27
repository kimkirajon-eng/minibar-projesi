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
        if (f === DB.products) return [{ name: 'Su' }, { name: 'Kola' }, { name: 'Bira' }];
        if (f === DB.structure) return [{name: "A", floors: [{name: "1", rooms: ["101", "102"]}]}];
        return [];
    }
    return JSON.parse(fs.readFileSync(f));
};

const writeDB = (f, d) => fs.writeFileSync(f, JSON.stringify(d, null, 2));

app.get('/sw.js', (req, res) => {
    res.setHeader('Content-Type', 'application/javascript');
    res.send(`self.addEventListener('install', e => self.skipWaiting()); self.addEventListener('fetch', e => e);`);
});

app.post('/api/login', (req, res) => {
    const users = readDB(DB.users);
    const u = users.find(u => u.user === req.body.user && u.pass === req.body.pass);
    if (u) {
        const sid = Math.random().toString(36).substring(7);
        u.sid = sid;
        writeDB(DB.users, users);
        res.json({ success: true, role: u.role, user: u.user, sid: sid });
    } else res.json({ success: false });
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
    res.setHeader('Content-Disposition', 'attachment; filename=minibar_rapor.xlsx');
    res.send(buf);
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
    <title>Smart Minibar v16.7</title>
    <style>
        :root { --p: #2c3e50; --g: #2ecc71; --y: #f1c40f; --r: #e74c3c; --b: #3498db; --gr: #95a5a6; }
        body { font-family: sans-serif; background: #f4f7f6; margin: 0; padding: 10px; }
        .page { display: none; max-width: 1100px; margin: auto; background: white; padding: 15px; border-radius: 12px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .active { display: block; }
        button { padding: 12px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; }
        input { width: 100%; padding: 12px; margin: 8px 0; border: 1px solid #ddd; border-radius: 8px; box-sizing: border-box; }
        .btn-p { background: var(--p); color: white; width: 100%; margin: 5px 0; }
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); gap: 10px; margin: 15px 0; }
        .btn-room { background: #fff; border: 1px solid #ddd; height: 55px; border-radius: 8px; display:flex; align-items:center; justify-content:center; cursor: pointer; font-weight: bold; }
        .room-done-Müsait { background: var(--g) !important; color: white; border:none; }
        .room-done-Sonra { background: var(--y) !important; color: #000; border:none; }
        .room-done-DND { background: var(--r) !important; color: white; border:none; }
        .admin-tabs { display: flex; gap: 5px; margin-bottom: 15px; overflow-x: auto; background: #eee; padding: 5px; border-radius: 8px; }
        .a-tab { flex: 1; padding: 10px; font-size: 12px; white-space: nowrap; background: none; }
        .a-tab.active { background: white; color: var(--p); box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 10px; }
        th, td { border: 1px solid #eee; padding: 10px; text-align: left; }
        th { background: #f8f9fa; }
        .sub-item { display: flex; justify-content: space-between; align-items: center; padding: 10px; border-bottom: 1px solid #eee; }
    </style>
</head>
<body>
    <div id="loginPage" class="page active">
        <h2 style="text-align:center; color:var(--p)">Smart Minibar</h2>
        <input type="text" id="lUser" placeholder="Kullanıcı Adı">
        <input type="password" id="lPass" placeholder="Şifre">
        <button class="btn-p" onclick="login()">GİRİŞ YAP</button>
    </div>

    <div id="staffPage" class="page">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px">
            <b id="sn"></b> <button onclick="logout()" style="background:#eee; padding:5px 10px">Çıkış</button>
        </div>
        
        <div id="staffContent">
            <div id="blockTabs" class="grid"></div>
            <div id="view_floors" class="grid" style="display:none"></div>
            <div id="view_rooms" class="grid" style="display:none"></div>
        </div>

        <div id="statusScreen" style="display:none; text-align:center;">
            <h2 id="targetRoomTitle"></h2>
            <button class="btn-p" style="background:var(--g); height:80px; font-size:18px" onclick="openProductMenu()">MÜSAİT / GİRİLDİ</button>
            <button class="btn-p" style="background:var(--y); color:#000; height:70px" onclick="submitLog('Sonra', '-')">SONRA BAKILACAK</button>
            <button class="btn-p" style="background:var(--r); height:70px" onclick="submitLog('DND', '-')">DND / GİRİLEMEZ</button>
            <button class="btn-p" style="background:var(--gr); margin-top:20px" onclick="showStaffMain()">⬅ GERİ DÖN</button>
        </div>

        <div id="productMenu" style="display:none">
            <h3 style="text-align:center">Harcama Girişi</h3>
            <div id="pGrid" class="grid"></div>
            <button class="btn-p" style="background:var(--g); height:70px; font-size:20px" onclick="processAndSubmit()">KAYDI TAMAMLA</button>
            <button class="btn-p" style="background:var(--gr); margin-top:10px" onclick="openStatusMenu(selRoom)">⬅ GERİ</button>
        </div>
    </div>

    <div id="adminPage" class="page">
        <div class="admin-tabs">
            <button class="a-tab active" onclick="switchAdminTab('t_live', this)">👁️ Canlı</button>
            <button class="a-tab" onclick="switchAdminTab('t_setup', this)">⚙️ Yapılandır</button>
            <button class="a-tab" onclick="switchAdminTab('t_end', this)">🧹 Gün Sonu</button>
            <button onclick="logout()" style="background:#ddd">Çıkış</button>
        </div>

        <div id="t_live" class="tab-content active">
            <div style="overflow-x:auto"><table id="liveLogTable"><thead><tr><th>Oda</th><th>Pers.</th><th>Durum</th><th>Harcamalar</th><th>Saat</th></tr></thead><tbody id="liveBody"></tbody></table></div>
        </div>

        <div id="t_setup" class="tab-content">
            <h4>Oda Yapısı Ekle</h4>
            <input type="text" id="inB" placeholder="Blok (Örn: A)">
            <input type="text" id="inF" placeholder="Kat (Örn: 1)">
            <input type="text" id="inR" placeholder="Odalar (Örn: 101,102,103)">
            <button class="btn-p" onclick="addStruct()">YAPIYI KAYDET</button>
            <hr>
            <h4>Ürün Ekle</h4>
            <input type="text" id="inP" placeholder="Ürün Adı">
            <button class="btn-p" onclick="addProd()">ÜRÜN EKLE</button>
            <div id="pList"></div>
        </div>

        <div id="t_end" class="tab-content" style="text-align:center">
            <button class="btn-p" onclick="window.open('/api/export')" style="background:var(--g); height:60px">📥 EXCEL RAPORU İNDİR</button>
            <button class="btn-p" onclick="endDay()" style="background:var(--r); margin-top:20px">🧹 TÜM KAYITLARI SIFIRLA</button>
        </div>
    </div>

    <script>
        let currentUser = null;
        let hotelData = [], products = [], logs = [], counts = {}, selRoom = "";

        // SAYFA YÜKLENDİĞİNDE OTURUM KONTROLÜ
        window.onload = () => {
            const savedUser = localStorage.getItem('minibar_user');
            if(savedUser) {
                currentUser = JSON.parse(savedUser);
                launchApp();
            }
        };

        async function login() {
            const user = document.getElementById('lUser').value, pass = document.getElementById('lPass').value;
            const res = await fetch('/api/login', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user,pass}) });
            const data = await res.json();
            if(data.success) {
                currentUser = data;
                localStorage.setItem('minibar_user', JSON.stringify(data));
                launchApp();
            } else alert("Kullanıcı adı veya şifre hatalı!");
        }

        function logout() {
            localStorage.removeItem('minibar_user');
            location.reload();
        }

        function launchApp() {
            document.getElementById('loginPage').classList.remove('active');
            if(currentUser.role === 'admin') {
                document.getElementById('adminPage').classList.add('active');
                initAdmin();
            } else {
                document.getElementById('staffPage').classList.add('active');
                document.getElementById('sn').innerText = "Personel: " + currentUser.user;
                initStaff();
            }
        }

        // PERSONEL FONKSİYONLARI
        async function initStaff() {
            hotelData = await (await fetch('/api/structure')).json();
            products = await (await fetch('/api/products')).json();
            logs = await (await fetch('/api/logs')).json();
            renderBlocks();
        }

        function renderBlocks() {
            document.getElementById('staffContent').style.display = 'block';
            document.getElementById('statusScreen').style.display = 'none';
            document.getElementById('productMenu').style.display = 'none';
            document.getElementById('view_floors').style.display = 'none';
            document.getElementById('view_rooms').style.display = 'none';
            document.getElementById('blockTabs').style.display = 'grid';
            document.getElementById('blockTabs').innerHTML = hotelData.map(b => '<button class="btn-room" onclick="selectBlock(\\''+b.name+'\\')">Blok '+b.name+'</button>').join('');
        }

        function selectBlock(n) {
            const b = hotelData.find(x => x.name === n);
            document.getElementById('blockTabs').style.display = 'none';
            document.getElementById('view_floors').style.display = 'grid';
            document.getElementById('view_floors').innerHTML = b.floors.map(f => '<button class="btn-room" style="background:var(--b); color:white" onclick="selectFloor(\\''+n+'\\',\\''+f.name+'\\')">Kat '+f.name+'</button>').join('') + '<button class="btn-p" onclick="renderBlocks()">⬅ GERİ</button>';
        }

        function selectFloor(bn, fn) {
            const f = hotelData.find(x => x.name === bn).floors.find(x => x.name === fn);
            document.getElementById('view_floors').style.display = 'none';
            document.getElementById('view_rooms').style.display = 'grid';
            document.getElementById('view_rooms').innerHTML = f.rooms.map(r => {
                const log = logs.find(l => String(l.room) === String(r));
                const cls = log ? 'room-done-' + log.status : '';
                return '<button class="btn-room '+cls+'" onclick="openStatusMenu(\\''+r+'\\')">'+r+'</button>';
            }).join('') + '<button class="btn-p" onclick="selectBlock(\\''+bn+'\\')">⬅ GERİ</button>';
        }

        function openStatusMenu(r) {
            selRoom = r;
            document.getElementById('staffContent').style.display = 'none';
            document.getElementById('productMenu').style.display = 'none';
            document.getElementById('statusScreen').style.display = 'block';
            document.getElementById('targetRoomTitle').innerText = "Oda " + r;
        }

        function openProductMenu() {
            document.getElementById('statusScreen').style.display = 'none';
            document.getElementById('productMenu').style.display = 'block';
            counts = {};
            products.forEach(p => counts[p.name] = 0);
            document.getElementById('pGrid').innerHTML = products.map((p,i) => \`
                <div style="border:1px solid #ddd; padding:10px; text-align:center; border-radius:8px;" onclick="counts['\${p.name}']++; document.getElementById('c\${i}').innerText=counts['\${p.name}']">
                    \${p.name}<br><b id="c\${i}" style="color:var(--b); font-size:20px;">0</b>
                </div>
            \`).join('');
        }

        function showStaffMain() { initStaff(); }

        async function submitLog(status, details) {
            const entry = { room: selRoom, status, details, staff: currentUser.user };
            await fetch('/api/logs', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(entry) });
            alert("Kaydedildi");
            initStaff(); 
        }

        function processAndSubmit() {
            let items = Object.entries(counts).filter(e => e[1] > 0).map(e => e[0] + " x" + e[1]);
            submitLog('Müsait', items.length > 0 ? items.join(", ") : "Kontrol Edildi");
        }

        // YÖNETİCİ FONKSİYONLARI
        async function initAdmin() {
            const lRes = await fetch('/api/logs'); logs = await lRes.json();
            const pRes = await fetch('/api/products'); products = await pRes.json();
            const sRes = await fetch('/api/structure'); hotelData = await sRes.json();
            
            document.getElementById('liveBody').innerHTML = logs.map(l => '<tr><td>'+l.room+'</td><td>'+l.staff+'</td><td>'+l.status+'</td><td>'+l.details+'</td><td>'+l.endTime+'</td></tr>').join('');
            document.getElementById('pList').innerHTML = products.map(p => '<div class="sub-item">'+p.name+'</div>').join('');
        }

        function switchAdminTab(id, btn) {
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            document.querySelectorAll('.a-tab').forEach(b => b.classList.remove('active'));
            document.getElementById(id).classList.add('active');
            btn.classList.add('active');
            initAdmin();
        }

        async function addStruct() {
            const b = document.getElementById('inB').value, f = document.getElementById('inF').value, r = document.getElementById('inR').value;
            if(!b || !f || !r) return alert("Eksik alan!");
            let current = hotelData;
            let block = current.find(x => x.name === b);
            if(!block) { block = {name: b, floors: []}; current.push(block); }
            block.floors.push({name: f, rooms: r.split(',').map(x => x.trim())});
            await fetch('/api/structure', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(current) });
            alert("Eklendi"); initAdmin();
        }

        async function addProd() {
            products.push({name: document.getElementById('inP').value});
            await fetch('/api/products', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(products) });
            document.getElementById('inP').value = ""; initAdmin();
        }

        async function endDay() {
            if(confirm("Tüm veriler temizlenecek, emin misiniz?")) {
                await fetch('/api/end-day', { method:'POST' });
                initAdmin();
            }
        }
    </script>
</body>
</html>
    `);
});

app.listen(PORT, '0.0.0.0', () => console.log(`Smart Minibar v16.7 Aktif: Port ${PORT}`));

