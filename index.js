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

app.post('/api/login', (req, res) => {
    const users = readDB(DB.users);
    const u = users.find(u => u.user === req.body.user && u.pass === req.body.pass);
    if (u) res.json({ success: true, role: u.role, user: u.user });
    else res.json({ success: false });
});

app.get('/api/logs', (req, res) => res.json(readDB(DB.logs)));
app.post('/api/logs', (req, res) => {
    const logs = readDB(DB.logs);
    logs.unshift({ 
        ...req.body, 
        date: new Date().toLocaleDateString('tr-TR'), 
        endTime: new Date().toLocaleTimeString('tr-TR') 
    });
    writeDB(DB.logs, logs); 
    res.json({ success: true });
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
    <title>Smart Minibar v17.0</title>
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
        
        /* DURUM RENKLERİ */
        .status-Müsait { background-color: var(--g) !important; color: white !important; }
        .status-Sonra { background-color: var(--y) !important; color: black !important; }
        .status-DND { background-color: var(--r) !important; color: white !important; }
        
        .admin-tabs { display: flex; gap: 5px; margin-bottom: 15px; background: #eee; padding: 5px; border-radius: 8px; overflow-x: auto; }
        .a-tab { flex: 1; padding: 10px; font-size: 12px; white-space: nowrap; background: none; }
        .a-tab.active { background: white; color: var(--p); box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        table { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 10px; background: white; }
        th, td { border: 1px solid #eee; padding: 10px; text-align: left; }
        th { background: #34495e; color: white; }
        .sub-item { display: flex; justify-content: space-between; align-items: center; padding: 8px; border-bottom: 1px solid #eee; font-size: 13px; }
        .delete-btn { background: var(--r); color: white; padding: 4px 8px; font-size: 10px; border-radius: 4px; }
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
            <button class="btn-p" style="background:var(--gr); margin-top:20px" onclick="renderBlocks()">⬅ GERİ DÖN</button>
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
            <button class="a-tab active" onclick="switchAdminTab('t_live', this)">👁️ Canlı Takip</button>
            <button class="a-tab" onclick="switchAdminTab('t_setup', this)">⚙️ Yapılandır</button>
            <button class="a-tab" onclick="switchAdminTab('t_end', this)">🧹 Gün Sonu</button>
            <button onclick="logout()" style="background:#ddd">Çıkış</button>
        </div>

        <div id="t_live" class="tab-content active">
            <div style="overflow-x:auto"><table><thead><tr><th>Oda</th><th>Personel</th><th>Durum</th><th>Harcamalar</th><th>Saat</th></tr></thead><tbody id="liveBody"></tbody></table></div>
        </div>

        <div id="t_setup" class="tab-content">
            <div style="background:#f9f9f9; padding:10px; border-radius:8px; margin-bottom:15px">
                <h4 style="margin:0 0 10px 0">Personel Ekle</h4>
                <input type="text" id="inUN" placeholder="Personel Adı">
                <input type="text" id="inUP" placeholder="Şifre">
                <button class="btn-p" onclick="addStaff()" style="background:var(--b)">PERSONELİ KAYDET</button>
                <div id="uList" style="margin-top:10px"></div>
            </div>
            <div style="background:#f9f9f9; padding:10px; border-radius:8px; margin-bottom:15px">
                <h4 style="margin:0 0 10px 0">Otel Yapısı Ekle</h4>
                <input type="text" id="inB" placeholder="Blok">
                <input type="text" id="inF" placeholder="Kat">
                <input type="text" id="inR" placeholder="Odalar (101,102)">
                <button class="btn-p" onclick="addStruct()">YAPIYI KAYDET</button>
            </div>
            <div style="background:#f9f9f9; padding:10px; border-radius:8px;">
                <h4 style="margin:0 0 10px 0">Ürün Ekle</h4>
                <input type="text" id="inP" placeholder="Ürün Adı">
                <button class="btn-p" onclick="addProd()">ÜRÜN EKLE</button>
                <div id="pList" style="margin-top:10px"></div>
            </div>
        </div>

        <div id="t_end" class="tab-content" style="text-align:center">
            <button class="btn-p" onclick="window.open('/api/export')" style="background:var(--g); height:60px">📥 EXCEL RAPORU İNDİR</button>
            <button class="btn-p" onclick="endDay()" style="background:var(--r); margin-top:20px">🧹 TÜM KAYITLARI SIFIRLA</button>
        </div>
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
            else alert("Hata!");
        }

        function logout() { 
            localStorage.removeItem('minibar_user'); 
            if(adminRefreshInterval) clearInterval(adminRefreshInterval);
            location.reload(); 
        }

        function launchApp() {
            document.getElementById('loginPage').classList.remove('active');
            if(currentUser.role === 'admin') { 
                document.getElementById('adminPage').classList.add('active'); 
                initAdmin();
                adminRefreshInterval = setInterval(refreshLiveLogs, 5000); // 5 saniyeye düşürüldü
            }
            else { 
                document.getElementById('staffPage').classList.add('active'); 
                document.getElementById('sn').innerText = "Personel: " + currentUser.user; 
                initStaff(); 
            }
        }

        // --- PERSONEL ---
        async function initStaff() {
            const [s, p, l] = await Promise.all([
                fetch('/api/structure').then(r => r.json()),
                fetch('/api/products').then(r => r.json()),
                fetch('/api/logs').then(r => r.json())
            ]);
            hotelData = s; products = p; logs = l;
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
            selBlock = n;
            const b = hotelData.find(x => x.name === n);
            document.getElementById('blockTabs').style.display = 'none';
            document.getElementById('view_floors').style.display = 'grid';
            document.getElementById('view_floors').innerHTML = b.floors.map(f => '<button class="btn-room" style="background:var(--b); color:white" onclick="selectFloor(\\''+f.name+'\\')">Kat '+f.name+'</button>').join('') + '<button class="btn-p" onclick="renderBlocks()">⬅ GERİ</button>';
        }

        function selectFloor(n) {
            selFloor = n;
            const f = hotelData.find(x => x.name === selBlock).floors.find(x => x.name === n);
            document.getElementById('view_floors').style.display = 'none';
            document.getElementById('view_rooms').style.display = 'grid';
            document.getElementById('view_rooms').innerHTML = f.rooms.map(r => {
                const log = logs.find(l => String(l.room) === String(r));
                const statusCls = log ? 'status-' + log.status : '';
                return '<button class="btn-room '+statusCls+'" onclick="openStatusMenu(\\''+r+'\\')">'+r+'</button>';
            }).join('') + '<button class="btn-p" onclick="selectBlock(\\''+selBlock+'\\')">⬅ GERİ</button>';
        }

        function openStatusMenu(r) {
            selRoom = r;
            document.getElementById('staffContent').style.display='none';
            document.getElementById('statusScreen').style.display='block';
            document.getElementById('targetRoomTitle').innerText="Oda "+r;
        }

        function openProductMenu() {
            document.getElementById('statusScreen').style.display='none';
            document.getElementById('productMenu').style.display='block';
            counts = {}; products.forEach(p => counts[p.name] = 0);
            document.getElementById('pGrid').innerHTML = products.map((p,i) => '<div style="border:1px solid #ddd; padding:10px; text-align:center; border-radius:8px;" onclick="counts[\\''+p.name+'\\']++; document.getElementById(\\'c'+i+'\\').innerText=counts[\\''+p.name+'\\']">'+p.name+'<br><b id="c'+i+'" style="color:var(--b); font-size:20px;">0</b></div>').join('');
        }

        async function submitLog(status, details) {
            await fetch('/api/logs', { 
                method:'POST', 
                headers:{'Content-Type':'application/json'}, 
                body:JSON.stringify({room:selRoom, status, details, staff:currentUser.user}) 
            });
            alert("Kaydedildi");
            const resLogs = await fetch('/api/logs');
            logs = await resLogs.json();
            document.getElementById('statusScreen').style.display='none';
            document.getElementById('productMenu').style.display='none';
            document.getElementById('staffContent').style.display='block';
            selectFloor(selFloor);
        }

        function processAndSubmit() {
            let items = Object.entries(counts).filter(e => e[1] > 0).map(e => e[0] + " x" + e[1]);
            submitLog('Müsait', items.length > 0 ? items.join(", ") : "Kontrol Edildi");
        }

        // --- ADMIN ---
        async function initAdmin() {
            await refreshLiveLogs();
            const [p, s, u] = await Promise.all([
                fetch('/api/products').then(r => r.json()),
                fetch('/api/structure').then(r => r.json()),
                fetch('/api/users').then(r => r.json())
            ]);
            products = p; hotelData = s;
            document.getElementById('pList').innerHTML = products.map(p => '<div class="sub-item">'+p.name+'</div>').join('');
            document.getElementById('uList').innerHTML = u.filter(u => u.role !== 'admin').map(u => '<div class="sub-item">'+u.user+' <button class="delete-btn" onclick="delUser(\\''+u.user+'\\')">SİL</button></div>').join('');
        }

        async function refreshLiveLogs() {
            const res = await fetch('/api/logs');
            const latestLogs = await res.json();
            document.getElementById('liveBody').innerHTML = latestLogs.map(l => {
                const rowClass = 'status-' + l.status; // DURUM SINIFI
                return '<tr class="'+rowClass+'"><td><b>'+l.room+'</b></td><td>'+l.staff+'</td><td>'+l.status+'</td><td>'+l.details+'</td><td>'+l.endTime+'</td></tr>';
            }).join('');
        }

        function switchAdminTab(id, btn) {
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            document.querySelectorAll('.a-tab').forEach(b => b.classList.remove('active'));
            document.getElementById(id).classList.add('active'); btn.classList.add('active');
            if(id === 't_live') refreshLiveLogs();
        }

        async function addStaff() {
            const user = document.getElementById('inUN').value, pass = document.getElementById('inUP').value;
            if(!user || !pass) return alert("Eksik!");
            await fetch('/api/users', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({user, pass}) });
            document.getElementById('inUN').value = ""; document.getElementById('inUP').value = ""; initAdmin();
        }

        async function delUser(name) {
            if(confirm('Silinsin mi?')) { await fetch('/api/users/'+name, { method:'DELETE' }); initAdmin(); }
        }

        async function addStruct() {
            const b = document.getElementById('inB').value, f = document.getElementById('inF').value, r = document.getElementById('inR').value;
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
            if(confirm("Tüm veriler temizlenecek?")) { await fetch('/api/end-day', { method:'POST' }); refreshLiveLogs(); }
        }
    </script>
</body>
</html>
    `);
});

app.listen(PORT, '0.0.0.0', () => console.log(`v17.0 Renkli Takip Aktif: ${PORT}`));
