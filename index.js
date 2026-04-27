```js
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
    try { return JSON.parse(fs.readFileSync(f)); }
    catch (e) { return f === DB.notes ? {} : []; }
};

const writeDB = (f, d) => fs.writeFileSync(f, JSON.stringify(d, null, 2));

// LOGIN
app.post('/api/login', (req, res) => {
    const users = readDB(DB.users);
    const u = users.find(u => u.user === req.body.user && u.pass === req.body.pass);
    if (u) res.json({ success: true, role: u.role, user: u.user });
    else res.json({ success: false });
});

// LOGS
app.get('/api/logs', (req, res) => res.json(readDB(DB.logs)));

app.post('/api/logs', (req, res) => {
    const logs = readDB(DB.logs);
    logs.unshift({
        ...req.body,
        date: new Date().toLocaleDateString('tr-TR'),
        endTime: new Date().toLocaleTimeString('tr-TR')
    });
    writeDB(DB.logs, logs);

    if (req.body.items && Array.isArray(req.body.items)) {
        let products = readDB(DB.products);
        req.body.items.forEach(item => {
            let p = products.find(x => x.name === item.name);
            if (p) {
                const count = Number(item.count) || 0;
                p.stok = (p.stok || 0) - count;
            }
        });
        writeDB(DB.products, products);
    }

    let notes = readDB(DB.notes);
    delete notes[req.body.room];
    writeDB(DB.notes, notes);

    res.json({ success: true });
});

// STOK
app.post('/api/stok-islem', (req, res) => {
    let products = readDB(DB.products);
    let stokLogs = readDB(DB.stokLogs);

    const { productName, amount, source, type } = req.body;
    let p = products.find(x => x.name === productName);

    if (p) {
        const islemMiktari = Number(amount) || 0;

        if (type === 'giris') {
            p.stok = (p.stok || 0) + islemMiktari;
        } else {
            p.stok = (p.stok || 0) - islemMiktari;
        }

        stokLogs.unshift({
            productName,
            amount: type === 'giris' ? '+' + islemMiktari : '-' + islemMiktari,
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
    const u = readDB(DB.users);
    u.push({ ...req.body, role: 'staff' });
    writeDB(DB.users, u);
    res.json({ success: true });
});

app.delete('/api/users/:name', (req, res) => {
    let u = readDB(DB.users).filter(x => x.user !== req.params.name);
    writeDB(DB.users, u);
    res.json({ success: true });
});

app.post('/api/end-day', (req, res) => {
    writeDB(DB.logs, []);
    writeDB(DB.notes, {});
    res.json({ success: true });
});

// EXPORT
app.get('/api/export', (req, res) => {
    const logs = readDB(DB.logs);
    const ws = XLSX.utils.json_to_sheet(logs);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rapor");
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=rapor.xlsx');
    res.send(buf);
});

// FRONTEND (SADE)
app.get('/', (req, res) => {
    res.send(`
    <html>
    <body>
        <h2>Smart Minibar Çalışıyor ✅</h2>
        <p>API hazır.</p>
    </body>
    </html>
    `);
});

app.listen(PORT, '0.0.0.0', () => {
    console.log("Çalışıyor: http://localhost:" + PORT);
});
```
