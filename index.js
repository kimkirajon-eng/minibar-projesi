```js
var express = require('express');
var bodyParser = require('body-parser');
var fs = require('fs');
var path = require('path');
var XLSX = require('xlsx');

var app = express();
var PORT = process.env.PORT || 3000;

var DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);

var DB = {
    logs: path.join(DATA_DIR, 'logs.json'),
    users: path.join(DATA_DIR, 'users.json'),
    products: path.join(DATA_DIR, 'products.json'),
    structure: path.join(DATA_DIR, 'structure.json'),
    notes: path.join(DATA_DIR, 'notes.json'),
    stokLogs: path.join(DATA_DIR, 'stokLogs.json')
};

app.use(bodyParser.json());

function readDB(f) {
    if (!fs.existsSync(f)) {
        if (f === DB.users) return [{ user: 'hakkı', pass: '2125', role: 'admin' }];
        if (f === DB.products) return [{ name: 'Su', stok: 100 }, { name: 'Kola', stok: 50 }];
        if (f === DB.stokLogs) return [];
        if (f === DB.structure) return [];
        return f === DB.notes ? {} : [];
    }
    try {
        return JSON.parse(fs.readFileSync(f));
    } catch (e) {
        return f === DB.notes ? {} : [];
    }
}

function writeDB(f, d) {
    fs.writeFileSync(f, JSON.stringify(d, null, 2));
}

// LOGIN
app.post('/api/login', function (req, res) {
    var users = readDB(DB.users);
    var u = null;
    for (var i = 0; i < users.length; i++) {
        if (users[i].user === req.body.user && users[i].pass === req.body.pass) {
            u = users[i];
            break;
        }
    }
    if (u) res.json({ success: true, role: u.role, user: u.user });
    else res.json({ success: false });
});

// LOGS
app.get('/api/logs', function (req, res) {
    res.json(readDB(DB.logs));
});

app.post('/api/logs', function (req, res) {
    var logs = readDB(DB.logs);

    logs.unshift({
        room: req.body.room,
        status: req.body.status,
        details: req.body.details,
        items: req.body.items,
        staff: req.body.staff,
        date: new Date().toLocaleDateString('tr-TR'),
        endTime: new Date().toLocaleTimeString('tr-TR')
    });

    writeDB(DB.logs, logs);

    if (req.body.items && Array.isArray(req.body.items)) {
        var products = readDB(DB.products);

        for (var i = 0; i < req.body.items.length; i++) {
            var item = req.body.items[i];
            for (var j = 0; j < products.length; j++) {
                if (products[j].name === item.name) {
                    var count = Number(item.count) || 0;
                    products[j].stok = (products[j].stok || 0) - count;
                }
            }
        }

        writeDB(DB.products, products);
    }

    var notes = readDB(DB.notes);
    delete notes[req.body.room];
    writeDB(DB.notes, notes);

    res.json({ success: true });
});

// STOK
app.post('/api/stok-islem', function (req, res) {
    var products = readDB(DB.products);
    var stokLogs = readDB(DB.stokLogs);

    var productName = req.body.productName;
    var amount = req.body.amount;
    var source = req.body.source;
    var type = req.body.type;

    var p = null;
    for (var i = 0; i < products.length; i++) {
        if (products[i].name === productName) {
            p = products[i];
            break;
        }
    }

    if (p) {
        var islemMiktari = Number(amount) || 0;

        if (type === 'giris') {
            p.stok = (p.stok || 0) + islemMiktari;
        } else {
            p.stok = (p.stok || 0) - islemMiktari;
        }

        stokLogs.unshift({
            productName: productName,
            amount: type === 'giris' ? '+' + islemMiktari : '-' + islemMiktari,
            source: source,
            date: new Date().toLocaleDateString('tr-TR'),
            time: new Date().toLocaleTimeString('tr-TR')
        });

        writeDB(DB.products, products);
        writeDB(DB.stokLogs, stokLogs);

        res.json({ success: true });
    } else {
        res.json({ success: false });
    }
});

app.get('/api/stok-logs', function (req, res) {
    res.json(readDB(DB.stokLogs));
});

app.get('/api/products', function (req, res) {
    res.json(readDB(DB.products));
});

app.post('/api/products', function (req, res) {
    writeDB(DB.products, req.body);
    res.json({ success: true });
});

app.get('/api/structure', function (req, res) {
    res.json(readDB(DB.structure));
});

app.post('/api/structure', function (req, res) {
    writeDB(DB.structure, req.body);
    res.json({ success: true });
});

app.get('/api/notes', function (req, res) {
    res.json(readDB(DB.notes));
});

app.post('/api/notes', function (req, res) {
    var notes = readDB(DB.notes);
    notes[req.body.room] = req.body.note;
    writeDB(DB.notes, notes);
    res.json({ success: true });
});

app.get('/api/users', function (req, res) {
    res.json(readDB(DB.users));
});

app.post('/api/users', function (req, res) {
    var u = readDB(DB.users);
    u.push({ user: req.body.user, pass: req.body.pass, role: 'staff' });
    writeDB(DB.users, u);
    res.json({ success: true });
});

app.delete('/api/users/:name', function (req, res) {
    var users = readDB(DB.users);
    var newUsers = [];
    for (var i = 0; i < users.length; i++) {
        if (users[i].user !== req.params.name) {
            newUsers.push(users[i]);
        }
    }
    writeDB(DB.users, newUsers);
    res.json({ success: true });
});

app.post('/api/end-day', function (req, res) {
    writeDB(DB.logs, []);
    writeDB(DB.notes, {});
    res.json({ success: true });
});

// EXPORT
app.get('/api/export', function (req, res) {
    var logs = readDB(DB.logs);
    var ws = XLSX.utils.json_to_sheet(logs);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rapor");

    var buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=rapor.xlsx');
    res.send(buf);
});

app.get('/', function (req, res) {
    res.send('<h2>Smart Minibar Çalışıyor ✅</h2>');
});

app.listen(PORT, '0.0.0.0', function () {
    console.log('Çalışıyor: http://localhost:' + PORT);
});
```
