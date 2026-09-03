const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Kết nối Database SQLite
const dbFile = path.join(__dirname, 'database.sqlite');
const db = new sqlite3.Database(dbFile, (err) => {
    if (err) console.error('Lỗi kết nối DB:', err.message);
    else console.log('Đã kết nối cơ sở dữ liệu SQLite.');
});

// Tạo bảng dữ liệu
db.run(`CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    route TEXT,
    name TEXT,
    status TEXT DEFAULT 'Chờ giao',
    price_policy TEXT,
    promo TEXT,
    amount REAL DEFAULT 0,
    note TEXT,
    ngo INTEGER DEFAULT 0, ngo_doi INTEGER DEFAULT 0, p_ngo REAL DEFAULT 8,
    thai INTEGER DEFAULT 0, thai_doi INTEGER DEFAULT 0, p_thai REAL DEFAULT 7,
    hong INTEGER DEFAULT 0, hong_doi INTEGER DEFAULT 0, p_hong REAL DEFAULT 7,
    dau INTEGER DEFAULT 0, dau_doi INTEGER DEFAULT 0, p_dau REAL DEFAULT 10,
    dua INTEGER DEFAULT 0, dua_doi INTEGER DEFAULT 0, p_dua REAL DEFAULT 10,
    chau INTEGER DEFAULT 0, chau_doi INTEGER DEFAULT 0, p_chau REAL DEFAULT 10,
    gao INTEGER DEFAULT 0, gao_doi INTEGER DEFAULT 0, p_gao REAL DEFAULT 8
)`);

// API lấy danh sách tuyến (suggestions)
app.get('/api/suggestions', (req, res) => {
    db.all("SELECT DISTINCT route FROM customers WHERE route IS NOT NULL AND route != ''", [], (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ routes: rows.map(r => r.route) });
    });
});

// API lấy danh sách khách hàng
app.get('/api/customers', (req, res) => {
    const route = req.query.route || 'ALL';
    let query = "SELECT * FROM customers";
    let params = [];
    if (route !== 'ALL') {
        query += " WHERE route = ?";
        params.push(route);
    }
    db.all(query, params, (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(rows);
    });
});

// API thêm hàng loạt khách hàng
app.post('/api/customers/batch', (req, res) => {
    const { route, names } = req.body;
    if (!names || !names.length) return res.status(400).json({ error: 'Danh sách trống' });

    const stmt = db.prepare("INSERT INTO customers (route, name) VALUES (?, ?)");
    db.serialize(() => {
        names.forEach(name => stmt.run(route || 'Chưa phân loại', name));
        stmt.finalize((err) => {
            if (err) return res.status(500).json({ error: err.message });
            res.json({ success: true });
        });
    });
});

// API cập nhật thông tin khách hàng
app.put('/api/customers/:id', (req, res) => {
    const id = req.params.id;
    const d = req.body;
    const query = `
        UPDATE customers SET name=?, status=?, price_policy=?, promo=?, amount=?, note=?,
        ngo=?, ngo_doi=?, p_ngo=?, thai=?, thai_doi=?, p_thai=?,
        hong=?, hong_doi=?, p_hong=?, dau=?, dau_doi=?, p_dau=?,
        dua=?, dua_doi=?, p_dua=?, chau=?, chau_doi=?, p_chau=?,
        gao=?, gao_doi=?, p_gao=? WHERE id=?
    `;
    const params = [
        d.name, d.status, d.price_policy, d.promo, d.amount, d.note,
        d.ngo, d.ngo_doi, d.p_ngo, d.thai, d.thai_doi, d.p_thai,
        d.hong, d.hong_doi, d.p_hong, d.dau, d.dau_doi, d.p_dau,
        d.dua, d.dua_doi, d.p_dua, d.chau, d.chau_doi, d.p_chau,
        d.gao, d.gao_doi, d.p_gao, id
    ];
    db.run(query, params, function(err) {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ success: true });
    });
});

// API xóa 1 khách hàng
app.delete('/api/customers/:id', (req, res) => {
    db.run("DELETE FROM customers WHERE id = ?", req.params.id, function(err) {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ success: true });
    });
});

// API xóa toàn bộ theo tuyến
app.delete('/api/customers/route/:route', (req, res) => {
    const route = req.params.route;
    let query = route === 'ALL' ? "DELETE FROM customers" : "DELETE FROM customers WHERE route = ?";
    let params = route === 'ALL' ? [] : [route];

    db.run(query, params, function(err) {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ success: true });
    });
});

// API Chốt sổ & xuất Excel (Đã căn chỉnh giao diện đẹp mắt)
app.post('/api/close-day', (req, res) => {
    db.all("SELECT * FROM customers", [], async (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Bao Cao Giao Hang');

        // Thiết lập tiêu đề chính to rõ ở dòng 1
        sheet.mergeCells('A1:V1');
        const titleCell = sheet.getCell('A1');
        titleCell.value = 'BÁO CÁO TỔNG KẾT GIAO HÀNG';
        titleCell.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFF' } };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1E293B' } }; // Màu nền xanh đậm slate-800
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        sheet.getRow(1).height = 35;

        // Dòng 2 trống tạo khoảng cách
        sheet.getRow(2).height = 10;

        // Định nghĩa các cột
        sheet.columns = [
            { header: 'ID', key: 'id', width: 8 },
            { header: 'Tuyến', key: 'route', width: 18 },
            { header: 'Khách hàng', key: 'name', width: 25 },
            { header: 'Trạng thái', key: 'status', width: 14 },
            { header: 'Chính sách', key: 'price_policy', width: 15 },
            { header: 'KM', key: 'promo', width: 12 },
            { header: 'Tổng tiền (k)', key: 'amount', width: 15 },
            { header: 'Ghi chú', key: 'note', width: 20 },
            { header: 'Ngô', key: 'ngo', width: 8 }, { header: 'Ngô Đổi', key: 'ngo_doi', width: 10 },
            { header: 'Thái', key: 'thai', width: 8 }, { header: 'Thái Đổi', key: 'thai_doi', width: 10 },
            { header: 'Hồng', key: 'hong', width: 8 }, { header: 'Hồng Đổi', key: 'hong_doi', width: 10 },
            { header: 'Đậu', key: 'dau', width: 8 }, { header: 'Đậu Đổi', key: 'dau_doi', width: 10 },
            { header: 'Dừa', key: 'dua', width: 8 }, { header: 'Dừa Đổi', key: 'dua_doi', width: 10 },
            { header: 'Châu', key: 'chau', width: 8 }, { header: 'Châu Đổi', key: 'chau_doi', width: 10 },
            { header: 'Gạo', key: 'gao', width: 8 }, { header: 'Gạo Đổi', key: 'gao_doi', width: 10 }
        ];

        // Đưa dòng Header vào hàng số 3
        const headerRow = sheet.insertRow(3, [
            'ID', 'Tuyến', 'Khách hàng', 'Trạng thái', 'Chính sách', 'KM', 'Tổng tiền (k)', 'Ghi chú',
            'Ngô', 'Ngô Đổi', 'Thái', 'Thái Đổi', 'Hồng', 'Hồng Đổi', 'Đậu', 'Đậu Đổi',
            'Dừa', 'Dừa Đổi', 'Châu', 'Châu Đổi', 'Gạo', 'Gạo Đổi'
        ]);
        
        sheet.getRow(3).height = 25;
        headerRow.eachCell((cell) => {
            cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '334155' } }; // Màu xanh slate-700
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
                top: { style: 'thin', color: { argb: 'CBD5E1' } },
                left: { style: 'thin', color: { argb: 'CBD5E1' } },
                bottom: { style: 'thin', color: { argb: 'CBD5E1' } },
                right: { style: 'thin', color: { argb: 'CBD5E1' } }
            };
        });

        // Thêm dữ liệu từng dòng
        rows.forEach((r, index) => {
            const rowData = [
                r.id, r.route, r.name, r.status, r.price_policy, r.promo, r.amount, r.note,
                r.ngo, r.ngo_doi, r.thai, r.thai_doi, r.hong, r.hong_doi, r.dau, r.dau_doi,
                r.dua, r.dua_doi, r.chau, r.chau_doi, r.gao, r.gao_doi
            ];
            const addedRow = sheet.addRow(rowData);
            addedRow.height = 20;

            // Kẻ bảng và định dạng từng ô dữ liệu
            addedRow.eachCell((cell, colNumber) => {
                cell.font = { name: 'Arial', size: 10 };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'E2E8F0' } },
                    left: { style: 'thin', color: { argb: 'E2E8F0' } },
                    bottom: { style: 'thin', color: { argb: 'E2E8F0' } },
                    right: { style: 'thin', color: { argb: 'E2E8F0' } }
                };

                // Canh lề thông minh theo cột
                if (colNumber === 1 || (colNumber >= 9 && colNumber <= 22)) {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' }; // Cột ID và Số lượng căn giữa
                } else if (colNumber === 7) {
                    cell.alignment = { horizontal: 'right', vertical: 'middle' }; // Cột tiền căn phải
                    cell.numFmt = '#,##0.0';
                } else {
                    cell.alignment = { horizontal: 'left', vertical: 'middle' }; // Tên, tuyến, ghi chú căn trái
                }

                // Tô màu nhẹ cho trạng thái "Đã giao" hoặc "Hủy" cho dễ nhìn
                if (colNumber === 4) {
                    if (r.status === 'Đã giao') {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DCFCE7' } }; // Xanh lá nhạt
                        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: '166534' } };
                    } else if (r.status === 'Hủy') {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE4E6' } }; // Đỏ nhạt
                        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: '9F1239' } };
                    }
                }
            });
        });

        // Xóa sạch database sau khi chốt sổ
        db.run("DELETE FROM customers", async () => {
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', 'attachment; filename=Bao_Cao_Giao_Hang.xlsx');
            await workbook.xlsx.write(res);
            res.end();
        });
    });
});app.listen(3000, () => {
    console.log('Server đang chay tai http://localhost:3000');
});