const express = require('express');
const cors = require('cors');
const path = require('path');
const xlsx = require('xlsx');
const db = require('./database');
const nodemailer = require('nodemailer');

const app = express();
const PORT = 3000;

// -------------------------
// SERVIR FRONTEND
// -------------------------
app.use(express.static(path.join(__dirname, '../frontend')));

// -------------------------
// MIDDLEWARE
// -------------------------
app.use(cors());
app.use(express.json());

// -------------------------
// CONFIGURAR NODEMAILER
// -------------------------
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_PASS,
    pass: process.env.EMAIL_USER
  }
});

// -------------------------
// RUTA DE PRUEBA
// -------------------------
app.get('/', (req, res) => {
  res.send('Servidor de pedidos de aceite funcionando 🛢️');
});

// -------------------------
// LISTAR PRODUCTOS (BUSCADOR)
// -------------------------
app.get('/productos', (req, res) => {
  const q = req.query.q || '';

  db.all(
    `SELECT codigo, descripcion, unidad, activo
     FROM productos
     WHERE activo = 1
       AND (codigo LIKE ? OR descripcion LIKE ?)
     ORDER BY descripcion
     LIMIT 50`,
    [`%${q}%`, `%${q}%`],
    (err, rows) => {
      if (err) return res.status(500).json({ error: 'Error consultando productos' });
      res.json(rows);
    }
  );
});

// -------------------------
// IMPORTAR PRODUCTOS DESDE EXCEL
// -------------------------
app.post('/importar-productos', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'productos.xlsx');
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet);

    const stmt = db.prepare(`
      INSERT INTO productos (codigo, descripcion, unidad, activo)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(codigo) DO UPDATE SET
        descripcion = excluded.descripcion,
        unidad = excluded.unidad,
        activo = excluded.activo
    `);

    rows.forEach(row => {
      stmt.run(
        row.codigo,
        row.descripcion,
        row.unidad || '',
        row.activo ?? 1
      );
    });

    stmt.finalize();
    res.json({ success: true, count: rows.length });

  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false });
  }
});

// -------------------------
// IMPORTAR CLIENTES DESDE EXCEL
// -------------------------
app.post('/importar-clientes', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'clientes.xlsx');
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet);

    const stmt = db.prepare(`
      INSERT INTO clientes (codigo, nombre, ciudad, activo)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(codigo) DO UPDATE SET
        nombre = excluded.nombre,
        ciudad = excluded.ciudad,
        activo = excluded.activo
    `);

    rows.forEach(row => {
      stmt.run(
        row.codigo,
        row.nombre,
        row.ciudad || '',
        row.activo ?? 1
      );
    });

    stmt.finalize();
    res.json({ success: true, count: rows.length });

  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false });
  }
});

// -------------------------
// LOGIN CLIENTE
// -------------------------
app.post('/login', (req, res) => {
  const { codigo } = req.body;

  db.get(
    `SELECT * FROM clientes WHERE codigo = ? AND activo = 1`,
    [codigo],
    (err, row) => {
      if (err) return res.status(500).json({ success: false });
      if (!row) return res.json({ success: false });

      res.json({ success: true, cliente: row });
    }
  );
});

// -------------------------
// GUARDAR PEDIDO (VARIOS PRODUCTOS)
// -------------------------
app.post('/pedido', (req, res) => {
  const { codigo, productos, obs } = req.body;

  db.run(
    `INSERT INTO pedidos (codigo, productos, obs, fecha)
     VALUES (?, ?, ?, datetime('now'))`,
    [codigo, JSON.stringify(productos), obs],
    function (err) {
      if (err) return res.status(500).json({ success: false });

      // Email
      const lista = productos
        .map(p => `${p.codigo} - ${p.descripcion} x ${p.cantidad}`)
        .join('\n');

      const mailOptions = {
        from: 'edward1897@gmail.com',
        to: 'edward1897@gmail.com',
        subject: `Pedido #${this.lastID} de ${codigo}`,
        text: `
Nuevo pedido recibido:

Número de pedido: ${this.lastID}
Código cliente: ${codigo}

Productos:
${lista}

Observaciones: ${obs || 'Ninguna'}

Fecha: ${new Date().toLocaleString()}
        `
      };

      transporter.sendMail(mailOptions, () => {});

      res.json({ success: true, id: this.lastID });
    }
  );
});

// -------------------------
// ARRANCAR SERVIDOR
// -------------------------
app.listen(PORT, () => {
  console.log(`Servidor escuchando en http://localhost:${PORT}`);
});
