const express = require('express');
const cors = require('cors');
const path = require('path');
const xlsx = require('xlsx');
const db = require('./database');
const nodemailer = require('nodemailer');
const PDFDocument = require('pdfkit');

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
  host: process.env.SMTP_HOST,
  port: process.env.SMTP_PORT,
  secure: false,
  requireTLS: true,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
  }
});

// -------------------------
// RUTA DE PRUEBA
// -------------------------
app.get('/', (req, res) => {
  res.redirect('/login.html');
});

// -------------------------
// LISTAR PRODUCTOS
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
// BUSCAR CLIENTES
// -------------------------
app.get('/clientes', (req, res) => {
  const q = req.query.q || '';

  db.all(
    `SELECT codigo, nombre, privincia, activo
     FROM clientes
     WHERE activo = 1
       AND (codigo LIKE ? OR nombre LIKE ? OR privincia LIKE ?)
     ORDER BY nombre
     LIMIT 50`,
    [`%${q}%`, `%${q}%`, `%${q}%`],
    (err, rows) => {
      if (err) return res.status(500).json({ error: 'Error consultando clientes' });
      res.json(rows);
    }
  );
});

// -------------------------
// IMPORTAR PRODUCTOS DESDE EXCEL
// -------------------------
app.get('/importar-productos', (req, res) => {
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
app.get('/importar-clientes', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'clientes.xlsx');
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet);

    const stmt = db.prepare(`
      INSERT INTO clientes (codigo, nombre, privincia, activo)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(codigo) DO UPDATE SET
        nombre = excluded.nombre,
        privincia = excluded.privincia,
        activo = excluded.activo
    `);

    rows.forEach(row => {
      stmt.run(
        row.codigo,
        row.nombre,
        row.privincia || '',
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
// GUARDAR PEDIDO
// -------------------------
app.post('/pedido', (req, res) => {
  const { codigo, productos, obs } = req.body;

  db.run(
    `INSERT INTO pedidos (codigo, productos, obs, fecha)
     VALUES (?, ?, ?, datetime('now'))`,
    [codigo, JSON.stringify(productos), obs],
    function (err) {
      if (err) return res.status(500).json({ success: false });

      const numeroPedido = `AU-P${String(this.lastID).padStart(7, '0')}`;

      const lista = productos
        .map(p => `${p.codigo} - ${p.descripcion} x ${p.cantidad} ${p.tipo === "cajas" ? "cajas" : "unidades sueltas"}`)
        .join('\n');

      const mailOptions = {
        from: process.env.EMAIL_USER,
        to: process.env.EMAIL_USER,
        subject: `Pedido ${numeroPedido} de ${codigo}`,
        text: `
Nuevo pedido recibido:

Número de pedido: ${numeroPedido}
Código cliente: ${codigo}

Productos:
${lista}

Observaciones: ${obs || 'Ninguna'}

Fecha: ${new Date().toLocaleString()}
        `
      };

      transporter.sendMail(mailOptions, (err) => {
        if (err) console.error("Error enviando email:", err);
      });

      res.json({ success: true, id: this.lastID, numeroPedido });
    }
  );
});

// -------------------------
// GENERAR PDF DEL PEDIDO
// -------------------------
app.post('/pdf', (req, res) => {
  const { id, numeroPedido, codigo, productos, obs } = req.body;

  const doc = new PDFDocument({ margin: 40 });

  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename=${numeroPedido}.pdf`);

  doc.pipe(res);

  // LOGO (ajusta la ruta si lo pones en otro sitio)
  try {
    doc.image(path.join(__dirname, 'logo.png'), 40, 40, { width: 120 });
  } catch (e) {
    console.log("Logo no encontrado, continuando sin logo.");
  }

  // CABECERA
  doc.fontSize(22).text(`Pedido ${numeroPedido}`, 200, 40);
  doc.moveDown(2);

  const fecha = new Date().toLocaleString();

  doc.fontSize(12).text(`Fecha: ${fecha}`);
  doc.moveDown();

  // DATOS DEL CLIENTE
  doc.fontSize(14).text("Datos del cliente", { underline: true });
  doc.moveDown(0.5);

  // Recuperar datos completos del cliente desde la BD
  db.get(
    `SELECT * FROM clientes WHERE codigo = ?`,
    [codigo],
    (err, cliente) => {

      doc.fontSize(12).text(`Código: ${cliente.codigo}`);
      doc.text(`Nombre: ${cliente.nombre}`);
      doc.text(`Provincia: ${cliente.privincia}`);
      doc.moveDown(2);

      // TABLA DE PRODUCTOS
      doc.fontSize(14).text("Productos", { underline: true });
      doc.moveDown(1);

      // Encabezados de tabla
      const tableTop = doc.y;
      const col1 = 40;
      const col2 = 140;
      const col3 = 360;
      const col4 = 430;

      doc.fontSize(12).text("Código", col1, tableTop);
      doc.text("Descripción", col2, tableTop);
      doc.text("Cantidad", col3, tableTop);
      doc.text("Tipo", col4, tableTop);

      doc.moveDown(0.5);

      // Línea debajo del encabezado
      doc.moveTo(40, doc.y).lineTo(550, doc.y).stroke();

      doc.moveDown(0.5);

      // Filas
      productos.forEach(p => {
        const tipoTexto = p.tipo === "cajas" ? "Cajas" : "Unidades sueltas";

        doc.text(p.codigo, col1, doc.y);
        doc.text(p.descripcion, col2, doc.y);
        doc.text(String(p.cantidad), col3, doc.y);
        doc.text(tipoTexto, col4, doc.y);

        doc.moveDown(0.5);
      });

      doc.moveDown(2);

      // OBSERVACIONES
      doc.fontSize(14).text("Observaciones", { underline: true });
      doc.moveDown(0.5);

      doc.fontSize(12)
        .rect(40, doc.y, 500, 60)
        .stroke()
        .text(obs || "Ninguna", 50, doc.y + 10, { width: 480 });

      doc.moveDown(4);

      // PIE DE PÁGINA
      doc.fontSize(10)
        .fillColor("#555")
        .text("Aceites Únicos · Documento generado automáticamente", 40, 760, {
          align: "center"
        });

      doc.end();
    }
  );
});

// -------------------------
// PROBAR ENVÍO DE EMAIL
// -------------------------
app.get('/test-email', async (req, res) => {
  try {
    await transporter.sendMail({
      from: process.env.EMAIL_USER,
      to: process.env.EMAIL_USER,
      subject: "Prueba de envío desde AuPedidos",
      text: "Este es un email de prueba para verificar la configuración SMTP."
    });

    res.json({ success: true, message: "Email enviado correctamente" });

  } catch (err) {
    console.error("Error enviando email:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// -------------------------
// ARRANCAR SERVIDOR
// -------------------------
app.listen(PORT, () => {
  console.log(`Servidor escuchando en http://localhost:${PORT}`);
});
