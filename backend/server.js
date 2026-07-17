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
