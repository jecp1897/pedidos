const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./clientes.db');

db.serialize(() => {

  // Tabla de clientes
  db.run(`
    CREATE TABLE IF NOT EXISTS clientes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codigo TEXT UNIQUE,
      nombre TEXT,
      ciudad TEXT,
      activo INTEGER DEFAULT 1,
      usos INTEGER DEFAULT 0
    );
  `);

  // Tabla de productos
  db.run(`
    CREATE TABLE IF NOT EXISTS productos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codigo TEXT UNIQUE,
      descripcion TEXT,
      unidad TEXT,
      activo INTEGER DEFAULT 1,
      usos INTEGER DEFAULT 0
    );
  `);

  // Tabla de pedidos (cabecera)
  db.run(`
    CREATE TABLE IF NOT EXISTS pedidos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codigoCliente TEXT,
      codigoPedido TEXT,
      productos TEXT,   -- JSON con array de productos
      obs TEXT,
      fecha TEXT
    );
  `);

});

module.exports = db;
