const mysql = require('mysql2/promise');

// Hostinger MySQL credentials from environment variables
const dbHost = process.env.DB_HOST || process.env.MYSQL_HOST || 'localhost';
const dbUser = process.env.DB_USER || process.env.MYSQL_USER || '';
const dbPassword = process.env.DB_PASSWORD || process.env.DB_PASS || process.env.MYSQL_PASSWORD || '';
const dbName = process.env.DB_NAME || process.env.DB_DATABASE || process.env.MYSQL_DATABASE || '';
const dbPort = Number(process.env.DB_PORT || process.env.MYSQL_PORT) || 3306;
const databaseUrl = process.env.DATABASE_URL || process.env.MYSQL_URI || process.env.MYSQL_URL;

let pool = null;

if (databaseUrl) {
  pool = mysql.createPool(databaseUrl);
  console.log("Hostinger MySQL configurado via DATABASE_URL");
} else if (dbUser && dbName) {
  pool = mysql.createPool({
    host: dbHost,
    user: dbUser,
    password: dbPassword,
    database: dbName,
    port: dbPort,
    connectTimeout: 8000,
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0,
    charset: 'utf8mb4'
  });
  console.log(`Hostinger MySQL configurado para ${dbUser}@${dbHost}:${dbPort}/${dbName}`);
} else {
  console.log("Variáveis de conexão MySQL da Hostinger (DB_HOST, DB_USER, DB_PASSWORD, DB_NAME) não configuradas.");
}

// Auto-create log_entries table if connected
if (pool) {
  pool.query(`
    CREATE TABLE IF NOT EXISTS log_entries (
      id INT AUTO_INCREMENT PRIMARY KEY,
      data VARCHAR(50),
      cidade VARCHAR(100),
      numero_chamado VARCHAR(100),
      incidente VARCHAR(255),
      descricao TEXT,
      status VARCHAR(50) DEFAULT 'Pendente',
      data_conclusao VARCHAR(50) NULL,
      created_at VARCHAR(100)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
  `).then(() => {
    console.log("Tabela 'log_entries' verificada/criada no MySQL da Hostinger com sucesso.");
  }).catch((err) => {
    console.error("Erro ao verificar/criar tabela 'log_entries' no MySQL:", err.message);
  });
}

module.exports = pool;
