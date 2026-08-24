import express from "express";
import path from "path";
import mysql, { Pool } from "mysql2/promise";
import dotenv from "dotenv";
import cors from "cors";

dotenv.config();

// Pool cache to reuse connections where possible
const poolsCache = new Map<string, Pool>();

interface DbConnectionParams {
  host?: string;
  user?: string;
  password?: string;
  database?: string;
  port?: number;
  uri?: string;
}

function resolveDbParams(req?: express.Request): DbConnectionParams {
  const headerHost = req?.headers['x-db-host'] as string | undefined;
  const headerUser = req?.headers['x-db-user'] as string | undefined;
  const headerPassword = req?.headers['x-db-password'] as string | undefined;
  const headerDatabase = req?.headers['x-db-name'] as string | undefined;
  const headerPort = req?.headers['x-db-port'] ? Number(req.headers['x-db-port']) : undefined;
  const headerUri = req?.headers['x-db-uri'] as string | undefined;

  const uri = headerUri || process.env.DATABASE_URL || process.env.MYSQL_URI || process.env.MYSQL_URL;
  const host = headerHost || process.env.DB_HOST || process.env.MYSQL_HOST || 'localhost';
  const user = headerUser || process.env.DB_USER || process.env.MYSQL_USER || '';
  const password = headerPassword || process.env.DB_PASSWORD || process.env.DB_PASS || process.env.MYSQL_PASSWORD || '';
  const database = headerDatabase || process.env.DB_NAME || process.env.DB_DATABASE || process.env.MYSQL_DATABASE || '';
  const port = headerPort || Number(process.env.DB_PORT || process.env.MYSQL_PORT) || 3306;

  return { host, user, password, database, port, uri };
}

async function getHostingerDbPool(req?: express.Request): Promise<Pool> {
  const params = resolveDbParams(req);

  const cacheKey = params.uri || `${params.user}:${params.host}:${params.port}/${params.database}`;

  if (!params.uri && (!params.user || !params.database)) {
    throw new Error("Credenciais do Banco de Dados Hostinger (MySQL) incompletas. Configure DB_HOST, DB_USER, DB_PASSWORD e DB_NAME.");
  }

  let pool = poolsCache.get(cacheKey);
  if (!pool) {
    if (params.uri) {
      pool = mysql.createPool(params.uri);
    } else {
      pool = mysql.createPool({
        host: params.host,
        user: params.user,
        password: params.password,
        database: params.database,
        port: params.port,
        connectTimeout: 5000,
        waitForConnections: true,
        connectionLimit: 10,
        queueLimit: 0,
        charset: 'utf8mb4',
        ssl: params.host === 'localhost' || params.host === '127.0.0.1' ? undefined : { rejectUnauthorized: false }
      });
    }
    poolsCache.set(cacheKey, pool);
  }

  // Ensure table exists
  await ensureLogTable(pool);

  return pool;
}

async function ensureLogTable(pool: Pool) {
  const createTableSql = `
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
  `;
  await pool.query(createTableSql);
}

async function startServer() {
  console.log("Starting server initialization with Hostinger MySQL...");
  const app = express();
  const PORT = Number(process.env.PORT) || 3000;

  app.use(cors());
  app.use(express.json());

  // Request logger helper
  app.use((req, res, next) => {
    res.setHeader('X-Backend-Server', 'Express-Hostinger-MySQL');
    if (req.url.startsWith('/api')) {
      console.log(`[API REQUEST] ${new Date().toISOString()} - ${req.method} ${req.url}`);
    }
    next();
  });

  // API Routes
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok", mode: process.env.NODE_ENV, type: "Hostinger MySQL", time: new Date().toISOString() });
  });

  // DB Test endpoint executing SELECT 1 exclusively from backend environment variables
  app.get("/api/db-test", async (req, res) => {
    const host = process.env.DB_HOST || 'localhost';
    const port = Number(process.env.DB_PORT) || 3306;
    const database = process.env.DB_NAME || '';
    const user = process.env.DB_USER || '';
    const hasPassword = Boolean(process.env.DB_PASSWORD || process.env.DB_PASS);

    if (!user || !database) {
      return res.status(400).json({
        success: false,
        message: "Variáveis de ambiente do MySQL não configuradas no servidor.",
        details: "Defina DB_HOST, DB_PORT, DB_NAME, DB_USER e DB_PASSWORD no ambiente do Node.js.",
        config: { host, port, database, user, hasPassword }
      });
    }

    try {
      const pool = await getHostingerDbPool(req);
      const [rows] = await pool.query("SELECT 1 AS connected");

      return res.json({
        success: true,
        message: "Conexão com o banco MySQL da Hostinger estabelecida com sucesso!",
        result: rows,
        config: { host, port, database, user, hasPassword }
      });
    } catch (err: any) {
      console.error("[DB TEST ERROR]", err);

      let hint = "Verifique as configurações de rede e credenciais.";
      if (err.code === 'ETIMEDOUT' || err.message?.includes('ETIMEDOUT')) {
        hint = "Tempo limite esgotado (ETIMEDOUT). Conexões externas diretas na porta 3306 costumam ser bloqueadas. Se a aplicação estiver hospedada na Hostinger, use DB_HOST=localhost.";
      } else if (err.code === 'ER_ACCESS_DENIED_ERROR' || err.message?.includes('Access denied')) {
        hint = "Acesso negado. Verifique se DB_USER e DB_PASSWORD correspondem ao usuário do MySQL criado no hPanel.";
      } else if (err.code === 'ER_BAD_DB_ERROR' || err.message?.includes('Unknown database')) {
        hint = `O banco de dados '${database}' não foi encontrado no MySQL da Hostinger.`;
      } else if (err.code === 'ECONNREFUSED' || err.code === 'ENOTFOUND') {
        hint = `Não foi possível alcançar o host MySQL '${host}'. Verifique o endereço ou use 'localhost'.`;
      }

      return res.status(500).json({
        success: false,
        message: "Falha ao conectar ao banco MySQL da Hostinger.",
        error: err.message,
        code: err.code || 'DB_CONNECTION_ERROR',
        hint,
        config: { host, port, database, user, hasPassword }
      });
    }
  });

  app.get("/api/db-status", async (req, res) => {
    console.log("Handling /api/db-status (Hostinger MySQL)");
    const params = resolveDbParams(req);
    const isConfigured = !!(params.uri || (params.user && params.database));

    if (!isConfigured) {
      return res.json({ 
        configured: false, 
        type: 'Hostinger MySQL', 
        error: 'Configuração incompleta. Preencha DB_HOST, DB_USER, DB_PASSWORD e DB_NAME nas variáveis da Hostinger ou no app.' 
      });
    }

    try {
      const pool = await getHostingerDbPool(req);
      const [rows] = await pool.query("SELECT COUNT(*) as count FROM log_entries");
      const count = (rows as any)?.[0]?.count || 0;

      res.json({ 
        configured: true, 
        type: 'Hostinger MySQL', 
        status: 'connected', 
        count 
      });
    } catch (err: any) {
      console.error("Hostinger MySQL connection check failed:", err.message);

      let friendlyError = err.message;
      if (err.code === 'ETIMEDOUT' || err.message.includes('ETIMEDOUT')) {
        if (params.host === 'localhost' || params.host === '127.0.0.1') {
          friendlyError = "O host está como 'localhost'. Para conectar a partir deste ambiente de Preview, clique no botão 'BANCO HOSTINGER' e informe o IP/Host público do servidor MySQL da Hostinger (ex: 147.79.84.15 ou seu domínio).";
        } else {
          friendlyError = `Tempo limite esgotado ao conectar a '${params.host}'. Verifique se o IP ou Host MySQL está correto e se a porta 3306 não está bloqueada pelo provedor.`;
        }
      } else if (err.code === 'ER_ACCESS_DENIED_ERROR' || err.message.includes('Access denied')) {
        friendlyError = "Falha de autenticação MySQL. Verifique o usuário (DB_USER) e a senha (DB_PASSWORD) no painel da Hostinger.";
      } else if (err.code === 'ER_BAD_DB_ERROR' || err.message.includes('Unknown database')) {
        friendlyError = `O banco de dados '${params.database}' não foi encontrado. Verifique se o nome confere exatamente com o painel da Hostinger.`;
      } else if (err.code === 'ECONNREFUSED' || err.code === 'ENOTFOUND') {
        friendlyError = `Não foi possível conectar ao host MySQL '${params.host}'. Verifique o endereço no painel da Hostinger.`;
      }

      res.json({ 
        configured: true, 
        type: 'Hostinger MySQL', 
        status: 'error', 
        error: err.message,
        details: friendlyError
      });
    }
  });

  app.get("/api/logs", async (req, res) => {
    try {
      const pool = await getHostingerDbPool(req);
      const [rows] = await pool.query("SELECT * FROM log_entries ORDER BY data DESC, id DESC");
      
      const formatted = (rows as any[]).map(row => ({
        id: String(row.id),
        data: row.data || '',
        cidade: row.cidade || '',
        numero_chamado: row.numero_chamado || '',
        incidente: row.incidente || '',
        descricao: row.descricao || '',
        status: row.status || 'Pendente',
        data_conclusao: row.data_conclusao || null,
        created_at: row.created_at || ''
      }));

      res.json(formatted);
    } catch (err: any) {
      console.error("GET /api/logs catch error:", err);
      res.status(500).json({ error: err.message || "Falha ao buscar logs no banco MySQL da Hostinger" });
    }
  });

  app.post("/api/logs", async (req, res) => {
    try {
      const pool = await getHostingerDbPool(req);
      const { data, cidade, numero_chamado, incidente, descricao, status, data_conclusao, created_at } = req.body;
      
      const creationDate = created_at || new Date().toISOString();
      const currentStatus = status || 'Pendente';

      const [result]: any = await pool.query(
        `INSERT INTO log_entries (data, cidade, numero_chamado, incidente, descricao, status, data_conclusao, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        [data || '', cidade || '', numero_chamado || '', incidente || '', descricao || '', currentStatus, data_conclusao || null, creationDate]
      );
      
      res.status(201).json({
        id: String(result.insertId),
        data,
        cidade,
        numero_chamado,
        incidente,
        descricao,
        status: currentStatus,
        data_conclusao: data_conclusao || null,
        created_at: creationDate
      });
    } catch (err: any) {
      console.error("POST /api/logs catch error:", err);
      res.status(500).json({ error: err.message || "Falha ao salvar registro no banco MySQL da Hostinger" });
    }
  });

  app.put("/api/logs/:id", async (req, res) => {
    const { id } = req.params;
    try {
      const pool = await getHostingerDbPool(req);
      const { data, cidade, numero_chamado, incidente, descricao, status, data_conclusao } = req.body;

      await pool.query(
        `UPDATE log_entries 
         SET data = ?, cidade = ?, numero_chamado = ?, incidente = ?, descricao = ?, status = ?, data_conclusao = ?
         WHERE id = ?`,
        [data, cidade, numero_chamado, incidente, descricao, status, data_conclusao || null, id]
      );

      res.json({
        id,
        data,
        cidade,
        numero_chamado,
        incidente,
        descricao,
        status,
        data_conclusao: data_conclusao || null
      });
    } catch (err: any) {
      console.error("PUT /api/logs/:id catch error:", err);
      res.status(500).json({ error: err.message || "Falha ao atualizar registro no banco MySQL da Hostinger" });
    }
  });

  app.delete("/api/logs/:id", async (req, res) => {
    const { id } = req.params;
    try {
      const pool = await getHostingerDbPool(req);
      await pool.query("DELETE FROM log_entries WHERE id = ?", [id]);
      res.json({ message: "Registro excluído com sucesso do banco MySQL da Hostinger" });
    } catch (err: any) {
      console.error("DELETE /api/logs/:id error:", err);
      res.status(500).json({ error: err.message || "Falha ao excluir registro do banco MySQL da Hostinger" });
    }
  });

  // Proxy GitHub Raw Files (avoids browser CORS or restrictive network constraints)
  app.get("/api/proxy-github", async (req, res) => {
    const rawUrl = req.query.url as string;
    if (!rawUrl) {
      return res.status(400).json({ error: "Parâmetro 'url' é obrigatório" });
    }
    try {
      let target = rawUrl.trim();
      if (target.includes("github.com") && !target.includes("raw.githubusercontent.com")) {
        target = target.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/").replace("/raw/", "/");
      }
      target = target.replace("/refs/heads/", "/").replace(/ /g, "%20");
      
      const response = await fetch(target);
      if (!response.ok) {
        return res.status(response.status).json({ error: `GitHub retornou status ${response.status}` });
      }
      const arrayBuffer = await response.arrayBuffer();
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.send(Buffer.from(arrayBuffer));
    } catch (err: any) {
      console.error("Proxy GitHub error:", err);
      res.status(500).json({ error: err.message || "Erro ao baixar arquivo do GitHub via proxy" });
    }
  });

  // PROTECT: Any other /api/* routes that are not defined should return JSON 404
  app.all("/api/*", (req, res) => {
    console.warn(`[404] API Route not found: ${req.method} ${req.url}`);
    res.status(404).json({ error: `Rota API ${req.url} não encontrada` });
  });

  // Handle SPA
  if (process.env.NODE_ENV !== "production") {
    console.log("Initializing Vite dev middleware...");
    try {
      const { createServer: createViteServer } = await import("vite");
      const vite = await createViteServer({
        server: { middlewareMode: true },
        appType: "spa",
      });
      app.use(vite.middlewares);
      console.log("Vite dev middleware initialized.");
    } catch (err) {
      console.error("Failed to initialize Vite middleware:", err);
      app.use(express.static(path.join(process.cwd(), "dist")));
    }
  } else {
    console.log("Running in production mode, serving dist...");
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server is listening on 0.0.0.0:${PORT} with Hostinger MySQL backend`);
  });
}

startServer().catch(err => {
  console.error("FATAL ERROR DURING SERVER STARTUP:", err);
  process.exit(1);
});
