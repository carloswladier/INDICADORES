// Entry point for Hostinger, PM2, cPanel, and Node.js hosting environments
const fs = require('fs');
const path = require('path');

// Load environment variables if dotenv is present
try {
  require('dotenv').config();
} catch (e) {}

const bundledServer = path.join(__dirname, 'dist', 'server.cjs');

if (fs.existsSync(bundledServer)) {
  console.log('Iniciando servidor Express a partir de dist/server.cjs...');
  require(bundledServer);
} else {
  console.error('Erro: arquivo dist/server.cjs não foi encontrado. Execute "npm run build" antes de iniciar.');
}
