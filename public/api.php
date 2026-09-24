<?php
/**
 * Conector MySQL Hostinger para o Dashboard de Indicadores
 * Este arquivo conecta diretamente ao MySQL local da Hostinger.
 */

ini_set('display_errors', 0);
error_reporting(E_ALL);

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: GET, POST, PUT, DELETE, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type, Authorization, x-db-host, x-db-user, x-db-password, x-db-name, x-db-port, x-db-uri");
header("Content-Type: application/json; charset=UTF-8");

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit();
}

// Configurações padrão do Banco de Dados na Hostinger
$db_host = "localhost";
$db_user = "u688072783_CW_INDICADORES";
$db_pass = "Cwrocha2026";
$db_name = "u688072783_INDICADORES";

// Permite sobrescrever via Headers HTTP se informados
if (!empty($_SERVER['HTTP_X_DB_USER'])) $db_user = $_SERVER['HTTP_X_DB_USER'];
if (!empty($_SERVER['HTTP_X_DB_PASSWORD'])) $db_pass = $_SERVER['HTTP_X_DB_PASSWORD'];
if (!empty($_SERVER['HTTP_X_DB_NAME'])) $db_name = $_SERVER['HTTP_X_DB_NAME'];
if (!empty($_SERVER['HTTP_X_DB_HOST'])) $db_host = $_SERVER['HTTP_X_DB_HOST'];

// Detecta rota / endpoint
$uri = $_SERVER['REQUEST_URI'] ?? '';
$endpoint = $_GET['endpoint'] ?? '';
$action = $_GET['action'] ?? '';
$id = isset($_GET['id']) ? intval($_GET['id']) : 0;

if (empty($endpoint)) {
    if (strpos($uri, '/api/proxy-github') !== false) {
        $endpoint = 'proxy-github';
    } elseif (strpos($uri, '/api/logs/') !== false) {
        $parts = explode('/api/logs/', $uri);
        if (!empty($parts[1])) {
            $id = intval(explode('?', $parts[1])[0]);
            $endpoint = 'logs';
        }
    } elseif (strpos($uri, '/api/logs') !== false) {
        $endpoint = 'logs';
    } elseif (strpos($uri, '/api/db-test') !== false) {
        $endpoint = 'db-test';
    } elseif (strpos($uri, '/api/db-status') !== false || strpos($uri, '/api/health') !== false) {
        $endpoint = 'status';
    }
} else {
    if (strpos($endpoint, 'logs/') === 0) {
        $id = intval(substr($endpoint, 5));
        $endpoint = 'logs';
    } elseif (strpos($endpoint, 'proxy-github') === 0) {
        $endpoint = 'proxy-github';
    }
}

// Endpoint de Proxy para o GitHub (não requer conexão MySQL)
if ($endpoint === 'proxy-github' || $action === 'proxy-github') {
    $rawUrl = $_GET['url'] ?? '';
    if (empty($rawUrl)) {
        http_response_code(400);
        header("Content-Type: application/json; charset=UTF-8");
        echo json_encode(["error" => "Parâmetro 'url' é obrigatório"]);
        exit();
    }
    
    // Normaliza URL do GitHub
    $target = trim($rawUrl);
    // Reescreve repositórios antigos para o atual INDICADORES2
    $target = str_replace(
        ['/carloswladier/DASH_AT1_G1/', '/carloswladier/INDICADORES_MANUT/', '/carloswladier/INDICADORES/'],
        '/carloswladier/INDICADORES2/',
        $target
    );

    if (strpos($target, 'github.com') !== false && strpos($target, 'raw.githubusercontent.com') === false) {
        $target = str_replace(['github.com', '/blob/', '/raw/'], ['raw.githubusercontent.com', '/', '/'], $target);
    }
    $target = str_replace('/refs/heads/', '/', $target);
    $target = str_replace(' ', '%20', $target);

    // Se o arquivo existir localmente na raiz da hospedagem (repositório clonado na Hostinger), serve diretamente
    $localName = basename(urldecode(parse_url($target, PHP_URL_PATH)));
    if (!empty($localName) && file_exists(__DIR__ . '/' . $localName) && filesize(__DIR__ . '/' . $localName) > 1000) {
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Length: " . filesize(__DIR__ . '/' . $localName));
        header("Content-Disposition: attachment; filename=\"" . $localName . "\"");
        readfile(__DIR__ . '/' . $localName);
        exit();
    }

    // Validação de segurança: apenas URLs do raw.githubusercontent.com ou github.com
    if (strpos($target, 'raw.githubusercontent.com') === false) {
        http_response_code(403);
        header("Content-Type: application/json; charset=UTF-8");
        echo json_encode(["error" => "Origem não permitida"]);
        exit();
    }

    $ch = curl_init();
    curl_setopt($ch, CURLOPT_URL, $target);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
    curl_setopt($ch, CURLOPT_USERAGENT, 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)');
    curl_setopt($ch, CURLOPT_TIMEOUT, 60);
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
    $content = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $errorMsg = curl_error($ch);
    curl_close($ch);

    if ($httpCode >= 200 && $httpCode < 300 && !empty($content)) {
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Length: " . strlen($content));
        header("Content-Disposition: attachment; filename=\"excel_data.xlsx\"");
        echo $content;
        exit();
    } else {
        http_response_code($httpCode >= 400 ? $httpCode : 502);
        header("Content-Type: application/json; charset=UTF-8");
        echo json_encode([
            "error" => "Falha ao baixar arquivo do GitHub via proxy",
            "httpCode" => $httpCode,
            "details" => $errorMsg ?: "Código HTTP $httpCode retornado pelo GitHub",
            "url" => $target
        ]);
        exit();
    }
}

// Conexão com o banco de dados via PDO
try {
    $pdo = new PDO("mysql:host={$db_host};dbname={$db_name};charset=utf8mb4", $db_user, $db_pass, [
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        PDO::ATTR_TIMEOUT => 5
    ]);

    // Cria a tabela automaticamente se não existir
    $pdo->exec("CREATE TABLE IF NOT EXISTS log_entries (
        id INT AUTO_INCREMENT PRIMARY KEY,
        data VARCHAR(50),
        cidade VARCHAR(100),
        numero_chamado VARCHAR(100),
        incidente VARCHAR(255),
        descricao TEXT,
        status VARCHAR(50) DEFAULT 'Pendente',
        data_conclusao VARCHAR(50) NULL,
        created_at VARCHAR(100)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;");

} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode([
        "success" => false,
        "configured" => false,
        "type" => "Hostinger MySQL (PHP Bridge)",
        "status" => "error",
        "error" => $e->getMessage(),
        "details" => "Erro ao conectar ao MySQL na Hostinger. Verifique se o banco de dados e usuário existem no hPanel.",
        "config" => [
            "host" => $db_host,
            "user" => $db_user,
            "database" => $db_name
        ]
    ]);
    exit();
}

$method = $_SERVER['REQUEST_METHOD'];

// Rota de Teste de Conexão (SELECT 1)
if ($endpoint === 'db-test' || $action === 'test') {
    try {
        $stmt = $pdo->query("SELECT 1 AS connected");
        $res = $stmt->fetch();
        echo json_encode([
            "success" => true,
            "message" => "Conexão com o banco MySQL da Hostinger estabelecida com sucesso!",
            "result" => $res,
            "config" => [
                "host" => $db_host,
                "user" => $db_user,
                "database" => $db_name
            ]
        ]);
    } catch (Exception $e) {
        http_response_code(500);
        echo json_encode([
            "success" => false,
            "message" => "Falha na query SELECT 1",
            "error" => $e->getMessage()
        ]);
    }
    exit();
}

// Rota de Status / Verificação
if ($endpoint === 'status' || $endpoint === 'db-status' || $endpoint === 'health' || $action === 'status') {
    try {
        $stmt = $pdo->query("SELECT COUNT(*) as count FROM log_entries");
        $res = $stmt->fetch();
        echo json_encode([
            "success" => true,
            "configured" => true,
            "type" => "Hostinger MySQL",
            "status" => "connected",
            "count" => intval($res['count'] ?? 0)
        ]);
    } catch (Exception $e) {
        echo json_encode([
            "success" => false,
            "status" => "error",
            "error" => $e->getMessage()
        ]);
    }
    exit();
}

// GET - Listar registros
if ($method === 'GET') {
    if (!empty($endpoint) && $endpoint !== 'logs') {
        http_response_code(404);
        echo json_encode(["error" => "Rota GET /api/{$endpoint} não encontrada"]);
        exit();
    }
    $stmt = $pdo->query("SELECT * FROM log_entries ORDER BY data DESC, id DESC");
    $rows = $stmt->fetchAll();
    $formatted = array_map(function($row) {
        return [
            "id" => strval($row['id']),
            "data" => $row['data'] ?? '',
            "cidade" => $row['cidade'] ?? '',
            "numero_chamado" => $row['numero_chamado'] ?? '',
            "incidente" => $row['incidente'] ?? '',
            "descricao" => $row['descricao'] ?? '',
            "status" => $row['status'] ?? 'Pendente',
            "data_conclusao" => $row['data_conclusao'] ?? null,
            "created_at" => $row['created_at'] ?? ''
        ];
    }, $rows);
    echo json_encode($formatted);
    exit();
}

// POST - Inserir novo registro
if ($method === 'POST') {
    $raw = file_get_contents('php://input');
    $input = json_decode($raw, true) ?: [];
    
    $data = $input['data'] ?? '';
    $cidade = $input['cidade'] ?? '';
    $numero_chamado = $input['numero_chamado'] ?? '';
    $incidente = $input['incidente'] ?? '';
    $descricao = $input['descricao'] ?? '';
    $status = $input['status'] ?? 'Pendente';
    $data_conclusao = !empty($input['data_conclusao']) ? $input['data_conclusao'] : null;
    $created_at = $input['created_at'] ?? date('c');

    $stmt = $pdo->prepare("INSERT INTO log_entries (data, cidade, numero_chamado, incidente, descricao, status, data_conclusao, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
    $stmt->execute([$data, $cidade, $numero_chamado, $incidente, $descricao, $status, $data_conclusao, $created_at]);
    
    $newId = $pdo->lastInsertId();
    http_response_code(201);
    echo json_encode([
        "id" => strval($newId),
        "data" => $data,
        "cidade" => $cidade,
        "numero_chamado" => $numero_chamado,
        "incidente" => $incidente,
        "descricao" => $descricao,
        "status" => $status,
        "data_conclusao" => $data_conclusao,
        "created_at" => $created_at
    ]);
    exit();
}

// PUT - Atualizar registro existente
if ($method === 'PUT') {
    $raw = file_get_contents('php://input');
    $input = json_decode($raw, true) ?: [];
    $targetId = $id > 0 ? $id : intval($input['id'] ?? 0);

    $stmt = $pdo->prepare("UPDATE log_entries SET data = ?, cidade = ?, numero_chamado = ?, incidente = ?, descricao = ?, status = ?, data_conclusao = ? WHERE id = ?");
    $stmt->execute([
        $input['data'] ?? '',
        $input['cidade'] ?? '',
        $input['numero_chamado'] ?? '',
        $input['incidente'] ?? '',
        $input['descricao'] ?? '',
        $input['status'] ?? 'Pendente',
        !empty($input['data_conclusao']) ? $input['data_conclusao'] : null,
        $targetId
    ]);

    echo json_encode(array_merge($input, ["id" => strval($targetId)]));
    exit();
}

// DELETE - Excluir registro
if ($method === 'DELETE') {
    $targetId = $id;
    if ($targetId <= 0) {
        $raw = file_get_contents('php://input');
        $input = json_decode($raw, true) ?: [];
        $targetId = intval($input['id'] ?? 0);
    }
    
    $stmt = $pdo->prepare("DELETE FROM log_entries WHERE id = ?");
    $stmt->execute([$targetId]);
    echo json_encode(["message" => "Registro excluído com sucesso"]);
    exit();
}
