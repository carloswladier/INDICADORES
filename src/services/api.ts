import { getDbConfig } from '../lib/database';

export interface LogEntry {
  id?: string | number;
  created_at?: string;
  data: string;
  cidade: string;
  numero_chamado: string;
  incidente: string;
  descricao: string;
  status: 'Pendente' | 'Em Andamento' | 'Resolvido';
  data_conclusao?: string | null;
}

const LOCAL_STORAGE_KEY = 'DIARIO_DE_BORDO_LOCAL_ENTRIES';

export const getLocalEntries = (): LogEntry[] => {
  try {
    const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
};

export const saveLocalEntries = (entries: LogEntry[]): void => {
  try {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(entries));
  } catch (e) {
    console.warn('Unable to save local entries:', e);
  }
};

const getHeaders = (): Record<string, string> => {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    'Accept': 'application/json'
  };
  const dbConfig = getDbConfig();
  if (dbConfig.host) {
    headers['x-db-host'] = dbConfig.host;
  }
  if (dbConfig.user) {
    headers['x-db-user'] = dbConfig.user;
  }
  if (dbConfig.password) {
    headers['x-db-password'] = dbConfig.password;
  }
  if (dbConfig.database) {
    headers['x-db-name'] = dbConfig.database;
  }
  if (dbConfig.port) {
    headers['x-db-port'] = String(dbConfig.port);
  }
  if (dbConfig.uri) {
    headers['x-db-uri'] = dbConfig.uri;
  }
  return headers;
};

// Safe JSON parser to handle HTML errors gracefully
async function parseJsonResponse(res: Response): Promise<{ ok: boolean; data: any; isHtml?: boolean }> {
  const contentType = res.headers.get('content-type') || '';
  const text = await res.text();
  
  if (!text || text.trim().startsWith('<') || (!contentType.includes('application/json') && text.includes('<!DOCTYPE'))) {
    return { ok: false, data: null, isHtml: true };
  }

  try {
    const parsed = JSON.parse(text);
    return { ok: res.ok, data: parsed };
  } catch (e) {
    return { ok: false, data: null };
  }
}

// Determines the base API endpoint
function getEndpoints(): string[] {
  const config = getDbConfig();
  const custom = config.apiEndpoint?.trim();
  const list: string[] = [];

  if (custom && custom.startsWith('http')) {
    list.push(custom);
  }

  if (typeof window !== 'undefined' && window.location.hostname.includes('hostinger')) {
    list.push('/api.php');
    list.push('/api');
  } else {
    list.push('/api');
    list.push('/api.php');
  }

  return list;
}

function buildUrl(base: string, path: string, query?: string): string {
  const isPhp = base.endsWith('.php') || base.includes('.php');
  if (isPhp) {
    if (path === 'logs') {
      return query ? `${base}?${query}` : base;
    }
    if (path.startsWith('logs/')) {
      const id = path.replace('logs/', '');
      return `${base}?id=${id}${query ? `&${query}` : ''}`;
    }
    if (path === 'status' || path === 'db-status') {
      return `${base}?action=status`;
    }
    if (path === 'test' || path === 'db-test') {
      return `${base}?action=test`;
    }
    return base;
  }

  const cleanBase = base.endsWith('/') ? base.slice(0, -1) : base;
  const cleanPath = path.startsWith('/') ? path.slice(1) : path;
  const url = `${cleanBase}/${cleanPath}`;
  return query ? `${url}?${query}` : url;
}

export const logApi = {
  async getAll(): Promise<LogEntry[]> {
    const endpoints = getEndpoints();
    let lastError: any = null;

    for (const base of endpoints) {
      try {
        const url = buildUrl(base, 'logs');
        const res = await fetch(url, { headers: getHeaders() });
        const result = await parseJsonResponse(res);

        if (result.ok && Array.isArray(result.data)) {
          saveLocalEntries(result.data);
          return result.data;
        }
      } catch (err) {
        lastError = err;
      }
    }

    // Fallback to local storage
    const local = getLocalEntries();
    if (local.length > 0) {
      console.warn('[API] Servidor indisponível, utilizando dados em cache local.');
      return local;
    }

    if (lastError) {
      console.warn('[API] Erro ao buscar logs:', lastError);
    }
    return [];
  },

  async create(entry: LogEntry): Promise<LogEntry> {
    const endpoints = getEndpoints();
    let created: LogEntry | null = null;

    for (const base of endpoints) {
      try {
        const url = buildUrl(base, 'logs');
        const res = await fetch(url, {
          method: 'POST',
          headers: getHeaders(),
          body: JSON.stringify(entry)
        });
        const result = await parseJsonResponse(res);
        if (result.ok && result.data && result.data.id) {
          created = result.data;
          break;
        }
      } catch (err) {
        console.warn('[API] Falha POST no endpoint:', base, err);
      }
    }

    const finalEntry: LogEntry = created || {
      ...entry,
      id: entry.id || `local_${Date.now()}`,
      created_at: entry.created_at || new Date().toISOString()
    };

    const current = getLocalEntries();
    saveLocalEntries([finalEntry, ...current.filter(e => String(e.id) !== String(finalEntry.id))]);
    return finalEntry;
  },

  async update(id: string | number, entry: LogEntry): Promise<LogEntry> {
    const endpoints = getEndpoints();
    let updated: LogEntry | null = null;

    for (const base of endpoints) {
      try {
        const url = buildUrl(base, `logs/${id}`);
        const res = await fetch(url, {
          method: 'PUT',
          headers: getHeaders(),
          body: JSON.stringify(entry)
        });
        const result = await parseJsonResponse(res);
        if (result.ok && result.data) {
          updated = result.data;
          break;
        }
      } catch (err) {
        console.warn('[API] Falha PUT no endpoint:', base, err);
      }
    }

    const finalEntry: LogEntry = updated || { ...entry, id };
    const current = getLocalEntries();
    const list = current.map(item => String(item.id) === String(id) ? finalEntry : item);
    saveLocalEntries(list);
    return finalEntry;
  },

  async delete(id: string | number): Promise<void> {
    const endpoints = getEndpoints();

    for (const base of endpoints) {
      try {
        const url = buildUrl(base, `logs/${id}`);
        await fetch(url, {
          method: 'DELETE',
          headers: getHeaders()
        });
      } catch (err) {
        // Ignora erro no servidor
      }
    }

    const current = getLocalEntries();
    saveLocalEntries(current.filter(item => String(item.id) !== String(id)));
  },

  async checkStatus(): Promise<{ configured: boolean; type: string; status?: string; error?: string; details?: string; count?: number }> {
    const endpoints = getEndpoints();

    for (const base of endpoints) {
      try {
        const url = buildUrl(base, 'status');
        const res = await fetch(url, { headers: getHeaders() });
        const result = await parseJsonResponse(res);

        if (result.ok && result.data) {
          return {
            configured: Boolean(result.data.configured ?? true),
            type: result.data.type || 'Hostinger MySQL',
            status: result.data.status || 'connected',
            count: result.data.count
          };
        }
      } catch (err: any) {
        console.warn('[API] Falha ao checar status no endpoint:', base, err);
      }
    }

    return {
      configured: false,
      type: 'Hostinger MySQL',
      status: 'offline',
      error: 'Servidor MySQL não respondeu com JSON.',
      details: 'Certifique-se de que os arquivos api.php e .htaccess estão na pasta public_html da Hostinger.'
    };
  }
};
