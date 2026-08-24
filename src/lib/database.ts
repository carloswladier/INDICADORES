// Helper to get Hostinger MySQL database configuration from localStorage or environment
export interface HostingerDbConfig {
  host: string;
  user: string;
  password?: string;
  database: string;
  port: number;
  uri?: string;
  apiEndpoint?: string;
}

export const getDbConfig = (): HostingerDbConfig => {
  try {
    const localHost = localStorage.getItem('DB_HOST') || '';
    const localUser = localStorage.getItem('DB_USER') || '';
    const localPassword = localStorage.getItem('DB_PASSWORD') || localStorage.getItem('DB_PASS') || '';
    const localDb = localStorage.getItem('DB_NAME') || localStorage.getItem('DB_DATABASE') || '';
    const localPort = Number(localStorage.getItem('DB_PORT')) || 3306;
    const localUri = localStorage.getItem('DATABASE_URL') || localStorage.getItem('MYSQL_URI') || '';
    const localApiEndpoint = localStorage.getItem('API_ENDPOINT') || localStorage.getItem('VITE_API_URL') || '';

    return {
      host: localHost,
      user: localUser,
      password: localPassword,
      database: localDb,
      port: localPort,
      uri: localUri,
      apiEndpoint: localApiEndpoint
    };
  } catch (e) {
    return {
      host: '',
      user: '',
      password: '',
      database: '',
      port: 3306,
      uri: '',
      apiEndpoint: ''
    };
  }
};

export const setDbConfig = (config: Partial<HostingerDbConfig>) => {
  try {
    if (config.host !== undefined) {
      if (config.host) localStorage.setItem('DB_HOST', config.host);
      else localStorage.removeItem('DB_HOST');
    }
    if (config.user !== undefined) {
      if (config.user) localStorage.setItem('DB_USER', config.user);
      else localStorage.removeItem('DB_USER');
    }
    if (config.password !== undefined) {
      if (config.password) localStorage.setItem('DB_PASSWORD', config.password);
      else localStorage.removeItem('DB_PASSWORD');
    }
    if (config.database !== undefined) {
      if (config.database) localStorage.setItem('DB_NAME', config.database);
      else localStorage.removeItem('DB_NAME');
    }
    if (config.port !== undefined) {
      localStorage.setItem('DB_PORT', String(config.port || 3306));
    }
    if (config.uri !== undefined) {
      if (config.uri) localStorage.setItem('DATABASE_URL', config.uri);
      else localStorage.removeItem('DATABASE_URL');
    }
    if (config.apiEndpoint !== undefined) {
      if (config.apiEndpoint) localStorage.setItem('API_ENDPOINT', config.apiEndpoint);
      else localStorage.removeItem('API_ENDPOINT');
    }
  } catch (err) {
    console.warn('Unable to set Database config in localStorage:', err);
  }
};
