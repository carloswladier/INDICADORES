export const DEFAULT_GITHUB_URLS = {
  at1: "https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/DASH%20AT1%20PERSONA_ATUALIZADO.xlsx",
  outage: "https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/OUTAGE_SGO.xlsx",
  revisita: "https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/REVISITA_30D_Jul_Dez.xlsx",
  revisitaJanJun: "https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/REVISITA_30D_Jan_Jun.xlsx",
  revisitaJulDez: "https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/REVISITA_30D_Jul_Dez.xlsx",
  at5: "https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/AT5_NORTE.xlsx",
  qoeGpon: "https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/QOE_GPON_NORTE.xlsx",
};

// Migrate deprecated repository URLs to current INDICADORES2 repository
function migrateDeprecatedRepoUrl(url: string): string {
  if (!url) return '';
  return url
    .replace('/carloswladier/DASH_AT1_G1/', '/carloswladier/INDICADORES2/')
    .replace('/carloswladier/INDICADORES_MANUT/', '/carloswladier/INDICADORES2/')
    .replace('/carloswladier/INDICADORES/', '/carloswladier/INDICADORES2/');
}

export function getEnvValue(key: string, altKeys: string[] = [], fallback = ''): string {
  const metaEnv = (import.meta as any).env || {};
  if (metaEnv[key] && String(metaEnv[key]).trim() !== '') {
    return migrateDeprecatedRepoUrl(String(metaEnv[key]).trim());
  }
  for (const alt of altKeys) {
    if (metaEnv[alt] && String(metaEnv[alt]).trim() !== '') {
      return migrateDeprecatedRepoUrl(String(metaEnv[alt]).trim());
    }
  }
  try {
    const localVal = localStorage.getItem(key);
    if (localVal && localVal.trim() !== '') {
      const migrated = migrateDeprecatedRepoUrl(localVal.trim());
      if (migrated !== localVal.trim()) {
        try { localStorage.setItem(key, migrated); } catch {}
      }
      return migrated;
    }
    for (const alt of altKeys) {
      const altLocal = localStorage.getItem(alt);
      if (altLocal && altLocal.trim() !== '') {
        const migrated = migrateDeprecatedRepoUrl(altLocal.trim());
        if (migrated !== altLocal.trim()) {
          try { localStorage.setItem(alt, migrated); } catch {}
        }
        return migrated;
      }
    }
  } catch {
    // Ignore localStorage exceptions
  }
  return fallback;
}

export function normalizeGithubRawUrl(targetUrl: string): string {
  if (!targetUrl) return '';
  let url = targetUrl.trim();
  
  // Transform github.com web URLs to raw.githubusercontent.com
  if (url.includes('github.com') && !url.includes('raw.githubusercontent.com')) {
    url = url
      .replace('github.com', 'raw.githubusercontent.com')
      .replace('/blob/', '/')
      .replace('/raw/', '/');
  }
  
  // Normalize /refs/heads/
  url = url.replace('/refs/heads/', '/');
  
  // Ensure spaces in file names are encoded for fetch
  return url.replace(/ /g, '%20');
}

export async function fetchGithubFileArrayBuffer(targetUrl: string): Promise<ArrayBuffer> {
  const primaryUrl = normalizeGithubRawUrl(targetUrl);
  
  // Array of URL candidates to attempt
  const candidates: string[] = [];

  // 1. HIGHEST PRIORITY: ALWAYS target the active INDICADORES2 repository on GitHub
  if (primaryUrl.includes('carloswladier')) {
    const indic2 = primaryUrl
      .replace('/carloswladier/DASH_AT1_G1/', '/carloswladier/INDICADORES2/')
      .replace('/carloswladier/INDICADORES_MANUT/', '/carloswladier/INDICADORES2/')
      .replace('/carloswladier/INDICADORES/', '/carloswladier/INDICADORES2/');
    candidates.push(indic2);
  }

  // 2. Add primaryUrl
  candidates.push(primaryUrl);
  
  // Also add original targetUrl in case it was already raw
  if (targetUrl && targetUrl !== primaryUrl) {
    candidates.push(targetUrl.trim().replace(/ /g, '%20'));
  }

  // Support file-specific local host fallbacks (especially for Hostinger Vite static hosting)
  const extractFilename = (u: string) => {
    try {
      const parts = u.split('/');
      return parts[parts.length - 1];
    } catch {
      return '';
    }
  };
  const fileName = extractFilename(primaryUrl);
  if (fileName) {
    candidates.push(`/${fileName}`);
    candidates.push(`./${fileName}`);
    candidates.push(fileName);
  }

  // Support INDICADORES_MANUT and DASH_AT1_G1 transitions
  if (primaryUrl.includes('INDICADORES_MANUT')) {
    candidates.push(primaryUrl.replace('INDICADORES_MANUT', 'INDICADORES2'));
    candidates.push(primaryUrl.replace('INDICADORES_MANUT', 'INDICADORES'));
  }
  if (primaryUrl.includes('DASH_AT1_G1')) {
    candidates.push(primaryUrl.replace('DASH_AT1_G1', 'INDICADORES2'));
    candidates.push(primaryUrl.replace('DASH_AT1_G1', 'INDICADORES_MANUT'));
  }
  
  // If targetUrl contains AT1 filename variations
  if (primaryUrl.includes('DASH') && primaryUrl.includes('AT1')) {
    candidates.push('https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/DASH%20AT1%20PERSONA_ATUALIZADO.xlsx');
    candidates.push('/DASH%20AT1%20PERSONA_ATUALIZADO.xlsx');
  }

  // If targetUrl contains revisita filename variations
  if (primaryUrl.includes('OUTAGE_SGO') || primaryUrl.includes('outage')) {
    candidates.push('https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/OUTAGE_SGO.xlsx');
    candidates.push('/OUTAGE_SGO.xlsx');
    candidates.push(primaryUrl.replace('OUTAGE_SGO.xlsx', 'OUTAGE_SGO_JAN_JUN.xlsx'));
    candidates.push(primaryUrl.replace('OUTAGE_SGO.xlsx', 'OUTAGE_SGO_JUL_DEZ.xlsx'));
    candidates.push(primaryUrl.replace('OUTAGE_SGO.xlsx', 'OUTAGE_SGO_Jan_Jun.xlsx'));
    candidates.push(primaryUrl.replace('OUTAGE_SGO.xlsx', 'OUTAGE_SGO_Jul_Dez.xlsx'));
    candidates.push(primaryUrl.replace('OUTAGE_SGO.xlsx', 'OUTAGE_SGO.XLSX'));
    candidates.push(primaryUrl.replace('OUTAGE_SGO.xlsx', 'OUTAGE.xlsx'));
  }
  if (primaryUrl.includes('REVISITA_30D_202608_Norte.xlsx')) {
    candidates.push(primaryUrl.replace('REVISITA_30D_202608_Norte.xlsx', 'REVISITA_30D_Norte.xlsx'));
  }
  if (primaryUrl.includes('REVISITA_30D_Norte.xlsx')) {
    candidates.push(primaryUrl.replace('REVISITA_30D_Norte.xlsx', 'REVISITA_30D_Jul_Dez.xlsx'));
    candidates.push('https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/REVISITA_30D_Jul_Dez.xlsx');
    candidates.push(primaryUrl.replace('REVISITA_30D_Norte.xlsx', 'REVISITA_30D_202608_Norte.xlsx'));
  }
  if (primaryUrl.includes('REVISITA_30D_Jan_Jun')) {
    candidates.push(primaryUrl.replace('REVISITA_30D_Jan_Jun', 'REVISITA_30D_JAN_JUN'));
    candidates.push(primaryUrl.replace('REVISITA_30D_Jan_Jun', 'REVISITA_30D_Jan-Jun'));
    candidates.push(primaryUrl.replace('REVISITA_30D_Jan_Jun', 'REVISITA_30D_Jan_a_Jun'));
    candidates.push(primaryUrl.replace('.xlsx', '.xls'));
  }
  if (primaryUrl.includes('REVISITA_30D_Jul_Dez')) {
    candidates.push('https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/REVISITA_30D_Jul_Dez.xlsx');
    candidates.push('/REVISITA_30D_Jul_Dez.xlsx');
    candidates.push(primaryUrl.replace('REVISITA_30D_Jul_Dez', 'REVISITA_30D_JUL_DEZ'));
    candidates.push(primaryUrl.replace('REVISITA_30D_Jul_Dez', 'REVISITA_30D_Jul-Dez'));
    candidates.push(primaryUrl.replace('REVISITA_30D_Jul_Dez', 'REVISITA_30D_Jul_a_Dez'));
    candidates.push(primaryUrl.replace('.xlsx', '.xls'));
  }
  if (primaryUrl.includes('QOE_GPON') || primaryUrl.includes('qoe_gpon') || primaryUrl.includes('QOE')) {
    candidates.push('https://raw.githubusercontent.com/carloswladier/INDICADORES2/main/QOE_GPON_NORTE.xlsx');
    candidates.push('/QOE_GPON_NORTE.xlsx');
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'QOE_GPON_NORTE.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON_NORTE.xlsx', 'QOE_GPON.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'QOE_GPON.XLSX'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'QOE%20GPON.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'BASE_QOE_GPON.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'DASH_QOE_GPON.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'qoe_gpon.xlsx'));
    candidates.push(primaryUrl.replace('.xlsx', '.xls'));
    candidates.push(primaryUrl.replace('.xlsx', '.csv'));
  }

  // Helper to ensure array buffer is real binary excel data, not an HTML error or empty JSON
  const isValidExcelOrDataBuffer = (buf: ArrayBuffer): boolean => {
    if (!buf || buf.byteLength < 50) return false;
    const b = new Uint8Array(buf.slice(0, 8));
    // ZIP / XLSX magic: PK (0x50, 0x4B, 0x03, 0x04)
    if (b[0] === 0x50 && b[1] === 0x4b) return true;
    // XLS (BIFF8) magic: 0xD0, 0xCF, 0x11, 0xE0
    if (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) return true;
    // Reject HTML ('<') and JSON ('{', '[')
    if (b[0] === 0x3c || b[0] === 0x7b || b[0] === 0x5b) return false;
    return true;
  };

  // Remove duplicates
  const uniqueCandidates = Array.from(new Set(candidates));
  
  // Try direct fetch first for all candidate URLs
  for (const url of uniqueCandidates) {
    try {
      const res = await fetch(url);
      if (res.ok) {
        const buf = await res.arrayBuffer();
        if (isValidExcelOrDataBuffer(buf)) {
          return buf;
        }
      }
    } catch {
      // Continue to next candidate / proxy
    }
  }
  
  // If direct fetch fails, attempt via server proxy endpoint
  for (const url of uniqueCandidates) {
    try {
      const proxyUrl = `/api/proxy-github?url=${encodeURIComponent(url)}`;
      const res = await fetch(proxyUrl);
      if (res.ok) {
        const buf = await res.arrayBuffer();
        if (isValidExcelOrDataBuffer(buf)) {
          return buf;
        }
      }
    } catch {
      // Continue
    }
  }
  
  throw new Error(`Não foi possível baixar o arquivo do GitHub. Verifique a URL: ${primaryUrl}`);
}

export function getGithubAt1Url(): string {
  return normalizeGithubRawUrl(
    getEnvValue('VITE_GITHUB_EXCEL_URL', ['VITE_GITHUB_AT1_URL', 'GITHUB_EXCEL', 'VITE_GITHUB_EXCEL', 'VITE_GITHUB_EXCEL_URL_1'], DEFAULT_GITHUB_URLS.at1)
  );
}

export function getGithubOutageUrl(): string {
  return normalizeGithubRawUrl(
    getEnvValue('VITE_GITHUB_OUTAGE_URL', ['VITE_GITHUB_EXCEL_OUTAGE', 'GITHUB_EXCEL_OUTAGE', 'GITHUB_OUTAGE_URL', 'VITE_GITHUB_EXCEL_URL_2'], DEFAULT_GITHUB_URLS.outage)
  );
}

export function getGithubRevisitaUrl(): string {
  return normalizeGithubRawUrl(
    getEnvValue(
      'VITE_GITHUB_REVISITA_URL',
      ['VITE_GITHUB_EXCEL_REVISITA', 'GITHUB_EXCEL_REVISITA', 'GITHUB_REVISITA_URL', 'VITE_GITHUB_EXCEL_URL_3', 'VITE_GITHUB_EXCEL_REVISITA_URL'],
      DEFAULT_GITHUB_URLS.revisita
    )
  );
}

export function getGithubRevisitaJanJunUrl(): string {
  return normalizeGithubRawUrl(
    getEnvValue(
      'VITE_GITHUB_REVISITA_JAN_JUN_URL',
      ['VITE_GITHUB_EXCEL_REVISITA_JAN_JUN', 'GITHUB_EXCEL_REVISITA_JAN_JUN', 'GITHUB_REVISITA_JAN_JUN', 'VITE_GITHUB_REVISITA_JAN_JUN'],
      DEFAULT_GITHUB_URLS.revisitaJanJun
    )
  );
}

export function getGithubRevisitaJulDezUrl(): string {
  return normalizeGithubRawUrl(
    getEnvValue(
      'VITE_GITHUB_REVISITA_JUL_DEZ_URL',
      ['VITE_GITHUB_EXCEL_REVISITA_JUL_DEZ', 'GITHUB_EXCEL_REVISITA_JUL_DEZ', 'GITHUB_REVISITA_JUL_DEZ', 'VITE_GITHUB_REVISITA_JUL_DEZ'],
      DEFAULT_GITHUB_URLS.revisitaJulDez
    )
  );
}

export function getGithubAt5Url(): string {
  return normalizeGithubRawUrl(
    getEnvValue(
      'VITE_GITHUB_EXCEL_URL_AT5',
      ['VITE_GITHUB_AT5_URL', 'GITHUB_EXCEL_AT5', 'VITE_GITHUB_EXCEL_URL_5', 'GITHUB_AT5_URL'],
      DEFAULT_GITHUB_URLS.at5
    )
  );
}

export function getGithubQoeGponUrl(): string {
  return normalizeGithubRawUrl(
    getEnvValue(
      'VITE_GITHUB_EXCEL_URL_QOE_GPON',
      ['VITE_GITHUB_QOE_GPON_URL', 'GITHUB_EXCEL_QOE_GPON', 'GITHUB_QOE_GPON_URL', 'VITE_GITHUB_EXCEL_URL_QOE'],
      DEFAULT_GITHUB_URLS.qoeGpon
    )
  );
}


