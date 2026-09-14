export const DEFAULT_GITHUB_URLS = {
  at1: "https://raw.githubusercontent.com/carloswladier/INDICADORES/main/DASH%20AT1%20PERSONA_ATUALIZADO.xlsx",
  outage: "https://raw.githubusercontent.com/carloswladier/INDICADORES/main/OUTAGE_SGO.xlsx",
  revisita: "https://raw.githubusercontent.com/carloswladier/INDICADORES/main/REVISITA_30D_Norte.xlsx",
  revisitaJanJun: "https://raw.githubusercontent.com/carloswladier/INDICADORES/main/REVISITA_30D_Jan_Jun.xlsx",
  revisitaJulDez: "https://raw.githubusercontent.com/carloswladier/INDICADORES/main/REVISITA_30D_Jul_Dez.xlsx",
  at5: "https://raw.githubusercontent.com/carloswladier/INDICADORES/main/AT5_NORTE.xlsx",
  qoeGpon: "https://raw.githubusercontent.com/carloswladier/INDICADORES/main/QOE_GPON.xlsx",
};

export function getEnvValue(key: string, altKeys: string[] = [], fallback = ''): string {
  const metaEnv = (import.meta as any).env || {};
  if (metaEnv[key] && String(metaEnv[key]).trim() !== '') {
    return String(metaEnv[key]).trim();
  }
  for (const alt of altKeys) {
    if (metaEnv[alt] && String(metaEnv[alt]).trim() !== '') {
      return String(metaEnv[alt]).trim();
    }
  }
  try {
    const localVal = localStorage.getItem(key);
    if (localVal && localVal.trim() !== '') return localVal.trim();
    for (const alt of altKeys) {
      const altLocal = localStorage.getItem(alt);
      if (altLocal && altLocal.trim() !== '') return altLocal.trim();
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
  const candidates: string[] = [primaryUrl];
  
  // Also add original targetUrl in case it was already raw
  if (targetUrl && targetUrl !== primaryUrl) {
    candidates.push(targetUrl.trim().replace(/ /g, '%20'));
  }

  // If URL has DASH_AT1_G1, also try INDICADORES
  if (primaryUrl.includes('DASH_AT1_G1')) {
    candidates.push(primaryUrl.replace('DASH_AT1_G1', 'INDICADORES'));
  }
  // If URL has INDICADORES, also try DASH_AT1_G1
  if (primaryUrl.includes('INDICADORES')) {
    candidates.push(primaryUrl.replace('INDICADORES', 'DASH_AT1_G1'));
  }
  
  // If targetUrl contains revisita filename variations
  if (primaryUrl.includes('REVISITA_30D_202608_Norte.xlsx')) {
    candidates.push(primaryUrl.replace('REVISITA_30D_202608_Norte.xlsx', 'REVISITA_30D_Norte.xlsx'));
  }
  if (primaryUrl.includes('REVISITA_30D_Norte.xlsx')) {
    candidates.push(primaryUrl.replace('REVISITA_30D_Norte.xlsx', 'REVISITA_30D_202608_Norte.xlsx'));
  }
  if (primaryUrl.includes('REVISITA_30D_Jan_Jun')) {
    candidates.push(primaryUrl.replace('REVISITA_30D_Jan_Jun', 'REVISITA_30D_JAN_JUN'));
    candidates.push(primaryUrl.replace('REVISITA_30D_Jan_Jun', 'REVISITA_30D_Jan-Jun'));
    candidates.push(primaryUrl.replace('REVISITA_30D_Jan_Jun', 'REVISITA_30D_Jan_a_Jun'));
    candidates.push(primaryUrl.replace('.xlsx', '.xls'));
  }
  if (primaryUrl.includes('REVISITA_30D_Jul_Dez')) {
    candidates.push(primaryUrl.replace('REVISITA_30D_Jul_Dez', 'REVISITA_30D_JUL_DEZ'));
    candidates.push(primaryUrl.replace('REVISITA_30D_Jul_Dez', 'REVISITA_30D_Jul-Dez'));
    candidates.push(primaryUrl.replace('REVISITA_30D_Jul_Dez', 'REVISITA_30D_Jul_a_Dez'));
    candidates.push(primaryUrl.replace('.xlsx', '.xls'));
  }
  if (primaryUrl.includes('QOE_GPON') || primaryUrl.includes('qoe_gpon') || primaryUrl.includes('QOE')) {
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'QOE_GPON.XLSX'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'QOE%20GPON.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'QOE_GPON_NORTE.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'BASE_QOE_GPON.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'DASH_QOE_GPON.xlsx'));
    candidates.push(primaryUrl.replace('QOE_GPON.xlsx', 'qoe_gpon.xlsx'));
    candidates.push(primaryUrl.replace('.xlsx', '.xls'));
    candidates.push(primaryUrl.replace('.xlsx', '.csv'));
  }

  // Remove duplicates
  const uniqueCandidates = Array.from(new Set(candidates));
  
  // Try direct fetch first for all candidate URLs
  for (const url of uniqueCandidates) {
    try {
      const res = await fetch(url);
      if (res.ok) {
        return await res.arrayBuffer();
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
        return await res.arrayBuffer();
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


