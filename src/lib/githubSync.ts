export const DEFAULT_GITHUB_URLS = {
  at1: "https://raw.githubusercontent.com/carloswladier/DASH_AT1_G1/main/DASH%20AT1%20PERSONA_ATUALIZADO.xlsx",
  outage: "https://raw.githubusercontent.com/carloswladier/DASH_AT1_G1/main/OUTAGE_SGO.xlsx",
  revisita: "https://raw.githubusercontent.com/carloswladier/DASH_AT1_G1/main/REVISITA_30D_Norte.xlsx",
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
  const candidates = [primaryUrl];
  
  // If targetUrl contains the old file name, also add the new file name candidate
  if (primaryUrl.includes('REVISITA_30D_202608_Norte.xlsx')) {
    candidates.push(primaryUrl.replace('REVISITA_30D_202608_Norte.xlsx', 'REVISITA_30D_Norte.xlsx'));
  }
  
  // Try direct fetch first for all candidate URLs
  for (const url of candidates) {
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
  for (const url of candidates) {
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

