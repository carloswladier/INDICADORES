import React, { useState, useMemo, useEffect, useRef } from 'react';
import { 
  FileSpreadsheet, 
  Upload, 
  Filter, 
  CheckCircle2, 
  RefreshCw, 
  XCircle, 
  X, 
  BarChart3, 
  TrendingUp, 
  AlertTriangle, 
  Layers, 
  Download, 
  SlidersHorizontal, 
  ChevronDown, 
  Database, 
  Search,
  Activity,
  Table,
  RotateCcw,
  MapPin,
  Building2,
  Check,
  Trash2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { MultiFilterSelect } from './MultiFilterSelect';
import { 
  BarChart, 
  Bar, 
  LineChart, 
  Line, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  Cell, 
  LabelList 
} from 'recharts';
import * as XLSX from 'xlsx';
import { cn, formatPercent, formatDecimal } from '../lib/utils';

export interface LOGRow {
  contrato: string;
  municipio: string;
  tipoOs: string;
  areaDespacho: string;
  statusOs: string;
  log: number;
  codigoBaixa: string;
  razaoCancelamento: string;
  node: string;
  data: string;
}

// Generate high quality initial realistic dataset based on official Excel distribution
const generateInitialMockLOGData = (): LOGRow[] => {
  const result: LOGRow[] = [];
  const dates = ['01/08', '02/08', '03/08', '04/08', '05/08', '06/08', '07/08', '08/08', '09/08', '10/08', '11/08'];
  
  const cityConfigs: Array<{
    name: string;
    area: string;
    counts: { [logLevel: number]: number };
  }> = [
    { name: 'ANANINDEUA', area: 'AREA ANANINDEUA', counts: { 0: 318, 1: 85, 2: 23, 3: 9, 4: 3 } },
    { name: 'BELÉM', area: 'AREA BELEM', counts: { 0: 3034, 1: 731, 2: 185, 3: 48, 4: 12, 5: 6, 6: 2, 7: 1 } },
    { name: 'CASTANHAL', area: 'AREA CASTANHAL', counts: { 0: 221, 1: 67, 2: 23, 3: 4, 4: 3, 5: 1 } },
    { name: 'CAXIAS', area: 'MA2', counts: { 0: 7, 1: 1 } },
    { name: 'IMPERATRIZ', area: 'AREA IMPERATRIZ', counts: { 0: 239, 1: 59, 2: 9, 3: 3 } },
    { name: 'MACAPÁ', area: 'AREA MACAPA', counts: { 0: 318, 1: 72, 2: 18, 3: 6, 4: 2 } },
    { name: 'MANAUS', area: 'MAS', counts: { 0: 4442, 1: 917, 2: 264, 3: 86, 4: 34, 5: 9 } },
    { name: 'MARABÁ', area: 'AREA 4', counts: { 0: 82, 1: 11, 2: 4, 3: 2, 4: 1, 5: 1 } },
    { name: 'PARAGOMINAS', area: 'AREA PARAGOMINAS', counts: { 0: 56, 1: 10, 2: 4, 3: 2 } },
    { name: 'PARAUAPEBAS', area: 'AREA PARAUAPEBA', counts: { 0: 159, 1: 23, 2: 5, 3: 2, 4: 2 } },
    { name: 'SANTANA', area: 'AREA SANTANA', counts: { 0: 163, 1: 27, 2: 15, 3: 2 } },
    { name: 'SÃO LUÍS', area: 'MA1', counts: { 0: 2931, 1: 791, 2: 254, 3: 84, 4: 23, 5: 11, 6: 2 } },
    { name: 'TIMON', area: 'MA3', counts: { 0: 18, 1: 2, 2: 1 } },
  ];

  const tiposOS = ['MANUTENCAO GPON', 'INSTALACAO BLC', 'CORRETIVA HFC', 'MUDANCA ENDERECO', 'REPARO FTTH'];
  const statusOSList = ['EXECUTADA', 'ENCERRADA', 'CANCELADA', 'EM ANDAMENTO'];
  const codigosBaixaExec = ['C01 - CONCLUIDO', 'C02 - SINAL NORMALIZADO', 'C05 - TROCA DE DROP', 'C08 - RECONFIGURACAO', 'C12 - TROCA DE ONT'];
  const codigosBaixaCancel = ['302', '40', '101', '301', '302 - CLIENTE AUSENTE', '40 - AREA DE RISCO'];
  const razaoCancelList = ['302', '40', '101', '301', '302 - CLIENTE AUSENTE', '40 - AREA DE RISCO'];

  let globalIdx = 100000;

  cityConfigs.forEach((cfg) => {
    Object.entries(cfg.counts).forEach(([logStr, count]) => {
      const logLevel = Number(logStr);
      for (let i = 0; i < count; i++) {
        globalIdx++;
        const dateStr = dates[i % dates.length];
        const tipoStr = tiposOS[i % tiposOS.length];
        const isExec = logLevel < 3 || i % 4 !== 0;
        const statusStr = isExec ? statusOSList[i % 2] : statusOSList[2 + (i % 2)];
        const baixaList = isExec ? codigosBaixaExec : codigosBaixaCancel;
        const razaoCancStr = !isExec || statusStr === 'CANCELADA' ? razaoCancelList[i % razaoCancelList.length] : '-';
        const codigoBaixaStr = isExec ? baixaList[i % baixaList.length] : razaoCancStr;
        
        const nodeIdx = (i % 15) + 1;
        const nodeStr = `${cfg.name.substring(0, 3)}${nodeIdx.toString().padStart(2, '0')}`;

        result.push({
          contrato: `${globalIdx}`,
          municipio: cfg.name,
          tipoOs: tipoStr,
          areaDespacho: cfg.area,
          statusOs: statusStr,
          log: logLevel,
          codigoBaixa: codigoBaixaStr,
          razaoCancelamento: razaoCancStr,
          node: nodeStr,
          data: dateStr,
        });
      }
    });
  });

  return result;
};

export default function LOGDashboard() {
  const [data, setData] = useState<LOGRow[]>([]);

  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searchFilter, setSearchFilter] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 15;

  // GitHub & Action refs
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [showGithubInput, setShowGithubInput] = useState(false);
  const [githubUrl, setGithubUrl] = useState('');
  const [isGithubLoading, setIsGithubLoading] = useState(false);

  const handleGithubLoad = async (urlToLoad?: string) => {
    const targetUrl = urlToLoad || githubUrl;
    if (!targetUrl.trim()) {
      setError('Por favor, informe a URL do arquivo no GitHub.');
      return;
    }
    setIsLoading(true);
    setIsGithubLoading(true);
    setError(null);
    try {
      let rawUrl = targetUrl.trim();
      if (rawUrl.includes('github.com') && rawUrl.includes('/blob/')) {
        rawUrl = rawUrl.replace('github.com', 'raw.githubusercontent.com').replace('/blob/', '/');
      }
      const response = await fetch(rawUrl);
      if (!response.ok) {
        throw new Error(`Falha ao baixar o arquivo: HTTP ${response.status} ${response.statusText}`);
      }
      const arrayBuffer = await response.arrayBuffer();
      const success = processExcelBuffer(arrayBuffer);
      if (success) {
        setShowGithubInput(false);
        setGithubUrl('');
      }
    } catch (err: any) {
      console.error(err);
      setError(`Erro ao carregar do GitHub: ${err?.message || 'Verifique se o link é válido e público.'}`);
      setIsLoading(false);
    } finally {
      setIsGithubLoading(false);
    }
  };

  const clearAllData = () => {
    setData([]);
    handleClearFilters();
  };

  // Clear legacy localStorage data on mount to ensure fresh state
  useEffect(() => {
    try {
      localStorage.removeItem('LOG_DASHBOARD_DATA');
    } catch (err) {
      // ignore
    }
  }, []);

  // 5 Primary Filters State
  const [filters, setFilters] = useState<{
    municipio: string[];
    tipoOs: string[];
    areaDespacho: string[];
    statusOs: string[];
    logLevel: string[];
  }>({
    municipio: [],
    tipoOs: [],
    areaDespacho: [],
    statusOs: [],
    logLevel: [],
  });

  // Calculate dynamic options list for dropdown filters
  const uniqueOptions = useMemo(() => {
    const getFilteredFor = (exceptKey: keyof typeof filters) => {
      return data.filter((item) => {
        return (Object.keys(filters) as Array<keyof typeof filters>).every((k) => {
          if (k === exceptKey) return true;
          const selected = filters[k];
          if (selected.length === 0) return true;
          if (k === 'logLevel') {
            return selected.includes(String(item.log ?? 0));
          }
          return selected.includes(String(item[k]));
        });
      });
    };

    const extractUnique = (dataset: LOGRow[], key: keyof LOGRow) => {
      const set = new Set<string>();
      dataset.forEach((row) => {
        const val = row[key];
        if (val && val !== 'N/A' && val !== '') {
          set.add(String(val).trim());
        }
      });
      return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'));
    };

    const extractUniqueLog = (dataset: LOGRow[]) => {
      const set = new Set<string>();
      dataset.forEach((row) => {
        const val = row.log ?? 0;
        set.add(String(val));
      });
      return Array.from(set).sort((a, b) => Number(a) - Number(b));
    };

    return {
      municipio: extractUnique(getFilteredFor('municipio'), 'municipio'),
      tipoOs: extractUnique(getFilteredFor('tipoOs'), 'tipoOs'),
      areaDespacho: extractUnique(getFilteredFor('areaDespacho'), 'areaDespacho'),
      statusOs: extractUnique(getFilteredFor('statusOs'), 'statusOs'),
      logLevel: extractUniqueLog(getFilteredFor('logLevel')),
    };
  }, [data, filters]);

  // Filtered dataset based on 5 filters
  const filteredData = useMemo(() => {
    return data.filter((item) => {
      if (filters.municipio.length > 0 && !filters.municipio.includes(item.municipio)) return false;
      if (filters.tipoOs.length > 0 && !filters.tipoOs.includes(item.tipoOs)) return false;
      if (filters.areaDespacho.length > 0 && !filters.areaDespacho.includes(item.areaDespacho)) return false;
      if (filters.statusOs.length > 0 && !filters.statusOs.includes(item.statusOs)) return false;
      if (filters.logLevel.length > 0 && !filters.logLevel.includes(String(item.log ?? 0))) return false;
      return true;
    });
  }, [data, filters]);

  // Handle Multi-select filter toggles
  const handleFilterToggle = (filterKey: keyof typeof filters, value: string) => {
    setFilters((prev) => {
      const current = prev[filterKey];
      const exists = current.includes(value);
      const updated = exists ? current.filter((v) => v !== value) : [...current, value];
      return { ...prev, [filterKey]: updated };
    });
  };

  const handleClearFilters = () => {
    setFilters({
      municipio: [],
      tipoOs: [],
      areaDespacho: [],
      statusOs: [],
      logLevel: [],
    });
    setSearchFilter('');
  };

  const activeFilterCount = useMemo(() => {
    return (
      filters.municipio.length +
      filters.tipoOs.length +
      filters.areaDespacho.length +
      filters.statusOs.length +
      filters.logLevel.length
    );
  }, [filters]);

  // Excel Importer Parser
  const processExcelBuffer = (arrayBuffer: ArrayBuffer): boolean => {
    try {
      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      
      // Look for a sheet containing 'LOG', 'ANALITICO', 'ANALÍTICO', 'AT5', or use the first sheet
      const logSheetName = wb.SheetNames.find((name) => {
        const n = name.toLowerCase();
        return n.includes('log') || n.includes('analitico') || n.includes('analítico') || n.includes('at5');
      });
      const wsname = logSheetName || wb.SheetNames[0];
      const ws = wb.Sheets[wsname];

      const rawRows = XLSX.utils.sheet_to_json(ws) as any[];

      if (!rawRows || rawRows.length === 0) {
        setError('O arquivo Excel selecionado está vazio ou não contém dados válidos.');
        setIsLoading(false);
        return false;
      }

      const getValue = (row: any, candidates: string[]): any => {
        const rowKeys = Object.keys(row);
        const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
        const normalizedCandidates = candidates.map(normalize);
        const found = rowKeys.find((rk) => normalizedCandidates.includes(normalize(rk)));
        return found ? row[found] : undefined;
      };

      const parsedRows: LOGRow[] = rawRows.map((r, idx) => {
        const municipioVal = String(
          getValue(r, ['MUNICIPIO', 'NM_MUNICIPIO', 'NM_MUNICIPIO_BI', 'Municipio', 'Cidade']) || 'N/A'
        ).trim().toUpperCase();

        const tipoOsVal = String(
          getValue(r, ['TIPO_OS', 'TIPO OS', 'TIPO', 'NM_TIPO_OS', 'TIPO_ORDEM_SERVICO']) || 'N/A'
        ).trim().toUpperCase();

        const areaVal = String(
          getValue(r, ['AREA_DESPACHO', 'AREA DESPACHO', 'AREA', 'NM_AREA_DESPACHO', 'DESPACHO']) || 'N/A'
        ).trim().toUpperCase();

        const statusVal = String(
          getValue(r, ['STATUS_OS', 'STATUS OS', 'STATUS', 'NM_STATUS_ORDEM_SERVICO', 'SITUACAO']) || 'N/A'
        ).trim().toUpperCase();

        const logValRaw = getValue(r, ['LOG', 'Log', 'QTD_LOG', 'QUANTIDADE_LOG', 'REINCIDENCIA', 'REINCIDÊNCIA', 'VALOR_LOG']);
        const parsedLog = Number(logValRaw);
        const logVal = !isNaN(parsedLog) && parsedLog >= 0 ? parsedLog : 0;

        const rawCodigoBaixa = String(
          getValue(r, ['CD_BAIXA', 'CODIGO_BAIXA', 'CODIGO BAIXA', 'BAIXA', 'CD_BAIXA_OS']) || ''
        ).trim();

        const razaoCancelamentoVal = String(
          getValue(r, [
            'RAZAO_CANCELAMENTO',
            'RAZAO CANCELAMENTO',
            'CD_RAZAO_CANC_ORDEM_SERVICO',
            'CD_RAZAO_CANC',
            'CD RAZAO CANC ORD SERVICO',
            'RAZAO_CANC',
            'MOTIVO_CANCELAMENTO',
            'CD_MOTIVO_CANCELAMENTO',
            'CODIGO_CANCELAMENTO',
            'CODIGO CANCELAMENTO'
          ]) || '-'
        ).trim();

        // Fallback: If CD_BAIXA is missing or N/A, use RAZAO_CANCELAMENTO if available
        let codigoBaixaVal = rawCodigoBaixa;
        if (
          !codigoBaixaVal ||
          codigoBaixaVal.toUpperCase() === 'N/A' ||
          codigoBaixaVal === '-' ||
          codigoBaixaVal.toLowerCase() === 'null'
        ) {
          if (
            razaoCancelamentoVal &&
            razaoCancelamentoVal !== '-' &&
            razaoCancelamentoVal.toUpperCase() !== 'N/A' &&
            razaoCancelamentoVal.toLowerCase() !== 'null'
          ) {
            codigoBaixaVal = razaoCancelamentoVal;
          } else {
            codigoBaixaVal = 'N/A';
          }
        }

        const nodeVal = String(
          getValue(r, ['NODE', 'NM_NODE', 'Node', 'NODE_ID', 'DESPACHO_NODE']) || 'N/A'
        ).trim().toUpperCase();

        const contratoVal = String(
          getValue(r, ['CONTRATO', 'NR_CONTRATO', 'Contrato', 'CD_CONTRATO']) || `${100000 + idx}`
        ).trim();

        // Date parsing
        let dateVal = '01/08';
        const rawDate = getValue(r, ['DATA', 'DATA_ABERTURA', 'DATA_EXECUCAO', 'Data', 'DATA_OS']);
        if (rawDate) {
          if (typeof rawDate === 'number' && rawDate > 40000 && rawDate < 50000) {
            const dObj = new Date((rawDate - 25569) * 86400 * 1000 + 12 * 60 * 60 * 1000);
            if (!isNaN(dObj.getTime())) {
              const d = String(dObj.getUTCDate()).padStart(2, '0');
              const m = String(dObj.getUTCMonth() + 1).padStart(2, '0');
              dateVal = `${d}/${m}`;
            }
          } else {
            const strD = String(rawDate).trim();
            const dmy = strD.match(/^(\d{1,2})[-/. ](\d{1,2})/);
            if (dmy) {
              dateVal = `${dmy[1].padStart(2, '0')}/${dmy[2].padStart(2, '0')}`;
            }
          }
        }

        return {
          contrato: contratoVal,
          municipio: municipioVal,
          tipoOs: tipoOsVal,
          areaDespacho: areaVal,
          statusOs: statusVal,
          log: logVal,
          codigoBaixa: codigoBaixaVal,
          razaoCancelamento: razaoCancelamentoVal,
          node: nodeVal,
          data: dateVal,
        };
      });

      setData(parsedRows);
      handleClearFilters();
      setIsLoading(false);
      setError(null);
      return true;
    } catch (err: any) {
      console.error(err);
      setError('Erro ao ler a planilha Excel. Verifique o formato do arquivo.');
      setIsLoading(false);
      return false;
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    setError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const ab = evt.target?.result;
      if (ab instanceof ArrayBuffer) {
        processExcelBuffer(ab);
      } else {
        setError('Falha ao ler o conteúdo do arquivo.');
        setIsLoading(false);
      }
    };
    reader.onerror = () => {
      setError('Erro ao carregar o arquivo.');
      setIsLoading(false);
    };
    reader.readAsArrayBuffer(file);
  };

  // KPI Calculations
  const totalOSCount = filteredData.length;
  const logIncidentCount = useMemo(() => {
    return filteredData.filter((r) => r.log >= 3).length;
  }, [filteredData]);

  const notaLOGPercent = useMemo(() => {
    if (totalOSCount === 0) return 0;
    return (logIncidentCount / totalOSCount) * 100;
  }, [totalOSCount, logIncidentCount]);

  const maxLogValue = useMemo(() => {
    if (filteredData.length === 0) return 0;
    return Math.max(...filteredData.map((r) => r.log || 0));
  }, [filteredData]);

  const avgLogValue = useMemo(() => {
    if (filteredData.length === 0) return 0;
    const sum = filteredData.reduce((acc, curr) => acc + (curr.log || 0), 0);
    return sum / filteredData.length;
  }, [filteredData]);

  // Chart 1: LOG por Município (Nota LOG %)
  const groupByMunicipioData = useMemo(() => {
    const map = new Map<string, { total: number; logCount: number }>();
    filteredData.forEach((row) => {
      const key = row.municipio || 'OUTROS';
      if (!map.has(key)) map.set(key, { total: 0, logCount: 0 });
      const item = map.get(key)!;
      item.total += 1;
      if (row.log >= 3) item.logCount += 1;
    });

    return Array.from(map.entries())
      .map(([name, stat]) => ({
        name,
        total: stat.total,
        logCount: stat.logCount,
        nota: Number(((stat.logCount / stat.total) * 100).toFixed(2)),
      }))
      .sort((a, b) => b.nota - a.nota);
  }, [filteredData]);

  // Chart 2: LOG por Dia (Evolução Diária da Nota LOG %)
  const groupByDayData = useMemo(() => {
    const map = new Map<string, { total: number; logCount: number }>();
    filteredData.forEach((row) => {
      const key = row.data || '01/08';
      if (!map.has(key)) map.set(key, { total: 0, logCount: 0 });
      const item = map.get(key)!;
      item.total += 1;
      if (row.log >= 3) item.logCount += 1;
    });

    return Array.from(map.entries())
      .map(([name, stat]) => ({
        name,
        total: stat.total,
        logCount: stat.logCount,
        nota: Number(((stat.logCount / stat.total) * 100).toFixed(2)),
      }))
      .sort((a, b) => {
        const [dA, mA] = a.name.split('/').map(Number);
        const [dB, mB] = b.name.split('/').map(Number);
        if (mA !== mB) return (mA || 0) - (mB || 0);
        return (dA || 0) - (dB || 0);
      });
  }, [filteredData]);

  // Chart 3: LOG por Área de Despacho
  const groupByAreaData = useMemo(() => {
    const map = new Map<string, { total: number; logCount: number }>();
    filteredData.forEach((row) => {
      const key = row.areaDespacho || 'SEM ÁREA';
      if (!map.has(key)) map.set(key, { total: 0, logCount: 0 });
      const item = map.get(key)!;
      item.total += 1;
      if (row.log >= 3) item.logCount += 1;
    });

    return Array.from(map.entries())
      .map(([name, stat]) => ({
        name,
        total: stat.total,
        logCount: stat.logCount,
        nota: Number(((stat.logCount / stat.total) * 100).toFixed(2)),
      }))
      .sort((a, b) => b.nota - a.nota);
  }, [filteredData]);

  // Chart 4: Códigos de Baixa (para casos LOG >= 3)
  const codigosBaixaData = useMemo(() => {
    const map = new Map<string, number>();
    filteredData
      .filter((r) => r.log >= 3)
      .forEach((row) => {
        let key = row.codigoBaixa;
        if (
          (!key || key.toUpperCase() === 'N/A' || key === 'SEM BAIXA' || key === '-' || key.toLowerCase() === 'null') &&
          row.razaoCancelamento &&
          row.razaoCancelamento !== '-' &&
          row.razaoCancelamento.toUpperCase() !== 'N/A' &&
          row.razaoCancelamento.toLowerCase() !== 'null'
        ) {
          key = row.razaoCancelamento;
        }
        if (!key || key === 'SEM BAIXA' || key === '-') key = 'N/A';
        map.set(key, (map.get(key) || 0) + 1);
      });

    return Array.from(map.entries())
      .map(([name, total]) => ({ name, total }))
      .sort((a, b) => b.total - a.total)
      .slice(0, 10);
  }, [filteredData]);

  // Chart 5: Top 15 Nodes com Maiores Logs (Volume de LOG >= 3)
  const top15Nodes = useMemo(() => {
    const map = new Map<string, { total: number; sumLog: number; logCount: number }>();
    filteredData.forEach((row) => {
      const key = row.node || 'N/A';
      if (key === 'N/A') return;
      if (!map.has(key)) map.set(key, { total: 0, sumLog: 0, logCount: 0 });
      const item = map.get(key)!;
      item.total += 1;
      item.sumLog += row.log || 0;
      if (row.log >= 3) item.logCount += 1;
    });

    return Array.from(map.entries())
      .map(([name, stat]) => ({
        name,
        total: stat.logCount,
        totalOS: stat.total,
        avgLog: Number((stat.sumLog / stat.total).toFixed(1)),
      }))
      .filter((item) => item.total > 0)
      .sort((a, b) => b.total - a.total)
      .slice(0, 15);
  }, [filteredData]);

  // Excel Pivot Table Breakdown View (Matching Screenshot)
  const pivotTableData = useMemo(() => {
    // Find highest LOG level in filtered dataset to dynamically list all columns
    let maxLvl = 5;
    filteredData.forEach((r) => {
      const l = r.log || 0;
      if (l > maxLvl) maxLvl = l;
    });

    const allLevels = Array.from({ length: maxLvl + 1 }, (_, i) => i);

    const map = new Map<
      string,
      {
        municipio: string;
        logs: { [logLevel: number]: number };
        total: number;
        logIncidentCount: number;
      }
    >();

    filteredData.forEach((row) => {
      const muni = row.municipio || 'OUTROS';
      if (!map.has(muni)) {
        map.set(muni, { municipio: muni, logs: {}, total: 0, logIncidentCount: 0 });
      }
      const entry = map.get(muni)!;
      const l = row.log || 0;
      entry.logs[l] = (entry.logs[l] || 0) + 1;
      entry.total += 1;
      if (l >= 3) {
        entry.logIncidentCount += 1;
      }
    });

    const rows = Array.from(map.values())
      .map((e) => ({
        ...e,
        notaLog: Number(((e.logIncidentCount / e.total) * 100).toFixed(2)),
      }))
      .sort((a, b) => a.municipio.localeCompare(b.municipio, 'pt-BR'));

    const totalsByLevel: { [l: number]: number } = {};
    let overallTotal = 0;
    let overallLogIncidents = 0;

    rows.forEach((r) => {
      overallTotal += r.total;
      overallLogIncidents += r.logIncidentCount;
      Object.entries(r.logs).forEach(([lvl, cnt]) => {
        const numLvl = Number(lvl);
        totalsByLevel[numLvl] = (totalsByLevel[numLvl] || 0) + cnt;
      });
    });

    const overallNotaLog = overallTotal > 0 ? Number(((overallLogIncidents / overallTotal) * 100).toFixed(2)) : 0;

    return {
      allLevels,
      rows,
      totalsByLevel,
      overallTotal,
      overallLogIncidents,
      overallNotaLog,
    };
  }, [filteredData]);

  // Table Searched & Paginated
  const searchedData = useMemo(() => {
    if (!searchFilter.trim()) return filteredData;
    const term = searchFilter.toLowerCase();
    return filteredData.filter((r) => {
      return (
        r.contrato.toLowerCase().includes(term) ||
        r.municipio.toLowerCase().includes(term) ||
        r.tipoOs.toLowerCase().includes(term) ||
        r.areaDespacho.toLowerCase().includes(term) ||
        r.statusOs.toLowerCase().includes(term) ||
        r.codigoBaixa.toLowerCase().includes(term) ||
        (r.razaoCancelamento && r.razaoCancelamento.toLowerCase().includes(term)) ||
        r.node.toLowerCase().includes(term) ||
        String(r.log).includes(term)
      );
    });
  }, [filteredData, searchFilter]);

  const totalPages = Math.ceil(searchedData.length / rowsPerPage) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return searchedData.slice(start, start + rowsPerPage);
  }, [searchedData, currentPage, rowsPerPage]);

  const handleResetToDefault = () => {
    setData(generateInitialMockLOGData());
    handleClearFilters();
    setError(null);
  };

  const handleExportCSV = () => {
    const csvHeader = 'CONTRATO;MUNICIPIO;TIPO_OS;AREA_DESPACHO;STATUS_OS;LOG;CODIGO_BAIXA;RAZAO_CANCELAMENTO;NODE;DATA\n';
    const csvRows = filteredData
      .map(
        (r) =>
          `"${r.contrato}";"${r.municipio}";"${r.tipoOs}";"${r.areaDespacho}";"${r.statusOs}";${r.log};"${r.codigoBaixa}";"${r.razaoCancelamento || '-'}";"${r.node}";"${r.data}"`
      )
      .join('\n');

    const blob = new Blob([csvHeader + csvRows], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `analise_log_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="w-full space-y-8 pb-16">
      {/* Action Header Card - Claro Corporate Standard */}
      <section className="max-w-7xl mx-auto">
        <div className="bg-white p-6 rounded-3xl shadow-md border border-slate-100 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-14 h-14 bg-red-50 rounded-2xl flex items-center justify-center text-[#EE1D23] shadow-inner">
              <Layers className="w-7 h-7" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded-md bg-[#EE1D23] text-white text-[10px] font-black uppercase tracking-wider">
                  MÓDULO
                </span>
                <h1 className="text-2xl font-black text-[#333333] tracking-tight uppercase italic">
                  PAINEL DE REINCIDÊNCIA & LOG OPERACIONAL
                </h1>
              </div>
              <p className="text-xs font-bold text-slate-400 mt-0.5">
                Acompanhamento e Análise de Reincidência Operacional (LOG 0 a LOG 7+) por Cidade e Área
              </p>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={() => {
                const preConfiguredUrl = (import.meta as any).env?.VITE_GITHUB_EXCEL_URL_LOG || (import.meta as any).env?.VITE_GITHUB_EXCEL_URL;
                if (preConfiguredUrl) {
                  handleGithubLoad(preConfiguredUrl);
                } else {
                  setShowGithubInput(!showGithubInput);
                }
              }}
              className="flex items-center gap-2 bg-[#EE1D23] hover:bg-red-600 text-white font-black py-2.5 px-4 rounded-xl transition-all shadow-md shadow-red-500/15 active:scale-95 uppercase italic text-xs cursor-pointer"
              title="Sincronizar planilha com o repositório GitHub"
            >
              <Activity className="w-3.5 h-3.5" />
              <span>Sincronizar GitHub</span>
            </button>

            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-800 font-black py-2.5 px-4 rounded-xl border border-slate-200 transition-all shadow-xs active:scale-95 uppercase italic text-xs cursor-pointer"
              title="Importar planilha Excel (.xlsx, .xls)"
            >
              <Upload className="w-3.5 h-3.5 text-[#EE1D23]" />
              <span>Importar Excel</span>
            </button>

            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx, .xls"
              onChange={handleFileUpload}
              className="hidden"
            />

            <button
              onClick={handleExportCSV}
              className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all shadow-xs active:scale-95 uppercase italic text-xs cursor-pointer"
              title="Exportar dados filtrados para CSV/Excel"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span>Exportar</span>
            </button>

            <button
              onClick={handleResetToDefault}
              className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all active:scale-95 text-xs cursor-pointer"
              title="Restaurar dados padrão de exemplo"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Restaurar</span>
            </button>

            {data.length > 0 && (
              <button
                onClick={clearAllData}
                className="flex items-center gap-2 bg-red-50 hover:bg-red-100 text-red-600 font-bold py-2.5 px-4 rounded-xl border border-red-100 transition-all active:scale-95 text-xs cursor-pointer"
                title="Limpar todos os dados carregados"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Limpar</span>
              </button>
            )}
          </div>
        </div>

        {/* GitHub Input Expansion */}
        <AnimatePresence>
          {showGithubInput && (
            <motion.div 
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              className="mt-4 overflow-hidden"
            >
              <div className="bg-white p-4 rounded-2xl border-2 border-slate-100 shadow-sm flex flex-col md:flex-row gap-3">
                <div className="flex-1 relative">
                  <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                  <input 
                    type="text" 
                    placeholder="Cole o link do arquivo Excel LOG no GitHub (ex: https://github.com/usuario/repo/blob/main/log.xlsx)"
                    value={githubUrl}
                    onChange={(e) => setGithubUrl(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl pl-11 pr-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
                  />
                </div>
                <button 
                  onClick={() => handleGithubLoad()}
                  disabled={!githubUrl || isGithubLoading}
                  className="bg-[#EE1D23] hover:bg-[#D1191F] disabled:bg-slate-300 text-white font-black py-3 px-8 rounded-xl transition-all shadow-lg shadow-red-500/20 active:scale-95 uppercase italic text-sm flex items-center justify-center gap-2 cursor-pointer"
                >
                  {isGithubLoading ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Check className="w-4 h-4" />}
                  Carregar Planilha
                </button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </section>

      {/* Error Banner */}
      {error && (
        <div className="bg-red-50 border border-red-200 p-5 rounded-[32px] text-red-700 flex items-start justify-between gap-4 shadow-sm">
          <div className="flex items-start gap-3">
            <XCircle className="w-5 h-5 shrink-0 mt-0.5 text-red-600" />
            <div>
              <h4 className="font-black uppercase tracking-wider text-xs mb-1">Aviso do Importador</h4>
              <p className="text-xs font-bold leading-relaxed">{error}</p>
            </div>
          </div>
          <button
            onClick={() => setError(null)}
            className="p-1.5 rounded-xl hover:bg-red-100 text-red-500 hover:text-red-700 transition-colors shrink-0"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* Loading Indicator */}
      {isLoading && (
        <div className="bg-white border border-slate-200 p-8 rounded-[32px] shadow-md flex flex-col items-center justify-center">
          <RefreshCw className="w-8 h-8 text-[#EE1D23] animate-spin mb-3" />
          <p className="text-sm font-black text-slate-700 uppercase italic">Carregando e Processando Dados LOG...</p>
        </div>
      )}

      {data.length === 0 ? (
        <div className="bg-white rounded-[32px] border border-slate-200/80 shadow-sm p-12 text-center flex flex-col items-center justify-center my-6">
          <div className="w-14 h-14 bg-red-50 rounded-2xl flex items-center justify-center mb-4">
            <FileSpreadsheet className="w-7 h-7 text-[#EE1D23]" />
          </div>
          <h3 className="text-xl font-black text-slate-800 uppercase italic tracking-tight mb-2">
            Nenhum Dado LOG Carregado
          </h3>
          <p className="text-xs text-slate-400 max-w-md mb-6 leading-relaxed font-bold uppercase tracking-wider">
            Sincronize com o GitHub ou importe a planilha Excel (BASE_LOG.xlsx) para visualizar as métricas de qualidade e log de chamados.
          </p>
          <div className="flex flex-wrap items-center justify-center gap-3">
            <button
              onClick={() => {
                const preConfiguredUrl = (import.meta as any).env?.VITE_GITHUB_EXCEL_URL_LOG || (import.meta as any).env?.VITE_GITHUB_EXCEL_URL;
                if (preConfiguredUrl) {
                  handleGithubLoad(preConfiguredUrl);
                } else {
                  setShowGithubInput(true);
                }
              }}
              className="flex items-center gap-2 bg-[#EE1D23] hover:bg-red-600 text-white font-black py-2.5 px-5 rounded-xl transition-all shadow-md shadow-red-500/15 active:scale-95 uppercase italic text-xs cursor-pointer"
            >
              <Activity className="w-3.5 h-3.5" />
              <span>Sincronizar GitHub</span>
            </button>
            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-800 font-black py-2.5 px-5 rounded-xl border border-slate-200 transition-all shadow-2xs active:scale-95 uppercase italic text-xs cursor-pointer"
            >
              <Upload className="w-3.5 h-3.5 text-[#EE1D23]" />
              <span>Importar Excel</span>
            </button>
          </div>
        </div>
      ) : (
        <>
      {/* 4 Core Filters Panel */}
      <div className="bg-white rounded-[32px] p-6 border border-slate-200/80 shadow-xs">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6 pb-4 border-b border-slate-100">
          <div className="flex items-center gap-2">
            <SlidersHorizontal className="w-5 h-5 text-[#EE1D23]" />
            <h3 className="text-sm font-black text-slate-800 uppercase italic tracking-tight">
              Filtros do Painel LOG
            </h3>
            {activeFilterCount > 0 && (
              <span className="bg-red-100 text-[#EE1D23] text-[10px] font-black px-2 py-0.5 rounded-full">
                {activeFilterCount} ativos
              </span>
            )}
          </div>

          {activeFilterCount > 0 && (
            <button
              onClick={handleClearFilters}
              className="text-xs font-bold text-slate-400 hover:text-[#EE1D23] uppercase tracking-wider flex items-center gap-1 transition-colors"
            >
              <X className="w-3.5 h-3.5" />
              Limpar Filtros
            </button>
          )}
        </div>

        {/* Filter Dropdowns Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-4">
          <MultiFilterSelect
            label="MUNICÍPIO"
            icon={<MapPin className="w-3.5 h-3.5" />}
            value={filters.municipio}
            options={uniqueOptions.municipio}
            onChange={(val) => setFilters((prev) => ({ ...prev, municipio: val.filter(v => v !== 'Todos') }))}
          />

          <MultiFilterSelect
            label="TIPO OS"
            icon={<Layers className="w-3.5 h-3.5" />}
            value={filters.tipoOs}
            options={uniqueOptions.tipoOs}
            onChange={(val) => setFilters((prev) => ({ ...prev, tipoOs: val.filter(v => v !== 'Todos') }))}
          />

          <MultiFilterSelect
            label="ÁREA DESPACHO"
            icon={<Building2 className="w-3.5 h-3.5" />}
            value={filters.areaDespacho}
            options={uniqueOptions.areaDespacho}
            onChange={(val) => setFilters((prev) => ({ ...prev, areaDespacho: val.filter(v => v !== 'Todos') }))}
          />

          <MultiFilterSelect
            label="STATUS OS"
            icon={<Activity className="w-3.5 h-3.5" />}
            value={filters.statusOs}
            options={uniqueOptions.statusOs}
            onChange={(val) => setFilters((prev) => ({ ...prev, statusOs: val.filter(v => v !== 'Todos') }))}
          />

          <MultiFilterSelect
            label="NÍVEL LOG"
            icon={<Database className="w-3.5 h-3.5" />}
            value={filters.logLevel}
            options={uniqueOptions.logLevel}
            getOptionLabel={(opt) => (opt === 'Todos' ? 'Todos' : `LOG ${opt}`)}
            onChange={(val) => setFilters((prev) => ({ ...prev, logLevel: val.filter(v => v !== 'Todos') }))}
          />
        </div>
      </div>

      {/* KPI Highlight Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* NOTA LOG (%) CARD */}
        <div className="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200/80 shadow-xs relative overflow-hidden">
          <div className="flex justify-between items-start mb-3">
            <div>
              <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">
                NOTA LOG (%)
              </span>
              <p className="text-[11px] text-slate-400 font-bold mt-0.5">Taxa LOG ≥ 3 / Total OS</p>
            </div>
            <div className="w-9 h-9 rounded-xl bg-red-50 flex items-center justify-center text-[#EE1D23]">
              <TrendingUp className="w-4 h-4" />
            </div>
          </div>
          <div className="flex items-baseline gap-2">
            <span className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight font-mono">
              {formatPercent(notaLOGPercent)}
            </span>
          </div>
          <div className="mt-2.5 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500 font-medium">Classificação</span>
            <span
              className={cn(
                'font-black px-2 py-0.5 rounded-full text-[9px] uppercase',
                notaLOGPercent <= 2.5
                  ? 'bg-emerald-100 text-emerald-800'
                  : notaLOGPercent <= 4.0
                  ? 'bg-amber-100 text-amber-800'
                  : 'bg-red-100 text-red-800'
              )}
            >
              {notaLOGPercent <= 2.5 ? 'Excelente' : notaLOGPercent <= 4.0 ? 'Atenção' : 'Crítico'}
            </span>
          </div>
        </div>

        {/* LOG INCIDENTS COUNT */}
        <div className="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200/80 shadow-xs">
          <div className="flex justify-between items-start mb-3">
            <div>
              <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">
                OS REINCIDENTES (LOG ≥ 3)
              </span>
              <p className="text-[11px] text-slate-400 font-bold mt-0.5">Qtd de Casos de Reincidência</p>
            </div>
            <div className="w-9 h-9 rounded-xl bg-amber-50 flex items-center justify-center text-amber-600">
              <AlertTriangle className="w-4 h-4" />
            </div>
          </div>
          <span className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight font-mono">
            {logIncidentCount.toLocaleString('pt-BR')}
          </span>
          <div className="mt-2.5 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500 font-medium">Proporção no Total</span>
            <span className="font-bold text-slate-700 font-mono">
              {totalOSCount > 0 ? ((logIncidentCount / totalOSCount) * 100).toFixed(1) : 0}%
            </span>
          </div>
        </div>

        {/* TOTAL OS ANALISADAS */}
        <div className="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200/80 shadow-xs">
          <div className="flex justify-between items-start mb-3">
            <div>
              <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">
                TOTAL OS ANALISADAS
              </span>
              <p className="text-[11px] text-slate-400 font-bold mt-0.5">Base Total Filtrada</p>
            </div>
            <div className="w-9 h-9 rounded-xl bg-slate-100 flex items-center justify-center text-slate-600">
              <Layers className="w-4 h-4" />
            </div>
          </div>
          <span className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight font-mono">
            {totalOSCount.toLocaleString('pt-BR')}
          </span>
          <div className="mt-2.5 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500 font-medium">Média de LOG por OS</span>
            <span className="font-bold text-slate-700 font-mono">{avgLogValue.toFixed(2)}</span>
          </div>
        </div>

        {/* LOG MÁXIMO */}
        <div className="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200/80 shadow-xs">
          <div className="flex justify-between items-start mb-3">
            <div>
              <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">
                REINCIDÊNCIA MÁXIMA
              </span>
              <p className="text-[11px] text-slate-400 font-bold mt-0.5">Maior Nível de LOG</p>
            </div>
            <div className="w-9 h-9 rounded-xl bg-red-50 flex items-center justify-center text-[#EE1D23]">
              <Database className="w-4 h-4" />
            </div>
          </div>
          <span className="text-2xl sm:text-3xl font-black text-[#EE1D23] tracking-tight font-mono">
            LOG {maxLogValue}
          </span>
          <div className="mt-2.5 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500 font-medium">Status da Operação</span>
            <span className="font-bold text-emerald-600">Ativa</span>
          </div>
        </div>
      </div>

      {/* PIVOT TABLE BREAKDOWN VIEW */}
      <div className="bg-white rounded-[28px] p-5 sm:p-6 border border-slate-200/80 shadow-sm">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-5 pb-4 border-b border-slate-100">
          <div>
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-red-50 border border-red-100 text-[#EE1D23] text-[10px] font-black uppercase tracking-wider mb-1.5">
              <Table className="w-3.5 h-3.5" />
              Resumo Analítico por Nível de LOG
            </div>
            <h3 className="text-lg sm:text-xl font-black text-slate-900 uppercase italic tracking-tight">
              LOG POR MUNICÍPIO
            </h3>
            <p className="text-xs text-slate-500 font-medium mt-0.5">
              Distribuição e contagem de Ordens de Serviço (OS) segmentadas por município e nível de reincidência.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <span className="inline-flex items-center gap-1.5 text-[11px] font-black text-red-700 bg-red-50 px-3.5 py-1.5 rounded-xl border border-red-200 shadow-2xs">
              <AlertTriangle className="w-3.5 h-3.5 text-[#EE1D23]" />
              LOG ≥ 3 = Reincidência Operacional
            </span>
          </div>
        </div>

        <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white shadow-2xs">
          <table className="w-full text-left border-collapse min-w-[800px]">
            <thead>
              <tr className="bg-slate-900 text-white text-[11px] font-black uppercase tracking-wider">
                <th className="py-3 px-4 border-b border-slate-800">CIDADE</th>
                {pivotTableData.allLevels.map((lvl) => (
                  <th
                    key={lvl}
                    className={cn(
                      "py-3 px-2 border-b border-slate-800 text-center min-w-[48px]",
                      lvl >= 3 ? "bg-red-950/90 text-red-300 font-black border-red-900/60" : "text-slate-200"
                    )}
                  >
                    LOG {lvl}
                  </th>
                ))}
                <th className="py-3 px-4 border-b border-slate-800 text-right bg-slate-950 text-slate-100 font-black">
                  Total Geral
                </th>
                <th className="py-3 px-4 border-b border-slate-800 text-right bg-[#EE1D23] text-white font-black">
                  LOG% (≥3)
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs font-bold text-slate-700">
              {pivotTableData.rows.map((row, idx) => (
                <tr
                  key={row.municipio}
                  className={cn(
                    "transition-colors",
                    idx % 2 === 0 ? "bg-white" : "bg-slate-50/60",
                    "hover:bg-red-50/20"
                  )}
                >
                  <td className="py-2.5 px-4 font-black text-slate-800 uppercase tracking-tight">
                    {row.municipio}
                  </td>
                  {pivotTableData.allLevels.map((lvl) => {
                    const val = row.logs[lvl] || 0;
                    const isHigh = lvl >= 3;
                    return (
                      <td
                        key={lvl}
                        className={cn(
                          "py-2.5 px-2 text-center font-mono text-xs",
                          isHigh && val > 0
                            ? "font-black text-red-700 bg-red-50/90"
                            : val > 0
                            ? "text-slate-800 font-bold"
                            : "text-slate-300 font-normal"
                        )}
                      >
                        {val > 0 ? val.toLocaleString('pt-BR') : '-'}
                      </td>
                    );
                  })}
                  <td className="py-2.5 px-4 text-right font-mono font-black text-slate-900 bg-slate-100/60 text-xs">
                    {row.total.toLocaleString('pt-BR')}
                  </td>
                  <td className="py-2.5 px-4 text-right font-mono font-black text-red-600 bg-red-50/80 text-xs">
                    {formatPercent(row.notaLog)}
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="bg-slate-900 text-white font-black text-xs uppercase tracking-wider">
                <td className="py-3 px-4">TOTAL GERAL</td>
                {pivotTableData.allLevels.map((lvl) => (
                  <td
                    key={lvl}
                    className={cn(
                      "py-3 px-2 text-center font-mono",
                      lvl >= 3 ? "text-red-400 font-black bg-red-950/40" : "text-slate-200"
                    )}
                  >
                    {(pivotTableData.totalsByLevel[lvl] || 0).toLocaleString('pt-BR')}
                  </td>
                ))}
                <td className="py-3 px-4 text-right font-mono text-sm text-amber-400 bg-slate-950">
                  {pivotTableData.overallTotal.toLocaleString('pt-BR')}
                </td>
                <td className="py-3 px-4 text-right font-mono text-sm bg-[#EE1D23] text-white">
                  {formatPercent(pivotTableData.overallNotaLog)}
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* CHARTS GRID SECTION */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* CHART 1: LOG POR MUNICÍPIO */}
        <div className="bg-white rounded-[32px] p-6 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="mb-4">
            <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
              AT5 LOG POR MUNICÍPIO
            </span>
            <h3 className="text-base font-black text-slate-800 uppercase italic tracking-tight mt-1">
              Nota LOG % por Município
            </h3>
          </div>

          <div className="h-[360px] w-full overflow-x-auto pb-2">
            <div style={{ width: '100%', minWidth: `${Math.max(500, groupByMunicipioData.length * 45)}px` }} className="h-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={groupByMunicipioData} margin={{ top: 25, right: 15, left: 5, bottom: 60 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis
                    dataKey="name"
                    stroke="#334155"
                    fontSize={11}
                    fontWeight="bold"
                    tickLine={false}
                    axisLine={false}
                    interval={0}
                    angle={-35}
                    textAnchor="end"
                    height={65}
                  />
                  <YAxis
                    stroke="#64748b"
                    fontSize={11}
                    fontWeight="bold"
                    tickLine={false}
                    axisLine={false}
                    tickFormatter={(v) => `${v}%`}
                  />
                  <Tooltip
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ backgroundColor: '#1e293b', borderRadius: '16px', border: 'none', color: '#fff' }}
                    labelStyle={{ fontWeight: 'black', textTransform: 'uppercase', fontSize: '12px', color: '#EE1D23' }}
                    formatter={(val: number) => [formatPercent(val), 'Nota LOG (%)']}
                  />
                  <Bar name="Nota LOG %" dataKey="nota" fill="#EE1D23" radius={[8, 8, 0, 0]} barSize={26}>
                    {groupByMunicipioData.map((d, i) => (
                      <Cell
                        key={`cell-muni-${i}`}
                        fill={d.nota >= 3.0 ? '#EE1D23' : d.nota >= 2.0 ? '#f59e0b' : '#10b981'}
                      />
                    ))}
                    <LabelList
                      dataKey="nota"
                      position="top"
                      formatter={(val: number) => formatPercent(val)}
                      style={{ fontSize: 10, fontWeight: '900', fill: '#0f172a' }}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {/* CHART 2: LOG POR DIA (EVOLUÇÃO DIÁRIA) */}
        <div className="bg-white rounded-[32px] p-6 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="mb-4">
            <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
              AT5 LOG POR DIA
            </span>
            <h3 className="text-base font-black text-slate-800 uppercase italic tracking-tight mt-1">
              Evolução Diária da Nota LOG %
            </h3>
          </div>

          <div className="h-[360px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={groupByDayData} margin={{ top: 25, right: 25, left: 5, bottom: 25 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="name" stroke="#334155" fontSize={11} fontWeight="bold" tickLine={false} axisLine={false} />
                <YAxis
                  stroke="#64748b"
                  fontSize={11}
                  fontWeight="bold"
                  tickLine={false}
                  axisLine={false}
                  tickFormatter={(v) => `${v}%`}
                />
                <Tooltip
                  contentStyle={{ backgroundColor: '#1e293b', borderRadius: '16px', border: 'none', color: '#fff' }}
                  labelStyle={{ fontWeight: 'black', textTransform: 'uppercase', fontSize: '12px', color: '#EE1D23' }}
                  formatter={(val: number) => [formatPercent(val), 'Nota LOG (%)']}
                />
                <Line
                  type="monotone"
                  dataKey="nota"
                  stroke="#EE1D23"
                  strokeWidth={3}
                  dot={{ r: 5, fill: '#EE1D23', strokeWidth: 2, stroke: '#fff' }}
                  activeDot={{ r: 8, fill: '#EE1D23' }}
                >
                  <LabelList
                    dataKey="nota"
                    position="top"
                    formatter={(val: number) => formatPercent(val)}
                    style={{ fontSize: 11, fontWeight: '900', fill: '#0f172a' }}
                  />
                </Line>
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* CHART 3: LOG POR ÁREA DE DESPACHO (FULL WIDTH DESTAQUE COM ESPAÇAMENTO PARA TODAS AS NOTAS) */}
        <div className="bg-white rounded-[32px] p-6 border border-slate-200/80 shadow-xs flex flex-col justify-between lg:col-span-2">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 mb-4">
            <div>
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                AT5 LOG POR ÁREA
              </span>
              <h3 className="text-lg font-black text-slate-800 uppercase italic tracking-tight mt-0.5">
                Nota LOG % por Área de Despacho (Visão Completa)
              </h3>
            </div>
            <span className="text-[11px] font-bold text-slate-400 bg-slate-50 px-3 py-1 rounded-xl border border-slate-200">
              Total de {groupByAreaData.length} áreas analisadas
            </span>
          </div>

          <div className="h-[380px] w-full overflow-x-auto pb-2">
            <div style={{ width: '100%', minWidth: `${Math.max(800, groupByAreaData.length * 58)}px` }} className="h-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={groupByAreaData} margin={{ top: 30, right: 20, left: 10, bottom: 75 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis
                    dataKey="name"
                    stroke="#334155"
                    fontSize={11}
                    fontWeight="bold"
                    tickLine={false}
                    axisLine={false}
                    interval={0}
                    angle={-38}
                    textAnchor="end"
                    height={75}
                  />
                  <YAxis
                    stroke="#64748b"
                    fontSize={11}
                    fontWeight="bold"
                    tickLine={false}
                    axisLine={false}
                    tickFormatter={(v) => `${v}%`}
                  />
                  <Tooltip
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ backgroundColor: '#1e293b', borderRadius: '16px', border: 'none', color: '#fff' }}
                    labelStyle={{ fontWeight: 'black', textTransform: 'uppercase', fontSize: '12px', color: '#EE1D23' }}
                    formatter={(val: number) => [formatPercent(val), 'Nota LOG (%)']}
                  />
                  <Bar name="Nota LOG %" dataKey="nota" fill="#EE1D23" radius={[8, 8, 0, 0]} barSize={28}>
                    {groupByAreaData.map((d, i) => (
                      <Cell
                        key={`cell-area-${i}`}
                        fill={d.nota >= 3.0 ? '#EE1D23' : d.nota >= 2.0 ? '#f59e0b' : '#10b981'}
                      />
                    ))}
                    <LabelList
                      dataKey="nota"
                      position="top"
                      dy={-6}
                      formatter={(val: number) => formatPercent(val)}
                      style={{ fontSize: 10, fontWeight: '900', fill: '#0f172a' }}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {/* CHART 4: CÓDIGOS DE BAIXA PARA LOG >= 3 */}
        <div className="bg-white rounded-[32px] p-6 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="mb-4">
            <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
              MOTIVOS DE REINCIDÊNCIA
            </span>
            <h3 className="text-base font-black text-slate-800 uppercase italic tracking-tight mt-1">
              Top Códigos de Baixa em OS com LOG ≥ 3
            </h3>
          </div>

          <div className="h-[360px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={codigosBaixaData} layout="vertical" margin={{ left: 10, right: 35, top: 5, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                <XAxis type="number" stroke="#64748b" fontSize={11} fontWeight="bold" tickLine={false} axisLine={false} />
                <YAxis
                  dataKey="name"
                  type="category"
                  stroke="#334155"
                  fontSize={11}
                  fontWeight="bold"
                  tickLine={false}
                  axisLine={false}
                  width={140}
                />
                <Tooltip
                  cursor={{ fill: '#f8fafc' }}
                  contentStyle={{ backgroundColor: '#1e293b', borderRadius: '16px', border: 'none', color: '#fff' }}
                  labelStyle={{ fontWeight: 'black', fontSize: '12px', color: '#EE1D23' }}
                  formatter={(val: number) => [val, 'Volume OS (LOG ≥ 3)']}
                />
                <Bar name="Volume Baixas" dataKey="total" fill="#EE1D23" radius={[0, 8, 8, 0]} barSize={20}>
                  <LabelList dataKey="total" position="right" style={{ fontSize: 11, fontWeight: 'bold', fill: '#0f172a' }} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* CHART 5: TOP 15 NODES COM MAIORES LOGS */}
      <div className="bg-white rounded-[32px] p-6 border border-slate-200/80 shadow-xs">
        <div className="mb-4">
          <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
            OFFENDERS DE REDE / NODES
          </span>
          <h3 className="text-base font-black text-slate-800 uppercase italic tracking-tight mt-1">
            Top 15 Nodes com Maior Volume de Reincidência (LOG ≥ 3)
          </h3>
        </div>

        <div className="h-[380px] w-full">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={top15Nodes} margin={{ top: 25, right: 10, left: 10, bottom: 40 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
              <XAxis
                dataKey="name"
                stroke="#334155"
                fontSize={11}
                fontWeight="bold"
                tickLine={false}
                axisLine={false}
                interval={0}
                angle={-35}
                textAnchor="end"
                height={55}
              />
              <YAxis stroke="#64748b" fontSize={11} fontWeight="bold" tickLine={false} axisLine={false} />
              <Tooltip
                cursor={{ fill: '#f8fafc' }}
                contentStyle={{ backgroundColor: '#1e293b', borderRadius: '16px', border: 'none', color: '#fff' }}
                labelStyle={{ fontWeight: 'black', textTransform: 'uppercase', fontSize: '12px', color: '#EE1D23' }}
                formatter={(val: number, name: string) => [val, 'Volume LOG ≥ 3']}
              />
              <Bar name="Volume Reincidência" dataKey="total" fill="#EE1D23" radius={[8, 8, 0, 0]} barSize={22}>
                <LabelList dataKey="total" position="top" style={{ fontSize: 11, fontWeight: '900', fill: '#0f172a' }} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* OS RECORDS DETAILED TABLE */}
      <div className="bg-white rounded-[32px] border border-slate-200/80 shadow-xs overflow-hidden">
        <div className="p-6 bg-slate-50 border-b border-slate-200/80 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
          <div>
            <h3 className="text-xs font-black text-slate-400 uppercase tracking-widest">Base de Registros LOG</h3>
            <p className="text-sm font-black text-slate-800 uppercase italic tracking-tight">
              {searchedData.length} registros correspondentes aos filtros
            </p>
          </div>

          <div className="relative w-full sm:w-72">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-3 pointer-events-none" />
            <input
              type="text"
              placeholder="Pesquisar por Contrato, Node, Baixa..."
              value={searchFilter}
              onChange={(e) => {
                setSearchFilter(e.target.value);
                setCurrentPage(1);
              }}
              className="w-full bg-white border border-slate-200 rounded-2xl pl-9 pr-4 py-2 text-xs font-bold text-slate-700 placeholder:text-slate-400 focus:outline-none focus:border-red-500 shadow-xs"
            />
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[800px]">
            <thead>
              <tr className="bg-slate-100 text-slate-500 text-[10px] font-black uppercase tracking-wider border-b border-slate-200">
                <th className="p-4">Contrato</th>
                <th className="p-4">Município</th>
                <th className="p-4">Tipo OS</th>
                <th className="p-4">Área Despacho</th>
                <th className="p-4">Status OS</th>
                <th className="p-4 text-center">Nível LOG</th>
                <th className="p-4">Código Baixa</th>
                <th className="p-4">Razão Cancelamento</th>
                <th className="p-4">Node</th>
                <th className="p-4">Data</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs font-bold text-slate-700">
              {paginatedData.map((row, idx) => {
                const isIncident = row.log >= 3;
                return (
                  <tr key={`${row.contrato}-${idx}`} className="hover:bg-slate-50/80 transition-colors">
                    <td className="p-4 font-mono text-slate-900 font-extrabold">{row.contrato}</td>
                    <td className="p-4 font-extrabold">{row.municipio}</td>
                    <td className="p-4 text-slate-600">{row.tipoOs}</td>
                    <td className="p-4 text-slate-600">{row.areaDespacho}</td>
                    <td className="p-4">
                      <span
                        className={cn(
                          'px-2.5 py-1 rounded-lg text-[10px] font-black uppercase',
                          row.statusOs.includes('EXECUTADA') || row.statusOs.includes('ENCERRADA')
                            ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                            : 'bg-amber-50 text-amber-700 border border-amber-200'
                        )}
                      >
                        {row.statusOs}
                      </span>
                    </td>
                    <td className="p-4 text-center">
                      <span
                        className={cn(
                          'px-2.5 py-1 rounded-full text-[11px] font-black font-mono inline-block min-w-[32px]',
                          isIncident
                            ? 'bg-red-600 text-white shadow-xs animate-pulse'
                            : 'bg-slate-100 text-slate-600'
                        )}
                      >
                        {row.log}
                      </span>
                    </td>
                    <td className="p-4 text-slate-600 text-[11px]">{row.codigoBaixa}</td>
                    <td className="p-4 text-slate-600 text-[11px]">
                      {row.razaoCancelamento && row.razaoCancelamento !== '-' && row.razaoCancelamento !== 'N/A' ? (
                        <span className="inline-block px-2.5 py-1 rounded-lg bg-red-50 text-red-700 border border-red-200/80 font-bold text-[10px]">
                          {row.razaoCancelamento}
                        </span>
                      ) : (
                        <span className="text-slate-300 font-normal">-</span>
                      )}
                    </td>
                    <td className="p-4 font-mono text-slate-800">{row.node}</td>
                    <td className="p-4 font-mono text-slate-500">{row.data}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Pagination Footer */}
        <div className="p-4 bg-slate-50 border-t border-slate-200/80 flex flex-col sm:flex-row justify-between items-center gap-4 text-xs font-bold text-slate-500">
          <span>
            Página <strong className="text-slate-800 font-mono">{currentPage}</strong> de{' '}
            <strong className="text-slate-800 font-mono">{totalPages}</strong>
          </span>

          <div className="flex gap-2">
            <button
              disabled={currentPage === 1}
              onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              className="px-4 py-2 rounded-xl bg-white border border-slate-200 hover:bg-slate-100 disabled:opacity-40 disabled:hover:bg-white text-slate-700 font-black uppercase text-[10px] transition-all"
            >
              Anterior
            </button>
            <button
              disabled={currentPage === totalPages}
              onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              className="px-4 py-2 rounded-xl bg-white border border-slate-200 hover:bg-slate-100 disabled:opacity-40 disabled:hover:bg-white text-slate-700 font-black uppercase text-[10px] transition-all"
            >
              Próxima
            </button>
          </div>
        </div>
      </div>
      </>
      )}
    </div>
  );
}
