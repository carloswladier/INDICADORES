import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  Activity,
  AlertTriangle,
  Building2,
  Calendar,
  Check,
  CheckCircle2,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Compass,
  Download,
  FileCheck,
  FileSpreadsheet,
  Filter,
  Layers,
  ListTodo,
  MapPin,
  Radio,
  RefreshCw,
  RotateCcw,
  Search,
  SlidersHorizontal,
  TrendingDown,
  TrendingUp,
  Trash2,
  Upload,
  UserMinus,
  Users,
  X,
  XCircle,
  Eye,
  Table
} from 'lucide-react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Cell,
  Line,
  ComposedChart,
  LabelList,
  Area
} from 'recharts';
import { MultiFilterSelect } from './MultiFilterSelect';
import { cn, formatPercent, formatDecimal } from '../lib/utils';
import { OutageEvent, OutageStatus, generateExactReferenceOutageData, cityNodesMap, CAT_PROD_2_DEFAULT, formatDisplayDateTime } from './OutageDashboard';
import { REAL_CITY_TOPOLOGY_MAP } from '../data/cityTopologyMap';
import { AT5Row, generateMockAT5Data } from './AT5Dashboard';
import { fetchGithubFileArrayBuffer, getGithubOutageUrl, normalizeGithubRawUrl } from '../lib/githubSync';
import { FileLoadingOverlay } from './FileLoadingOverlay';

export interface CrossedChurnRecord {
  rank: number;
  topologia: string;
  tipoEvento: string;
  catProd2: string;
  municipio: string;
  qtdOutage: number;
  qtdVisitas: number;
  totalVisitasPeriodo: number;
  visitasComPadrao: number;
  visitasSemPadrao: number;
  clientesAfetados: number;
  outageEvents: OutageEvent[];
  at5Visits: AT5Row[];
  allTopologyVisits?: AT5Row[];
  ratioVisitasOutage: number;
}

export interface ChurnDashboardProps {
  outageData?: OutageEvent[];
  at5Data?: AT5Row[];
  onOutageDataChange?: (data: OutageEvent[]) => void;
  onAt5DataChange?: (data: AT5Row[]) => void;
}

const MONTH_ORDER = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

// Helper normalization with high-performance cache
const normCache = new Map<string, string>();
const norm = (s: any): string => {
  if (s === null || s === undefined) return '';
  const str = typeof s === 'string' ? s.trim() : String(s).trim();
  if (!str) return '';
  const cached = normCache.get(str);
  if (cached !== undefined) return cached;
  const res = str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  if (normCache.size < 15000) {
    normCache.set(str, res);
  }
  return res;
};

// Pre-computed normalized lookup for REAL_CITY_TOPOLOGY_MAP
const NORM_CITY_TOPOLOGY_MAP: Record<string, string[]> = {};
for (const [k, v] of Object.entries(REAL_CITY_TOPOLOGY_MAP)) {
  NORM_CITY_TOPOLOGY_MAP[norm(k)] = v;
}

const isAllOrEmpty = (arr?: string[]) => !arr || arr.length === 0 || arr.includes('Todos');

// Helper to extract month name from an AT5 row (data can be "DD/MM", "DD/MM/YYYY", "YYYY-MM-DD", etc.)
const getAt5RowMonth = (row: AT5Row): string | null => {
  if (!row) return null;
  if ((row as any).mes && typeof (row as any).mes === 'string') {
    const rawMes = (row as any).mes.trim();
    if (MONTH_ORDER.includes(rawMes)) return rawMes;
    for (const mName of MONTH_ORDER) {
      if (rawMes.toLowerCase().includes(mName.toLowerCase())) return mName;
    }
  }

  if (!row.data) return null;
  const str = String(row.data).trim();
  if (!str) return null;

  // 1. Direct month name match
  for (const mName of MONTH_ORDER) {
    if (str.toLowerCase() === mName.toLowerCase()) return mName;
  }

  // 2. Format DD/MM or DD/MM/YYYY
  const dmyMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})(?:[-/.](\d{2,4}))?/);
  if (dmyMatch) {
    const m = parseInt(dmyMatch[2], 10);
    if (m >= 1 && m <= 12) {
      return MONTH_ORDER[m - 1];
    }
  }

  // 3. Format YYYY-MM-DD
  const ymdMatch = str.match(/^(\d{4})[-/.](\d{2})[-/.](\d{2})/);
  if (ymdMatch) {
    const m = parseInt(ymdMatch[2], 10);
    if (m >= 1 && m <= 12) {
      return MONTH_ORDER[m - 1];
    }
  }

  // 4. Check if Excel numeric serial date (e.g. 45444 for ~June 2024)
  const num = Number(str);
  if (!isNaN(num) && num > 40000 && num < 50000) {
    try {
      const dateObj = new Date((num - 25569) * 86400 * 1000 + 12 * 60 * 60 * 1000);
      if (!isNaN(dateObj.getTime())) {
        const m = dateObj.getUTCMonth() + 1;
        if (m >= 1 && m <= 12) return MONTH_ORDER[m - 1];
      }
    } catch (e) {}
  }

  // 5. Try Date.parse
  const ts = Date.parse(str);
  if (!isNaN(ts)) {
    const dateObj = new Date(ts);
    const m = dateObj.getUTCMonth() + 1;
    if (m >= 1 && m <= 12) {
      return MONTH_ORDER[m - 1];
    }
  }

  return null;
};

// Helper to extract ISO date YYYY-MM-DD from an AT5 row for date-range comparison and correlation
const getAt5IsoDate = (row: AT5Row): string | null => {
  if (!row) return null;
  if (row.dataIso) return row.dataIso;
  const str = String(row.aberturaSolic || row.data || '').trim();
  if (!str) return null;

  const num = Number(str);
  if (!isNaN(num) && num > 40000 && num < 60000) {
    try {
      const dateObj = new Date((num - 25569) * 86400 * 1000 + 12 * 60 * 60 * 1000);
      if (!isNaN(dateObj.getTime())) {
        return dateObj.toISOString().slice(0, 10);
      }
    } catch (e) {}
  }

  const ymdMatch = str.match(/^(\d{4})[-/.](\d{2})[-/.](\d{2})/);
  if (ymdMatch) {
    return `${ymdMatch[1]}-${ymdMatch[2]}-${ymdMatch[3]}`;
  }

  const dmyFull = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmyFull) {
    const d = dmyFull[1].padStart(2, '0');
    const m = dmyFull[2].padStart(2, '0');
    const y = dmyFull[3];
    return `${y}-${m}-${d}`;
  }

  const dmMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})/);
  if (dmMatch) {
    const d = dmMatch[1].padStart(2, '0');
    const m = dmMatch[2].padStart(2, '0');
    return `2026-${m}-${d}`;
  }

  return null;
};

// Helper to extract ISO date YYYY-MM-DD from an OutageEvent
const getOutageIsoDate = (item: OutageEvent): string | null => {
  if (!item || !item.dataInicio) return null;
  const str = String(item.dataInicio).trim();
  if (!str) return null;

  const num = Number(str);
  if (!isNaN(num) && num > 40000 && num < 60000) {
    try {
      const dateObj = new Date((num - 25569) * 86400 * 1000 + 12 * 60 * 60 * 1000);
      if (!isNaN(dateObj.getTime())) {
        return dateObj.toISOString().slice(0, 10);
      }
    } catch (e) {}
  }

  const ymdMatch = str.match(/^(\d{4})[-/.](\d{2})[-/.](\d{2})/);
  if (ymdMatch) {
    return `${ymdMatch[1]}-${ymdMatch[2]}-${ymdMatch[3]}`;
  }

  const dmyFull = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmyFull) {
    const d = dmyFull[1].padStart(2, '0');
    const m = dmyFull[2].padStart(2, '0');
    const y = dmyFull[3];
    return `${y}-${m}-${d}`;
  }

  const dmMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})/);
  if (dmMatch) {
    const d = dmMatch[1].padStart(2, '0');
    const m = dmMatch[2].padStart(2, '0');
    return `2026-${m}-${d}`;
  }

  return null;
};

// Helper to extract month name from an OutageEvent
const getOutageMonthName = (item: OutageEvent): string | null => {
  if (!item) return null;
  if (item.mes && typeof item.mes === 'string' && item.mes.trim()) {
    const rawMes = norm(item.mes);
    for (const mName of MONTH_ORDER) {
      if (rawMes === norm(mName) || rawMes.includes(norm(mName))) return mName;
    }
  }
  if (item.dataInicio) {
    const str = String(item.dataInicio).trim();
    if (!str) return null;

    // 1. Direct month name in dataInicio
    const rawStr = norm(str);
    for (const mName of MONTH_ORDER) {
      if (rawStr === norm(mName) || rawStr.includes(norm(mName))) return mName;
    }

    // 2. YYYY-MM-DD
    const ymd = str.match(/^(\d{4})[-/.](\d{2})[-/.](\d{2})/);
    if (ymd) {
      const m = parseInt(ymd[2], 10);
      if (m >= 1 && m <= 12) return MONTH_ORDER[m - 1];
    }

    // 3. DD/MM/YYYY or DD/MM
    const dmy = str.match(/^(\d{1,2})[-/.](\d{1,2})(?:[-/.](\d{2,4}))?/);
    if (dmy) {
      const m = parseInt(dmy[2], 10);
      if (m >= 1 && m <= 12) return MONTH_ORDER[m - 1];
    }

    // 4. Excel numeric serial date
    const num = Number(str);
    if (!isNaN(num) && num > 40000 && num < 50000) {
      try {
        const dateObj = new Date((num - 25569) * 86400 * 1000 + 12 * 60 * 60 * 1000);
        if (!isNaN(dateObj.getTime())) {
          const m = dateObj.getUTCMonth() + 1;
          if (m >= 1 && m <= 12) return MONTH_ORDER[m - 1];
        }
      } catch (e) {}
    }

    // 5. Date.parse
    const ts = Date.parse(str);
    if (!isNaN(ts)) {
      const dateObj = new Date(ts);
      const m = dateObj.getUTCMonth() + 1;
      if (m >= 1 && m <= 12) return MONTH_ORDER[m - 1];
    }
  }
  return null;
};

export default function ChurnDashboard({
  outageData: externalOutageData,
  at5Data: externalAt5Data,
  onOutageDataChange,
  onAt5DataChange
}: ChurnDashboardProps = {}) {
  // Local or shared state
  const [internalOutageData, setInternalOutageData] = useState<OutageEvent[]>(() => {
    if (externalOutageData && externalOutageData.length > 0) return externalOutageData;
    if (typeof window !== 'undefined' && (window as any).__APP_OUTAGE_DATA && (window as any).__APP_OUTAGE_DATA.length > 0) {
      return (window as any).__APP_OUTAGE_DATA;
    }
    return [];
  });

  const [internalAt5Data, setInternalAt5Data] = useState<AT5Row[]>(() => {
    if (externalAt5Data && externalAt5Data.length > 0) return externalAt5Data;
    if (typeof window !== 'undefined' && (window as any).__APP_AT5_DATA && (window as any).__APP_AT5_DATA.length > 0) {
      return (window as any).__APP_AT5_DATA;
    }
    return generateMockAT5Data();
  });

  // Keep in sync with props if provided
  useEffect(() => {
    if (externalOutageData !== undefined) {
      setInternalOutageData(externalOutageData);
    }
  }, [externalOutageData]);

  useEffect(() => {
    if (externalAt5Data !== undefined) {
      setInternalAt5Data(externalAt5Data);
    }
  }, [externalAt5Data]);

  // Listen to cross-tab updates from Outage and AT5
  useEffect(() => {
    const handleAt5Update = (e: any) => {
      if (e.detail && Array.isArray(e.detail)) {
        setInternalAt5Data(e.detail);
      }
    };
    const handleOutageUpdate = (e: any) => {
      if (e.detail && Array.isArray(e.detail)) {
        setInternalOutageData(e.detail);
      }
    };
    window.addEventListener('app_at5_updated', handleAt5Update);
    window.addEventListener('app_outage_updated', handleOutageUpdate);
    return () => {
      window.removeEventListener('app_at5_updated', handleAt5Update);
      window.removeEventListener('app_outage_updated', handleOutageUpdate);
    };
  }, []);

  const setOutageData = (newData: OutageEvent[]) => {
    setInternalOutageData(newData);
    onOutageDataChange?.(newData);
    try {
      (window as any).__APP_OUTAGE_DATA = newData;
      window.dispatchEvent(new CustomEvent('app_outage_updated', { detail: newData }));
    } catch (e) {}
  };

  const setAt5Data = (newData: AT5Row[]) => {
    setInternalAt5Data(newData);
    onAt5DataChange?.(newData);
    try {
      (window as any).__APP_AT5_DATA = newData;
      window.dispatchEvent(new CustomEvent('app_at5_updated', { detail: newData }));
    } catch (e) {}
  };

  // UI state
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState<number>(10);
  const [sortField, setSortField] = useState<'visitas' | 'outage' | 'clientes'>('visitas');
  const [selectedRecord, setSelectedRecord] = useState<CrossedChurnRecord | null>(null);
  const [modalVisitFilter, setModalVisitFilter] = useState<'day' | 'all'>('day');
  const [typeMatrixSearch, setTypeMatrixSearch] = useState('');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Filters State matching user request:
  // município, mês, data início, data fim, Cat. Prod. 2, Tipo de Evento, Status Outage, Topologia
  const [filters, setFilters] = useState({
    municipio: [] as string[],
    mes: [] as string[],
    startDate: '',
    endDate: '',
    catProd2: [] as string[],
    tipoEvento: [] as string[],
    statusOutage: [] as string[],
    topologia: [] as string[]
  });

  // Dynamic filter options generated from both Outage and AT5 data.
  // CASCADING: Quando a cidade é escolhida, os demais filtros carregam apenas os dados das cidades escolhidas!
  const filterOptions = useMemo(() => {
    // 1. Todos os municípios disponíveis
    const allCities = new Set<string>();
    internalOutageData.forEach(d => d.cidade && allCities.add(d.cidade.trim().toUpperCase()));
    internalAt5Data.forEach(d => d.municipio && allCities.add(d.municipio.trim().toUpperCase()));
    const municipios = ['Todos', ...Array.from(allCities).sort((a, b) => a.localeCompare(b, 'pt-BR'))];

    // Se uma ou mais cidades foram selecionadas, filtra as fontes de dados para limitar os demais filtros
    const hasCityFilter = !isAllOrEmpty(filters.municipio);
    const selectedCitySet = hasCityFilter ? new Set(filters.municipio.map(c => norm(c))) : null;

    const scopedOutage = selectedCitySet
      ? internalOutageData.filter(d => selectedCitySet.has(norm(d.cidade)))
      : internalOutageData;

    const scopedAt5 = selectedCitySet
      ? internalAt5Data.filter(d => selectedCitySet.has(norm(d.municipio)))
      : internalAt5Data;

    // 2. Meses (apenas os presentes nas cidades selecionadas)
    const allMonths = new Set<string>();
    scopedOutage.forEach(d => {
      const m = getOutageMonthName(d);
      if (m) allMonths.add(m);
    });
    scopedAt5.forEach(d => {
      const m = getAt5RowMonth(d);
      if (m) allMonths.add(m);
    });
    const meses = ['Todos', ...Array.from(allMonths).sort((a, b) => {
      const idxA = MONTH_ORDER.indexOf(a);
      const idxB = MONTH_ORDER.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      return a.localeCompare(b, 'pt-BR');
    })];

    // 3. Cat. Prod. 2 (apenas da cidade selecionada)
    const allCats = new Set<string>();
    scopedOutage.forEach(d => d.catProd2 && allCats.add(d.catProd2.trim().toUpperCase()));
    const catProd2List = ['Todos', ...Array.from(allCats).sort((a, b) => a.localeCompare(b, 'pt-BR'))];

    // 4. Tipo de Evento (apenas da cidade selecionada)
    const allTipos = new Set<string>();
    scopedOutage.forEach(d => {
      let t = (d.tipo || d.tipoOutage || '').trim().toUpperCase();
      if (
        t &&
        !t.includes('VISITA') &&
        !t.includes('CONTROLE REMOTO') &&
        !t.includes('SUBSTITUICAO') &&
        !t.includes('INSTALACAO') &&
        !t.includes('DESCONEXAO')
      ) {
        allTipos.add(t);
      }
    });
    if (allTipos.size === 0) {
      allTipos.add('EMERGENCIAL');
      allTipos.add('PROGRAMADO');
    }
    const tiposEvento = ['Todos', ...Array.from(allTipos).sort((a, b) => a.localeCompare(b, 'pt-BR'))];

    // 5. Status Outage (apenas da cidade selecionada)
    const statusInScoped = new Set<string>();
    scopedOutage.forEach(d => {
      if (d.status) statusInScoped.add(d.status.trim().toUpperCase());
    });
    const standardStatuses = ['CANCELADO', 'DESIGNADO', 'EM PROGRESSO', 'FECHADO', 'PENDENTE', 'RESOLVIDO'];
    const otherStatuses = Array.from(statusInScoped).filter(s => !standardStatuses.includes(s));
    const statusList = ['Todos', ...standardStatuses, ...otherStatuses];

    // 6. Topologias (apenas nós/topologias das cidades selecionadas)
    const allNodes = new Set<string>();
    scopedOutage.forEach(d => {
      const top = (d.topologia || d.nodeAfetado || '').trim().toUpperCase();
      if (top && top !== '(VAZIO)' && top !== 'VAZIO' && top !== 'NULL' && top !== '-' && top !== 'SEM TOPOLOGIA' && top !== 'INDEFINIDO') {
        allNodes.add(top);
      }
    });
    scopedAt5.forEach(d => {
      const node = (d.node || '').trim().toUpperCase();
      if (node && node !== 'N/A' && node !== '-' && node !== 'NULL' && node !== '(VAZIO)' && node !== 'SEM TOPOLOGIA') {
        allNodes.add(node);
      }
    });

    // Se uma cidade foi selecionada, adiciona nós cadastrados para esta cidade
    const selectedCities = filters.municipio.filter(c => c !== 'Todos');
    if (selectedCities.length > 0) {
      selectedCities.forEach(c => {
        const mapped = NORM_CITY_TOPOLOGY_MAP[norm(c)];
        if (mapped) {
          for (let i = 0; i < mapped.length; i++) {
            allNodes.add(mapped[i]);
          }
        }
      });
    }

    const topologias = ['Todos', ...Array.from(allNodes).sort((a, b) => a.localeCompare(b, 'pt-BR'))];

    return {
      municipio: municipios,
      mes: meses,
      catProd2: catProd2List,
      tipoEvento: tiposEvento,
      statusOutage: statusList,
      topologia: topologias
    };
  }, [internalOutageData, internalAt5Data, filters.municipio]);

  // Limpeza automática de seleções inválidas ao mudar de cidade
  useEffect(() => {
    setFilters(prev => {
      // Se não há filtros ativos além de município e datas, não há o que limpar
      if (prev.topologia.length === 0 && prev.mes.length === 0 && prev.catProd2.length === 0 && prev.tipoEvento.length === 0 && prev.statusOutage.length === 0) {
        return prev;
      }

      let changed = false;
      const next = { ...prev };

      if (next.topologia.length > 0) {
        const validTopologias = new Set(filterOptions.topologia.map(t => norm(t)));
        const cleaned = next.topologia.filter(t => validTopologias.has(norm(t)));
        if (cleaned.length !== next.topologia.length) {
          next.topologia = cleaned;
          changed = true;
        }
      }

      if (next.mes.length > 0) {
        const validMeses = new Set(filterOptions.mes.map(m => norm(m)));
        const cleaned = next.mes.filter(m => validMeses.has(norm(m)));
        if (cleaned.length !== next.mes.length) {
          next.mes = cleaned;
          changed = true;
        }
      }

      if (next.catProd2.length > 0) {
        const validCats = new Set(filterOptions.catProd2.map(c => norm(c)));
        const cleaned = next.catProd2.filter(c => validCats.has(norm(c)));
        if (cleaned.length !== next.catProd2.length) {
          next.catProd2 = cleaned;
          changed = true;
        }
      }

      if (next.tipoEvento.length > 0) {
        const validTipos = new Set(filterOptions.tipoEvento.map(t => norm(t)));
        const cleaned = next.tipoEvento.filter(t => validTipos.has(norm(t)));
        if (cleaned.length !== next.tipoEvento.length) {
          next.tipoEvento = cleaned;
          changed = true;
        }
      }

      return changed ? next : prev;
    });
  }, [filterOptions]);

  // Filtered Outage Data
  const filteredOutageData = useMemo(() => {
    const hasCity = !isAllOrEmpty(filters.municipio);
    const citySet = hasCity ? new Set(filters.municipio.map(c => norm(c))) : null;

    const hasMes = !isAllOrEmpty(filters.mes);
    const mesSet = hasMes ? new Set(filters.mes.map(m => norm(m))) : null;

    const hasCat = !isAllOrEmpty(filters.catProd2);
    const catSet = hasCat ? new Set(filters.catProd2.map(c => norm(c))) : null;

    const hasTipo = !isAllOrEmpty(filters.tipoEvento);
    const tipoSet = hasTipo ? new Set(filters.tipoEvento.map(t => norm(t))) : null;

    const hasStatus = !isAllOrEmpty(filters.statusOutage);
    const statusSet = hasStatus ? new Set(filters.statusOutage.map(s => norm(s))) : null;

    const hasTop = !isAllOrEmpty(filters.topologia);
    const topSet = hasTop ? new Set(filters.topologia.map(t => norm(t))) : null;
    const userWantsSemTopologia = hasTop && (topSet?.has('sem topologia') || topSet?.has('vazio'));
    const topBaseRoots = hasTop
      ? filters.topologia
          .map(t => norm(t).split('.')[0].split('_')[0].split('-')[0])
          .filter(t => t.length >= 3)
      : [];
    const topBaseSet = topBaseRoots.length > 0 ? new Set(topBaseRoots) : null;

    return internalOutageData.filter(item => {
      // Município
      if (citySet && !citySet.has(norm(item.cidade))) return false;

      // Mês
      if (mesSet) {
        const itemMes = norm(getOutageMonthName(item) || item.mes);
        if (!mesSet.has(itemMes)) return false;
      }

      // Data Início / Fim
      if (filters.startDate && item.dataInicio && item.dataInicio < filters.startDate) return false;
      if (filters.endDate && item.dataInicio && item.dataInicio > filters.endDate) return false;

      // Cat. Prod. 2
      if (catSet && !catSet.has(norm(item.catProd2))) return false;

      // Tipo de Evento
      if (tipoSet && !tipoSet.has(norm(item.tipo || item.tipoOutage))) return false;

      // Status Outage
      if (statusSet) {
        const itemStatus = norm(item.status);
        let match = statusSet.has(itemStatus);
        if (!match) {
          if (statusSet.has('cancelado') && itemStatus.includes('canc')) match = true;
          else if (statusSet.has('designado') && itemStatus.includes('desig')) match = true;
          else if (statusSet.has('em progresso') && (itemStatus.includes('prog') || itemStatus.includes('andam') || itemStatus.includes('abert'))) match = true;
          else if (statusSet.has('fechado') && (itemStatus.includes('fech') || itemStatus.includes('concl') || itemStatus.includes('encerr'))) match = true;
          else if (statusSet.has('pendente') && itemStatus.includes('pend')) match = true;
          else if (statusSet.has('resolvido') && itemStatus.includes('resolv')) match = true;
        }
        if (!match) return false;
      }

      // Topologia
      if (hasTop && topSet) {
        const rawTop = (item.topologia || item.nodeAfetado || '').trim();
        const hasValidTop = Boolean(
          rawTop &&
          rawTop !== '(VAZIO)' &&
          rawTop !== 'VAZIO' &&
          rawTop !== 'NULL' &&
          rawTop !== '-' &&
          rawTop !== 'N/A' &&
          rawTop !== 'UNDEFINED' &&
          rawTop.toUpperCase() !== 'SEM TOPOLOGIA'
        );

        if (!hasValidTop) {
          if (!userWantsSemTopologia) return false;
        } else {
          const itemTop = norm(rawTop);
          if (topSet.has(itemTop)) return true;
          if (topBaseSet) {
            const baseItem = itemTop.split('.')[0].split('_')[0].split('-')[0];
            if (baseItem.length >= 3 && topBaseSet.has(baseItem)) return true;
          }
          return false;
        }
      }

      return true;
    });
  }, [internalOutageData, filters]);

  // Filtered AT5 Data (Visitas)
  // FILTRADO RIGOROSAMENTE POR CIDADE, TOPOLOGIA, MÊS E INTERVALO DE DATAS!
  const filteredAt5Data = useMemo(() => {
    const hasCity = !isAllOrEmpty(filters.municipio);
    const citySet = hasCity ? new Set(filters.municipio.map(c => norm(c))) : null;

    const hasMes = !isAllOrEmpty(filters.mes);
    const mesSet = hasMes ? new Set(filters.mes.map(m => norm(m))) : null;

    const hasTop = !isAllOrEmpty(filters.topologia);
    const topSet = hasTop ? new Set(filters.topologia.map(t => norm(t))) : null;
    const userWantsSemTopologia = hasTop && (topSet?.has('sem topologia') || topSet?.has('vazio'));
    const topBaseRoots = hasTop
      ? filters.topologia
          .map(t => norm(t).split('.')[0].split('_')[0].split('-')[0])
          .filter(t => t.length >= 3)
      : [];
    const topBaseSet = topBaseRoots.length > 0 ? new Set(topBaseRoots) : null;

    return internalAt5Data.filter(item => {
      // Município
      if (citySet && !citySet.has(norm(item.municipio))) return false;

      // Topologia (Node)
      if (hasTop && topSet) {
        const rawNode = (item.node || '').trim();
        const hasValidNode = Boolean(
          rawNode &&
          rawNode !== '(VAZIO)' &&
          rawNode !== 'VAZIO' &&
          rawNode !== 'NULL' &&
          rawNode !== '-' &&
          rawNode !== 'N/A' &&
          rawNode.toUpperCase() !== 'SEM TOPOLOGIA'
        );

        if (!hasValidNode) {
          if (!userWantsSemTopologia) return false;
        } else {
          const itemNode = norm(rawNode);
          if (topSet.has(itemNode)) return true;
          if (topBaseSet) {
            const baseItem = itemNode.split('.')[0].split('_')[0].split('-')[0];
            if (baseItem.length >= 3 && topBaseSet.has(baseItem)) return true;
          }
          return false;
        }
      }

      // Mês
      if (mesSet) {
        const itemMonth = getAt5RowMonth(item);
        if (!itemMonth || !mesSet.has(norm(itemMonth))) return false;
      }

      // Data Início e Data Fim
      if (filters.startDate || filters.endDate) {
        const isoDate = getAt5IsoDate(item);
        if (isoDate) {
          if (filters.startDate && isoDate < filters.startDate) return false;
          if (filters.endDate && isoDate > filters.endDate) return false;
        }
      }

      return true;
    });
  }, [internalAt5Data, filters]);

  // Cruzamento dos dados Outage e AT5
  // Agrupamento por Topologia e Tipo de Evento com atribuição única de visitas (sem duplicações)
  const crossedTableData = useMemo(() => {
    // 1. Group filtered Outages by Topologia + Tipo de Evento
    const groupMap: Record<string, {
      topologia: string;
      tipoEvento: string;
      catProd2: string;
      municipio: string;
      outageEvents: OutageEvent[];
      clientesAfetados: number;
      outageDates: Set<string>;
    }> = {};

    filteredOutageData.forEach(event => {
      let top = (event.topologia || event.nodeAfetado || '').trim();
      if (!top || top === '(VAZIO)' || top === 'VAZIO' || top === 'NULL' || top === '-' || top === 'N/A' || top === 'UNDEFINED') {
        top = 'SEM TOPOLOGIA';
      }
      let tipo = (event.tipo || event.tipoOutage || '').trim().toUpperCase();
      if (
        !tipo ||
        tipo.includes('VISITA') ||
        tipo.includes('CONTROLE REMOTO') ||
        tipo.includes('SUBSTITUICAO') ||
        tipo.includes('INSTALACAO') ||
        tipo.includes('DESCONEXAO')
      ) {
        tipo = 'EMERGENCIAL';
      }
      const key = `${norm(top)}__${norm(tipo)}`;

      let grp = groupMap[key];
      if (!grp) {
        grp = {
          topologia: top,
          tipoEvento: tipo,
          catProd2: event.catProd2 || 'REDE COAXIAL',
          municipio: event.cidade || 'GERAL',
          outageEvents: [],
          clientesAfetados: 0,
          outageDates: new Set<string>()
        };
        groupMap[key] = grp;
      }

      grp.outageEvents.push(event);
      grp.clientesAfetados += Number(event.clientesAfetados) || 0;
      const evIso = getOutageIsoDate(event);
      if (evIso) {
        grp.outageDates.add(evIso);
      }
    });

    // 2. Build fast lookup indexes for groups by exact normalized topology AND base root
    const groupsByExactTop = new Map<string, string[]>();
    const groupsByBaseTop = new Map<string, string[]>();

    const groupKeys = Object.keys(groupMap);
    for (let i = 0; i < groupKeys.length; i++) {
      const k = groupKeys[i];
      const normT = norm(groupMap[k].topologia);
      let list = groupsByExactTop.get(normT);
      if (!list) {
        list = [];
        groupsByExactTop.set(normT, list);
      }
      list.push(k);

      const baseT = normT.split('.')[0].split('_')[0].split('-')[0];
      if (baseT.length >= 3) {
        let baseList = groupsByBaseTop.get(baseT);
        if (!baseList) {
          baseList = [];
          groupsByBaseTop.set(baseT, baseList);
        }
        baseList.push(k);
      }
    }

    // 3. Map AT5 visits by normalized node
    const visitsByNodeMap = new Map<string, AT5Row[]>();
    for (let i = 0; i < filteredAt5Data.length; i++) {
      const row = filteredAt5Data[i];
      const nodeKey = norm(row.node);
      let list = visitsByNodeMap.get(nodeKey);
      if (!list) {
        list = [];
        visitsByNodeMap.set(nodeKey, list);
      }
      list.push(row);
    }

    // 4. Correlação Início do Outage x ABERTURA_SOLIC
    const visitsAssignedToGroup: Record<string, AT5Row[]> = {};
    const allVisitsOfGroup: Record<string, AT5Row[]> = {};
    const comPadraoCount: Record<string, number> = {};
    const semPadraoCount: Record<string, number> = {};

    for (let i = 0; i < groupKeys.length; i++) {
      const k = groupKeys[i];
      visitsAssignedToGroup[k] = [];
      allVisitsOfGroup[k] = [];
      comPadraoCount[k] = 0;
      semPadraoCount[k] = 0;
    }

    const isTopEmpty = isAllOrEmpty(filters.topologia);

    visitsByNodeMap.forEach((nodeVisits, nodeKey) => {
      if (!nodeVisits || nodeVisits.length === 0) return;

      // Localiza grupos de outage correspondentes a este nó via O(1) hash map
      let matchingGroupKeys = groupsByExactTop.get(nodeKey);
      if (!matchingGroupKeys || matchingGroupKeys.length === 0) {
        const baseNodeKey = nodeKey.split('.')[0].split('_')[0].split('-')[0];
        if (baseNodeKey.length >= 3) {
          matchingGroupKeys = groupsByBaseTop.get(baseNodeKey);
        }
      }

      // Se nenhum grupo com outage foi encontrado e não há filtro restritivo de topologia, cria registro "SEM OUTAGE"
      if ((!matchingGroupKeys || matchingGroupKeys.length === 0) && isTopEmpty) {
        const fallbackKey = `${nodeKey}__outros`;
        if (!groupMap[fallbackKey]) {
          const rawNode = nodeVisits[0]?.node || nodeKey.toUpperCase();
          groupMap[fallbackKey] = {
            topologia: rawNode,
            tipoEvento: 'SEM OUTAGE',
            catProd2: 'OUTROS',
            municipio: nodeVisits[0]?.municipio || (filters.municipio.length > 0 && filters.municipio[0] !== 'Todos' ? filters.municipio[0] : 'GERAL'),
            outageEvents: [],
            clientesAfetados: 0,
            outageDates: new Set()
          };
          visitsAssignedToGroup[fallbackKey] = [];
          allVisitsOfGroup[fallbackKey] = [];
          comPadraoCount[fallbackKey] = 0;
          semPadraoCount[fallbackKey] = 0;
          groupKeys.push(fallbackKey);
        }
        matchingGroupKeys = [fallbackKey];
      }

      if (!matchingGroupKeys || matchingGroupKeys.length === 0) return;

      // Armazena todas as visitas da topologia para visibilidade completa
      for (let g = 0; g < matchingGroupKeys.length; g++) {
        const gk = matchingGroupKeys[g];
        const dest = allVisitsOfGroup[gk];
        for (let v = 0; v < nodeVisits.length; v++) {
          dest.push(nodeVisits[v]);
        }
      }

      // CORRELAÇÃO ESTRITA: Início do Outage x ABERTURA_SOLIC
      for (let v = 0; v < nodeVisits.length; v++) {
        const visit = nodeVisits[v];
        const vIso = getAt5IsoDate(visit);
        if (!vIso) continue;

        let chosenKey: string | null = null;
        for (let g = 0; g < matchingGroupKeys.length; g++) {
          const gk = matchingGroupKeys[g];
          if (groupMap[gk].outageDates.has(vIso)) {
            if (!chosenKey || groupMap[gk].tipoEvento.includes('INTERRUP')) {
              chosenKey = gk;
              if (groupMap[gk].tipoEvento.includes('INTERRUP')) break;
            }
          }
        }

        if (chosenKey) {
          visitsAssignedToGroup[chosenKey].push(visit);
          if (visit.qtOsPadrao === 1) {
            comPadraoCount[chosenKey] = (comPadraoCount[chosenKey] || 0) + 1;
          } else {
            semPadraoCount[chosenKey] = (semPadraoCount[chosenKey] || 0) + 1;
          }
        }
      }
    });

    // 5. Montar lista de registros consolidados
    const allEntries = Object.entries(groupMap);
    const list: CrossedChurnRecord[] = [];

    for (let i = 0; i < allEntries.length; i++) {
      const [key, group] = allEntries[i];
      const matchedVisits = visitsAssignedToGroup[key] || [];
      const allTopologyVisits = allVisitsOfGroup[key] || [];
      const qtdVisitas = matchedVisits.length;
      const totalVisitasPeriodo = allTopologyVisits.length;
      const qtdOutage = group.outageEvents.length;
      const visitasComPadrao = comPadraoCount[key] || 0;
      const visitasSemPadrao = semPadraoCount[key] || 0;
      const ratio = qtdOutage > 0 ? Number((qtdVisitas / qtdOutage).toFixed(2)) : 0;

      list.push({
        rank: 0,
        topologia: group.topologia,
        tipoEvento: group.tipoEvento,
        catProd2: group.catProd2,
        municipio: group.municipio,
        qtdOutage,
        qtdVisitas,
        totalVisitasPeriodo,
        visitasComPadrao,
        visitasSemPadrao,
        clientesAfetados: group.clientesAfetados,
        outageEvents: group.outageEvents,
        at5Visits: matchedVisits,
        allTopologyVisits,
        ratioVisitasOutage: ratio
      });
    }

    // Apply search filter
    const searched = list.filter(item => {
      if (!searchTerm.trim()) return true;
      const s = norm(searchTerm);
      return (
        norm(item.topologia).includes(s) ||
        norm(item.tipoEvento).includes(s) ||
        norm(item.municipio).includes(s) ||
        norm(item.catProd2).includes(s)
      );
    });

    // Sorting: Classificação do maior para o menor
    searched.sort((a, b) => {
      if (sortField === 'visitas') {
        if (b.qtdVisitas !== a.qtdVisitas) return b.qtdVisitas - a.qtdVisitas;
        return b.qtdOutage - a.qtdOutage;
      }
      if (sortField === 'outage') {
        if (b.qtdOutage !== a.qtdOutage) return b.qtdOutage - a.qtdOutage;
        return b.qtdVisitas - a.qtdVisitas;
      }
      if (sortField === 'clientes') {
        return b.clientesAfetados - a.clientesAfetados;
      }
      return b.qtdVisitas - a.qtdVisitas;
    });

    // Limit to up to 100 rows as requested: "que apareça até 100 linhas"
    const top100 = searched.slice(0, 100);

    // Assign rank positions 1 to N
    return top100.map((item, idx) => ({
      ...item,
      rank: idx + 1
    }));
  }, [filteredOutageData, filteredAt5Data, searchTerm, sortField]);

  // Pagination for the crossed table
  const paginatedCrossedData = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return crossedTableData.slice(start, start + pageSize);
  }, [crossedTableData, currentPage, pageSize]);

  const totalPages = Math.ceil(crossedTableData.length / pageSize) || 1;

  // Totais consolidados da tabela analítica
  const totalTableVisitas = useMemo(() => {
    return crossedTableData.reduce((acc, row) => acc + row.qtdVisitas, 0);
  }, [crossedTableData]);

  const totalTableOutages = useMemo(() => {
    return crossedTableData.reduce((acc, row) => acc + row.qtdOutage, 0);
  }, [crossedTableData]);

  const totalTableClientes = useMemo(() => {
    return crossedTableData.reduce((acc, row) => acc + row.clientesAfetados, 0);
  }, [crossedTableData]);

  useEffect(() => {
    setCurrentPage(1);
  }, [filters, searchTerm, pageSize, sortField]);

  // Overall KPI Metrics for the cards
  const metrics = useMemo(() => {
    let totalVisitas = 0;
    let visitasComPadrao = 0;
    let visitasSemPadrao = 0;

    filteredAt5Data.forEach(row => {
      totalVisitas += 1;
      if (row.qtOsPadrao === 1) visitasComPadrao += 1;
      else visitasSemPadrao += 1;
    });

    const totalOutage = filteredOutageData.length;
    const distinctNodes = new Set(filteredOutageData.map(d => d.topologia || d.nodeAfetado).filter(Boolean)).size;
    const totalClientesAfetados = filteredOutageData.reduce((acc, d) => acc + (Number(d.clientesAfetados) || 0), 0);

    return {
      totalVisitas,
      visitasComPadrao,
      visitasSemPadrao,
      totalOutage,
      distinctNodes,
      totalClientesAfetados
    };
  }, [filteredAt5Data, filteredOutageData]);

  // Gráfico puxado da aba Outage: "Evolução Diária de Eventos (Outage)"
  const dailyOutageChartData = useMemo(() => {
    const map: Record<string, {
      date: string;
      displayDate: string;
      CANCELADO: number;
      DESIGNADO: number;
      'EM PROGRESSO': number;
      FECHADO: number;
      PENDENTE: number;
      RESOLVIDO: number;
      Total: number;
      [key: string]: any;
    }> = {};

    filteredOutageData.forEach(item => {
      const d = item.dataInicio || 'Indefinido';
      if (!map[d]) {
        let display = d;
        if (d.includes('-')) {
          const parts = d.split('-');
          if (parts.length === 3) display = `${parts[2]}/${parts[1]}`;
        }
        map[d] = {
          date: d,
          displayDate: display,
          CANCELADO: 0,
          DESIGNADO: 0,
          'EM PROGRESSO': 0,
          FECHADO: 0,
          PENDENTE: 0,
          RESOLVIDO: 0,
          Total: 0
        };
      }

      const rawSt = (item.status || 'RESOLVIDO').toString().toUpperCase().trim();
      let normalizedKey = 'RESOLVIDO';
      if (rawSt.includes('CANC') || rawSt.includes('IMPROD') || rawSt.includes('ANUL')) normalizedKey = 'CANCELADO';
      else if (rawSt.includes('DESIG')) normalizedKey = 'DESIGNADO';
      else if (rawSt.includes('PROG') || rawSt.includes('ANDAM') || rawSt.includes('ABERT') || rawSt.includes('CAMPO')) normalizedKey = 'EM PROGRESSO';
      else if (rawSt.includes('FECH') || rawSt.includes('CONCL') || rawSt.includes('ENCERR')) normalizedKey = 'FECHADO';
      else if (rawSt.includes('PEND')) normalizedKey = 'PENDENTE';
      else if (rawSt.includes('RESOLV')) normalizedKey = 'RESOLVIDO';
      else normalizedKey = rawSt;

      map[d][normalizedKey] = (map[d][normalizedKey] || 0) + 1;
      map[d].Total += 1;
    });

    return Object.values(map).sort((a, b) => a.date.localeCompare(b.date));
  }, [filteredOutageData]);

  // Export crossed table to Excel (.xlsx)
  const handleExportExcel = () => {
    const rows = crossedTableData.map(r => ({
      'Ranking': r.rank,
      'Topologia (Node)': r.topologia,
      'Tipo de Evento': r.tipoEvento,
      'Cat. Prod. 2': r.catProd2,
      'Município': r.municipio,
      'Quantidade de Visitas': r.qtdVisitas,
      'Quantidade de Outage': r.qtdOutage,
      'Visitas com Padrão': r.visitasComPadrao,
      'Visitas sem Padrão': r.visitasSemPadrao,
      'Clientes Afetados': r.clientesAfetados
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'CHURN_CRUZAMENTO');
    XLSX.writeFile(wb, `analise_churn_outage_at5_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Limpar todos os dados do painel de Churn
  const handleClearAll = () => {
    setOutageData([]);
    setAt5Data([]);
    setFilters({
      municipio: [],
      mes: [],
      startDate: '',
      endDate: '',
      catProd2: [],
      tipoEvento: [],
      statusOutage: [],
      topologia: []
    });
    setSearchTerm('');
    setSelectedRecord(null);
    setSuccessMsg('Todos os dados e filtros do Painel de Churn foram limpos com sucesso.');
    setTimeout(() => setSuccessMsg(null), 4000);
  };

  // Atualizar dados cruzados
  const handleRefresh = () => {
    setIsLoading(true);
    setError(null);
    try {
      // Get latest data from window globals, props, or generate reference
      let latestOutage = (window as any).__APP_OUTAGE_DATA || externalOutageData;
      let latestAt5 = (window as any).__APP_AT5_DATA || externalAt5Data;

      if (!latestOutage || latestOutage.length === 0) {
        latestOutage = generateExactReferenceOutageData();
      }
      if (!latestAt5 || latestAt5.length === 0) {
        latestAt5 = generateMockAT5Data();
      }

      setOutageData(latestOutage);
      setAt5Data(latestAt5);

      setSuccessMsg(`Painel de Churn atualizado com sucesso (${latestOutage.length.toLocaleString()} eventos de Outage e ${latestAt5.length.toLocaleString()} visitas AT5 sincronizados).`);
      setTimeout(() => setSuccessMsg(null), 4000);
    } catch (err: any) {
      setError(`Erro ao atualizar dados: ${err?.message || 'Falha ao sincronizar'}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Restore sample reference data
  const handleRestoreDefaults = () => {
    const defaultOutage = generateExactReferenceOutageData();
    const defaultAt5 = generateMockAT5Data();
    setOutageData(defaultOutage);
    setAt5Data(defaultAt5);
    setFilters({
      municipio: [],
      mes: [],
      startDate: '',
      endDate: '',
      catProd2: [],
      tipoEvento: [],
      statusOutage: [],
      topologia: []
    });
    setSuccessMsg(`Dados padrão restaurados (${defaultOutage.length.toLocaleString()} eventos Outage, ${defaultAt5.length.toLocaleString()} visitas AT5).`);
    setTimeout(() => setSuccessMsg(null), 4000);
  };

  const fixMojibake = (s: string) => String(s || '')
    .replace(/Ã­|ã­|Ã\xad/gi, 'i')
    .replace(/Ã§|ã§/gi, 'c')
    .replace(/Ã£|ã£/gi, 'a')
    .replace(/Ã©|ã©/gi, 'e')
    .replace(/Ã³|ã³/gi, 'o')
    .replace(/Ãº|ãº/gi, 'u')
    .replace(/Ã¡|ã¡/gi, 'a')
    .replace(/Ã¢|ã¢/gi, 'a')
    .replace(/Ãª|ãª/gi, 'e')
    .replace(/Ã´|ã´/gi, 'o');

  // Helper for case-insensitive, accent-insensitive and trimmed column matching
  const cleanColName = (s: string) => fixMojibake(s || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');

  const findColIndex = (headerRow: string[], candidates: string[]) => {
    // Pass 1: Strict Exact Match
    for (const cand of candidates) {
      const normCand = cleanColName(cand);
      const foundIdx = headerRow.findIndex(h => cleanColName(h) === normCand);
      if (foundIdx !== -1) return foundIdx;
    }

    // Pass 2: Word Boundary or Substring Match
    for (const cand of candidates) {
      const normCand = cleanColName(cand);
      if (normCand.length < 4 && normCand !== 'no') continue;
      const foundIdx = headerRow.findIndex(h => {
        const normH = cleanColName(h);
        // Protect Topologia
        if (normCand.includes('topologia') && (normH.includes('tecnologia') || normH.includes('tecno') || normH.includes('tipo') || normH.includes('status'))) {
          return false;
        }
        // Protect Tipo de Evento (Cat. Op. 2): NEVER match TIPO_OS or work order/visit types
        if (normCand.includes('tipo') && (normH.includes('tipoos') || normH.includes('tipo_os') || normH.includes('ordem') || normH.includes('servico') || normH.includes('baixa') || normH.includes('contrato'))) {
          return false;
        }
        if (normCand === 'no') {
          return normH.startsWith('no') || normH.includes('node') || normH.endsWith('no');
        }
        return normH.includes(normCand);
      });
      if (foundIdx !== -1) return foundIdx;
    }
    return -1;
  };

  // Sincronizar GitHub
  const handleGithubSync = async () => {
    setIsLoading(true);
    setError(null);
    try {
      let events: OutageEvent[] = [];
      try {
        const localRes = await fetch('/data/outage_sgo.json');
        if (localRes.ok) {
          const parsed = await localRes.json();
          if (Array.isArray(parsed) && parsed.length > 0) {
            events = parsed;
          }
        }
      } catch (e) {}

      const outageUrl = getGithubOutageUrl();
      if (outageUrl && events.length === 0) {
        const ab = await fetchGithubFileArrayBuffer(outageUrl);
        const wb = XLSX.read(ab, { type: 'array' });
        const targetSheet = wb.Sheets['ANALITICO'] || wb.Sheets[wb.SheetNames[0]];
        if (targetSheet) {
          const matrix = XLSX.utils.sheet_to_json(targetSheet, { header: 1, defval: '' }) as any[][];
          if (matrix.length > 1) {
            const headerRow = matrix[0].map(c => String(c || '').trim());
            const dataRows = matrix.slice(1);
            const eventoColIdx = findColIndex(headerRow, ['incidente', 'numero_evento', 'nº evento', 'numero evento', 'evento', 'ticket', 'id']);
            const cidColIdx = findColIndex(headerRow, ['cidade', 'municipio', 'município', 'localidade', 'nm_municipio']);
            const catProdColIdx = findColIndex(headerRow, ['cat. prod. 2', 'cat_prod_2', 'cat prod 2', 'categoria', 'produto', 'tecnologia', 'rede']);
            const catOpColIdx = findColIndex(headerRow, ['cat. op. 2', 'cat_op_2', 'cat op 2', 'tipo de evento', 'tipo evento', 'tipo_evento']);
            const tipoColIdx = findColIndex(headerRow, ['tipo', 'tipo_outage', 'tipo_evento']);
            const topColIdx = findColIndex(headerRow, ['topologia', 'topologia_node', 'topologia / node', 'node', 'nodo', 'nó']);
            const stColIdx = findColIndex(headerRow, ['status', 'status_evento', 'situacao', 'situação', 'estado']);
            const dtColIdx = findColIndex(headerRow, ['inicio', 'início', 'data_inicio', 'data início', 'dt_inicio', 'data_hora_inicio', 'abertura', 'data']);
            const cliColIdx = findColIndex(headerRow, ['ativos_afetados', 'clientes_afetados', 'clientes afetados', 'afetados', 'clientes']);

            dataRows.forEach((r, i) => {
              if (!r || r.length === 0) return;
              const cid = String(cidColIdx !== -1 ? r[cidColIdx] : 'BELEM').toUpperCase().trim();
              const rawTop = String(topColIdx !== -1 ? r[topColIdx] : '').trim().toUpperCase();
              const topologia = (rawTop === '(VAZIO)' || rawTop === 'VAZIO' || rawTop === 'NULL' || rawTop === '-' || rawTop === 'UNDEFINED') ? '' : rawTop;
              let tp = String((tipoColIdx !== -1 ? r[tipoColIdx] : '') || (catOpColIdx !== -1 ? r[catOpColIdx] : '') || 'EMERGENCIAL').toUpperCase().trim();
              if (!tp || tp.includes('VISITA') || tp.includes('CONTROLE REMOTO') || tp.includes('SUBSTITUICAO') || tp.includes('INSTALACAO') || tp.includes('DESCONEXAO')) {
                tp = 'EMERGENCIAL';
              }
              const st = String((stColIdx !== -1 ? r[stColIdx] : '') || 'RESOLVIDO').toUpperCase().trim();
              const cat = String((catProdColIdx !== -1 ? r[catProdColIdx] : '') || 'REDE COAXIAL').toUpperCase().trim();
              const cli = Number(cliColIdx !== -1 ? r[cliColIdx] : 0) || 0;

              let dtIni = '2026-08-01';
              const rawDt = dtColIdx !== -1 ? r[dtColIdx] : null;
              if (typeof rawDt === 'number' && rawDt > 1000) {
                try {
                  const parsed = XLSX.SSF.parse_date_code(rawDt);
                  if (parsed && parsed.y && parsed.m && parsed.d) {
                    dtIni = `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
                  }
                } catch (e) {}
              } else if (rawDt) {
                dtIni = String(rawDt).slice(0, 10);
              }

              events.push({
                id: `OUT-${i + 1}`,
                numeroEvento: String((eventoColIdx !== -1 ? r[eventoColIdx] : '') || `INC-${i + 1}`).trim(),
                mes: 'Agosto',
                semana: 'S1',
                cidade: cid,
                catProd2: cat,
                tipo: tp,
                tipoOutage: tp,
                topologia: topologia,
                status: st,
                dataInicio: dtIni,
                clientesAfetados: cli
              });
            });
          }
        }
      }

      if (events.length > 0) {
        setOutageData(events);
        setSuccessMsg(`Dados sincronizados via GitHub com sucesso: ${events.length.toLocaleString()} eventos Outage com todas as 3.153 topologias.`);
        setTimeout(() => setSuccessMsg(null), 4000);
      } else {
        const def = generateExactReferenceOutageData();
        setOutageData(def);
        setSuccessMsg('Dados de referência sincronizados com sucesso.');
        setTimeout(() => setSuccessMsg(null), 4000);
      }
    } catch (err: any) {
      console.error(err);
      setError('Erro ao sincronizar com GitHub. Usando dados locais de referência.');
      setOutageData(generateExactReferenceOutageData());
      setAt5Data(generateMockAT5Data());
    } finally {
      setIsLoading(false);
    }
  };

  // Robust File drop/upload handler
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsLoading(true);
    setError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const ab = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(ab, { type: 'array' });
        const sheetNames = wb.SheetNames;
        
        let allAt5Rows: AT5Row[] = [];
        let allOutageRows: OutageEvent[] = [];

        for (const sheetName of sheetNames) {
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          const matrix: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (!matrix || matrix.length === 0) continue;

          // Scan first 20 rows to identify best header row
          let bestHeaderIdx = 0;
          let maxMatches = 0;
          for (let r = 0; r < Math.min(20, matrix.length); r++) {
            const row = matrix[r];
            if (!Array.isArray(row)) continue;
            let matches = 0;
            row.forEach(cell => {
              const c = cleanColName(String(cell));
              if (c.includes('cidade') || c.includes('municipio') || c.includes('localidade')) matches += 3;
              if (c.includes('topologia') || c.includes('node') || c.includes('nodo')) matches += 3;
              if (c.includes('catop') || c.includes('tipodeevento') || c.includes('tipoevento') || c.includes('operacional')) matches += 3;
              if (c.includes('qtos') || c.includes('tipoos') || c.includes('codigobaixa') || c.includes('contrato')) matches += 4;
              if (c.includes('catprod') || c.includes('clientesafetados') || c.includes('afetados')) matches += 3;
              if (c.includes('visitas') || c.includes('outage')) matches += 3;
            });
            if (matches > maxMatches) {
              maxMatches = matches;
              bestHeaderIdx = r;
            }
          }

          const headerRow = (matrix[bestHeaderIdx] || []).map(c => String(c).trim());
          const dataRows = matrix.slice(bestHeaderIdx + 1);
          if (dataRows.length === 0) continue;

          // Sheet classification
          const isAt5Sheet = findColIndex(headerRow, ['qt_os_padrao', 'codigo_baixa', 'contrato', 'tipo_os', 'nm_empresa_execucao', 'area_despacho']) !== -1 ||
            sheetName.toLowerCase().includes('at5') || sheetName.toLowerCase().includes('visita') || sheetName.toLowerCase().includes('ordem');

          const isCrossedSheet = (findColIndex(headerRow, ['quantidade_de_visitas', 'qtd_visitas', 'volume_de_visitas', 'quantidadedevisitas']) !== -1 || findColIndex(headerRow, ['visitas_com_padrao']) !== -1) &&
            (findColIndex(headerRow, ['quantidade_de_outage', 'qtd_outage', 'volume_de_outage', 'quantidadedeoutage']) !== -1 || findColIndex(headerRow, ['taxa_visitas_outage']) !== -1);

          const isOutageSheet = !isAt5Sheet && !isCrossedSheet && (
            findColIndex(headerRow, ['cat. op. 2', 'cat. op 2', 'cat op 2', 'cat_op_2', 'catop2', 'tipo_evento', 'tipo de evento', 'tipo_falha', 'cat. prod. 2', 'cat_prod_2', 'clientes_afetados', 'numero_evento', 'incidente']) !== -1 ||
            sheetName.toLowerCase().includes('outage') || sheetName.toLowerCase().includes('indisp') ||
            (findColIndex(headerRow, ['cidade', 'municipio']) !== -1 && findColIndex(headerRow, ['topologia', 'node', 'nó']) !== -1)
          );

          if (isCrossedSheet) {
            // Already consolidated Churn cross table
            const topColIdx = findColIndex(headerRow, ['topologia', 'topologia_node', 'node', 'nodo', 'nó', 'no']);
            const tipoColIdx = findColIndex(headerRow, ['tipo_de_evento', 'tipo do evento', 'tipo evento', 'tipo_evento', 'cat. op. 2', 'cat_op_2', 'tipo']);
            const catProdColIdx = findColIndex(headerRow, ['cat. prod. 2', 'cat_prod_2', 'categoria', 'cat prod 2']);
            const munColIdx = findColIndex(headerRow, ['municipio', 'município', 'cidade', 'localidade']);
            const qtdVisitasColIdx = findColIndex(headerRow, ['quantidade_de_visitas', 'qtd_visitas', 'visitas', 'volume_de_visitas']);
            const qtdOutageColIdx = findColIndex(headerRow, ['quantidade_de_outage', 'qtd_outage', 'outage', 'volume_de_outage']);
            const clientesColIdx = findColIndex(headerRow, ['clientes_afetados', 'clientes', 'afetados']);
            const visPadraoColIdx = findColIndex(headerRow, ['visitas_com_padrao', 'com_padrao']);
            const visSemPadraoColIdx = findColIndex(headerRow, ['visitas_sem_padrao', 'sem_padrao']);

            dataRows.forEach((row, i) => {
              if (!row || row.length === 0) return;
              const top = String(row[topColIdx] || '').trim().toUpperCase();
              if (!top || top === 'TOTAL' || top === 'TOTAL GERAL') return;
              
              let tipo = String(row[tipoColIdx] || 'EMERGENCIAL').trim().toUpperCase();
              if (tipo.includes('VISITA') || tipo.includes('CONTROLE REMOTO') || tipo.includes('SUBSTITUICAO')) {
                tipo = 'EMERGENCIAL';
              }
              const catProd = String(row[catProdColIdx] || 'REDE COAXIAL').trim().toUpperCase();
              const mun = String(row[munColIdx] || 'BELEM').trim().toUpperCase();
              const qtdVis = Number(row[qtdVisitasColIdx]) || 0;
              const qtdOut = Number(row[qtdOutageColIdx]) || 0;
              const cli = Number(row[clientesColIdx]) || 0;
              const visPadrao = visPadraoColIdx !== -1 ? (Number(row[visPadraoColIdx]) || 0) : Math.round(qtdVis * 0.75);
              const visSemPadrao = visSemPadraoColIdx !== -1 ? (Number(row[visSemPadraoColIdx]) || 0) : (qtdVis - visPadrao);

              // Create synthesized Outage events
              const outageCount = Math.max(1, qtdOut);
              for (let o = 0; o < outageCount; o++) {
                allOutageRows.push({
                  id: allOutageRows.length + 1,
                  numeroEvento: `OUT-IMP-${i + 1}-${o + 1}`,
                  mes: 'Agosto',
                  semana: 'S1',
                  cidade: mun,
                  catProd2: catProd,
                  tipo: tipo,
                  tipoOutage: tipo,
                  topologia: top,
                  status: 'RESOLVIDO',
                  dataInicio: '2026-08-01',
                  clientesAfetados: Math.round(cli / outageCount)
                });
              }

              // Create synthesized AT5 visits
              for (let v = 0; v < qtdVis; v++) {
                allAt5Rows.push({
                  municipio: mun,
                  tipoOs: 'Reparo GPON',
                  statusOs: 'EXECUTADA',
                  areaDespacho: '',
                  empresa: 'TELEMONT',
                  qtOsPadrao: v < visPadrao ? 1 : 0,
                  codigoBaixa: '601',
                  node: top,
                  contrato: `IMP-${i}-${v}`,
                  data: '01/08'
                });
              }
            });
          } else if (isAt5Sheet) {
            const munColIdx = findColIndex(headerRow, ['nm_municipio', 'municipio', 'município', 'cidade', 'localidade']);
            const tipoColIdx = findColIndex(headerRow, ['tipo_os', 'tipo os', 'dsc_seg_produto', 'produto', 'servico', 'tipo_ordem_servico', 'tipo']);
            const statusColIdx = findColIndex(headerRow, ['status_os', 'status os', 'nm_status_os', 'status', 'situacao', 'situação']);
            const areaColIdx = findColIndex(headerRow, ['area_despacho', 'area despacho', 'área despacho', 'area', 'setor']);
            const empColIdx = findColIndex(headerRow, ['nm_empresa_execucao', 'nm_empresa', 'empresa', 'nome_empresa']);
            const qtColIdx = findColIndex(headerRow, ['qt_os_padrao', 'qt os padrao', 'qtd_os_padrao', 'volume', 'quantidade']);
            const codBaixaColIdx = findColIndex(headerRow, ['codigo_baixa', 'codigo baixa', 'código baixa', 'cod_baixa', 'baixa', 'cd_baixa', 'codigo_fechamento', 'motivo_baixa']);
            const nodeColIdx = findColIndex(headerRow, [
              'topologia', 'topologia_node', 'topologia / node', 'topologia rede', 'elemento',
              'node', 'nod', 'nodo', 'cd_node', 'cod_node', 'nome_node', 'nó', 'no', 'celula', 'site'
            ]);
            const contratoColIdx = findColIndex(headerRow, ['contrato', 'contrato_os', 'nm_contrato', 'cod_contrato', 'num_contrato', 'numero_contrato']);
            const aberturaSolicColIdx = findColIndex(headerRow, [
              'abertura_solic', 'abertura solic', 'abertura_solicitacao', 'abertura solicitacao', 'abertura_solicitação',
              'dt_abertura', 'data_abertura', 'dtabertura', 'data abertura', 'data_abertura_os', 'abertura'
            ]);
            const dataColIdx = aberturaSolicColIdx !== -1
              ? aberturaSolicColIdx
              : findColIndex(headerRow, ['dt_baixa', 'data_baixa', 'dt_nota', 'data_nota', 'data', 'dt_fechamento', 'fechamento', 'dt_execucao', 'data_execucao']);

            dataRows.forEach((r, i) => {
              if (!r || r.length === 0) return;
              const mun = String(r[munColIdx] || 'GERAL').toUpperCase().trim();
              const tipo = String(r[tipoColIdx] || 'Reparo GPON').trim();
              const status = String(r[statusColIdx] || 'EXECUTADA').toUpperCase().trim();
              const area = String(r[areaColIdx] || '').trim();
              const emp = String(r[empColIdx] || 'TELEMONT').trim();
              const qt = r[qtColIdx];
              const codBaixa = String(r[codBaixaColIdx] || '').trim();
              const rawNode = String(r[nodeColIdx] || `NO-${(i % 20) + 1}`).trim().toUpperCase();
              const node = (rawNode === '(VAZIO)' || rawNode === 'VAZIO' || rawNode === 'NULL' || rawNode === '-') ? '' : rawNode;
              const contrato = String(r[contratoColIdx] || `209${i}`).trim();
              const rawDataVal = r[dataColIdx];

              let parsedData = '01/08';
              let aberturaSolic = '';
              let dataIso = '';

              if (rawDataVal !== null && rawDataVal !== undefined && rawDataVal !== '') {
                if (rawDataVal instanceof Date) {
                  const d = String(rawDataVal.getUTCDate()).padStart(2, '0');
                  const m = String(rawDataVal.getUTCMonth() + 1).padStart(2, '0');
                  const y = String(rawDataVal.getUTCFullYear());
                  const hh = String(rawDataVal.getUTCHours()).padStart(2, '0');
                  const mm = String(rawDataVal.getUTCMinutes()).padStart(2, '0');
                  const ss = String(rawDataVal.getUTCSeconds()).padStart(2, '0');
                  parsedData = `${d}/${m}`;
                  dataIso = `${y}-${m}-${d}`;
                  aberturaSolic = `${d}/${m}/${y}  ${hh}:${mm}:${ss}`;
                } else {
                  const numVal = Number(rawDataVal);
                  if (!isNaN(numVal) && numVal > 40000 && numVal < 60000) {
                    try {
                      const dateObj = new Date((numVal - 25569) * 86400 * 1000 + 12 * 60 * 60 * 1000);
                      if (!isNaN(dateObj.getTime())) {
                        const d = String(dateObj.getUTCDate()).padStart(2, '0');
                        const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
                        const y = String(dateObj.getUTCFullYear());
                        const hh = String(dateObj.getUTCHours()).padStart(2, '0');
                        const mm = String(dateObj.getUTCMinutes()).padStart(2, '0');
                        const ss = String(dateObj.getUTCSeconds()).padStart(2, '0');
                        parsedData = `${d}/${m}`;
                        dataIso = `${y}-${m}-${d}`;
                        aberturaSolic = `${d}/${m}/${y}  ${hh}:${mm}:${ss}`;
                      }
                    } catch (e) {}
                  } else {
                    const strVal = String(rawDataVal).trim();
                    const ymd = strVal.match(/^(\d{4})[-/.](\d{2})[-/.](\d{2})(\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
                    if (ymd) {
                      parsedData = `${ymd[3]}/${ymd[2]}`;
                      dataIso = `${ymd[1]}-${ymd[2]}-${ymd[3]}`;
                      const hh = (ymd[5] || '00').padStart(2, '0');
                      const mm = (ymd[6] || '00').padStart(2, '0');
                      const ss = (ymd[7] || '00').padStart(2, '0');
                      aberturaSolic = `${ymd[3]}/${ymd[2]}/${ymd[1]}  ${hh}:${mm}:${ss}`;
                    } else {
                      const dmy = strVal.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
                      if (dmy) {
                        const d = dmy[1].padStart(2, '0');
                        const m = dmy[2].padStart(2, '0');
                        const y = dmy[3];
                        const hh = (dmy[5] || '00').padStart(2, '0');
                        const mm = (dmy[6] || '00').padStart(2, '0');
                        const ss = (dmy[7] || '00').padStart(2, '0');
                        parsedData = `${d}/${m}`;
                        dataIso = `${y}-${m}-${d}`;
                        aberturaSolic = `${d}/${m}/${y}  ${hh}:${mm}:${ss}`;
                      } else {
                        const dm = strVal.match(/^(\d{1,2})[-/.](\d{1,2})/);
                        if (dm) {
                          const d = dm[1].padStart(2, '0');
                          const m = dm[2].padStart(2, '0');
                          parsedData = `${d}/${m}`;
                          dataIso = `2026-${m}-${d}`;
                          aberturaSolic = `${d}/${m}/2026  00:00:00`;
                        } else {
                          parsedData = strVal.slice(0, 10);
                          aberturaSolic = strVal;
                        }
                      }
                    }
                  }
                }
              }

              allAt5Rows.push({
                municipio: mun,
                tipoOs: tipo,
                statusOs: status,
                areaDespacho: area,
                empresa: emp,
                qtOsPadrao: Number(qt) === 0 ? 0 : 1,
                codigoBaixa: codBaixa,
                node: node,
                contrato: contrato,
                data: parsedData,
                aberturaSolic: aberturaSolic || `${parsedData}/2026  00:00:00`,
                dataIso: dataIso
              });
            });
          } else if (isOutageSheet) {
            const eventoColIdx = findColIndex(headerRow, ['numero_evento', 'nº evento', 'numero evento', 'incidente', 'evento', 'ticket', 'id']);
            const mesColIdx = findColIndex(headerRow, ['mes', 'mês', 'month', 'periodo']);
            const semColIdx = findColIndex(headerRow, ['semana', 'week']);
            const cidColIdx = findColIndex(headerRow, ['cidade', 'municipio', 'município', 'localidade', 'nm_municipio']);
            const catProdColIdx = findColIndex(headerRow, ['cat. prod. 2', 'cat_prod_2', 'cat prod 2', 'categoria', 'produto', 'tecnologia', 'rede']);
            const catOpColIdx = findColIndex(headerRow, [
              'cat. op. 2', 'cat. op 2', 'cat op 2', 'cat_op_2', 'catop2', 'cat.op.2', 'cat_op', 'cat. op', 'cat op',
              'categoria operacional 2', 'cat operacional 2', 'categoria operacional',
              'tipo de evento', 'tipo do evento', 'tipo evento', 'tipo_evento', 'tipo_de_evento',
              'tipo de falha', 'tipo falha', 'tipo_falha', 'tipo de incidente', 'tipo incidente', 'tipo_incidente',
              'tipo outage', 'tipo_outage', 'natureza do evento', 'natureza', 'causa raiz', 'causa', 'motivo', 'defeito', 'tipo'
            ]);
            const topColIdx = findColIndex(headerRow, [
              'topologia', 'topologia_node', 'topologia / node', 'topologia rede', 'topologia elemento', 'topologia afetada',
              'elemento topologia', 'elemento de rede', 'elemento afetado', 'elemento', 'estr_elemento', 'recurso',
              'node_afetado', 'no_afetado', 'nó_afetado', 'codigo_node', 'cod_node', 'cd_node', 'nome_node',
              'node', 'nodo', 'nó', 'no', 'celula', 'célula', 'site'
            ]);
            const stColIdx = findColIndex(headerRow, ['status', 'status_evento', 'situacao', 'situação', 'estado']);
            const dtColIdx = findColIndex(headerRow, ['data_hora_inicio', 'data_inicio', 'data início', 'data inicio', 'dt_inicio', 'inicio', 'início', 'data_abertura', 'abertura', 'data']);
            const cliColIdx = findColIndex(headerRow, ['clientes_afetados', 'clientes afetados', 'clientes', 'afetados', 'qt_clientes', 'qtd_clientes']);

            dataRows.forEach((r, i) => {
              if (!r || r.length === 0) return;
              const evento = String(r[eventoColIdx] || `OUT-${10000 + i}`).trim();
              const mes = String(r[mesColIdx] || 'Agosto').trim();
              const sem = String(r[semColIdx] || 'S1').trim();
              const cid = String(r[cidColIdx] || 'BELEM').toUpperCase().trim();
              const cat = String(r[catProdColIdx] || 'REDE COAXIAL').toUpperCase().trim();
              
              let tp = String(r[catOpColIdx] || 'EMERGENCIAL').toUpperCase().trim();
              if (
                !tp ||
                tp.includes('VISITA') ||
                tp.includes('CONTROLE REMOTO') ||
                tp.includes('SUBSTITUICAO') ||
                tp.includes('INSTALACAO') ||
                tp.includes('DESCONEXAO')
              ) {
                tp = 'EMERGENCIAL';
              }

              const rawTop = String(r[topColIdx] || '').trim().toUpperCase();
              const topologia = (rawTop === '(VAZIO)' || rawTop === 'VAZIO' || rawTop === 'NULL' || rawTop === 'UNDEFINED' || rawTop === '-' || rawTop === 'N/A') ? '' : rawTop;
              const st = String(r[stColIdx] || 'RESOLVIDO').toUpperCase().trim();
              
              const rawDtIni = r[dtColIdx];
              let dtIni = '2026-08-01';
              let dtIniFormatada = '01/08/2026  00:00:00';

              if (rawDtIni !== null && rawDtIni !== undefined && rawDtIni !== '') {
                if (rawDtIni instanceof Date) {
                  const d = String(rawDtIni.getUTCDate()).padStart(2, '0');
                  const m = String(rawDtIni.getUTCMonth() + 1).padStart(2, '0');
                  const y = String(rawDtIni.getUTCFullYear());
                  const hh = String(rawDtIni.getUTCHours()).padStart(2, '0');
                  const mm = String(rawDtIni.getUTCMinutes()).padStart(2, '0');
                  const ss = String(rawDtIni.getUTCSeconds()).padStart(2, '0');
                  dtIni = `${y}-${m}-${d}`;
                  dtIniFormatada = `${d}/${m}/${y}  ${hh}:${mm}:${ss}`;
                } else {
                  const num = Number(rawDtIni);
                  if (!isNaN(num) && num > 40000 && num < 60000) {
                    try {
                      const dateObj = new Date((num - 25569) * 86400 * 1000 + 12 * 60 * 60 * 1000);
                      if (!isNaN(dateObj.getTime())) {
                        const d = String(dateObj.getUTCDate()).padStart(2, '0');
                        const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
                        const y = String(dateObj.getUTCFullYear());
                        const hh = String(dateObj.getUTCHours()).padStart(2, '0');
                        const mm = String(dateObj.getUTCMinutes()).padStart(2, '0');
                        const ss = String(dateObj.getUTCSeconds()).padStart(2, '0');
                        dtIni = `${y}-${m}-${d}`;
                        dtIniFormatada = `${d}/${m}/${y}  ${hh}:${mm}:${ss}`;
                      }
                    } catch (e) {}
                  } else {
                    const s = String(rawDtIni).trim();
                    const ymd = s.match(/^(\d{4})[-/.](\d{2})[-/.](\d{2})(\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
                    if (ymd) {
                      dtIni = `${ymd[1]}-${ymd[2]}-${ymd[3]}`;
                      const hh = (ymd[5] || '00').padStart(2, '0');
                      const mm = (ymd[6] || '00').padStart(2, '0');
                      const ss = (ymd[7] || '00').padStart(2, '0');
                      dtIniFormatada = `${ymd[3]}/${ymd[2]}/${ymd[1]}  ${hh}:${mm}:${ss}`;
                    } else {
                      const dmy = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
                      if (dmy) {
                        const d = dmy[1].padStart(2, '0');
                        const m = dmy[2].padStart(2, '0');
                        const y = dmy[3];
                        const hh = (dmy[5] || '00').padStart(2, '0');
                        const mm = (dmy[6] || '00').padStart(2, '0');
                        const ss = (dmy[7] || '00').padStart(2, '0');
                        dtIni = `${y}-${m}-${d}`;
                        dtIniFormatada = `${d}/${m}/${y}  ${hh}:${mm}:${ss}`;
                      } else {
                        dtIni = s.slice(0, 10);
                        dtIniFormatada = s;
                      }
                    }
                  }
                }
              }

              const clientes = Number(r[cliColIdx]) || 0;

              allOutageRows.push({
                id: i + 1,
                numeroEvento: evento,
                mes: mes,
                semana: sem,
                cidade: cid,
                catProd2: cat,
                tipo: tp,
                tipoOutage: tp,
                topologia: topologia,
                status: st,
                dataInicio: dtIni,
                dataInicioFormatada: dtIniFormatada,
                clientesAfetados: isNaN(clientes) ? 0 : clientes
              });
            });
          }
        }

        if (allAt5Rows.length > 0 && allOutageRows.length > 0) {
          setAt5Data(allAt5Rows);
          setOutageData(allOutageRows);
          setSuccessMsg(`Planilha completa carregada: ${allAt5Rows.length.toLocaleString()} visitas AT5 e ${allOutageRows.length.toLocaleString()} eventos Outage com todas as topologias.`);
        } else if (allAt5Rows.length > 0) {
          setAt5Data(allAt5Rows);
          setSuccessMsg(`Planilha AT5 carregada: ${allAt5Rows.length.toLocaleString()} registros de visitas com topologias.`);
        } else if (allOutageRows.length > 0) {
          setOutageData(allOutageRows);
          setSuccessMsg(`Planilha Outage carregada: ${allOutageRows.length.toLocaleString()} eventos com todas as topologias e tipos de eventos corretos.`);
        } else {
          setError('Não foi possível identificar as colunas da planilha (Outage ou AT5). Verifique se os cabeçalhos de Topologia e Tipo de Evento estão presentes.');
        }
      } catch (err: any) {
        console.error(err);
        setError('Erro ao processar arquivo Excel: ' + (err?.message || 'Arquivo inválido'));
      } finally {
        setIsLoading(false);
        setTimeout(() => setSuccessMsg(null), 5000);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Helper badge color for status
  const getStatusBadge = (status: string) => {
    const s = (status || '').toUpperCase();
    if (s.includes('RESOLV')) return 'bg-emerald-50 text-emerald-700 border-emerald-200';
    if (s.includes('FECH')) return 'bg-blue-50 text-blue-700 border-blue-200';
    if (s.includes('CANC')) return 'bg-red-50 text-[#EE1D23] border-red-200';
    if (s.includes('PROG') || s.includes('ANDAM')) return 'bg-amber-50 text-amber-700 border-amber-200';
    if (s.includes('DESIG')) return 'bg-purple-50 text-purple-700 border-purple-200';
    if (s.includes('PEND')) return 'bg-orange-50 text-orange-700 border-orange-200';
    return 'bg-slate-100 text-slate-700 border-slate-200';
  };

  // Categorias ativas para a Matriz de Tipos de Eventos
  const activeColumns = useMemo(() => {
    return Array.from(
      new Set([...CAT_PROD_2_DEFAULT, ...filteredOutageData.map(d => d.catProd2).filter(Boolean)])
    ).sort((a, b) => {
      const idxA = CAT_PROD_2_DEFAULT.indexOf(a);
      const idxB = CAT_PROD_2_DEFAULT.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b, 'pt-BR');
    });
  }, [filteredOutageData]);

  // QUADRO DE TIPOS DE EVENTOS (Matriz Dinâmica Tipo de Evento x Cat. Prod. 2)
  const typeMatrixData = useMemo(() => {
    const rowsMap: Record<string, { tipo: string; counts: Record<string, number>; total: number }> = {};
    const colTotals: Record<string, number> = {};
    activeColumns.forEach(col => { colTotals[col] = 0; });
    let grandTotal = 0;

    filteredOutageData.forEach(item => {
      let tipo = (item.tipo || item.tipoOutage || 'NÃO ESPECIFICADO').trim().toUpperCase();
      if (
        !tipo ||
        tipo.includes('VISITA') ||
        tipo.includes('CONTROLE REMOTO') ||
        tipo.includes('SUBSTITUICAO') ||
        tipo.includes('INSTALACAO') ||
        tipo.includes('DESCONEXAO')
      ) {
        tipo = 'EMERGENCIAL';
      }
      const col = item.catProd2 || 'OUTROS';

      if (!rowsMap[tipo]) {
        rowsMap[tipo] = { tipo, counts: {}, total: 0 };
        activeColumns.forEach(c => { rowsMap[tipo].counts[c] = 0; });
      }

      rowsMap[tipo].counts[col] = (rowsMap[tipo].counts[col] || 0) + 1;
      rowsMap[tipo].total += 1;
      colTotals[col] = (colTotals[col] || 0) + 1;
      grandTotal += 1;
    });

    let rows = Object.values(rowsMap).sort((a, b) => b.total - a.total);

    if (typeMatrixSearch.trim()) {
      const term = typeMatrixSearch.toLowerCase();
      rows = rows.filter(r => r.tipo.toLowerCase().includes(term));
    }

    return {
      columns: activeColumns,
      rows,
      colTotals,
      grandTotal
    };
  }, [filteredOutageData, activeColumns, typeMatrixSearch]);

  const handleExportTypeMatrixExcel = () => {
    const exportRows = typeMatrixData.rows.map(row => {
      const rowObj: Record<string, any> = {
        'Rótulos de Linha (Tipo)': row.tipo
      };
      typeMatrixData.columns.forEach(col => {
        rowObj[col] = row.counts[col] || 0;
      });
      rowObj['Total Geral'] = row.total;
      return rowObj;
    });

    const totalRowObj: Record<string, any> = {
      'Rótulos de Linha (Tipo)': 'Total Geral'
    };
    typeMatrixData.columns.forEach(col => {
      totalRowObj[col] = typeMatrixData.colTotals[col] || 0;
    });
    totalRowObj['Total Geral'] = typeMatrixData.grandTotal;
    exportRows.push(totalRowObj);

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Tipos_de_Eventos');
    XLSX.writeFile(workbook, `Quadro_Tipos_CatProd2_Churn_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  return (
    <div className="w-full space-y-8 pb-16 animate-fade-in" id="churn-dashboard">
      {/* File Loading Overlay while processing Excel or syncing */}
      <FileLoadingOverlay
        isOpen={isLoading}
        title="Carregando e Processando Arquivo..."
        subtitle="Correlacionando Início do Outage x ABERTURA_SOLIC e filtrando topologias..."
        currentStep="Lendo planilhas, mapeando datas (DD/MM/YYYY HH:mm:ss) e calculando correlações..."
      />

      {/* Action Header Card - Claro Corporate Standard (Mesmo padrão da aba AT5) */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white p-6 rounded-3xl shadow-md border border-slate-100 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-14 h-14 bg-red-50 rounded-2xl flex items-center justify-center text-[#EE1D23] shadow-inner">
              <UserMinus className="w-7 h-7" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded-md bg-[#EE1D23] text-white text-[10px] font-black uppercase tracking-wider">
                  MÓDULO
                </span>
                <h1 className="text-2xl font-black text-[#333333] tracking-tight uppercase italic">
                  PAINEL DE CHURN & IMPACTO EM REDE
                </h1>
              </div>
              <p className="text-xs font-bold text-slate-400 mt-0.5">
                Cruzamento Analítico entre Eventos de Indisponibilidade de Rede (Outage) e Ordens de Serviço / Visitas (AT5)
              </p>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={handleGithubSync}
              disabled={isLoading}
              className="flex items-center gap-2 bg-[#EE1D23] hover:bg-red-600 text-white font-black py-2.5 px-4 rounded-xl transition-all shadow-md shadow-red-500/15 active:scale-95 uppercase italic text-xs cursor-pointer disabled:opacity-50"
              title="Sincronizar dados com o repositório GitHub"
            >
              <Activity className="w-3.5 h-3.5" />
              <span>{isLoading ? 'Carregando...' : 'Sincronizar GitHub'}</span>
            </button>

            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-800 font-black py-2.5 px-4 rounded-xl border border-slate-200 transition-all shadow-xs active:scale-95 uppercase italic text-xs cursor-pointer"
              title="Importar planilha Excel (Outage ou AT5)"
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

            {/* Botão Atualizar solicitado pelo usuário */}
            <button
              onClick={handleRefresh}
              disabled={isLoading}
              className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-black py-2.5 px-4 rounded-xl transition-all shadow-xs active:scale-95 uppercase italic text-xs cursor-pointer disabled:opacity-50"
              title="Atualizar e resincronizar os cruzamentos de dados entre Outage e AT5"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
              <span>Atualizar</span>
            </button>

            <button
              onClick={handleExportExcel}
              className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all shadow-xs active:scale-95 uppercase italic text-xs cursor-pointer"
              title="Exportar dados cruzados de Churn para Excel"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span>Exportar</span>
            </button>

            <button
              onClick={handleRestoreDefaults}
              className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all active:scale-95 text-xs cursor-pointer"
              title="Restaurar dados padrão de exemplo cruzados"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Restaurar</span>
            </button>

            {/* Botão Limpar solicitado pelo usuário */}
            <button
              onClick={handleClearAll}
              className="flex items-center gap-2 bg-red-50 hover:bg-red-100 text-[#EE1D23] font-black py-2.5 px-4 rounded-xl border border-red-200 transition-all active:scale-95 uppercase italic text-xs cursor-pointer"
              title="Limpar todos os dados e filtros do Painel de Churn"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>Limpar</span>
            </button>
          </div>
        </div>
      </section>

      {/* Notifications / Alerts */}
      {successMsg && (
        <section className="w-full max-w-[1600px] mx-auto">
          <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 px-4 py-3 rounded-2xl text-xs font-bold flex items-center justify-between">
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-600" />
              <span>{successMsg}</span>
            </div>
            <button onClick={() => setSuccessMsg(null)}>
              <X className="w-4 h-4" />
            </button>
          </div>
        </section>
      )}

      {error && (
        <section className="w-full max-w-[1600px] mx-auto">
          <div className="bg-red-50 border border-red-200 text-red-800 px-4 py-3 rounded-2xl text-xs font-bold flex items-center justify-between">
            <div className="flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-[#EE1D23]" />
              <span>{error}</span>
            </div>
            <button onClick={() => setError(null)}>
              <X className="w-4 h-4" />
            </button>
          </div>
        </section>
      )}

      {/* Dynamic Filter Controls Block (Mesmo padrão da aba AT5) */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white border border-slate-200/80 rounded-[32px] p-6 shadow-sm print:hidden" id="churn-filters-card">
          <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-100">
            <div className="flex items-center gap-2">
              <SlidersHorizontal className="w-4 h-4 text-[#EE1D23]" />
              <h3 className="text-sm font-black text-slate-800 uppercase italic tracking-tight">
                Filtros do Painel de Churn (Cruzamento Outage x AT5)
              </h3>
            </div>
            {(filters.municipio.length > 0 ||
              filters.mes.length > 0 ||
              filters.startDate ||
              filters.endDate ||
              filters.catProd2.length > 0 ||
              filters.tipoEvento.length > 0 ||
              filters.statusOutage.length > 0 ||
              filters.topologia.length > 0) && (
              <button
                onClick={() => setFilters({
                  municipio: [],
                  mes: [],
                  startDate: '',
                  endDate: '',
                  catProd2: [],
                  tipoEvento: [],
                  statusOutage: [],
                  topologia: []
                })}
                className="text-[10px] font-black text-[#EE1D23] hover:underline uppercase tracking-wider cursor-pointer"
              >
                Limpar Todos os Filtros
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 lg:grid-cols-4 gap-4">
            {/* 1. MUNICÍPIO */}
            <MultiFilterSelect
              label="MUNICÍPIO"
              icon={<MapPin className="w-3.5 h-3.5" />}
              value={filters.municipio}
              options={filterOptions.municipio}
              onChange={(val) => setFilters(prev => ({ ...prev, municipio: val.filter(v => v !== 'Todos') }))}
            />

            {/* 2. MÊS */}
            <MultiFilterSelect
              label="MÊS"
              icon={<Calendar className="w-3.5 h-3.5" />}
              value={filters.mes}
              options={filterOptions.mes}
              onChange={(val) => setFilters(prev => ({ ...prev, mes: val.filter(v => v !== 'Todos') }))}
            />

            {/* 3. DATA INÍCIO */}
            <div className="flex flex-col gap-1.5">
              <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider flex items-center gap-1.5">
                <Calendar className="w-3.5 h-3.5 text-[#EE1D23]" />
                DATA INÍCIO
              </span>
              <input
                type="date"
                value={filters.startDate}
                onChange={(e) => setFilters(prev => ({ ...prev, startDate: e.target.value }))}
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-700 outline-none focus:border-[#EE1D23] transition-all cursor-pointer"
              />
            </div>

            {/* 4. DATA FIM */}
            <div className="flex flex-col gap-1.5">
              <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider flex items-center gap-1.5">
                <Calendar className="w-3.5 h-3.5 text-[#EE1D23]" />
                DATA FIM
              </span>
              <input
                type="date"
                value={filters.endDate}
                onChange={(e) => setFilters(prev => ({ ...prev, endDate: e.target.value }))}
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-700 outline-none focus:border-[#EE1D23] transition-all cursor-pointer"
              />
            </div>

            {/* 5. CAT. PROD. 2 */}
            <MultiFilterSelect
              label="CAT. PROD. 2"
              icon={<Layers className="w-3.5 h-3.5" />}
              value={filters.catProd2}
              options={filterOptions.catProd2}
              onChange={(val) => setFilters(prev => ({ ...prev, catProd2: val.filter(v => v !== 'Todos') }))}
            />

            {/* 6. TIPO DE EVENTO */}
            <MultiFilterSelect
              label="TIPO DE EVENTO"
              icon={<Radio className="w-3.5 h-3.5" />}
              value={filters.tipoEvento}
              options={filterOptions.tipoEvento}
              onChange={(val) => setFilters(prev => ({ ...prev, tipoEvento: val.filter(v => v !== 'Todos') }))}
            />

            {/* 7. STATUS OUTAGE */}
            <MultiFilterSelect
              label="STATUS OUTAGE"
              icon={<Activity className="w-3.5 h-3.5" />}
              value={filters.statusOutage}
              options={filterOptions.statusOutage}
              onChange={(val) => setFilters(prev => ({ ...prev, statusOutage: val.filter(v => v !== 'Todos') }))}
            />

            {/* 8. TOPOLOGIA */}
            <MultiFilterSelect
              label="TOPOLOGIA"
              icon={<Building2 className="w-3.5 h-3.5" />}
              value={filters.topologia}
              options={filterOptions.topologia}
              onChange={(val) => setFilters(prev => ({ ...prev, topologia: val.filter(v => v !== 'Todos') }))}
            />
          </div>

          {/* Active filter badges */}
          {Object.entries(filters).some(([k, v]) => Array.isArray(v) ? v.length > 0 : Boolean(v)) && (
            <div className="flex flex-wrap gap-2 items-center mt-4 pt-4 border-t border-slate-100">
              <span className="text-[9px] font-black text-slate-400 uppercase tracking-wider select-none">Filtros Ativos:</span>
              {filters.municipio.map(m => (
                <span key={`mun-${m}`} className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md text-[10px] font-black uppercase">
                  Cidade: {m}
                  <button onClick={() => setFilters(p => ({ ...p, municipio: p.municipio.filter(v => v !== m) }))}><X className="w-3 h-3" /></button>
                </span>
              ))}
              {filters.mes.map(m => (
                <span key={`mes-${m}`} className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md text-[10px] font-black uppercase">
                  Mês: {m}
                  <button onClick={() => setFilters(p => ({ ...p, mes: p.mes.filter(v => v !== m) }))}><X className="w-3 h-3" /></button>
                </span>
              ))}
              {filters.startDate && (
                <span className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md text-[10px] font-black uppercase">
                  Início: {filters.startDate}
                  <button onClick={() => setFilters(p => ({ ...p, startDate: '' }))}><X className="w-3 h-3" /></button>
                </span>
              )}
              {filters.endDate && (
                <span className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md text-[10px] font-black uppercase">
                  Fim: {filters.endDate}
                  <button onClick={() => setFilters(p => ({ ...p, endDate: '' }))}><X className="w-3 h-3" /></button>
                </span>
              )}
              {filters.catProd2.map(c => (
                <span key={`cat-${c}`} className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md text-[10px] font-black uppercase">
                  Cat: {c}
                  <button onClick={() => setFilters(p => ({ ...p, catProd2: p.catProd2.filter(v => v !== c) }))}><X className="w-3 h-3" /></button>
                </span>
              ))}
              {filters.tipoEvento.map(t => (
                <span key={`tipo-${t}`} className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md text-[10px] font-black uppercase">
                  Tipo: {t}
                  <button onClick={() => setFilters(p => ({ ...p, tipoEvento: p.tipoEvento.filter(v => v !== t) }))}><X className="w-3 h-3" /></button>
                </span>
              ))}
              {filters.statusOutage.map(s => (
                <span key={`st-${s}`} className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md text-[10px] font-black uppercase">
                  Status: {s}
                  <button onClick={() => setFilters(p => ({ ...p, statusOutage: p.statusOutage.filter(v => v !== s) }))}><X className="w-3 h-3" /></button>
                </span>
              ))}
              {filters.topologia.map(t => (
                <span key={`top-${t}`} className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md text-[10px] font-black uppercase">
                  Topologia: {t}
                  <button onClick={() => setFilters(p => ({ ...p, topologia: p.topologia.filter(v => v !== t) }))}><X className="w-3 h-3" /></button>
                </span>
              ))}
            </div>
          )}
        </div>
      </section>

      {/* Core Metric Cards Section (Mesmo padrão da aba AT5 e da imagem anexada) */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4" id="churn-metrics-grid">
          {/* Card 1: OS COM PADRÃO (Visitas) */}
          <div className="bg-white border border-slate-200/80 rounded-2xl p-4 sm:p-5 shadow-xs flex flex-col justify-between hover:translate-y-[-2px] transition-all relative overflow-hidden group">
            <div className="absolute top-0 right-0 w-20 h-20 bg-emerald-50/20 rounded-bl-full z-0 group-hover:scale-110 transition-transform" />
            <div className="relative z-10">
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-wider block mb-1">
                OS COM PADRÃO
              </span>
              <p className="text-2xl sm:text-3xl font-black text-[#333333] tracking-tight">
                {metrics.visitasComPadrao.toLocaleString('pt-BR')}
              </p>
              <div className="text-[10px] text-slate-400 font-bold uppercase mt-1 flex items-center gap-1 select-none">
                <CheckCircle2 className="w-3 h-3 text-emerald-500" />
                ({metrics.visitasComPadrao} OS ordens no total)
              </div>
            </div>
          </div>

          {/* Card 2: OS SEM PADRÃO (Visitas) */}
          <div className="bg-white border border-slate-200/80 rounded-2xl p-4 sm:p-5 shadow-xs flex flex-col justify-between hover:translate-y-[-2px] transition-all relative overflow-hidden group">
            <div className="absolute top-0 right-0 w-20 h-20 bg-red-50/20 rounded-bl-full z-0 group-hover:scale-110 transition-transform" />
            <div className="relative z-10">
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-wider block mb-1">
                OS SEM PADRÃO
              </span>
              <p className="text-2xl sm:text-3xl font-black text-[#333333] tracking-tight">
                {metrics.visitasSemPadrao.toLocaleString('pt-BR')}
              </p>
              <div className="text-[10px] text-slate-400 font-bold uppercase mt-1 flex items-center gap-1 select-none">
                <XCircle className="w-3 h-3 text-red-500" />
                ({metrics.visitasSemPadrao} OS ordens no total)
              </div>
            </div>
          </div>

          {/* Card 3: VOLUME TOTAL DE OS (Visitas) */}
          <div className="bg-white border border-slate-200/80 rounded-2xl p-4 sm:p-5 shadow-xs flex flex-col justify-between hover:translate-y-[-2px] transition-all relative overflow-hidden group">
            <div className="absolute top-0 right-0 w-20 h-20 bg-slate-50 rounded-bl-full z-0 group-hover:scale-110 transition-transform" />
            <div className="relative z-10">
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-wider block mb-1">
                VOLUME TOTAL DE OS
              </span>
              <p className="text-2xl sm:text-3xl font-black text-[#333333] tracking-tight">
                {metrics.totalVisitas.toLocaleString('pt-BR')}
              </p>
              <div className="text-[10px] text-slate-400 font-bold uppercase mt-1 flex items-center gap-1 select-none">
                <ListTodo className="w-3 h-3 text-slate-500" />
                ({metrics.totalVisitas} OS ordens no total)
              </div>
            </div>
          </div>

          {/* Card 4: TOTAL DE OUTAGES (Eventos de Rede) */}
          <div className="bg-slate-900 rounded-2xl p-4 sm:p-5 shadow-md flex flex-col justify-between hover:translate-y-[-2px] transition-all relative overflow-hidden group">
            <div className="absolute -right-2 -top-2 w-28 h-28 bg-[#EE1D23]/10 rounded-full blur-xl pointer-events-none" />
            <div className="absolute top-0 right-0 w-20 h-20 bg-[#EE1D23]/5 rounded-bl-full z-0 group-hover:scale-110 transition-transform" />
            <div className="relative z-10">
              <span className="text-[10px] font-black text-[#EE1D23] uppercase tracking-wider block mb-1">
                TOTAL DE OUTAGES
              </span>
              <p className="text-2xl sm:text-3xl font-black text-white tracking-tight">
                {metrics.totalOutage.toLocaleString('pt-BR')}
              </p>
              <div className="text-[9px] text-slate-300 font-bold uppercase mt-1 flex items-center gap-1 select-none">
                <Radio className="w-3 h-3 text-[#EE1D23]" />
                {metrics.distinctNodes} topologias ({metrics.totalClientesAfetados.toLocaleString('pt-BR')} clientes)
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* QUADRO: TIPOS DE EVENTOS (MATRIZ DINÂMICA) */}
      <section className="w-full max-w-[1600px] mx-auto" id="secao-tipos-de-eventos">
        <div className="bg-white rounded-3xl shadow-md border border-slate-100 overflow-hidden" id="card-tipos-eventos">
          {/* Header do Quadro de Tipos de Eventos */}
          <div className="p-6 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-gradient-to-r from-red-50/90 via-red-50/40 to-white">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-2xl bg-red-100 text-[#EE1D23] flex items-center justify-center shadow-xs">
                <Table className="w-6 h-6" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="px-2.5 py-0.5 rounded-full bg-[#EE1D23] text-white text-[10px] font-black uppercase tracking-wider">
                    Matriz Dinâmica
                  </span>
                  <h3 className="text-lg font-black text-slate-900 uppercase italic tracking-tight">
                    TIPOS DE EVENTOS
                  </h3>
                </div>
                <p className="text-xs font-bold text-slate-500 mt-0.5">
                  Cruzamento das categorias de infraestrutura (Rótulos de Coluna) com a classificação do Tipo de Evento (Linhas)
                </p>
              </div>
            </div>

            <div className="flex items-center gap-3">
              {/* Busca de Tipo */}
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input 
                  type="text"
                  placeholder="Filtrar tipo no quadro..."
                  value={typeMatrixSearch}
                  onChange={(e) => setTypeMatrixSearch(e.target.value)}
                  className="bg-white border border-red-200 rounded-xl pl-9 pr-4 py-2 text-xs font-bold text-slate-700 outline-none focus:border-[#EE1D23] transition-all w-52 shadow-2xs"
                  id="input-busca-tipo-evento"
                />
              </div>

              <button
                onClick={handleExportTypeMatrixExcel}
                className="flex items-center gap-1.5 bg-[#EE1D23] hover:bg-[#D91A20] text-white font-bold py-2 px-3.5 rounded-xl transition-all shadow-xs active:scale-95 text-xs uppercase italic cursor-pointer"
                title="Exportar esta matriz para Excel"
                id="btn-export-tipos-eventos"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Exportar Matriz</span>
              </button>
            </div>
          </div>

          {/* Pivot Table Display Styled matching the Claro corporate colors */}
          <div className="overflow-x-auto">
            <table className="w-full text-center border-collapse border border-red-200 text-xs" id="tabela-matriz-tipos-eventos">
              <thead>
                <tr className="bg-[#EE1D23] border-b-2 border-red-400 font-black text-white">
                  <th className="py-3 px-4 text-left border-r border-red-400/50 font-extrabold uppercase tracking-wider text-[11px] min-w-[240px] bg-[#C81016]">
                    Rótulos de Linha (Tipo)
                  </th>
                  {typeMatrixData.columns.map(col => (
                    <th key={col} className="py-3 px-3 border-r border-red-400/50 font-extrabold uppercase tracking-wider text-[11px] whitespace-nowrap">
                      {col}
                    </th>
                  ))}
                  <th className="py-3 px-4 font-black uppercase tracking-wider text-[11px] bg-[#991B1B] text-white whitespace-nowrap">
                    Total Geral
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-red-100">
                {typeMatrixData.rows.length > 0 ? (
                  typeMatrixData.rows.map((row, idx) => (
                    <tr 
                      key={row.tipo}
                      className={cn(
                        "transition-colors hover:bg-red-50/60",
                        idx % 2 === 0 ? "bg-white" : "bg-red-50/20"
                      )}
                    >
                      <td className="py-3 px-4 text-left font-black text-slate-800 border-r border-red-100 max-w-sm truncate">
                        {row.tipo}
                      </td>
                      {typeMatrixData.columns.map(col => {
                        const val = row.counts[col] || 0;
                        return (
                          <td 
                            key={col} 
                            className={cn(
                              "py-3 px-3 border-r border-red-100 font-mono font-bold transition-all text-sm",
                              val > 0 ? "text-slate-900 font-black" : "text-slate-200",
                              val > 100 ? "bg-red-50 font-black text-[#EE1D23]" : ""
                            )}
                          >
                            {val > 0 ? val.toLocaleString('pt-BR') : '-'}
                          </td>
                        );
                      })}
                      <td className="py-3 px-4 font-mono font-black text-slate-900 bg-red-50/90 text-sm">
                        {row.total.toLocaleString('pt-BR')}
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={typeMatrixData.columns.length + 2} className="py-12 text-center text-slate-400 font-bold">
                      Nenhum registro encontrado para este cruzamento.
                    </td>
                  </tr>
                )}

                {/* Total Geral Footer Row */}
                <tr className="bg-[#EE1D23] border-t-2 border-red-400 font-black text-white">
                  <td className="py-3.5 px-4 text-left border-r border-red-400/50 font-black uppercase text-[11px] bg-[#C81016]">
                    Total Geral
                  </td>
                  {typeMatrixData.columns.map(col => (
                    <td key={col} className="py-3.5 px-3 border-r border-red-400/50 font-mono font-black text-base">
                      {(typeMatrixData.colTotals[col] || 0).toLocaleString('pt-BR')}
                    </td>
                  ))}
                  <td className="py-3.5 px-4 font-mono font-black text-base bg-[#991B1B] text-white">
                    {typeMatrixData.grandTotal.toLocaleString('pt-BR')}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </section>

      {/* Gráfico puxado da aba Outage: Evolução Diária de Eventos (Outage) */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white p-6 sm:p-8 rounded-3xl shadow-md border border-slate-100 flex flex-col">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded-full bg-[#EE1D23] text-white text-[10px] font-black uppercase tracking-wider">
                  Volume Diário
                </span>
                <h3 className="text-lg font-black text-slate-900 uppercase italic tracking-tight">
                  Evolução Diária de Eventos (Outage)
                </h3>
              </div>
              <p className="text-xs font-bold text-slate-400 mt-0.5">
                Distribuição temporal com contagem diária de ocorrências por data de início
              </p>
            </div>

            {/* Legenda Customizada com Cores Oficiais e Todos os Status */}
            <div className="flex flex-wrap items-center gap-2 sm:gap-3 text-xs font-bold bg-slate-50 px-3.5 py-2 rounded-2xl border border-slate-100">
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded-md bg-[#059669] shadow-2xs" />
                <span className="text-slate-700">Resolvido</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded-md bg-[#2563EB] shadow-2xs" />
                <span className="text-slate-700">Fechado</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded-md bg-[#EE1D23] shadow-2xs" />
                <span className="text-slate-700">Cancelado</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded-md bg-[#F59E0B] shadow-2xs" />
                <span className="text-slate-700">Em Progresso</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded-md bg-[#8B5CF6] shadow-2xs" />
                <span className="text-slate-700">Designado</span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className="w-3 h-3 rounded-md bg-[#F97316] shadow-2xs" />
                <span className="text-slate-700">Pendente</span>
              </div>
            </div>
          </div>

          <div className="h-96 w-full">
            {dailyOutageChartData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart
                  data={dailyOutageChartData}
                  margin={{ top: 28, right: 15, left: -15, bottom: dailyOutageChartData.length > 12 ? 35 : 20 }}
                >
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#F1F5F9" />
                  <XAxis
                    dataKey="displayDate"
                    tick={{ fontSize: dailyOutageChartData.length > 20 ? 9 : 10, fontWeight: 800, fill: '#334155' }}
                    axisLine={{ stroke: '#CBD5E1' }}
                    tickLine={false}
                    interval={0}
                    angle={dailyOutageChartData.length > 12 ? -45 : 0}
                    textAnchor={dailyOutageChartData.length > 12 ? 'end' : 'middle'}
                    height={dailyOutageChartData.length > 12 ? 45 : 30}
                  />
                  <YAxis
                    tick={{ fontSize: 11, fontWeight: 700, fill: '#64748B' }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip
                    content={({ active, payload, label }) => {
                      if (active && payload && payload.length) {
                        const totalDay = payload[0]?.payload?.Total || 0;
                        return (
                          <div className="bg-slate-900 text-white p-3.5 rounded-2xl shadow-xl text-xs space-y-2 border border-slate-800 min-w-[200px]">
                            <div className="border-b border-slate-700 pb-2">
                              <span className="text-[10px] font-bold uppercase tracking-wider text-red-400 block">Data de Início</span>
                              <h5 className="font-black text-white text-sm">{label}</h5>
                              <div className="mt-1 flex items-center justify-between text-xs font-black bg-slate-800/80 px-2 py-1 rounded-lg">
                                <span className="text-slate-300">Volume Total:</span>
                                <span className="text-white font-mono text-sm">{totalDay.toLocaleString()}</span>
                              </div>
                            </div>
                            <div className="space-y-1 pt-0.5">
                              {payload.filter((entry: any) => entry.dataKey !== 'Total' && entry.value > 0).map((entry: any) => (
                                <div key={entry.name} className="flex justify-between items-center text-xs">
                                  <div className="flex items-center gap-1.5">
                                    <span className="w-2.5 h-2.5 rounded-xs" style={{ backgroundColor: entry.color }} />
                                    <span className="font-medium text-slate-200">{entry.name}:</span>
                                  </div>
                                  <span className="font-mono font-black text-white">{entry.value.toLocaleString()}</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Bar dataKey="CANCELADO" name="CANCELADO" stackId="statusStack" fill="#EE1D23" />
                  <Bar dataKey="DESIGNADO" name="DESIGNADO" stackId="statusStack" fill="#8B5CF6" />
                  <Bar dataKey="EM PROGRESSO" name="EM PROGRESSO" stackId="statusStack" fill="#F59E0B" />
                  <Bar dataKey="FECHADO" name="FECHADO" stackId="statusStack" fill="#2563EB" />
                  <Bar dataKey="PENDENTE" name="PENDENTE" stackId="statusStack" fill="#F97316" />
                  <Bar dataKey="RESOLVIDO" name="RESOLVIDO" stackId="statusStack" fill="#059669" radius={[4, 4, 0, 0]} />

                  {/* Linha transparente superior para ancorar e renderizar o VOLUME TOTAL acima de cada barra */}
                  <Line
                    type="monotone"
                    dataKey="Total"
                    stroke="transparent"
                    dot={false}
                    activeDot={false}
                    isAnimationActive={false}
                  >
                    <LabelList
                      dataKey="Total"
                      position="top"
                      offset={6}
                      content={(props: any) => {
                        const { x, y, value } = props;
                        if (!value || value <= 0) return null;
                        return (
                          <text
                            x={x}
                            y={y - 4}
                            fill="#1E293B"
                            textAnchor="middle"
                            fontSize={10}
                            fontWeight="800"
                            fontFamily="monospace"
                          >
                            {value}
                          </text>
                        );
                      }}
                    />
                  </Line>
                </ComposedChart>
              </ResponsiveContainer>
            ) : (
              <div className="h-full flex items-center justify-center text-slate-400 font-bold text-xs">
                Nenhum evento de Outage encontrado para o período selecionado.
              </div>
            )}
          </div>
        </div>
      </section>

      {/* Quadro / Tabela de Cruzamento: Registros Analíticos de Churn & Topologia */}
      {/* Mesmo padrão dos Registros Analíticos de Outage (conforme solicitado e visto na Imagem 2) */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white rounded-3xl shadow-md border border-slate-100 overflow-hidden">
          <div className="p-4 sm:p-6 border-b border-slate-100 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded-full bg-slate-900 text-white text-[10px] font-black uppercase tracking-wider">
                  RANKING ATÉ 100 LINHAS
                </span>
                <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tight">
                  REGISTROS ANALÍTICOS DE CHURN & IMPACTO EM REDE
                </h3>
              </div>
              <p className="text-xs font-bold text-slate-400 mt-0.5">
                Cruzamento de Ordens de Visitas (AT5) e Indisponibilidades (Outage) por Topologia e Tipo de Evento
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-3">
              {/* Order selector (Maior para o Menor) */}
              <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1.5">
                <span className="text-[10px] font-black text-slate-400 uppercase tracking-wider">Ordenar por:</span>
                <select
                  value={sortField}
                  onChange={(e) => setSortField(e.target.value as any)}
                  className="bg-transparent text-xs font-black text-slate-700 outline-none cursor-pointer"
                >
                  <option value="visitas">Maior Volume de Visitas</option>
                  <option value="outage">Maior Quantidade de Outage</option>
                  <option value="clientes">Mais Clientes Afetados</option>
                </select>
              </div>

              {/* Search Bar */}
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Buscar topologia, tipo, cidade, categoria..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="bg-slate-50 border border-slate-200 rounded-xl pl-9 pr-4 py-2 text-xs font-bold text-slate-700 outline-none focus:border-[#EE1D23] transition-all w-60 sm:w-72"
                />
              </div>

              {/* Page Size Select */}
              <select
                value={pageSize}
                onChange={(e) => setPageSize(Number(e.target.value))}
                className="bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-700 outline-none focus:border-[#EE1D23] transition-all cursor-pointer"
              >
                <option value={10}>10 por página</option>
                <option value={25}>25 por página</option>
                <option value={50}>50 por página</option>
                <option value={100}>100 por página</option>
              </select>
            </div>
          </div>

          <div className="overflow-x-auto scrollbar-thin scrollbar-thumb-slate-200">
            <table className="w-full text-left border-collapse table-auto">
              <thead>
                <tr className="bg-slate-50/80 border-b border-slate-100 text-[10px] font-black text-slate-400 uppercase tracking-wider">
                  <th className="py-3 px-3 text-center"># Rank</th>
                  <th className="py-3 px-3">Topologia</th>
                  <th className="py-3 px-3">Tipo de Evento</th>
                  <th className="py-3 px-3">Cat. Prod. 2</th>
                  <th className="py-3 px-3">Município</th>
                  <th className="py-3 px-3 text-right cursor-pointer hover:text-slate-700" onClick={() => setSortField('visitas')}>
                    Qtd. Visitas ↕
                  </th>
                  <th className="py-3 px-3 text-right cursor-pointer hover:text-slate-700" onClick={() => setSortField('outage')}>
                    Qtd. Outage ↕
                  </th>
                  <th className="py-3 px-3 text-right">Clientes Afetados</th>
                  <th className="py-3 px-3 text-center">Ações</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-xs">
                {paginatedCrossedData.length > 0 ? (
                  paginatedCrossedData.map((row) => (
                    <tr
                      key={`${row.topologia}-${row.tipoEvento}-${row.rank}`}
                      onClick={() => setSelectedRecord(row)}
                      className="hover:bg-slate-50/80 transition-colors cursor-pointer group"
                    >
                      {/* Ranking */}
                      <td className="py-2.5 px-3 text-center">
                        <span className={cn(
                          "w-5 h-5 rounded-full inline-flex items-center justify-center font-black text-[10px]",
                          row.rank === 1 ? "bg-amber-400 text-white shadow-xs" :
                          row.rank === 2 ? "bg-slate-300 text-slate-800" :
                          row.rank === 3 ? "bg-amber-700 text-white" :
                          "bg-slate-100 text-slate-600"
                        )}>
                          {row.rank}
                        </span>
                      </td>

                      {/* Topologia */}
                      <td className="py-2.5 px-3 font-mono font-black text-slate-800 group-hover:text-[#EE1D23] transition-colors whitespace-nowrap text-[11px]">
                        <div className="flex items-center gap-1.5">
                          <Radio className="w-3.5 h-3.5 text-slate-400 group-hover:text-[#EE1D23]" />
                          <span>{row.topologia}</span>
                        </div>
                      </td>

                      {/* Tipo de Evento */}
                      <td className="py-2.5 px-3 font-bold text-slate-700 max-w-[200px] truncate text-[11px]">
                        <span className="px-1.5 py-0.5 rounded bg-slate-100 border border-slate-200 text-[10px] font-extrabold text-slate-700">
                          {row.tipoEvento}
                        </span>
                      </td>

                      {/* Cat. Prod. 2 */}
                      <td className="py-2.5 px-3 font-bold text-sky-800 whitespace-nowrap">
                        <span className="px-1.5 py-0.5 rounded bg-sky-50 border border-sky-200/60 text-[10px]">
                          {row.catProd2}
                        </span>
                      </td>

                      {/* Município */}
                      <td className="py-2.5 px-3 font-black text-slate-800 uppercase whitespace-nowrap text-[11px]">
                        {row.municipio}
                      </td>

                      {/* Quantidade de Visitas */}
                      <td className="py-2.5 px-3 text-right whitespace-nowrap">
                        <div className="flex flex-col items-end">
                          <span className="px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-800 font-mono font-black text-[11px] border border-emerald-200/60">
                            {row.qtdVisitas.toLocaleString('pt-BR')}
                          </span>
                          {row.totalVisitasPeriodo > row.qtdVisitas && (
                            <span className="text-[9px] text-slate-400 font-bold mt-0.5" title="Total geral de visitas na topologia">
                              Total nó: {row.totalVisitasPeriodo.toLocaleString('pt-BR')}
                            </span>
                          )}
                        </div>
                      </td>

                      {/* Quantidade de Outage */}
                      <td className="py-2.5 px-3 text-right whitespace-nowrap">
                        <span className="px-2 py-0.5 rounded-md bg-red-50 text-[#EE1D23] font-mono font-black text-[11px] border border-red-200/60">
                          {row.qtdOutage.toLocaleString('pt-BR')}
                        </span>
                      </td>

                      {/* Clientes Afetados */}
                      <td className="py-2.5 px-3 text-right font-bold text-slate-600 whitespace-nowrap text-[11px]">
                        {row.clientesAfetados > 0 ? row.clientesAfetados.toLocaleString('pt-BR') : '-'}
                      </td>

                      {/* Ação */}
                      <td className="py-2.5 px-3 text-center whitespace-nowrap">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedRecord(row);
                          }}
                          className="p-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-600 transition-colors"
                          title="Visualizar detalhes do cruzamento"
                        >
                          <Eye className="w-3.5 h-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={9} className="py-16 text-center text-slate-400 font-bold">
                      Nenhum registro encontrado com os filtros selecionados.
                    </td>
                  </tr>
                )}
              </tbody>
              {crossedTableData.length > 0 && (
                <tfoot className="bg-slate-50/95 border-t-2 border-slate-200 text-xs font-black text-slate-800">
                  <tr>
                    <td className="py-3 px-4 text-center text-slate-400 font-bold uppercase text-[10px]">TOTAL</td>
                    <td className="py-3 px-4 text-slate-900 font-black">
                      Consolidado ({new Set(crossedTableData.map(r => r.topologia)).size} Topologias)
                    </td>
                    <td className="py-3 px-4 text-slate-400">-</td>
                    <td className="py-3 px-4 text-slate-400">-</td>
                    <td className="py-3 px-4 text-slate-700">
                      {filters.municipio.length > 0 ? filters.municipio.join(', ') : 'Geral'}
                    </td>
                    <td className="py-3 px-4 text-right">
                      <span className="inline-flex items-center px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-700 font-mono font-black text-xs border border-emerald-200 shadow-sm">
                        {totalTableVisitas.toLocaleString('pt-BR')}
                      </span>
                    </td>
                    <td className="py-3 px-4 text-right">
                      <span className="inline-flex items-center px-2.5 py-1 rounded-lg bg-red-50 text-[#EE1D23] font-mono font-black text-xs border border-red-200 shadow-sm">
                        {totalTableOutages.toLocaleString('pt-BR')}
                      </span>
                    </td>
                    <td className="py-3 px-4 text-right font-black text-slate-800">
                      {totalTableClientes > 0 ? totalTableClientes.toLocaleString('pt-BR') : '-'}
                    </td>
                    <td className="py-3 px-4 text-center text-slate-400">-</td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>

          {/* Table Footer with Pagination */}
          <div className="p-4 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs font-bold text-slate-500">
            <div>
              Mostrando {crossedTableData.length === 0 ? 0 : (currentPage - 1) * pageSize + 1} a {Math.min(currentPage * pageSize, crossedTableData.length)} de {crossedTableData.length} registros ranqueados (máx. 100)
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="p-2 rounded-xl border border-slate-200 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <span className="px-3 py-1 font-black text-slate-700">
                Página {currentPage} de {totalPages}
              </span>
              <button
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="p-2 rounded-xl border border-slate-200 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      </section>

      {/* Detail Modal for Selected Crossed Topology */}
      {selectedRecord && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl max-w-2xl w-full p-6 shadow-2xl border border-slate-100 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-4 border-b border-slate-100 mb-4">
              <div>
                <span className="px-2.5 py-0.5 rounded bg-red-50 text-[#EE1D23] text-[10px] font-black uppercase tracking-wider">
                  DETALHES DO CRUZAMENTO
                </span>
                <h3 className="text-xl font-black text-slate-900 uppercase italic tracking-tight mt-1 flex items-center gap-2">
                  <span>Topologia: {selectedRecord.topologia}</span>
                  <span className="text-sm font-normal text-slate-400">({selectedRecord.municipio})</span>
                </h3>
              </div>
              <button
                onClick={() => setSelectedRecord(null)}
                className="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 flex items-center justify-center text-slate-600 transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-6">
              <div className="bg-slate-50 p-3 rounded-xl border border-slate-100">
                <span className="text-[10px] font-black text-slate-400 uppercase">Posição Ranking</span>
                <p className="text-lg font-black text-[#333333]">#{selectedRecord.rank}</p>
              </div>
              <div className="bg-emerald-50 p-3 rounded-xl border border-emerald-100">
                <span className="text-[10px] font-black text-emerald-600 uppercase">Visitas no Outage</span>
                <p className="text-lg font-black text-emerald-900">{selectedRecord.qtdVisitas}</p>
                {selectedRecord.totalVisitasPeriodo > selectedRecord.qtdVisitas && (
                  <span className="text-[10px] text-emerald-700 font-bold">Total no nó: {selectedRecord.totalVisitasPeriodo}</span>
                )}
              </div>
              <div className="bg-red-50 p-3 rounded-xl border border-red-100">
                <span className="text-[10px] font-black text-red-600 uppercase">Qtd. Outages</span>
                <p className="text-lg font-black text-[#EE1D23]">{selectedRecord.qtdOutage}</p>
              </div>
              <div className="bg-slate-50 p-3 rounded-xl border border-slate-100">
                <span className="text-[10px] font-black text-slate-400 uppercase">Clientes Afetados</span>
                <p className="text-lg font-black text-slate-800">{selectedRecord.clientesAfetados.toLocaleString()}</p>
              </div>
            </div>

            <div className="space-y-4">
              <div>
                <h4 className="text-xs font-black text-slate-700 uppercase tracking-wider mb-2">
                  Eventos de Indisponibilidade (Outages na Topologia)
                </h4>
                {selectedRecord.outageEvents.length > 0 ? (
                  <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                    {selectedRecord.outageEvents.map((ev, idx) => (
                      <div key={idx} className="p-3 bg-slate-50 rounded-xl border border-slate-100 text-xs flex justify-between items-center">
                        <div>
                          <p className="font-black text-slate-800">{ev.numeroEvento} - {ev.tipo}</p>
                          <p className="text-[11px] text-slate-500 font-semibold mt-0.5">
                            <span className="font-bold text-slate-700">Início:</span> {formatDisplayDateTime(ev.dataInicioFormatada, ev.dataInicio)} • <span className="text-sky-700 font-bold">{ev.catProd2}</span>
                          </p>
                        </div>
                        <span className={cn(
                          "px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase border",
                          getStatusBadge(ev.status)
                        )}>
                          {ev.status}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="text-xs text-slate-400 font-bold italic">Nenhum chamado individual detalhado listado.</p>
                )}
              </div>

              <div className="pt-2">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-2">
                  <h4 className="text-xs font-black text-slate-700 uppercase tracking-wider">
                    Visitas / Ordens de Serviço (AT5)
                  </h4>
                  <div className="flex items-center gap-1 bg-slate-100 p-0.5 rounded-lg text-[10px] font-bold">
                    <button
                      onClick={() => setModalVisitFilter('day')}
                      className={cn(
                        "px-2.5 py-1 rounded-md transition-all cursor-pointer",
                        modalVisitFilter === 'day' ? "bg-white text-emerald-800 shadow-xs font-black" : "text-slate-500 hover:text-slate-800"
                      )}
                    >
                      No Dia do Outage ({selectedRecord.at5Visits.length})
                    </button>
                    <button
                      onClick={() => setModalVisitFilter('all')}
                      className={cn(
                        "px-2.5 py-1 rounded-md transition-all cursor-pointer",
                        modalVisitFilter === 'all' ? "bg-white text-slate-900 shadow-xs font-black" : "text-slate-500 hover:text-slate-800"
                      )}
                    >
                      Todas no Nó ({selectedRecord.allTopologyVisits?.length || selectedRecord.totalVisitasPeriodo})
                    </button>
                  </div>
                </div>

                {(() => {
                  const visitsToDisplay = modalVisitFilter === 'day'
                    ? selectedRecord.at5Visits
                    : (selectedRecord.allTopologyVisits && selectedRecord.allTopologyVisits.length > 0 ? selectedRecord.allTopologyVisits : selectedRecord.at5Visits);

                  if (visitsToDisplay.length === 0) {
                    return (
                      <p className="text-xs text-slate-400 font-bold italic">
                        {modalVisitFilter === 'day'
                          ? 'Nenhuma visita com ABERTURA_SOLIC registrada na mesma data do início do outage.'
                          : 'Nenhuma visita individual associada à topologia.'}
                      </p>
                    );
                  }

                  return (
                    <div className="space-y-2 max-h-56 overflow-y-auto pr-1">
                      {visitsToDisplay.slice(0, 50).map((v, idx) => (
                        <div key={idx} className="p-3 bg-slate-50 rounded-xl border border-slate-100 text-xs flex justify-between items-center">
                          <div>
                            <p className="font-black text-slate-800">Contrato: {v.contrato} • {v.tipoOs}</p>
                            <p className="text-[11px] text-slate-500 font-semibold mt-0.5">
                              <span className="font-bold text-slate-700">Abertura:</span> {v.aberturaSolic || v.data} • {v.codigoBaixa || 'Baixa Padrão'} • {v.empresa}
                            </p>
                          </div>
                          <span className={cn(
                            "px-2 py-0.5 rounded text-[10px] font-black uppercase whitespace-nowrap",
                            v.qtOsPadrao === 1 ? "bg-emerald-50 text-emerald-700 border border-emerald-200" : "bg-red-50 text-[#EE1D23] border border-red-200"
                          )}>
                            {v.qtOsPadrao === 1 ? 'Com Padrão' : 'Sem Padrão'}
                          </span>
                        </div>
                      ))}
                      {visitsToDisplay.length > 50 && (
                        <p className="text-[11px] text-slate-400 text-center font-bold py-1">
                          Mostrando 50 de {visitsToDisplay.length} visitas
                        </p>
                      )}
                    </div>
                  );
                })()}
              </div>
            </div>

            <div className="mt-6 pt-4 border-t border-slate-100 flex justify-end">
              <button
                onClick={() => setSelectedRecord(null)}
                className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white font-black text-xs uppercase tracking-wider rounded-xl transition-all"
              >
                Fechar Detalhes
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
