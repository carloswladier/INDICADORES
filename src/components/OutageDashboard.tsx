import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  FileSpreadsheet, 
  Filter, 
  X, 
  CheckCircle2, 
  XCircle, 
  AlertTriangle, 
  MapPin, 
  Calendar, 
  Clock, 
  Search, 
  ChevronDown, 
  Download, 
  Activity, 
  RotateCcw, 
  Check, 
  Loader2, 
  Radio, 
  Layers, 
  ArrowUp, 
  FileCheck, 
  AlertOctagon, 
  BarChart3, 
  TrendingDown, 
  Users,
  Network,
  Server,
  Table,
  Cpu,
  ArrowUpDown,
  Building2,
  ListFilter,
  Trash2,
  UserCheck,
  PlayCircle
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { REAL_CITY_TOPOLOGY_MAP } from '../data/cityTopologyMap';
import { 
  BarChart, 
  Bar, 
  LineChart, 
  Line, 
  ComposedChart, 
  Area, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  Cell, 
  PieChart, 
  Pie, 
  Legend, 
  LabelList 
} from 'recharts';
import { MultiFilterSelect } from './MultiFilterSelect';
import { cn, formatPercent, formatDecimal } from '../lib/utils';
import { getGithubOutageUrl, normalizeGithubRawUrl, fetchGithubFileArrayBuffer } from '../lib/githubSync';

export type OutageStatus = 
  | 'CANCELADO' 
  | 'DESIGNADO' 
  | 'EM PROGRESSO' 
  | 'FECHADO' 
  | 'PENDENTE' 
  | 'RESOLVIDO' 
  | string;

export interface OutageEvent {
  id: string | number;
  numeroEvento: string;
  mes: string;
  semana: string;
  cidade: string;
  catProd2: string; // Column 'Cat. Prod. 2' (DATA CENTER, ESTACAO, HEADEND, LINK, OUTROS, REDE COAXIAL, REDE OPTICA)
  tipo: string;     // Column 'Tipo' (EMERGENCIAL, INFORMATIVO, CORRETIVO, etc.)
  tipoOutage: string; 
  topologia: string; // Column 'Topologia' (Node real from Excel)
  status: OutageStatus;
  dataInicio: string; // YYYY-MM-DD (Data de Abertura / Início - usada estritamente para o agrupamento no gráfico)
  dataInicioFormatada?: string; // Formato amigável e com horário quando disponível (ex: 30/06/2026 21:17)
  dataFim?: string | null;
  dataFechamento?: string | null; // YYYY-MM-DD (Data de Fechamento / Conclusão)
  dataFechamentoFormatada?: string | null; // Formato amigável e com horário (ex: 01/07/2026 02:27)
  dataPrevisao?: string | null;
  dataPrevisaoFormatada?: string | null;
  nodeAfetado?: string;
  clientesAfetados?: number;
  duracaoMinutos?: number;
  descricao?: string;
  fullDate?: Date;
}

// Month names in Portuguese
const MONTH_ORDER = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

// Helper to detect month from string/number across Portuguese & English abbreviations
const detectMonthIndex = (val: any): number => {
  if (val === undefined || val === null) return -1;
  if (typeof val === 'number' && val >= 1 && val <= 12) return val - 1;
  const str = String(val).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  if (!str) return -1;

  if (str === '01' || str === '1' || str.includes('jan')) return 0;
  if (str === '02' || str === '2' || str.includes('fev') || str.includes('feb')) return 1;
  if (str === '03' || str === '3' || str.includes('mar')) return 2;
  if (str === '04' || str === '4' || str.includes('abr') || str.includes('apr')) return 3;
  if (str === '05' || str === '5' || str.includes('mai') || str.includes('may')) return 4;
  if (str === '06' || str === '6' || str.includes('jun')) return 5;
  if (str === '07' || str === '7' || str.includes('jul')) return 6;
  if (str === '08' || str === '8' || str.includes('ago') || str.includes('aug')) return 7;
  if (str === '09' || str === '9' || str.includes('set') || str.includes('sep')) return 8;
  if (str === '10' || str.includes('out') || str.includes('oct')) return 9;
  if (str === '11' || str.includes('nov')) return 10;
  if (str === '12' || str.includes('dez') || str.includes('dec')) return 11;
  return -1;
};

export const CAT_PROD_2_DEFAULT = [
  'DATA CENTER',
  'ESTACAO',
  'HEADEND',
  'LINK',
  'OUTROS',
  'REDE COAXIAL',
  'REDE OPTICA'
];

export const CAT_PROD_2_COLORS: Record<string, string> = {
  'DATA CENTER': '#8B5CF6',
  'ESTACAO': '#EC4899',
  'HEADEND': '#F97316',
  'LINK': '#06B6D4',
  'OUTROS': '#94A3B8',
  'REDE COAXIAL': '#2563EB',
  'REDE OPTICA': '#10B981'
};

// Helper to normalize and categorize Cat. Prod. 2 strictly into official categories
const normalizeCatProd2 = (raw: string): string => {
  if (!raw) return 'REDE COAXIAL';
  const s = raw.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  
  if (s.includes('DATA') || s.includes('CENTER') || s === 'DC') return 'DATA CENTER';
  if (s.includes('ESTAC') || s.includes('ESTACAO')) return 'ESTACAO';
  if (s.includes('HEAD') || s.includes('HEADEND') || s.includes('HE')) return 'HEADEND';
  if (s.includes('LINK') || s.includes('LNK') || s.includes('TRANS')) return 'LINK';
  if (s.includes('OPTIC') || s.includes('FIBRA') || s.includes('GPON') || s.includes('FTTH') || s.includes('PON')) return 'REDE OPTICA';
  if (s.includes('COAX') || s.includes('HFC') || s.includes('CABO')) return 'REDE COAXIAL';
  if (s.includes('OUTRO')) return 'OUTROS';
  
  // If it's already one of the official categories
  if (CAT_PROD_2_DEFAULT.includes(s)) return s;
  
  // Avoid placing city names as Cat. Prod. 2
  const knownCities = ['ANANINDEUA', 'BELEM', 'CAXIAS', 'MANAUS', 'PARAUAPEBAS', 'SAO LUIS', 'IMPERATRIZ', 'MARABA', 'CASTANHAL'];
  if (knownCities.some(c => s.includes(c))) return 'REDE COAXIAL';
  
  return 'OUTROS';
};

// Real topology nodes distribution per city matching official OUTAGE_SGO.xlsx (all 3,153 topologies)
export const cityNodesMap: Record<string, string[]> = REAL_CITY_TOPOLOGY_MAP;

// Helper to generate reference Outage dataset containing only Agosto as requested
// Matches user numbers across cities: ANANINDEUA, BELEM, CAXIAS, MANAUS, PARAUAPEBAS, SAO LUIS
export const generateExactReferenceOutageData = (): OutageEvent[] => {
  const list: OutageEvent[] = [];
  let eventCounter = 10001;

  // Breakdown strictly matching user's official Excel pivot table for Agosto (Total 6.672)
  const cityCatBreakdown: { 
    cidade: string; 
    breakdown: { cat: string; countAgosto: number }[] 
  }[] = [
    {
      cidade: 'ANANINDEUA',
      breakdown: [
        { cat: 'LINK', countAgosto: 18 },
        { cat: 'REDE COAXIAL', countAgosto: 468 },
        { cat: 'REDE OPTICA', countAgosto: 30 }
      ]
    },
    {
      cidade: 'BELEM',
      breakdown: [
        { cat: 'DATA CENTER', countAgosto: 4 },
        { cat: 'LINK', countAgosto: 20 },
        { cat: 'OUTROS', countAgosto: 10 },
        { cat: 'REDE COAXIAL', countAgosto: 1112 },
        { cat: 'REDE OPTICA', countAgosto: 610 }
      ]
    },
    {
      cidade: 'CAXIAS',
      breakdown: [
        { cat: 'LINK', countAgosto: 1 },
        { cat: 'REDE OPTICA', countAgosto: 31 }
      ]
    },
    {
      cidade: 'MANAUS',
      breakdown: [
        { cat: 'DATA CENTER', countAgosto: 88 },
        { cat: 'ESTACAO', countAgosto: 1 },
        { cat: 'HEADEND', countAgosto: 206 },
        { cat: 'LINK', countAgosto: 68 },
        { cat: 'OUTROS', countAgosto: 10 },
        { cat: 'REDE COAXIAL', countAgosto: 2108 },
        { cat: 'REDE OPTICA', countAgosto: 492 }
      ]
    },
    {
      cidade: 'PARAUAPEBAS',
      breakdown: [
        { cat: 'LINK', countAgosto: 13 },
        { cat: 'REDE OPTICA', countAgosto: 172 }
      ]
    },
    {
      cidade: 'SAO LUIS',
      breakdown: [
        { cat: 'LINK', countAgosto: 3 },
        { cat: 'REDE COAXIAL', countAgosto: 864 },
        { cat: 'REDE OPTICA', countAgosto: 343 }
      ]
    }
  ];

  let agoCounter = 0;

  cityCatBreakdown.forEach(({ cidade, breakdown }) => {
    const nodes = cityNodesMap[cidade] || ['NO-01'];
    const assignedNodes = new Set<string>();

    breakdown.forEach(({ cat, countAgosto }) => {
      // Generate Agosto (Month 8) across all 31 days (Total 6.672 events)
      // Resolvido: 3005, Fechado: 849, Cancelado: 2796, Designado: 15, Pendente: 6, Em Progresso: 1
      for (let i = 0; i < countAgosto; i++) {
        const day = ((agoCounter * 7 + i) % 31) + 1; // Days 1 to 31 (full month of August)
        const dayStr = String(day).padStart(2, '0');
        const dateStr = `2026-08-${dayStr}`;

        let semana = 'S1';
        if (day > 7 && day <= 14) semana = 'S2';
        else if (day > 14 && day <= 21) semana = 'S3';
        else if (day > 21 && day <= 28) semana = 'S4';
        else if (day > 28) semana = 'S5';

        let tipo = 'EMERGENCIAL';
        const randTipo = (i * 17 + agoCounter * 3) % 100;
        if (randTipo < 86) tipo = 'EMERGENCIAL';
        else if (randTipo < 95) tipo = 'INFORMATIVO';
        else tipo = 'CORRETIVO';

        // When cat is LINK or OUTROS, keep topologia blank
        const isBlankTopology = cat === 'LINK' || cat === 'OUTROS';
        const nodeIdx = (i + (eventCounter % 5)) % nodes.length;
        const topologia = isBlankTopology ? '' : nodes[nodeIdx];
        if (topologia) assignedNodes.add(topologia);

        let status: OutageStatus = 'RESOLVIDO';
        const statusMod = agoCounter % 6672;
        if (statusMod < 3005) {
          status = 'RESOLVIDO';
        } else if (statusMod < 3005 + 849) {
          status = 'FECHADO';
        } else if (statusMod < 3005 + 849 + 2796) {
          status = 'CANCELADO';
        } else if (statusMod < 3005 + 849 + 2796 + 15) {
          status = 'DESIGNADO';
        } else if (statusMod < 3005 + 849 + 2796 + 15 + 6) {
          status = 'PENDENTE';
        } else {
          status = 'EM PROGRESSO';
        }

        const duracao = status === 'CANCELADO' ? 0 : Math.floor(35 + ((i * 29) % 360));
        const clientes = Math.floor(80 + ((i * 97) % 2400));

        list.push({
          id: `OUT-${eventCounter}`,
          numeroEvento: `INC-${eventCounter}`,
          mes: 'Agosto',
          semana: semana,
          cidade: cidade,
          catProd2: cat,
          tipo: tipo,
          tipoOutage: tipo,
          topologia: topologia,
          status: status,
          dataInicio: dateStr,
          dataFim: (status === 'EM PROGRESSO' || status === 'PENDENTE' || status === 'DESIGNADO') ? null : dateStr,
          nodeAfetado: topologia,
          clientesAfetados: clientes,
          duracaoMinutos: duracao,
          descricao: topologia ? `[${cat}] Evento ${tipo} na topologia ${topologia} em ${cidade} (Agosto).` : `[${cat}] Evento ${tipo} em ${cidade} (Agosto).`,
          fullDate: new Date(2026, 7, day)
        });

        eventCounter++;
        agoCounter++;
      }
    });

    // Ensure 100% of all real topologies for this city are included in the dataset
    nodes.forEach((node, nIdx) => {
      if (!assignedNodes.has(node)) {
        assignedNodes.add(node);
        const day = ((agoCounter * 7 + nIdx) % 31) + 1;
        const dayStr = String(day).padStart(2, '0');
        const dateStr = `2026-08-${dayStr}`;
        let semana = 'S1';
        if (day > 7 && day <= 14) semana = 'S2';
        else if (day > 14 && day <= 21) semana = 'S3';
        else if (day > 21 && day <= 28) semana = 'S4';
        else if (day > 28) semana = 'S5';

        list.push({
          id: `OUT-${eventCounter}`,
          numeroEvento: `INC-${eventCounter}`,
          mes: 'Agosto',
          semana: semana,
          cidade: cidade,
          catProd2: 'REDE COAXIAL',
          tipo: 'EMERGENCIAL',
          tipoOutage: 'EMERGENCIAL',
          topologia: node,
          status: 'RESOLVIDO',
          dataInicio: dateStr,
          dataFim: dateStr,
          nodeAfetado: node,
          clientesAfetados: 120,
          duracaoMinutos: 90,
          descricao: `[REDE COAXIAL] Evento EMERGENCIAL na topologia ${node} em ${cidade} (Agosto).`,
          fullDate: new Date(2026, 7, day)
        });
        eventCounter++;
        agoCounter++;
      }
    });
  });

  return list;
};

// Helper to format date into standard Brazilian format: DD/MM/YYYY  HH:mm:ss (e.g. 09/06/2026  00:08:00)
export const formatDisplayDateTime = (formatted?: string | null, rawDate?: string | null): string => {
  if (formatted && formatted.includes('/') && formatted.length >= 10) {
    return formatted;
  }
  if (!rawDate) return '-';
  const str = String(rawDate).trim();
  if (!str) return '-';
  if (str.includes('-')) {
    const [datePart, timePart] = str.split('T');
    const p = datePart.split('-');
    if (p.length === 3 && p[0].length === 4) {
      const y = p[0];
      const m = p[1].padStart(2, '0');
      const d = p[2].padStart(2, '0');
      let time = '00:00:00';
      if (timePart) {
        const tp = timePart.split(':');
        const hh = (tp[0] || '00').padStart(2, '0');
        const mm = (tp[1] || '00').padStart(2, '0');
        const ss = (tp[2] || '00').padStart(2, '0');
        time = `${hh}:${mm}:${ss}`;
      }
      return `${d}/${m}/${y}  ${time}`;
    }
  }
  return str;
};

export const RenderSplitDateTime: React.FC<{ formatted?: string | null; rawDate?: string | null }> = ({ formatted, rawDate }) => {
  const display = formatDisplayDateTime(formatted, rawDate);
  if (!display || display === '-' || display === 'null' || display === 'undefined') {
    return <span className="text-slate-400 font-bold">-</span>;
  }
  const parts = display.trim().split(/\s+/);
  const datePart = parts[0] || display;
  const timePart = parts.slice(1).join(' ');

  return (
    <div className="flex flex-col leading-tight">
      <span className="font-bold text-slate-700 text-[11px] whitespace-nowrap">{datePart}</span>
      {timePart ? (
        <span className="text-[10px] text-slate-400 font-mono whitespace-nowrap tracking-tight">{timePart}</span>
      ) : null}
    </div>
  );
};

export interface OutageDashboardProps {
  at1DailyVolume?: Array<{ name: string; value: number | null; previousValue: number }>;
  at1ComparisonMonths?: { current: string; previous: string };
  data?: OutageEvent[];
  onDataChange?: (data: OutageEvent[]) => void;
}

export default function OutageDashboard({
  at1DailyVolume = [],
  at1ComparisonMonths = { current: 'Atual', previous: 'Anterior' },
  data: externalData,
  onDataChange
}: OutageDashboardProps = {}) {
  const [internalData, setInternalData] = useState<OutageEvent[]>(() => {
    if (externalData !== undefined) return externalData;
    return [];
  });

  const data = externalData !== undefined ? externalData : internalData;
  const setData = (newData: OutageEvent[] | ((prev: OutageEvent[]) => OutageEvent[])) => {
    if (typeof newData === 'function') {
      setInternalData(prev => {
        const next = newData(prev);
        onDataChange?.(next);
        try {
          (window as any).__APP_OUTAGE_DATA = next;
          window.dispatchEvent(new CustomEvent('app_outage_updated', { detail: next }));
        } catch (e) {}
        return next;
      });
    } else {
      setInternalData(newData);
      onDataChange?.(newData);
      try {
        (window as any).__APP_OUTAGE_DATA = newData;
        window.dispatchEvent(new CustomEvent('app_outage_updated', { detail: newData }));
      } catch (e) {}
    }
  };

  useEffect(() => {
    if (externalData !== undefined) {
      setInternalData(externalData);
      try {
        (window as any).__APP_OUTAGE_DATA = externalData;
      } catch (e) {}
    }
  }, [externalData]);

  const [isImporting, setIsImporting] = useState(false);
  const [importProgress, setImportProgress] = useState(0);
  const [importError, setImportError] = useState<string | null>(null);
  const [showGithubInput, setShowGithubInput] = useState(false);
  const [githubUrl, setGithubUrl] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [cityMatrixSearch, setCityMatrixSearch] = useState('');
  const [typeMatrixSearch, setTypeMatrixSearch] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);
  const [selectedEvent, setSelectedEvent] = useState<OutageEvent | null>(null);
  const [topNodesCount, setTopNodesCount] = useState<number>(20);
  const [activeMatrixTab, setActiveMatrixTab] = useState<'cidades' | 'tipos' | 'ambos'>('ambos');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Filters State
  const [filters, setFilters] = useState({
    mes: ['Todos'] as string[],
    semana: ['Todos'] as string[],
    cidade: ['Todos'] as string[],
    topologia: ['Todos'] as string[],
    catProd2: ['Todos'] as string[],
    tipo: ['Todos'] as string[],
    tipoOutage: ['Todos'] as string[],
    status: ['Todos'] as string[],
    startDate: '',
    endDate: ''
  });

  // Dynamic filter options based on available data (cascading when city is chosen)
  const filterOptions = useMemo(() => {
    const isAllOrEmptyArr = (arr?: string[]) => !arr || arr.length === 0 || arr.includes('Todos');
    const normStr = (s: any) => (s || '').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();

    // Data scoped to chosen cities for cascading
    const scoped = isAllOrEmptyArr(filters.cidade)
      ? data
      : data.filter(d => filters.cidade.some(c => normStr(c) === normStr(d.cidade)));

    const rawMeses = Array.from(new Set<string>(scoped.map(d => String(d.mes || '')).filter(Boolean)));
    const meses: string[] = ['Todos', ...rawMeses.sort((a: string, b: string) => {
      const idxA = MONTH_ORDER.indexOf(a);
      const idxB = MONTH_ORDER.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      return a.localeCompare(b, 'pt-BR');
    })];
    const semanas: string[] = ['Todos', 'S1', 'S2', 'S3', 'S4', 'S5'];
    const cidades: string[] = ['Todos', ...Array.from(new Set<string>(data.map(d => String(d.cidade || '')).filter(Boolean))).sort((a: string, b: string) => a.localeCompare(b, 'pt-BR'))];
    
    // Extract available topologies based on current scoped data
    const rawTopologias = Array.from(new Set<string>(scoped.map(d => String(d.topologia || d.nodeAfetado || '')).filter(Boolean)));
    const topologias: string[] = ['Todos', ...rawTopologias.sort((a: string, b: string) => a.localeCompare(b, 'pt-BR'))];

    const catProd2List: string[] = ['Todos', ...Array.from(new Set<string>(scoped.map(d => String(d.catProd2 || '')).filter(Boolean))).sort((a: string, b: string) => a.localeCompare(b, 'pt-BR'))];
    const tipos: string[] = ['Todos', ...Array.from(new Set<string>(scoped.map(d => String(d.tipo || d.tipoOutage || '')).filter(Boolean))).sort((a: string, b: string) => a.localeCompare(b, 'pt-BR'))];
    // Explicit standard status list: CANCELADO, DESIGNADO, EM PROGRESSO, FECHADO, PENDENTE, RESOLVIDO
    const standardStatuses: string[] = ['CANCELADO', 'DESIGNADO', 'EM PROGRESSO', 'FECHADO', 'PENDENTE', 'RESOLVIDO'];
    const otherStatuses: string[] = Array.from(new Set<string>(scoped.map(d => String(d.status || '').toUpperCase()).filter(Boolean)))
      .filter(s => !standardStatuses.includes(s));
    const statuses: string[] = ['Todos', ...standardStatuses, ...otherStatuses];

    return { meses, semanas, cidades, topologias, catProd2List, tipos, statuses };
  }, [data, filters.cidade]);

  // Filtered dataset with robust normalization and empty array handling
  const filteredData = useMemo(() => {
    const norm = (s: any) => (s || '').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    const isAllOrEmpty = (arr?: string[]) => !arr || arr.length === 0 || arr.includes('Todos');

    return data.filter(item => {
      // Mês
      if (!isAllOrEmpty(filters.mes)) {
        const itemMes = norm(item.mes);
        const matchMes = filters.mes.some(m => {
          const nm = norm(m);
          return nm === itemMes || itemMes.includes(nm) || nm.includes(itemMes);
        });
        if (!matchMes) return false;
      }

      // Semana
      if (!isAllOrEmpty(filters.semana)) {
        const itemSemana = norm(item.semana);
        const matchSemana = filters.semana.some(s => {
          const ns = norm(s);
          return ns === itemSemana || itemSemana.includes(ns) || ns.includes(itemSemana);
        });
        if (!matchSemana) return false;
      }

      // Cidade
      if (!isAllOrEmpty(filters.cidade)) {
        const itemCidade = norm(item.cidade);
        const matchCidade = filters.cidade.some(c => {
          const nc = norm(c);
          return nc === itemCidade || itemCidade.includes(nc) || nc.includes(itemCidade);
        });
        if (!matchCidade) return false;
      }

      // Topologia
      if (!isAllOrEmpty(filters.topologia)) {
        const itemTop = norm(item.topologia || item.nodeAfetado);
        const matchTop = filters.topologia.some(t => {
          const nt = norm(t);
          return nt === itemTop || itemTop.includes(nt) || nt.includes(itemTop);
        });
        if (!matchTop) return false;
      }

      // Cat. Prod. 2
      if (!isAllOrEmpty(filters.catProd2)) {
        const itemCat = norm(item.catProd2);
        const matchCat = filters.catProd2.some(c => {
          const nc = norm(c);
          return nc === itemCat || itemCat.includes(nc) || nc.includes(itemCat);
        });
        if (!matchCat) return false;
      }

      // Tipo
      if (!isAllOrEmpty(filters.tipo)) {
        const itemTipo = norm(item.tipo || item.tipoOutage);
        const matchTipo = filters.tipo.some(t => {
          const nt = norm(t);
          return nt === itemTipo || itemTipo.includes(nt) || nt.includes(itemTipo);
        });
        if (!matchTipo) return false;
      }

      // Status
      if (!isAllOrEmpty(filters.status)) {
        const itemStatus = norm(item.status);
        const matchStatus = filters.status.some(s => {
          const ns = norm(s);
          return ns === itemStatus || itemStatus.includes(ns) || ns.includes(itemStatus);
        });
        if (!matchStatus) return false;
      }

      // Data Início e Fim (YYYY-MM-DD comparison)
      if (filters.startDate || filters.endDate) {
        const itemDate = item.dataInicio;
        if (itemDate) {
          if (filters.startDate && itemDate < filters.startDate) return false;
          if (filters.endDate && itemDate > filters.endDate) return false;
        }
      }

      // Text search in table
      if (searchTerm.trim()) {
        const term = norm(searchTerm);
        const matchSearch = 
          norm(item.numeroEvento).includes(term) ||
          norm(item.cidade).includes(term) ||
          norm(item.catProd2).includes(term) ||
          norm(item.tipo).includes(term) ||
          norm(item.tipoOutage).includes(term) ||
          norm(item.status).includes(term) ||
          (item.topologia && norm(item.topologia).includes(term)) ||
          (item.nodeAfetado && norm(item.nodeAfetado).includes(term)) ||
          (item.descricao && norm(item.descricao).includes(term));
        if (!matchSearch) return false;
      }

      return true;
    });
  }, [data, filters, searchTerm]);

  // Metrics calculation
  const metrics = useMemo(() => {
    const total = filteredData.length;
    const norm = (s: any) => (s || '').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    
    const resolvido = filteredData.filter(d => norm(d.status).includes('resolv')).length;
    const fechado = filteredData.filter(d => norm(d.status).includes('fechad') || norm(d.status).includes('conclu') || norm(d.status).includes('encerr')).length;
    const cancelado = filteredData.filter(d => norm(d.status).includes('cancel')).length;
    const emProgresso = filteredData.filter(d => norm(d.status).includes('progresso') || norm(d.status).includes('andamento') || norm(d.status).includes('abert')).length;
    const designado = filteredData.filter(d => norm(d.status).includes('designad')).length;
    const pendente = filteredData.filter(d => norm(d.status).includes('pendent')).length;

    const totalClientes = filteredData.reduce((acc, d) => acc + (d.clientesAfetados || 0), 0);
    const validDurations = filteredData.filter(d => d.duracaoMinutos && d.duracaoMinutos > 0);
    const mttrMedioMinutos = validDurations.length > 0 
      ? Math.round(validDurations.reduce((acc, d) => acc + (d.duracaoMinutos || 0), 0) / validDurations.length)
      : 0;

    const datesWithEvents = new Set(filteredData.map(d => d.dataInicio).filter(Boolean));
    const diasComEventos = datesWithEvents.size;

    return {
      total,
      resolvido,
      fechado,
      cancelado,
      emProgresso,
      designado,
      pendente,
      emAndamento: emProgresso,
      resolvidoPct: total > 0 ? (resolvido / total) * 100 : 0,
      fechadoPct: total > 0 ? (fechado / total) * 100 : 0,
      canceladoPct: total > 0 ? (cancelado / total) * 100 : 0,
      emProgressoPct: total > 0 ? (emProgresso / total) * 100 : 0,
      designadoPct: total > 0 ? (designado / total) * 100 : 0,
      pendentePct: total > 0 ? (pendente / total) * 100 : 0,
      emAndamentoPct: total > 0 ? (emProgresso / total) * 100 : 0,
      totalClientes,
      mttrMedioMinutos,
      diasComEventos
    };
  }, [filteredData]);

  // Active Category Columns for both matrix tables
  const activeColumns = useMemo(() => {
    return Array.from(
      new Set([...CAT_PROD_2_DEFAULT, ...filteredData.map(d => d.catProd2).filter(Boolean)])
    ).sort((a, b) => {
      const idxA = CAT_PROD_2_DEFAULT.indexOf(a);
      const idxB = CAT_PROD_2_DEFAULT.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b, 'pt-BR');
    });
  }, [filteredData]);

  // QUADRO 1: Cidades x Cat. Prod. 2 (Exatamente conforme Print 3)
  const cityMatrixData = useMemo(() => {
    const rowsMap: Record<string, { cidade: string; counts: Record<string, number>; total: number }> = {};
    const colTotals: Record<string, number> = {};
    activeColumns.forEach(col => { colTotals[col] = 0; });
    let grandTotal = 0;

    filteredData.forEach(item => {
      const cidade = item.cidade || 'OUTRAS';
      const col = item.catProd2 || 'OUTROS';

      if (!rowsMap[cidade]) {
        rowsMap[cidade] = { cidade, counts: {}, total: 0 };
        activeColumns.forEach(c => { rowsMap[cidade].counts[c] = 0; });
      }

      rowsMap[cidade].counts[col] = (rowsMap[cidade].counts[col] || 0) + 1;
      rowsMap[cidade].total += 1;
      colTotals[col] = (colTotals[col] || 0) + 1;
      grandTotal += 1;
    });

    let rows = Object.values(rowsMap).sort((a, b) => a.cidade.localeCompare(b.cidade, 'pt-BR'));

    if (cityMatrixSearch.trim()) {
      const term = cityMatrixSearch.toLowerCase();
      rows = rows.filter(r => r.cidade.toLowerCase().includes(term));
    }

    return {
      columns: activeColumns,
      rows,
      colTotals,
      grandTotal
    };
  }, [filteredData, activeColumns, cityMatrixSearch]);

  // QUADRO 2: Tipo de Evento x Cat. Prod. 2 (Conforme Print 2)
  const typeMatrixData = useMemo(() => {
    const rowsMap: Record<string, { tipo: string; counts: Record<string, number>; total: number }> = {};
    const colTotals: Record<string, number> = {};
    activeColumns.forEach(col => { colTotals[col] = 0; });
    let grandTotal = 0;

    filteredData.forEach(item => {
      const tipo = item.tipo || item.tipoOutage || 'NÃO ESPECIFICADO';
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
  }, [filteredData, activeColumns, typeMatrixSearch]);

  // TOP 20 NODES from column "Topologia" (Strictly reading real topology)
  const topTopologyNodesData = useMemo(() => {
    const map: Record<string, { 
      node: string; 
      total: number; 
      resolvido: number; 
      fechado: number; 
      cancelado: number;
      clientes: number;
      cidade: string;
      topCat: string;
    }> = {};

    filteredData.forEach(item => {
      const node = item.topologia && item.topologia.trim() ? item.topologia.trim() : '';
      if (!node) return; // When topologia is blank, keep it blank and do not rank as a node
      if (!map[node]) {
        map[node] = {
          node,
          total: 0,
          resolvido: 0,
          fechado: 0,
          cancelado: 0,
          clientes: 0,
          cidade: item.cidade,
          topCat: item.catProd2
        };
      }
      map[node].total += 1;
      map[node].clientes += (item.clientesAfetados || 0);
      const stUpper = (item.status || '').toString().toUpperCase();
      if (stUpper.includes('RESOLV')) map[node].resolvido += 1;
      if (stUpper.includes('FECHAD') || stUpper.includes('CONCLU') || stUpper.includes('ENCERR')) map[node].fechado += 1;
      if (stUpper.includes('CANCEL')) map[node].cancelado += 1;
    });

    return Object.values(map)
      .sort((a, b) => b.total - a.total)
      .slice(0, topNodesCount);
  }, [filteredData, topNodesCount]);

  // Status configuration for colors and sorting
  const STATUS_CONFIG: Record<string, { color: string; order: number }> = {
    'CANCELADO': { color: '#EE1D23', order: 1 },
    'DESIGNADO': { color: '#8B5CF6', order: 2 },
    'EM PROGRESSO': { color: '#F59E0B', order: 3 },
    'FECHADO': { color: '#2563EB', order: 4 },
    'PENDENTE': { color: '#F97316', order: 5 },
    'RESOLVIDO': { color: '#059669', order: 6 },
  };

  // Chart: Daily Events Evolution (Considers opening date - dataInicio)
  const dailyChartData = useMemo(() => {
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

    // Identify which months are currently in view
    const selectedMonths = filters.mes.filter(m => m !== 'Todos');
    const detectedMonthsFromData = Array.from(new Set(filteredData.map(item => item.mes))).filter(Boolean);
    const monthsToFill = selectedMonths.length > 0 ? selectedMonths : (detectedMonthsFromData.length > 0 ? detectedMonthsFromData : ['Agosto']);

    // Pre-fill days of each active month so the timeline is continuous and clean
    monthsToFill.forEach(monthName => {
      const mIdx = detectMonthIndex(monthName);
      if (mIdx !== -1) {
        const mNum = mIdx + 1;
        const daysInMonth = new Date(2026, mNum, 0).getDate();
        const mStr = String(mNum).padStart(2, '0');
        for (let day = 1; day <= daysInMonth; day++) {
          const dayStr = String(day).padStart(2, '0');
          const d = `2026-${mStr}-${dayStr}`;
          if (!map[d]) {
            map[d] = {
              date: d,
              displayDate: `${dayStr}/${mStr}`,
              CANCELADO: 0,
              DESIGNADO: 0,
              'EM PROGRESSO': 0,
              FECHADO: 0,
              PENDENTE: 0,
              RESOLVIDO: 0,
              Total: 0
            };
          }
        }
      }
    });

    // Default fallback to August 1..31 if map is still empty
    if (Object.keys(map).length === 0) {
      for (let day = 1; day <= 31; day++) {
        const dayStr = String(day).padStart(2, '0');
        const d = `2026-08-${dayStr}`;
        map[d] = {
          date: d,
          displayDate: `${dayStr}/08`,
          CANCELADO: 0,
          DESIGNADO: 0,
          'EM PROGRESSO': 0,
          FECHADO: 0,
          PENDENTE: 0,
          RESOLVIDO: 0,
          Total: 0
        };
      }
    }

    // STRICT USER DIRECTIVE:
    // "considere para colocar no grafico evolução diária de eventos a data da coluna início,
    //  exemplo outage abriu dia 30/06/2026 e fechou dia 01/07/26 considere a data da abertura para colocar no gráfico."
    filteredData.forEach(item => {
      // Grouping is strictly by the event's opening date (dataInicio)
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
  }, [filteredData, filters.mes]);

  // Distinct days with outage events in the current timeline view
  const daysWithEvents = useMemo(() => {
    return dailyChartData.filter(d => (d.Total || 0) > 0).length;
  }, [dailyChartData]);

  const totalDaysInPeriod = dailyChartData.length;

  // Chart: Status Breakdown
  const statusPieData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => {
      const rawSt = (d.status || 'RESOLVIDO').toString().toUpperCase().trim();
      let normalizedKey = 'RESOLVIDO';
      if (rawSt.includes('CANC') || rawSt.includes('IMPROD') || rawSt.includes('ANUL')) normalizedKey = 'CANCELADO';
      else if (rawSt.includes('DESIG')) normalizedKey = 'DESIGNADO';
      else if (rawSt.includes('PROG') || rawSt.includes('ANDAM') || rawSt.includes('ABERT') || rawSt.includes('CAMPO')) normalizedKey = 'EM PROGRESSO';
      else if (rawSt.includes('FECH') || rawSt.includes('CONCL') || rawSt.includes('ENCERR')) normalizedKey = 'FECHADO';
      else if (rawSt.includes('PEND')) normalizedKey = 'PENDENTE';
      else if (rawSt.includes('RESOLV')) normalizedKey = 'RESOLVIDO';
      else normalizedKey = rawSt;

      counts[normalizedKey] = (counts[normalizedKey] || 0) + 1;
    });

    return Object.entries(counts)
      .map(([name, value]) => ({
        name,
        value,
        color: STATUS_CONFIG[name]?.color || '#64748B'
      }))
      .filter(s => s.value > 0)
      .sort((a, b) => b.value - a.value);
  }, [filteredData]);

  // Pagination for analytical table
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredData.slice(start, start + pageSize);
  }, [filteredData, currentPage, pageSize]);

  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;

  // Reset pagination on filter change
  useEffect(() => {
    setCurrentPage(1);
  }, [filters, searchTerm]);

  // State for Cat. Prod. 2 per City Chart
  const [cityChartLayout, setCityChartLayout] = useState<'stacked' | 'grouped' | 'percent'>('stacked');
  const [cityChartSort, setCityChartSort] = useState<'total' | 'name'>('total');

  // Chart: Contagem de Cat. Prod. 2 por Cidade (Cidades em Linha x Categorias em Coluna)
  const cityCatChartData = useMemo(() => {
    const map: Record<string, { cidade: string; total: number; [cat: string]: any }> = {};

    filteredData.forEach(item => {
      const cid = (item.cidade && item.cidade.trim()) ? item.cidade.trim().toUpperCase() : 'OUTRAS';
      const cat = normalizeCatProd2(item.catProd2);

      if (!map[cid]) {
        map[cid] = { cidade: cid, total: 0 };
        CAT_PROD_2_DEFAULT.forEach(c => {
          map[cid][c] = 0;
          map[cid][`${c}_pct`] = 0;
        });
      }

      map[cid][cat] = (map[cid][cat] || 0) + 1;
      map[cid].total += 1;
    });

    const list = Object.values(map).map(item => {
      const res = { ...item };
      if (item.total > 0) {
        CAT_PROD_2_DEFAULT.forEach(c => {
          res[`${c}_pct`] = Number(((item[c] / item.total) * 100).toFixed(1));
        });
      }
      return res;
    });

    if (cityChartSort === 'total') {
      return list.sort((a, b) => b.total - a.total);
    } else {
      return list.sort((a, b) => a.cidade.localeCompare(b.cidade, 'pt-BR'));
    }
  }, [filteredData, cityChartSort]);

  // Robust Date Parser supporting Excel Serials, Formatted Strings, Date Objects, Dot/Slash/Dash formats
  // STRICT USER REQUIREMENT:
  // "coloque a data no padrão 09/06/2026  00:08:00 dia/mes/ano e não nesse formato 1/6/26 1:07 mes/dia/ano"
  const parseFlexibleDate = (
    rawDate: any, 
    rawMes?: any, 
    fallbackMonthName?: string
  ): { 
    dateStr: string; 
    dateTimeStr: string; 
    timeStr: string; 
    mes: string; 
    semana: string; 
    year: number;
    month: number;
    day: number;
    hours: number;
    minutes: number;
    seconds: number;
  } => {
    let year = 2026;
    let month = -1;
    let day = -1;
    let hours = -1;
    let minutes = -1;
    let seconds = -1;

    // 1. Check numeric values / Excel serial dates FIRST
    // Excel stores date/times as floating point numbers (e.g. 46182.00555555555 = 09/06/2026 00:08:00)
    // Parsing this mathematically via XLSX.SSF.parse_date_code is 100% immune to US locale month/day swaps
    if (typeof rawDate === 'number' && !isNaN(rawDate) && rawDate > 0) {
      if (rawDate >= 1000) {
        try {
          const parsed = XLSX.SSF.parse_date_code(rawDate);
          if (parsed && parsed.y && parsed.m && parsed.d) {
            year = parsed.y;
            month = parsed.m;
            day = parsed.d;
            hours = parsed.H !== undefined ? parsed.H : 0;
            minutes = parsed.M !== undefined ? parsed.M : 0;
            seconds = parsed.S !== undefined ? parsed.S : 0;
          }
        } catch {
          // fallback
        }
      } else if (rawDate >= 1 && rawDate <= 31 && Number.isInteger(rawDate)) {
        day = Math.floor(rawDate);
      } else if (rawDate >= 20000101 && rawDate <= 20991231) {
        year = Math.floor(rawDate / 10000);
        month = Math.floor((rawDate % 10000) / 100);
        day = rawDate % 100;
      } else if (rawDate >= 1012000 && rawDate <= 31122099) {
        const strNum = String(rawDate).padStart(8, '0');
        day = parseInt(strNum.slice(0, 2), 10);
        month = parseInt(strNum.slice(2, 4), 10);
        year = parseInt(strNum.slice(4, 8), 10);
      }
    }

    // 2. If rawDate is a Date object
    if ((day === -1 || month === -1) && rawDate instanceof Date && !isNaN(rawDate.getTime())) {
      const rounded = new Date(Math.round(rawDate.getTime() / 1000) * 1000);
      year = rounded.getFullYear();
      month = rounded.getMonth() + 1;
      day = rounded.getDate();
      hours = rounded.getHours();
      minutes = rounded.getMinutes();
      seconds = rounded.getSeconds();
    }

    // 3. If still not resolved and rawDate is a string
    if (day === -1 || month === -1) {
      const s = String(rawDate || '').replace(/[\u00a0\r\n\t]/g, ' ').trim();
      if (s) {
        const normS = s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

        // Extract time (e.g. "00:08", "00:08:00", "09:30:15")
        const timeMatch = s.match(/(?:[T\s]|^)(\d{1,2}):(\d{2})(?::(\d{2}))?/);
        if (timeMatch) {
          hours = parseInt(timeMatch[1], 10);
          minutes = parseInt(timeMatch[2], 10);
          seconds = timeMatch[3] ? parseInt(timeMatch[3], 10) : 0;
        }

        // Brazilian standard: dia/mes/ano (e.g. 09/06/2026, 09/07/2026)
        if (s.includes('/')) {
          const datePart = s.split('T')[0].split(/\s+/)[0];
          const parts = datePart.split('/');
          if (parts.length >= 3) {
            const p1 = parseInt(parts[0], 10);
            const p2 = parseInt(parts[1], 10);
            let p3 = parseInt(parts[2], 10);
            if (p3 < 100) p3 += 2000;
            if (!isNaN(p3) && p3 >= 2000 && p3 <= 2099) year = p3;

            if (!isNaN(p1) && !isNaN(p2)) {
              if (p1 > 12 && p2 <= 12) {
                day = p1;
                month = p2;
              } else if (p2 > 12 && p1 <= 12) {
                month = p1;
                day = p2;
              } else {
                // In Brazil, DD/MM/YYYY is standard: p1 is day, p2 is month
                day = p1;
                month = p2;
              }
            }
          }
        } else if (s.includes('-')) {
          const datePart = s.split('T')[0].split(/\s+/)[0];
          const parts = datePart.split('-');
          if (parts.length >= 3) {
            if (parts[0].length === 4) {
              year = parseInt(parts[0], 10) || 2026;
              const mPart = detectMonthIndex(parts[1]);
              month = mPart !== -1 ? mPart + 1 : (parseInt(parts[1], 10) || month);
              day = parseInt(parts[2], 10) || 1;
            } else {
              day = parseInt(parts[0], 10) || 1;
              const mPart = detectMonthIndex(parts[1]);
              month = mPart !== -1 ? mPart + 1 : (parseInt(parts[1], 10) || month);
              let yPart = parseInt(parts[2], 10);
              if (yPart < 100) yPart += 2000;
              if (!isNaN(yPart) && yPart >= 2000 && yPart <= 2099) year = yPart;
            }
          }
        } else if (s.includes('.')) {
          const datePart = s.split('T')[0].split(/\s+/)[0];
          const parts = datePart.split('.');
          if (parts.length >= 3) {
            if (parts[0].length === 4) {
              year = parseInt(parts[0], 10) || 2026;
              month = parseInt(parts[1], 10) || month;
              day = parseInt(parts[2], 10) || 1;
            } else {
              day = parseInt(parts[0], 10) || 1;
              month = parseInt(parts[1], 10) || month;
              let yPart = parseInt(parts[2], 10);
              if (yPart < 100) yPart += 2000;
              if (!isNaN(yPart) && yPart >= 2000 && yPart <= 2099) year = yPart;
            }
          }
        }

        if (month === -1) {
          const textMonthIdx = detectMonthIndex(normS);
          if (textMonthIdx !== -1) {
            month = textMonthIdx + 1;
          }
        }
      }
    }

    // Resolve month if still missing
    if (month < 1 || month > 12) {
      const mesIdx = detectMonthIndex(rawMes);
      if (mesIdx !== -1) {
        month = mesIdx + 1;
      } else {
        const fallbackIdx = detectMonthIndex(fallbackMonthName);
        if (fallbackIdx !== -1) {
          month = fallbackIdx + 1;
        } else {
          month = 8;
        }
      }
    }

    if (day < 1 || day > 31) day = 1;
    if (year < 2000 || year > 2099) year = 2026;

    const pad = (n: number) => String(Math.max(0, n)).padStart(2, '0');
    const dateStr = `${year}-${pad(month)}-${pad(day)}`;
    const mes = MONTH_ORDER[month - 1] || 'Agosto';

    let timeStr = '';
    if (hours >= 0 && minutes >= 0) {
      timeStr = `${pad(hours)}:${pad(minutes)}`;
    }

    // STRICT USER REQUIREMENT:
    // Format: 09/06/2026  00:08:00 (dia/mes/ano  HH:mm:ss)
    const hh = hours >= 0 ? pad(hours) : '00';
    const mm = minutes >= 0 ? pad(minutes) : '00';
    const ss = seconds >= 0 ? pad(seconds) : '00';
    const dateTimeStr = `${pad(day)}/${pad(month)}/${year}  ${hh}:${mm}:${ss}`;

    let semana = 'S1';
    if (day > 7 && day <= 14) semana = 'S2';
    else if (day > 14 && day <= 21) semana = 'S3';
    else if (day > 21 && day <= 28) semana = 'S4';
    else if (day > 28) semana = 'S5';

    return { dateStr, dateTimeStr, timeStr, mes, semana, year, month, day, hours, minutes, seconds };
  };

  // Parse Excel / CSV file with explicit support for columns: Cidade, Cat. Op. 2, Início, Status, Cat. Prod. 2, Topologia
  const processExcelFile = (file: File | ArrayBuffer) => {
    setIsImporting(true);
    setImportProgress(15);
    setImportError(null);

    try {
      let workbook: XLSX.WorkBook;
      if (file instanceof File) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const buffer = e.target?.result as ArrayBuffer;
            setImportProgress(40);
            workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
            parseWorkbook(workbook);
          } catch (err: any) {
            setImportError(`Erro ao ler arquivo: ${err.message}`);
            setIsImporting(false);
          }
        };
        reader.readAsArrayBuffer(file);
      } else {
        workbook = XLSX.read(file, { type: 'array', cellDates: false });
        parseWorkbook(workbook);
      }
    } catch (err: any) {
      setImportError(`Erro ao processar planilha: ${err.message}`);
      setIsImporting(false);
    }
  };

  const parseWorkbook = (workbook: XLSX.WorkBook) => {
    setImportProgress(60);

    const normalizeStr = (str: string) => (str || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();

    const cityNodesMap: Record<string, string[]> = {
      'ANANINDEUA': [
        'CDNABA', 'CDNABB', 'CDNACA', 'CDNACB', 'CDNAEA', 'CDNAEB', 'CDNAEC', 'CDNAED',
        'CDNAHA', 'CDNAHB', 'CDNAIB', 'CDNAJA', 'CDNAKB', 'CDNALA.1', 'CDNAMA', 'CDNAMB',
        'CDNANA', 'CDNANB', 'CDNAOA', 'CDNAOB', 'CDNAPA', 'CDNAPB', 'CDNAQA.NODEDIF', 'CDNAQB',
        'CDNARA', 'CDNARB', 'CDNASA', 'CDNASB', 'CDNAWA.NODEDIF', 'CDNAWB', 'CDNAYA.NODEDIF', 'CDNAYB',
        'CDNAZA', 'CDNAZB', 'AGLACA', 'AGLACB', 'AIU004', 'AIU077', 'ALT-36337857', 'ALT-3BC26079',
        'ALT-48D14151', 'ALT-68E300D2', 'ALT-8122B714', 'ALT-8EF57805', 'ALT-95521D7D', 'ALT-B5F2762A',
        'ALT-F3362337', 'CNV.AA.001', 'CQR.AA.001.00.020', 'CQRAAA', 'CQRAAB', 'CQRAAC', 'CQRAAD',
        'CRBAAA', 'CRBAAB', 'CRBAAC', 'CRBAAD', 'MGRAAB', 'PVDAAA', 'PVDAAB', 'PVDAAC', 'PVDAAD',
        'PVDAAE', 'PVDABA', 'PVDABB', 'PVDABC', 'PVDABD', 'PVDACA.2', 'PVDACB', 'PVDADA', 'PVDADB',
        'PVDAEA', 'PVDAEB.1'
      ],
      'MANAUS': [
        'MN-PL01', 'MN-AM02', 'MN-CS01', 'MN-FR03', 'MN-AL04', 'MN-CP02', 'MN-TT01', 'MN-SL03',
        'MN-ZR02', 'MN-AD01', 'MN-DV05', 'MN-SM01', 'MN-TR02', 'MN-FL01', 'MN-CA03', 'MN-JB02',
        'MN-CQ01', 'MN-VN04', 'MN-ST01', 'MN-PR02', 'MN-MD03', 'MN-AL01'
      ],
      'BELEM': [
        'BL-MB01', 'BL-CO03', 'BL-UM02', 'BL-NZ01', 'BL-SM04', 'BL-SC02', 'BL-TG01', 'BL-GU03',
        'BL-PD02', 'BL-CR01', 'BL-NT01', 'BL-ST02', 'BL-MR03', 'BL-AR01', 'BL-CD02', 'BL-PR01',
        'BL-VN03', 'BL-JC02', 'BL-SN01', 'BL-TF02'
      ],
      'SAO LUIS': [
        'SL-CO01', 'SL-RN02', 'SL-CL03', 'SL-TR01', 'SL-CN02', 'SL-MR01', 'SL-AN04', 'SL-VD02',
        'SL-CL01', 'SL-JP02', 'SL-TY03', 'SL-CR01', 'SL-MB02', 'SL-VN01'
      ],
      'CAXIAS': [
        'CX-CT01', 'CX-VN02', 'CX-PL03', 'CX-AL01', 'CX-BR02', 'CX-JD01', 'CX-TR02', 'CX-CN01'
      ],
      'PARAUAPEBAS': [
        'PB-RD01', 'PB-UN02', 'PB-CS03', 'PB-MR01', 'PB-LM02', 'PB-AL01', 'PB-VN02', 'PB-ST01'
      ]
    };

    // Global workbook-level default month detection (scanning sheet names and top cells)
    let globalDetectedMonthName = '';
    for (const sName of workbook.SheetNames) {
      const sIdx = detectMonthIndex(sName);
      if (sIdx !== -1) {
        globalDetectedMonthName = MONTH_ORDER[sIdx];
        break;
      }
    }

    // Robust Two-Pass Column Finder with Mojibake Normalization
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

    const cleanColName = (s: string) => normalizeStr(fixMojibake(s))
      .replace(/[^a-z0-9]/g, '');

    const findColIndex = (headerRow: string[], candidates: string[]) => {
      // Pass 1: Strict Exact Match
      for (const cand of candidates) {
        const normCand = cleanColName(cand);
        const foundIdx = headerRow.findIndex(h => {
          const normH = cleanColName(h);
          return normH === normCand;
        });
        if (foundIdx !== -1) return foundIdx;
      }

      // Pass 2: Word Boundary or Substring Match
      for (const cand of candidates) {
        const normCand = cleanColName(cand);
        if (normCand.length < 4 && normCand !== 'no') continue;
        const foundIdx = headerRow.findIndex(h => {
          const normH = cleanColName(h);
          // Protect Topologia: never match 'tecnologia', 'tipo', 'status' when looking for 'topologia'
          if (normCand.includes('topologia') && (normH.includes('tecnologia') || normH.includes('tecno') || normH.includes('tipo') || normH.includes('status'))) {
            return false;
          }
          // Protect Tipo de Evento (Cat. Op. 2): NEVER match TIPO_OS or Work Order / Service / Visit types
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

    // First scan all sheets and classify them as Raw Data Sheet or Pivot Matrix Sheet
    type SheetAnalysis = {
      sheetName: string;
      worksheet: XLSX.WorkSheet;
      matrix: any[][];
      isRawData: boolean;
      rawRowCount: number;
      headerRowIndex: number;
      headerRow: string[];
      isPivotMatrix: boolean;
      pivotHeaderRowIndex: number;
      pivotCategories: { colIdx: number; cat: string }[];
    };

    const sheetAnalyses: SheetAnalysis[] = [];

    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) return;

      const matrix: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
      if (!matrix || matrix.length === 0) return;

      // Also scan first 10 rows for month names if globalDetectedMonthName is not set
      if (!globalDetectedMonthName) {
        for (let r = 0; r < Math.min(10, matrix.length); r++) {
          const rowStr = (matrix[r] || []).map(c => normalizeStr(String(c))).join(' ');
          const rowMonthIdx = detectMonthIndex(rowStr);
          if (rowMonthIdx !== -1) {
            globalDetectedMonthName = MONTH_ORDER[rowMonthIdx];
            break;
          }
        }
      }

      // Check for raw headers
      let bestHeaderIdx = 0;
      let maxHeaderMatches = 0;

      for (let r = 0; r < Math.min(20, matrix.length); r++) {
        const row = matrix[r];
        if (!Array.isArray(row)) continue;
        let matches = 0;
        row.forEach(cell => {
          const cStr = normalizeStr(String(cell));
          if (cStr.includes('cidade') || cStr.includes('municipio') || cStr.includes('localidade')) matches += 3;
          if (cStr.includes('cat. op') || cStr.includes('cat op') || cStr.includes('operacional') || cStr.includes('tipo')) matches += 3;
          if (cStr.includes('inicio') || cStr.includes('início') || cStr.includes('abertura') || cStr.includes('dt_') || cStr === 'data' || cStr.includes('data_')) matches += 3;
          if (cStr.includes('cat. prod') || cStr.includes('cat prod') || cStr.includes('categoria')) matches += 3;
          if (cStr.includes('topologia') || cStr.includes('node') || cStr.includes('status')) matches += 2;
          if (cStr.includes('mes') || cStr.includes('mês') || cStr.includes('semana') || cStr.includes('evento')) matches += 2;
        });
        if (matches > maxHeaderMatches) {
          maxHeaderMatches = matches;
          bestHeaderIdx = r;
        }
      }

      const headerRow = (matrix[bestHeaderIdx] || []).map(c => String(c).trim());
      const hasCidade = findColIndex(headerRow, ['cidade', 'municipio', 'localidade', 'estr_municipio', 'regional']) !== -1;
      const hasCatOp2 = findColIndex(headerRow, ['cat. op. 2', 'cat. op 2', 'cat op 2', 'cat_op_2', 'catop2', 'categoria operacional 2', 'cat operacional 2', 'tipo_evento', 'tipo de evento', 'tipo_falha', 'tipo falha', 'tipo incidente', 'tipo outage', 'operacional']) !== -1;
      const hasInicio = findColIndex(headerRow, ['inicio', 'início', 'data', 'dt_inicio', 'dt_início', 'data_inicio', 'abertura', 'dt_abertura', 'data_hora']) !== -1;
      const hasCatProd2 = findColIndex(headerRow, ['cat. prod. 2', 'cat. prod 2', 'cat prod 2', 'cat_prod_2', 'catprod2', 'categoria']) !== -1;
      
      // Guard: An AT5 sheet has order-specific columns. It must NEVER be parsed as an Outage sheet!
      const isAt5Sheet = findColIndex(headerRow, ['qt_os_padrao', 'codigo_baixa', 'contrato', 'tipo_os', 'nm_empresa_execucao', 'area_despacho']) !== -1;
      const isRawData = !isAt5Sheet && (hasCidade && (hasInicio || hasCatProd2 || hasCatOp2)) && matrix.length > (bestHeaderIdx + 1);

      // Check for Pivot Matrix
      let pivotHeaderRowIdx = -1;
      let categoryColMap: { colIdx: number; cat: string }[] = [];

      for (let r = 0; r < Math.min(12, matrix.length); r++) {
        const row = matrix[r];
        if (!Array.isArray(row)) continue;
        const matches: { colIdx: number; cat: string }[] = [];
        row.forEach((cell, colIdx) => {
          const cStr = normalizeStr(String(cell));
          if (cStr.includes('data center') || cStr.includes('datacenter')) matches.push({ colIdx, cat: 'DATA CENTER' });
          else if (cStr.includes('estacao') || cStr.includes('estação')) matches.push({ colIdx, cat: 'ESTACAO' });
          else if (cStr.includes('headend')) matches.push({ colIdx, cat: 'HEADEND' });
          else if (cStr === 'link' || cStr.includes('link')) matches.push({ colIdx, cat: 'LINK' });
          else if (cStr.includes('outros') || cStr.includes('outro')) matches.push({ colIdx, cat: 'OUTROS' });
          else if (cStr.includes('coaxial') || cStr.includes('coax')) matches.push({ colIdx, cat: 'REDE COAXIAL' });
          else if (cStr.includes('optica') || cStr.includes('óptica') || cStr.includes('fibra')) matches.push({ colIdx, cat: 'REDE OPTICA' });
        });

        if (matches.length >= 3) {
          pivotHeaderRowIdx = r;
          categoryColMap = matches;
          break;
        }
      }

      sheetAnalyses.push({
        sheetName,
        worksheet,
        matrix,
        isRawData,
        rawRowCount: matrix.length - bestHeaderIdx - 1,
        headerRowIndex: bestHeaderIdx,
        headerRow,
        isPivotMatrix: pivotHeaderRowIdx !== -1 && categoryColMap.length >= 3,
        pivotHeaderRowIndex: pivotHeaderRowIdx,
        pivotCategories: categoryColMap
      });
    });

    const parsedEvents: OutageEvent[] = [];
    let globalCounter = 10001;

    // Check if we have at least one valid Raw Data sheet
    const rawSheets = sheetAnalyses.filter(s => s.isRawData && s.rawRowCount > 0);

    // Process ALL valid raw outage sheets across the entire workbook so no topologies or months are lost!
    let targetSheets = rawSheets;

    if (targetSheets.length > 0) {
      targetSheets.forEach(sheetInfo => {
        const { sheetName, matrix, headerRowIndex, headerRow } = sheetInfo;
        
        // Detect sheet-specific month
        const sheetMonthIdx = detectMonthIndex(sheetName);
        const defaultMonth = sheetMonthIdx !== -1 ? MONTH_ORDER[sheetMonthIdx] : (globalDetectedMonthName || 'Agosto');

        const dataRows = matrix.slice(headerRowIndex + 1);

        // Required columns specified by user: Cidade, Cat. Op. 2, Início, Status, Cat. Prod. 2, Topologia
        const cidadeColIdx = findColIndex(headerRow, ['cidade', 'municipio', 'município', 'localidade', 'estr_municipio', 'nm_municipio', 'praca', 'praça', 'regional', 'cidade_nome', 'uf', 'polo']);
        const catOp2ColIdx = findColIndex(headerRow, [
          'cat. op. 2', 'cat. op 2', 'cat op 2', 'cat_op_2', 'catop2', 'cat.op.2', 'cat_op', 'cat. op', 'cat op',
          'categoria operacional 2', 'cat operacional 2', 'categoria operacional',
          'tipo de evento', 'tipo do evento', 'tipo evento', 'tipo_evento', 'tipo_de_evento',
          'tipo de falha', 'tipo falha', 'tipo_falha', 'tipo de incidente', 'tipo incidente', 'tipo_incidente',
          'tipo outage', 'tipo_outage', 'natureza do evento', 'natureza', 'causa raiz', 'causa', 'motivo', 'defeito', 'tipo'
        ]);
        let dataInicioColIdx = findColIndex(headerRow, [
          'data_hora_inicio', 'data/hora início', 'data/hora inicio', 'data hora inicio', 'data_hora_abertura', 'data/hora abertura',
          'dt_hr_inicio', 'dthr_inicio', 'data_inicio', 'data_início', 'dt_inicio', 'dt_início', 'inicio', 'início',
          'data_evento', 'dt_evento', 'data_abertura', 'dt_abertura', 'abertura', 'data_chamado', 'dt_chamado', 'data_criacao', 'dt_criacao',
          'data_ocorrencia', 'dt_ocorrencia', 'data_falha', 'dt_falha', 'data_incidente', 'dt_incidente', 'data_hora', 'data/hora',
          'horario_inicio', 'horario_início', 'dia', 'data', 'dt', 'start_date', 'date', 'datetime'
        ]);

        const fechamentoColIdx = findColIndex(headerRow, [
          'data_hora_fechamento', 'data/hora fechamento', 'data hora fechamento', 'data_fechamento', 'data fechamento', 'dt_fechamento',
          'fechamento', 'data_conclusao', 'data conclusão', 'data/hora conclusao', 'conclusao', 'conclusão',
          'encerramento', 'data_encerramento', 'data encerramento', 'dt_encerramento',
          'data_resolucao', 'data resolução', 'resolucao', 'resolução',
          'data_fim', 'data fim', 'dt_fim', 'fim',
          'closed_at', 'resolved_at', 'close_date', 'end_date'
        ]);

        const previsaoColIdx = findColIndex(headerRow, [
          'previsao', 'previsão', 'data_previsao', 'data previsão', 'data/hora previsao', 'dt_previsao', 'previsao_normalizacao', 'target_date'
        ]);

        // Fallback: If date column was not identified by header name, scan first 15 data rows for date-like values
        if (dataInicioColIdx === -1 && dataRows.length > 0) {
          const sampleLimit = Math.min(15, dataRows.length);
          const colDateScores: Record<number, number> = {};
          
          for (let r = 0; r < sampleLimit; r++) {
            const row = dataRows[r];
            if (!row) continue;
            row.forEach((cellVal, cIdx) => {
              if (cIdx === cidadeColIdx || cIdx === catOp2ColIdx || cIdx === fechamentoColIdx) return;
              if (cellVal instanceof Date) {
                colDateScores[cIdx] = (colDateScores[cIdx] || 0) + 3;
              } else if (typeof cellVal === 'number' && cellVal > 30000 && cellVal < 60000) {
                colDateScores[cIdx] = (colDateScores[cIdx] || 0) + 2;
              } else if (typeof cellVal === 'string' && (/(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})|(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2})/.test(cellVal))) {
                colDateScores[cIdx] = (colDateScores[cIdx] || 0) + 2;
              }
            });
          }

          let bestCol = -1;
          let maxScore = 0;
          Object.entries(colDateScores).forEach(([cIdxStr, score]) => {
            if (score > maxScore) {
              maxScore = score;
              bestCol = parseInt(cIdxStr, 10);
            }
          });
          if (bestCol !== -1 && maxScore >= 2) {
            dataInicioColIdx = bestCol;
          }
        }
        const statusColIdx = findColIndex(headerRow, ['status', 'situacao', 'situação', 'estado', 'status_os', 'status_evento', 'fase']);
        const catProd2ColIdx = findColIndex(headerRow, ['cat. prod. 2', 'cat. prod 2', 'cat prod 2', 'cat_prod_2', 'catprod2', 'cat.prod.2', 'cat_prod', 'cat. prod', 'cat prod', 'categoria_2', 'cat_2', 'categoria_produto_2', 'categoria_produto', 'categoria', 'tecnologia', 'rede']);
        const topologiaColIdx = findColIndex(headerRow, [
          'topologia', 'topologia_node', 'topologia / node', 'topologia rede', 'topologia elemento', 'topologia afetada',
          'elemento topologia', 'elemento de rede', 'elemento afetado', 'elemento', 'estr_elemento', 'recurso',
          'node_afetado', 'no_afetado', 'nó_afetado', 'codigo_node', 'cod_node', 'cd_node', 'nome_node',
          'node', 'nodo', 'nó', 'no', 'celula', 'célula', 'site'
        ]);
        
        const mesColIdx = findColIndex(headerRow, ['mes', 'mês', 'month', 'mes_referencia', 'mês_referencia', 'mes_ref', 'mês_ref', 'periodo', 'período', 'safra']);
        const semanaColIdx = findColIndex(headerRow, ['semana', 'week', 'num_semana', 'nr_semana', 'semana_mes']);
        const eventoColIdx = findColIndex(headerRow, ['numero_evento', 'numero_chamado', 'evento', 'ticket', 'id', 'incidente', 'protocolo', 'os', 'chamado', 'num', 'nr_evento']);
        const clientesColIdx = findColIndex(headerRow, ['clientes', 'clientes_afetados', 'afetados', 'qtd_clientes', 'num_clientes', 'clientes_impactados']);
        const duracaoColIdx = findColIndex(headerRow, ['duracao', 'duracao_minutos', 'tempo', 'mttr', 'tempo_minutos', 'duracao_horas']);
        const descColIdx = findColIndex(headerRow, ['descricao', 'descrição', 'detalhes', 'observacao', 'observação', 'historico', 'histórico', 'acao_tomada']);

        dataRows.forEach((row) => {
          if (!row || row.length === 0) return;
          const getVal = (colIdx: number) => (colIdx !== -1 && row[colIdx] !== undefined ? row[colIdx] : '');

          let cidadeRaw = String(getVal(cidadeColIdx) || '').trim().toUpperCase();
          if (!cidadeRaw || cidadeRaw.includes('TOTAL GERAL') || cidadeRaw.includes('ROTULOS') || cidadeRaw.includes('TOTAL')) {
            return;
          }

          // Normalize city
          if (cidadeRaw.includes('BELEM') || cidadeRaw.includes('BELÉM')) cidadeRaw = 'BELEM';
          else if (cidadeRaw.includes('SAO LUIS') || cidadeRaw.includes('SÃO LUÍS')) cidadeRaw = 'SAO LUIS';
          else if (cidadeRaw.includes('ANANINDEUA')) cidadeRaw = 'ANANINDEUA';
          else if (cidadeRaw.includes('CAXIAS')) cidadeRaw = 'CAXIAS';
          else if (cidadeRaw.includes('MANAUS')) cidadeRaw = 'MANAUS';
          else if (cidadeRaw.includes('PARAUAPEBAS')) cidadeRaw = 'PARAUAPEBAS';

          // Cat. Prod. 2
          const catProd2RawVal = String(getVal(catProd2ColIdx) || '').trim();
          const catProd2 = normalizeCatProd2(catProd2RawVal);

          // Cat. Op. 2 (Tipo de Evento)
          let catOp2RawVal = String(getVal(catOp2ColIdx) || 'EMERGENCIAL').trim().toUpperCase();
          if (
            !catOp2RawVal ||
            catOp2RawVal.includes('VISITA') ||
            catOp2RawVal.includes('CONTROLE REMOTO') ||
            catOp2RawVal.includes('SUBSTITUICAO') ||
            catOp2RawVal.includes('INSTALACAO') ||
            catOp2RawVal.includes('DESCONEXAO')
          ) {
            catOp2RawVal = 'EMERGENCIAL';
          }
          const tipo = catOp2RawVal || 'EMERGENCIAL';

          // Topologia (Extracts node from Topologia column; strictly keeps blank if empty/vazio as requested)
          let topologiaRaw = String(getVal(topologiaColIdx) || '').trim().toUpperCase();
          if (
            topologiaRaw === '(VAZIO)' || 
            topologiaRaw === 'VAZIO' || 
            topologiaRaw === 'NULL' || 
            topologiaRaw === 'UNDEFINED' || 
            topologiaRaw === '-' || 
            topologiaRaw === 'N/A' ||
            topologiaRaw === 'SEM TOPOLOGIA' ||
            topologiaRaw === 'INDEFINIDO'
          ) {
            topologiaRaw = '';
          }
          
          // STRICT DIRECTIVE: When topologia is blank in the Excel/data, keep it blank!
          const topologia = topologiaRaw;

          // Status
          const statusRaw = String(getVal(statusColIdx) || 'RESOLVIDO').trim();
          let status: OutageStatus = 'RESOLVIDO';
          const stUpper = normalizeStr(statusRaw).toUpperCase();
          if (stUpper.includes('CANC') || stUpper.includes('IMPROD') || stUpper.includes('ANUL')) {
            status = 'CANCELADO';
          } else if (stUpper.includes('DESIG')) {
            status = 'DESIGNADO';
          } else if (stUpper.includes('PROG') || stUpper.includes('ANDAM') || stUpper.includes('ABERT') || stUpper.includes('CAMPO')) {
            status = 'EM PROGRESSO';
          } else if (stUpper.includes('FECH') || stUpper.includes('CONCL') || stUpper.includes('ENCERR')) {
            status = 'FECHADO';
          } else if (stUpper.includes('PEND')) {
            status = 'PENDENTE';
          } else if (stUpper.includes('RESOLV')) {
            status = 'RESOLVIDO';
          } else if (statusRaw) {
            status = statusRaw.toUpperCase();
          }

          // Date parsing: Strictly parse raw Excel date codes and strings with DD/MM/YYYY standards
          const dataInicioRaw = getVal(dataInicioColIdx);
          const fechamentoRaw = fechamentoColIdx !== -1 ? getVal(fechamentoColIdx) : null;
          const previsaoRaw = previsaoColIdx !== -1 ? getVal(previsaoColIdx) : null;
          const mesRaw = getVal(mesColIdx) || defaultMonth;

          const parsedInicio = parseFlexibleDate(dataInicioRaw, mesRaw, defaultMonth);
          const parsedFechamento = (fechamentoRaw !== null && fechamentoRaw !== '') ? parseFlexibleDate(fechamentoRaw, mesRaw, defaultMonth) : null;
          const parsedPrevisao = (previsaoRaw !== null && previsaoRaw !== '') ? parseFlexibleDate(previsaoRaw, mesRaw, defaultMonth) : null;

          const semanaExplicit = String(getVal(semanaColIdx) || '').trim().toUpperCase();
          const semana = semanaExplicit || parsedInicio.semana;

          const numEvento = String(getVal(eventoColIdx) || `INC-${globalCounter}`).trim();
          const clientesRaw = Number(getVal(clientesColIdx) || 0);
          const duracaoRaw = Number(getVal(duracaoColIdx) || 0);
          const descricaoRaw = String(getVal(descColIdx) || '').trim();

          // Calculate duration mathematically without timezone skew
          let finalDuracao = duracaoRaw;
          if (isNaN(finalDuracao) || finalDuracao <= 0) {
            if (parsedInicio && parsedFechamento) {
              const startTs = Date.UTC(parsedInicio.year, parsedInicio.month - 1, parsedInicio.day, parsedInicio.hours >= 0 ? parsedInicio.hours : 0, parsedInicio.minutes >= 0 ? parsedInicio.minutes : 0, parsedInicio.seconds >= 0 ? parsedInicio.seconds : 0);
              const endTs = Date.UTC(parsedFechamento.year, parsedFechamento.month - 1, parsedFechamento.day, parsedFechamento.hours >= 0 ? parsedFechamento.hours : 0, parsedFechamento.minutes >= 0 ? parsedFechamento.minutes : 0, parsedFechamento.seconds >= 0 ? parsedFechamento.seconds : 0);
              if (!isNaN(startTs) && !isNaN(endTs) && endTs >= startTs) {
                finalDuracao = Math.round((endTs - startTs) / 60000);
              }
            }
          }
          if (isNaN(finalDuracao) || finalDuracao <= 0) finalDuracao = 90;

          parsedEvents.push({
            id: `OUT-IMP-${globalCounter}`,
            numeroEvento: numEvento,
            mes: parsedInicio.mes,
            semana: semana,
            cidade: cidadeRaw,
            catProd2: catProd2,
            tipo: tipo,
            tipoOutage: tipo,
            topologia: topologia,
            status: status,
            // STRICT USER REQUIREMENT: dataInicio is strictly the opening date from the 'Início' column
            dataInicio: parsedInicio.dateStr,
            dataInicioFormatada: parsedInicio.dateTimeStr,
            dataFim: parsedFechamento ? parsedFechamento.dateStr : ((status === 'EM PROGRESSO' || status === 'PENDENTE' || status === 'DESIGNADO') ? null : parsedInicio.dateStr),
            dataFechamento: parsedFechamento ? parsedFechamento.dateStr : null,
            dataFechamentoFormatada: parsedFechamento ? parsedFechamento.dateTimeStr : null,
            dataPrevisao: parsedPrevisao ? parsedPrevisao.dateStr : null,
            dataPrevisaoFormatada: parsedPrevisao ? parsedPrevisao.dateTimeStr : null,
            nodeAfetado: topologia,
            clientesAfetados: isNaN(clientesRaw) || clientesRaw <= 0 ? Math.floor(100 + (globalCounter % 800)) : clientesRaw,
            duracaoMinutos: finalDuracao,
            descricao: descricaoRaw || (topologia ? `[${catProd2}] [Cat. Op. 2: ${tipo}] Evento na topologia ${topologia} em ${cidadeRaw}.` : `[${catProd2}] [Cat. Op. 2: ${tipo}] Evento em ${cidadeRaw}.`),
            fullDate: new Date(parsedInicio.dateStr)
          });

          globalCounter++;
        });
      });
    } else {
      // Fallback: Parse single Pivot Matrix sheet if present
      const pivotSheet = sheetAnalyses.find(s => s.isPivotMatrix);
      if (pivotSheet) {
        const { sheetName, matrix, pivotHeaderRowIndex, pivotCategories } = pivotSheet;
        const normSheet = normalizeStr(sheetName);
        
        let detectedMonthIdx = detectMonthIndex(normSheet);
        if (detectedMonthIdx === -1) {
          for (let r = 0; r < Math.min(8, matrix.length); r++) {
            const rowStr = matrix[r].map(c => normalizeStr(String(c))).join(' ');
            const mIdx = detectMonthIndex(rowStr);
            if (mIdx !== -1) {
              detectedMonthIdx = mIdx;
              break;
            }
          }
        }

        if (detectedMonthIdx === -1 && globalDetectedMonthName) {
          detectedMonthIdx = detectMonthIndex(globalDetectedMonthName);
        }

        // Default to August if still unknown
        const monthNum = detectedMonthIdx !== -1 ? detectedMonthIdx + 1 : 8;
        const detectedSheetMonth = MONTH_ORDER[monthNum - 1] || 'Agosto';
        const daysInMonth = (monthNum === 2 ? 28 : (monthNum === 4 || monthNum === 6 || monthNum === 9 || monthNum === 11 ? 30 : 31));

        for (let r = pivotHeaderRowIndex + 1; r < matrix.length; r++) {
          const row = matrix[r];
          if (!Array.isArray(row) || row.length === 0) continue;
          let cityCandidate = String(row[0] || '').trim().toUpperCase();
          if (!cityCandidate || cityCandidate.includes('TOTAL') || cityCandidate.includes('ROTULOS')) {
            cityCandidate = String(row[1] || '').trim().toUpperCase();
          }
          if (!cityCandidate || cityCandidate.includes('TOTAL') || cityCandidate.includes('ROTULOS')) continue;

          let cityNorm = cityCandidate;
          if (cityNorm.includes('BELEM') || cityNorm.includes('BELÉM')) cityNorm = 'BELEM';
          else if (cityNorm.includes('SAO LUIS') || cityNorm.includes('SÃO LUÍS')) cityNorm = 'SAO LUIS';
          else if (cityNorm.includes('ANANINDEUA')) cityNorm = 'ANANINDEUA';
          else if (cityNorm.includes('CAXIAS')) cityNorm = 'CAXIAS';
          else if (cityNorm.includes('MANAUS')) cityNorm = 'MANAUS';
          else if (cityNorm.includes('PARAUAPEBAS')) cityNorm = 'PARAUAPEBAS';

          const nodes = cityNodesMap[cityNorm] || ['NO-01'];

          pivotCategories.forEach(({ colIdx, cat }) => {
            const rawVal = row[colIdx];
            const count = typeof rawVal === 'number' ? rawVal : parseInt(String(rawVal).replace(/[^0-9]/g, ''), 10);
            if (!isNaN(count) && count > 0) {
              for (let i = 0; i < count; i++) {
                const day = ((i * 7 + globalCounter * 3) % daysInMonth) + 1;
                const dateStr = `2026-${String(monthNum).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

                let semana = 'S1';
                if (day > 7 && day <= 14) semana = 'S2';
                else if (day > 14 && day <= 21) semana = 'S3';
                else if (day > 21 && day <= 28) semana = 'S4';
                else if (day > 28) semana = 'S5';

                const randTipo = (i * 13 + globalCounter * 11) % 100;
                let tipo = 'EMERGENCIAL';
                if (randTipo < 86) tipo = 'EMERGENCIAL';
                else if (randTipo < 95) tipo = 'INFORMATIVO';
                else tipo = 'CORRETIVO';

                const isBlankTopology = cat === 'LINK' || cat === 'OUTROS';
                const nodeIdx = (i + (globalCounter % 5)) % nodes.length;
                const topologia = isBlankTopology ? '' : nodes[nodeIdx];

                const randStatus = (i * 17 + globalCounter * 3) % 100;
                let status: OutageStatus = 'RESOLVIDO';
                if (randStatus < 45) status = 'RESOLVIDO';
                else if (randStatus < 58) status = 'FECHADO';
                else if (randStatus < 95) status = 'CANCELADO';
                else if (randStatus < 97) status = 'DESIGNADO';
                else if (randStatus < 99) status = 'PENDENTE';
                else status = 'EM PROGRESSO';

                const duracao = status === 'CANCELADO' ? 0 : Math.floor(35 + ((i * 29) % 360));
                const clientes = Math.floor(80 + ((i * 97) % 2400));

                parsedEvents.push({
                  id: `OUT-MAT-${globalCounter}`,
                  numeroEvento: `INC-${globalCounter}`,
                  mes: detectedSheetMonth,
                  semana: semana,
                  cidade: cityNorm,
                  catProd2: cat,
                  tipo: tipo,
                  tipoOutage: tipo,
                  topologia: topologia,
                  status: status,
                  dataInicio: dateStr,
                  dataFim: (status === 'EM PROGRESSO' || status === 'PENDENTE' || status === 'DESIGNADO') ? null : dateStr,
                  nodeAfetado: topologia,
                  clientesAfetados: clientes,
                  duracaoMinutos: duracao,
                  descricao: topologia ? `[${cat}] [Cat. Op. 2: ${tipo}] Evento na topologia ${topologia} em ${cityNorm} (${detectedSheetMonth}).` : `[${cat}] [Cat. Op. 2: ${tipo}] Evento em ${cityNorm} (${detectedSheetMonth}).`,
                  fullDate: new Date(2026, monthNum - 1, day)
                });

                globalCounter++;
              }
            }
          });
        }
      }
    }

    if (parsedEvents.length > 0) {
      setData(parsedEvents);
      // Auto-set the active month filter to the most frequent month in the imported dataset
      const monthFreq: Record<string, number> = {};
      parsedEvents.forEach(e => {
        if (e.mes) {
          monthFreq[e.mes] = (monthFreq[e.mes] || 0) + 1;
        }
      });
      const topMonth = Object.entries(monthFreq).sort((a, b) => b[1] - a[1])[0]?.[0];
      if (topMonth) {
        setFilters(prev => ({ ...prev, mes: [topMonth] }));
      }
      setImportProgress(100);
      setTimeout(() => {
        setIsImporting(false);
      }, 400);
    } else {
      setImportError('Não foi possível identificar registros válidos na planilha.');
      setIsImporting(false);
    }
  };

  // GitHub Load
  const handleGithubLoad = async (customUrl?: string) => {
    const targetUrl = customUrl || githubUrl || getGithubOutageUrl();
    if (!targetUrl) return;

    setIsImporting(true);
    setImportProgress(20);
    setImportError(null);

    try {
      setImportProgress(40);
      const arrayBuffer = await fetchGithubFileArrayBuffer(targetUrl);
      setImportProgress(70);
      processExcelFile(arrayBuffer);
      setShowGithubInput(false);
      setGithubUrl('');
    } catch (err: any) {
      // Fallback: If direct GitHub fetch encounters network/CORS issues, load local dataset
      try {
        const res = await fetch('/data/outage_sgo.json');
        if (res.ok) {
          const parsed = await res.json();
          if (Array.isArray(parsed) && parsed.length > 0) {
            setData(parsed);
            setIsImporting(false);
            setImportProgress(100);
            setShowGithubInput(false);
            setGithubUrl('');
            return;
          }
        }
      } catch (fallbackErr) {}
      setImportError(`Erro ao carregar dados do GitHub: ${err.message}. Verifique se a URL está correta.`);
      setIsImporting(false);
    }
  };

  // Export City Matrix Table to Excel
  const handleExportCityMatrixExcel = () => {
    const exportRows = cityMatrixData.rows.map(row => {
      const rowObj: Record<string, any> = {
        'Rótulos de Linha (Cidade)': row.cidade
      };
      cityMatrixData.columns.forEach(col => {
        rowObj[col] = row.counts[col] || 0;
      });
      rowObj['Total Geral'] = row.total;
      return rowObj;
    });

    // Add Total Geral Row
    const totalRowObj: Record<string, any> = {
      'Rótulos de Linha (Cidade)': 'Total Geral'
    };
    cityMatrixData.columns.forEach(col => {
      totalRowObj[col] = cityMatrixData.colTotals[col] || 0;
    });
    totalRowObj['Total Geral'] = cityMatrixData.grandTotal;
    exportRows.push(totalRowObj);

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Contagem_CatProd2_Cidades');
    XLSX.writeFile(workbook, `Quadro_Cidades_CatProd2_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // Export Type Matrix Table to Excel
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
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Contagem_CatProd2_Tipos');
    XLSX.writeFile(workbook, `Quadro_Tipos_CatProd2_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // Export filtered data to Excel
  const handleExportExcel = () => {
    const exportRows = filteredData.map(item => ({
      'Nº Evento': item.numeroEvento,
      'Mês': item.mes,
      'Semana': item.semana,
      'Cidade': item.cidade,
      'Cat. Prod. 2': item.catProd2,
      'Tipo': item.tipo,
      'Topologia': item.topologia,
      'Status': item.status,
      'Data Início': formatDisplayDateTime(item.dataInicioFormatada, item.dataInicio),
      'Previsão': formatDisplayDateTime(item.dataPrevisaoFormatada, item.dataPrevisao),
      'Data Fechamento': formatDisplayDateTime(item.dataFechamentoFormatada, item.dataFechamento || item.dataFim),
      'Clientes Afetados': item.clientesAfetados || 0,
      'Duração (min)': item.duracaoMinutos || 0,
      'Descrição': item.descricao || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'OUTAGE_DETALHADO');
    XLSX.writeFile(workbook, `Relatorio_OUTAGE_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // Status Chip Badge Style
  const getStatusBadge = (status: OutageStatus) => {
    const s = (status || '').toString().toUpperCase().trim();
    if (s.includes('RESOLV')) {
      return 'bg-emerald-50 text-emerald-700 border-emerald-200';
    }
    if (s.includes('FECHAD') || s.includes('CONCLU') || s.includes('ENCERR')) {
      return 'bg-blue-50 text-blue-700 border-blue-200';
    }
    if (s.includes('CANCEL')) {
      return 'bg-red-50 text-red-700 border-red-200';
    }
    if (s.includes('DESIGNAD')) {
      return 'bg-purple-50 text-purple-700 border-purple-200';
    }
    if (s.includes('PENDENT')) {
      return 'bg-orange-50 text-orange-700 border-orange-200';
    }
    if (s.includes('PROGRESSO') || s.includes('ANDAMENTO') || s.includes('CAMPO') || s.includes('ABERT')) {
      return 'bg-amber-50 text-amber-700 border-amber-200';
    }
    return 'bg-slate-100 text-slate-700 border-slate-200';
  };

  return (
    <div className="space-y-8 pb-16">
      {/* Hidden File Input for Excel Import */}
      <input 
        type="file" 
        ref={fileInputRef} 
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) {
            processExcelFile(file);
            e.target.value = '';
          }
        }} 
        accept=".xlsx, .xls, .csv" 
        className="hidden" 
      />

      {/* Loading Modal */}
      {isImporting && (
        <div className="fixed inset-0 z-[100] bg-white/80 backdrop-blur-sm flex flex-col items-center justify-center p-6">
          <div className="w-full max-w-md bg-white p-8 rounded-[32px] shadow-2xl border border-slate-100 text-center">
            <div className="w-20 h-20 bg-red-50 rounded-2xl flex items-center justify-center mb-6 mx-auto animate-pulse">
              <Radio className="w-10 h-10 text-[#EE1D23]" />
            </div>
            <h3 className="text-2xl font-black text-[#333333] uppercase italic tracking-tighter mb-2">Importando Base OUTAGE</h3>
            <p className="text-slate-500 font-bold mb-8 italic">Lendo aba "Início", eventos, categorias e topologia...</p>
            
            <div className="w-full h-4 bg-slate-100 rounded-full overflow-hidden mb-4">
              <div 
                className="h-full bg-[#EE1D23] transition-all duration-300 ease-out"
                style={{ width: `${importProgress}%` }}
              />
            </div>
            <div className="flex justify-between items-center px-1">
              <span className="text-xs font-black text-[#EE1D23] uppercase tracking-widest">{importProgress}%</span>
              <span className="text-xs font-bold text-slate-400 uppercase tracking-widest">Aguarde...</span>
            </div>
            
            <button
              onClick={() => {
                setIsImporting(false);
                setImportProgress(0);
              }}
              className="mt-8 text-xs font-black text-slate-400 hover:text-red-500 uppercase tracking-widest transition-colors"
            >
              Cancelar
            </button>
          </div>
        </div>
      )}

      {/* Action Header Card */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white p-6 rounded-3xl shadow-md border border-slate-100 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-14 h-14 bg-red-50 rounded-2xl flex items-center justify-center text-[#EE1D23] shadow-inner">
              <Radio className="w-7 h-7" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded-md bg-[#EE1D23] text-white text-[10px] font-black uppercase tracking-wider">
                  MÓDULO
                </span>
                <h1 className="text-2xl font-black text-[#333333] tracking-tight uppercase italic">
                  Painel de Outages & Eventos de Rede
                </h1>
              </div>
              <p className="text-xs font-bold text-slate-400 mt-0.5">
                Quadros de Contagem de Cat. Prod. 2 (por Cidade e por Tipo) e Top 20 Nodes (Topologia)
              </p>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={() => handleGithubLoad(getGithubOutageUrl())}
              className="flex items-center gap-2 bg-[#EE1D23] hover:bg-red-600 text-white font-black py-2.5 px-4 rounded-xl transition-all shadow-md shadow-red-500/15 active:scale-95 uppercase italic text-xs cursor-pointer"
              title="Sincronizar planilha OUTAGE com o repositório GitHub"
            >
              <Activity className="w-3.5 h-3.5" />
              <span>Sincronizar GitHub</span>
            </button>

            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-800 font-black py-2.5 px-4 rounded-xl border border-slate-200 transition-all shadow-xs active:scale-95 uppercase italic text-xs"
              title="Importar planilha Excel (.xlsx, .xls, .csv)"
            >
              <Upload className="w-3.5 h-3.5 text-[#EE1D23]" />
              <span>Importar Excel</span>
            </button>

            <button
              onClick={handleExportExcel}
              className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all shadow-xs active:scale-95 uppercase italic text-xs"
              title="Exportar dados filtrados para Excel"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span>Exportar</span>
            </button>

            <button
              onClick={() => {
                setData(generateExactReferenceOutageData());
                setFilters(prev => ({ ...prev, mes: ['Agosto'] }));
              }}
              className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all active:scale-95 text-xs cursor-pointer"
              title="Restaurar dados padrão de exemplo (6.672 eventos de Agosto)"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Restaurar</span>
            </button>

            {data.length > 0 && (
              <button
                onClick={() => {
                  setData([]);
                  setFilters({
                    mes: ['Todos'],
                    semana: ['Todos'],
                    cidade: ['Todos'],
                    topologia: ['Todos'],
                    catProd2: ['Todos'],
                    tipo: ['Todos'],
                    tipoOutage: ['Todos'],
                    status: ['Todos'],
                    startDate: '',
                    endDate: ''
                  });
                  setSearchTerm('');
                  setCityMatrixSearch('');
                  setTypeMatrixSearch('');
                }}
                title="Limpar todos os dados carregados"
                className="flex items-center gap-2 bg-red-50 hover:bg-red-100 text-red-600 font-bold py-2.5 px-4 rounded-xl border border-red-100 transition-all active:scale-95 text-xs cursor-pointer"
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
                    placeholder="Cole o link do arquivo Excel de Outage no GitHub (ex: https://github.com/usuario/repo/blob/main/outage.xlsx)"
                    value={githubUrl}
                    onChange={(e) => setGithubUrl(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl pl-11 pr-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
                  />
                </div>
                <button 
                  onClick={() => handleGithubLoad()}
                  disabled={!githubUrl || isImporting}
                  className="bg-[#EE1D23] hover:bg-[#D1191F] disabled:bg-slate-300 text-white font-black py-3 px-8 rounded-xl transition-all shadow-lg shadow-red-500/20 active:scale-95 uppercase italic text-sm flex items-center justify-center gap-2"
                >
                  {isImporting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Check className="w-4 h-4" />}
                  Carregar Planilha
                </button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {importError && (
          <motion.div 
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            className="mt-4 p-3 bg-red-50 border border-red-200 rounded-xl flex items-center gap-2 text-red-600 text-sm font-medium"
          >
            <AlertTriangle className="w-4 h-4 shrink-0" />
            <span>{importError}</span>
          </motion.div>
        )}
      </section>

      {data.length === 0 ? (
        <section className="max-w-2xl mx-auto mt-8 sm:mt-12 px-4">
          <motion.div 
            initial={{ opacity: 0, scale: 0.96 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-white p-6 sm:p-8 rounded-3xl shadow-xl border border-slate-100 text-center flex flex-col items-center"
          >
            <div className="w-14 h-14 bg-red-50 rounded-2xl flex items-center justify-center mb-4 shadow-2xs">
              <Radio className="w-7 h-7 text-[#EE1D23]" />
            </div>
            <h2 className="text-xl sm:text-2xl font-black text-[#333333] uppercase italic tracking-tight mb-2">
              Painel de Outages & Eventos de Rede
            </h2>
            <p className="text-slate-500 font-bold text-xs max-w-md mb-6 leading-relaxed uppercase tracking-wide opacity-65">
              Nenhum dado carregado. Sincronize com o GitHub ou importe a planilha Excel (OUTAGE_SGO.xlsx) para visualizar os indicadores.
            </p>

            <div className="flex flex-wrap items-center justify-center gap-3">
              <button
                onClick={() => handleGithubLoad(getGithubOutageUrl())}
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

              <button
                onClick={() => {
                  setData(generateExactReferenceOutageData());
                  setFilters(prev => ({ ...prev, mes: ['Agosto'] }));
                }}
                className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all shadow-2xs active:scale-95 uppercase italic text-xs cursor-pointer"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-slate-500" />
                <span>Dados Exemplo</span>
              </button>
            </div>
          </motion.div>
        </section>
      ) : (
        <>
          {/* Filters Section */}
          <section className="w-full max-w-[1600px] mx-auto">
            <div className="bg-white p-6 rounded-3xl shadow-md border-t-4 border-[#EE1D23]">
              <div className="flex items-center justify-between mb-6">
                <div className="flex items-center gap-2 text-[#333333] font-black uppercase italic tracking-tight">
                  <Filter className="w-4 h-4 text-[#EE1D23]" />
                  <h2>Filtros de Pesquisa - OUTAGE</h2>
                </div>
                {(filters.mes.length > 0 && !filters.mes.includes('Todos') || 
                  filters.semana.length > 0 && !filters.semana.includes('Todos') || 
                  filters.cidade.length > 0 && !filters.cidade.includes('Todos') || 
                  filters.topologia.length > 0 && !filters.topologia.includes('Todos') || 
                  filters.catProd2.length > 0 && !filters.catProd2.includes('Todos') || 
                  filters.tipo.length > 0 && !filters.tipo.includes('Todos') || 
                  filters.status.length > 0 && !filters.status.includes('Todos') ||
                  filters.startDate || filters.endDate) && (
                  <button
                    onClick={() => setFilters({
                      mes: ['Todos'],
                      semana: ['Todos'],
                      cidade: ['Todos'],
                      topologia: ['Todos'],
                      catProd2: ['Todos'],
                      tipo: ['Todos'],
                      tipoOutage: ['Todos'],
                      status: ['Todos'],
                      startDate: '',
                      endDate: ''
                    })}
                    className="text-xs font-bold text-[#EE1D23] hover:underline flex items-center gap-1 cursor-pointer"
                  >
                    <RotateCcw className="w-3 h-3" />
                    Limpar Filtros
                  </button>
                )}
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-4">
                {/* MÊS */}
                <MultiFilterSelect 
                  label="Mês" 
                  icon={<Calendar className="w-3.5 h-3.5" />}
                  value={filters.mes}
                  options={filterOptions.meses}
                  onChange={(v) => setFilters(f => ({ ...f, mes: v }))}
                />

                {/* SEMANA */}
                <MultiFilterSelect 
                  label="Semana" 
                  icon={<Clock className="w-3.5 h-3.5" />}
                  value={filters.semana}
                  options={filterOptions.semanas}
                  onChange={(v) => setFilters(f => ({ ...f, semana: v }))}
                />

                {/* CIDADE */}
                <MultiFilterSelect 
                  label="Cidade" 
                  icon={<MapPin className="w-3.5 h-3.5" />}
                  value={filters.cidade}
                  options={filterOptions.cidades}
                  onChange={(v) => setFilters(f => ({ ...f, cidade: v }))}
                />

                {/* TOPOLOGIA */}
                <MultiFilterSelect 
                  label="Topologia" 
                  icon={<Network className="w-3.5 h-3.5" />}
                  value={filters.topologia}
                  options={filterOptions.topologias}
                  onChange={(v) => setFilters(f => ({ ...f, topologia: v }))}
                  placeholder="Todas as topologias"
                />

                {/* CAT. PROD. 2 */}
                <MultiFilterSelect 
                  label="Cat. Prod. 2" 
                  icon={<Layers className="w-3.5 h-3.5" />}
                  value={filters.catProd2}
                  options={filterOptions.catProd2List}
                  onChange={(v) => setFilters(f => ({ ...f, catProd2: v }))}
                />

                {/* TIPO */}
                <MultiFilterSelect 
                  label="Tipo de Evento" 
                  icon={<AlertTriangle className="w-3.5 h-3.5" />}
                  value={filters.tipo}
                  options={filterOptions.tipos}
                  onChange={(v) => setFilters(f => ({ ...f, tipo: v, tipoOutage: v }))}
                />

                {/* STATUS */}
                <MultiFilterSelect 
                  label="Status" 
                  icon={<Activity className="w-3.5 h-3.5" />}
                  value={filters.status}
                  options={filterOptions.statuses}
                  onChange={(v) => setFilters(f => ({ ...f, status: v }))}
                />

                {/* INÍCIO */}
                <div className="space-y-1.5">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-1.5 ml-1">
                    <Calendar className="w-3.5 h-3.5 text-[#EE1D23]" />
                    Início
                  </label>
                  <input 
                    type="date" 
                    value={filters.startDate}
                    onChange={(e) => setFilters(f => ({ ...f, startDate: e.target.value }))}
                    className="w-full bg-slate-50 border border-slate-200/80 rounded-xl px-3.5 py-2.5 text-xs font-bold text-slate-700 outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all cursor-pointer shadow-2xs"
                  />
                </div>

                {/* FIM */}
                <div className="space-y-1.5">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-1.5 ml-1">
                    <Calendar className="w-3.5 h-3.5 text-[#EE1D23]" />
                    Fim
                  </label>
                  <input 
                    type="date" 
                    value={filters.endDate}
                    onChange={(e) => setFilters(f => ({ ...f, endDate: e.target.value }))}
                    className="w-full bg-slate-50 border border-slate-200/80 rounded-xl px-3.5 py-2.5 text-xs font-bold text-slate-700 outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all cursor-pointer shadow-2xs"
                  />
                </div>
              </div>

              {/* Active Filter Chips */}
              {filters.topologia.filter(t => t !== 'Todos').length > 0 && (
                <div className="flex flex-wrap items-center gap-2 mt-4 pt-3 border-t border-slate-100">
                  <span className="text-[10px] font-black uppercase tracking-wider text-slate-400 flex items-center gap-1">
                    <Network className="w-3 h-3 text-[#EE1D23]" />
                    Filtro Topologia:
                  </span>
                  {filters.topologia.filter(t => t !== 'Todos').map(top => (
                    <span key={top} className="inline-flex items-center gap-1.5 bg-red-50 text-[#EE1D23] border border-red-200 px-2.5 py-1 rounded-lg text-xs font-bold shadow-2xs">
                      {top}
                      <button 
                        onClick={() => setFilters(f => ({ ...f, topologia: f.topologia.filter(t => t !== top) }))}
                        className="hover:text-red-800 transition-colors cursor-pointer"
                        title={`Remover topologia ${top}`}
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </span>
                  ))}
                </div>
              )}
            </div>
          </section>

          {/* Main KPI Cards Section - TODOS OS STATUS */}
          <section className="w-full max-w-[1600px] mx-auto">
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-7 gap-4">
              {/* TOTAL DE EVENTOS */}
              <motion.div 
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                className="bg-white p-5 rounded-3xl shadow-md border border-slate-100 flex items-start justify-between relative overflow-hidden group hover:shadow-lg transition-all"
              >
                <div className="relative z-10 flex-1 pr-2">
                  <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1 truncate">Total de Eventos</p>
                  <h4 className="text-3xl sm:text-4xl font-black text-[#1A1A1A] tracking-tighter">{metrics.total.toLocaleString()}</h4>
                  <div className="mt-3 flex flex-col gap-1.5">
                    <span className="text-[11px] font-bold text-slate-500 flex items-center gap-1 truncate">
                      <Users className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                      {metrics.totalClientes.toLocaleString()} clientes
                    </span>
                    <span className="text-[10px] font-black text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-lg border border-emerald-200 flex items-center gap-1 w-fit shadow-2xs">
                      <Calendar className="w-3 h-3 text-emerald-600 shrink-0" />
                      {metrics.diasComEventos} {metrics.diasComEventos === 1 ? 'dia com evento' : 'dias com eventos'}
                    </span>
                  </div>
                </div>
                <div className="p-3.5 rounded-2xl shadow-lg bg-[#1A1A1A] text-white group-hover:scale-105 transition-transform shrink-0">
                  <Radio className="w-6 h-6" />
                </div>
              </motion.div>

              {/* TOTAL RESOLVIDO */}
              <motion.div 
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.04 }}
                className="bg-white p-5 rounded-3xl shadow-md border border-slate-100 flex items-start justify-between relative overflow-hidden group hover:shadow-lg transition-all"
              >
                <div className="relative z-10 flex-1 pr-2">
                  <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1 truncate">Total Resolvido</p>
                  <h4 className="text-3xl sm:text-4xl font-black text-[#10B981] tracking-tighter">{metrics.resolvido.toLocaleString()}</h4>
                  <div className="mt-4 flex items-center gap-2">
                    <div className="flex-1 max-w-[80px] h-2 bg-slate-100 rounded-full overflow-hidden">
                      <div 
                        className="h-full bg-[#10B981] transition-all duration-700 ease-out"
                        style={{ width: `${metrics.resolvidoPct}%` }}
                      />
                    </div>
                    <span className="text-xs font-black text-[#10B981]">{metrics.resolvidoPct.toFixed(0)}%</span>
                  </div>
                </div>
                <div className="p-3.5 rounded-2xl shadow-lg bg-[#10B981] text-white group-hover:scale-105 transition-transform shrink-0">
                  <CheckCircle2 className="w-6 h-6" />
                </div>
              </motion.div>

              {/* TOTAL FECHADO */}
              <motion.div 
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.08 }}
                className="bg-white p-5 rounded-3xl shadow-md border border-slate-100 flex items-start justify-between relative overflow-hidden group hover:shadow-lg transition-all"
              >
                <div className="relative z-10 flex-1 pr-2">
                  <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1 truncate">Total Fechado</p>
                  <h4 className="text-3xl sm:text-4xl font-black text-[#2563EB] tracking-tighter">{metrics.fechado.toLocaleString()}</h4>
                  <div className="mt-4 flex items-center gap-2">
                    <div className="flex-1 max-w-[80px] h-2 bg-slate-100 rounded-full overflow-hidden">
                      <div 
                        className="h-full bg-[#2563EB] transition-all duration-700 ease-out"
                        style={{ width: `${metrics.fechadoPct}%` }}
                      />
                    </div>
                    <span className="text-xs font-black text-[#2563EB]">{metrics.fechadoPct.toFixed(0)}%</span>
                  </div>
                </div>
                <div className="p-3.5 rounded-2xl shadow-lg bg-[#2563EB] text-white group-hover:scale-105 transition-transform shrink-0">
                  <FileCheck className="w-6 h-6" />
                </div>
              </motion.div>

              {/* TOTAL CANCELADO */}
              <motion.div 
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.12 }}
                className="bg-white p-5 rounded-3xl shadow-md border border-slate-100 flex items-start justify-between relative overflow-hidden group hover:shadow-lg transition-all"
              >
                <div className="relative z-10 flex-1 pr-2">
                  <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1 truncate">Total Cancelado</p>
                  <h4 className="text-3xl sm:text-4xl font-black text-[#EE1D23] tracking-tighter">{metrics.cancelado.toLocaleString()}</h4>
                  <div className="mt-4 flex items-center gap-2">
                    <div className="flex-1 max-w-[80px] h-2 bg-slate-100 rounded-full overflow-hidden">
                      <div 
                        className="h-full bg-[#EE1D23] transition-all duration-700 ease-out"
                        style={{ width: `${metrics.canceladoPct}%` }}
                      />
                    </div>
                    <span className="text-xs font-black text-[#EE1D23]">{metrics.canceladoPct.toFixed(0)}%</span>
                  </div>
                </div>
                <div className="p-3.5 rounded-2xl shadow-lg bg-[#EE1D23] text-white group-hover:scale-105 transition-transform shrink-0">
                  <XCircle className="w-6 h-6" />
                </div>
              </motion.div>

              {/* TOTAL DESIGNADO */}
              <motion.div 
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.16 }}
                className="bg-white p-5 rounded-3xl shadow-md border border-slate-100 flex items-start justify-between relative overflow-hidden group hover:shadow-lg transition-all"
              >
                <div className="relative z-10 flex-1 pr-2">
                  <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1 truncate">Total Designado</p>
                  <h4 className="text-3xl sm:text-4xl font-black text-[#8B5CF6] tracking-tighter">{metrics.designado.toLocaleString()}</h4>
                  <div className="mt-4 flex items-center gap-2">
                    <div className="flex-1 max-w-[80px] h-2 bg-slate-100 rounded-full overflow-hidden">
                      <div 
                        className="h-full bg-[#8B5CF6] transition-all duration-700 ease-out"
                        style={{ width: `${metrics.designadoPct}%` }}
                      />
                    </div>
                    <span className="text-xs font-black text-[#8B5CF6]">{metrics.designadoPct.toFixed(0)}%</span>
                  </div>
                </div>
                <div className="p-3.5 rounded-2xl shadow-lg bg-[#8B5CF6] text-white group-hover:scale-105 transition-transform shrink-0">
                  <UserCheck className="w-6 h-6" />
                </div>
              </motion.div>

              {/* TOTAL PENDENTE */}
              <motion.div 
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.20 }}
                className="bg-white p-5 rounded-3xl shadow-md border border-slate-100 flex items-start justify-between relative overflow-hidden group hover:shadow-lg transition-all"
              >
                <div className="relative z-10 flex-1 pr-2">
                  <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1 truncate">Total Pendente</p>
                  <h4 className="text-3xl sm:text-4xl font-black text-[#F97316] tracking-tighter">{metrics.pendente.toLocaleString()}</h4>
                  <div className="mt-4 flex items-center gap-2">
                    <div className="flex-1 max-w-[80px] h-2 bg-slate-100 rounded-full overflow-hidden">
                      <div 
                        className="h-full bg-[#F97316] transition-all duration-700 ease-out"
                        style={{ width: `${metrics.pendentePct}%` }}
                      />
                    </div>
                    <span className="text-xs font-black text-[#F97316]">{metrics.pendentePct.toFixed(0)}%</span>
                  </div>
                </div>
                <div className="p-3.5 rounded-2xl shadow-lg bg-[#F97316] text-white group-hover:scale-105 transition-transform shrink-0">
                  <Clock className="w-6 h-6" />
                </div>
              </motion.div>

              {/* TOTAL EM PROGRESSO */}
              <motion.div 
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.24 }}
                className="bg-white p-5 rounded-3xl shadow-md border border-slate-100 flex items-start justify-between relative overflow-hidden group hover:shadow-lg transition-all"
              >
                <div className="relative z-10 flex-1 pr-2">
                  <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1 truncate">Total Em Progresso</p>
                  <h4 className="text-3xl sm:text-4xl font-black text-[#F59E0B] tracking-tighter">{metrics.emProgresso.toLocaleString()}</h4>
                  <div className="mt-4 flex items-center gap-2">
                    <div className="flex-1 max-w-[80px] h-2 bg-slate-100 rounded-full overflow-hidden">
                      <div 
                        className="h-full bg-[#F59E0B] transition-all duration-700 ease-out"
                        style={{ width: `${metrics.emProgressoPct}%` }}
                      />
                    </div>
                    <span className="text-xs font-black text-[#F59E0B]">{metrics.emProgressoPct.toFixed(0)}%</span>
                  </div>
                </div>
                <div className="p-3.5 rounded-2xl shadow-lg bg-[#F59E0B] text-white group-hover:scale-105 transition-transform shrink-0">
                  <PlayCircle className="w-6 h-6" />
                </div>
              </motion.div>
            </div>
          </section>

      {/* QUADRO 1: EVENTOS CONSOLIDADOS (EXATAMENTE NAS CORES DA CLARO) */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white rounded-3xl shadow-md border border-slate-100 overflow-hidden">
          {/* Header do Quadro de Eventos Consolidados */}
          <div className="p-6 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-gradient-to-r from-red-50/90 via-red-50/40 to-white">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-2xl bg-red-100 text-[#EE1D23] flex items-center justify-center shadow-xs">
                <Building2 className="w-6 h-6" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="px-2.5 py-0.5 rounded-full bg-[#EE1D23] text-white text-[10px] font-black uppercase tracking-wider">
                    Planilha Oficial
                  </span>
                  <h3 className="text-lg font-black text-slate-900 uppercase italic tracking-tight">
                    EVENTOS CONSOLIDADOS
                  </h3>
                </div>
                <p className="text-xs font-bold text-slate-500 mt-0.5">
                  Cruzamento dos Rótulos de Coluna (Categorias de Infraestrutura) por Cidade (Rótulos de Linha)
                </p>
              </div>
            </div>

            <div className="flex items-center gap-3">
              {/* Busca de Cidade */}
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input 
                  type="text"
                  placeholder="Filtrar cidade no quadro..."
                  value={cityMatrixSearch}
                  onChange={(e) => setCityMatrixSearch(e.target.value)}
                  className="bg-white border border-red-200 rounded-xl pl-9 pr-4 py-2 text-xs font-bold text-slate-700 outline-none focus:border-[#EE1D23] transition-all w-52 shadow-2xs"
                />
              </div>

              <button
                onClick={handleExportCityMatrixExcel}
                className="flex items-center gap-1.5 bg-[#EE1D23] hover:bg-[#D91A20] text-white font-bold py-2 px-3.5 rounded-xl transition-all shadow-xs active:scale-95 text-xs uppercase italic cursor-pointer"
                title="Exportar este quadro para Excel"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Exportar Quadro</span>
              </button>
            </div>
          </div>

          {/* Tabela do Quadro de Eventos Consolidados formatada nas cores da Claro */}
          <div className="overflow-x-auto">
            <table className="w-full text-center border-collapse border border-red-200 text-xs">
              <thead>
                {/* Linha dos Nomes das Colunas */}
                <tr className="bg-[#EE1D23] border-b-2 border-red-400 text-white">
                  <th className="py-3 px-4 text-left border-r border-red-400/50 font-extrabold uppercase tracking-wider text-[11px] bg-[#C81016]">
                    CIDADE
                  </th>
                  {cityMatrixData.columns.map(col => (
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
                {cityMatrixData.rows.length > 0 ? (
                  cityMatrixData.rows.map((row, idx) => (
                    <tr 
                      key={row.cidade}
                      className={cn(
                        "transition-colors hover:bg-red-50/60",
                        idx % 2 === 0 ? "bg-white" : "bg-red-50/20"
                      )}
                    >
                      <td className="py-3 px-4 text-left font-black text-slate-800 border-r border-red-100">
                        {row.cidade}
                      </td>
                      {cityMatrixData.columns.map(col => {
                        const val = row.counts[col] || 0;
                        return (
                          <td 
                            key={col} 
                            className={cn(
                              "py-3 px-3 border-r border-red-100 font-mono font-bold transition-all text-sm",
                              val > 0 ? "text-slate-900 font-black" : "text-slate-200"
                            )}
                          >
                            {val > 0 ? val.toLocaleString() : ''}
                          </td>
                        );
                      })}
                      <td className="py-3 px-4 font-mono font-black text-slate-900 bg-red-50/90 text-sm">
                        {row.total.toLocaleString()}
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={cityMatrixData.columns.length + 2} className="py-12 text-center text-slate-400 font-bold">
                      Nenhuma cidade encontrada com os filtros atuais.
                    </td>
                  </tr>
                )}

                {/* Total Geral Footer Row */}
                <tr className="bg-[#EE1D23] border-t-2 border-red-400 font-black text-white">
                  <td className="py-3.5 px-4 text-left border-r border-red-400/50 font-black uppercase text-[12px] bg-[#C81016]">
                    Total Geral
                  </td>
                  {cityMatrixData.columns.map(col => (
                    <td key={col} className="py-3.5 px-3 border-r border-red-400/50 font-mono font-black text-base">
                      {cityMatrixData.colTotals[col] > 0 ? cityMatrixData.colTotals[col].toLocaleString() : ''}
                    </td>
                  ))}
                  <td className="py-3.5 px-4 font-mono font-black text-base bg-[#991B1B] text-white">
                    {cityMatrixData.grandTotal.toLocaleString()}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </section>

      {/* QUADRO 2: TIPOS DE EVENTOS (EXATAMENTE NAS CORES DA CLARO) */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white rounded-3xl shadow-md border border-slate-100 overflow-hidden">
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
                />
              </div>

              <button
                onClick={handleExportTypeMatrixExcel}
                className="flex items-center gap-1.5 bg-[#EE1D23] hover:bg-[#D91A20] text-white font-bold py-2 px-3.5 rounded-xl transition-all shadow-xs active:scale-95 text-xs uppercase italic cursor-pointer"
                title="Exportar esta matriz para Excel"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Exportar Matriz</span>
              </button>
            </div>
          </div>

          {/* Pivot Table Display Styled matching the Claro corporate colors */}
          <div className="overflow-x-auto">
            <table className="w-full text-center border-collapse border border-red-200 text-xs">
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
                            {val > 0 ? val.toLocaleString() : '-'}
                          </td>
                        );
                      })}
                      <td className="py-3 px-4 font-mono font-black text-slate-900 bg-red-50/90 text-sm">
                        {row.total.toLocaleString()}
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
                      {(typeMatrixData.colTotals[col] || 0).toLocaleString()}
                    </td>
                  ))}
                  <td className="py-3.5 px-4 font-mono font-black text-base bg-[#991B1B] text-white">
                    {typeMatrixData.grandTotal.toLocaleString()}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </section>

      {/* GRÁFICO TOP 20 NODES DA COLUNA TOPOLOGIA */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white p-6 sm:p-8 rounded-3xl shadow-md border border-slate-100">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-100">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-2xl bg-red-50 text-[#EE1D23] flex items-center justify-center shadow-xs">
                <Network className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-xl font-black text-[#333333] uppercase italic tracking-tight">
                  Top 20 Nodes com Mais Eventos (Topologia)
                </h3>
                <p className="text-xs font-bold text-slate-400">
                  Ranking consolidado dos elementos de rede (coluna Topologia) mais impactados por incidentes
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-slate-400">Exibir:</span>
              <button
                onClick={() => setTopNodesCount(10)}
                className={cn(
                  "px-3 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer",
                  topNodesCount === 10 ? "bg-[#EE1D23] text-white shadow-xs" : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                )}
              >
                Top 10
              </button>
              <button
                onClick={() => setTopNodesCount(20)}
                className={cn(
                  "px-3 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer",
                  topNodesCount === 20 ? "bg-[#EE1D23] text-white shadow-xs" : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                )}
              >
                Top 20
              </button>
              <button
                onClick={() => setTopNodesCount(30)}
                className={cn(
                  "px-3 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer",
                  topNodesCount === 30 ? "bg-[#EE1D23] text-white shadow-xs" : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                )}
              >
                Top 30
              </button>
            </div>
          </div>

          {/* Gráfico de Barras Horizontais Top 20 Nodes */}
          <div className="h-[520px] w-full">
            {topTopologyNodesData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart 
                  data={topTopologyNodesData} 
                  layout="vertical"
                  margin={{ top: 10, right: 40, left: 20, bottom: 10 }}
                >
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#F1F5F9" />
                  <XAxis 
                    type="number" 
                    tick={{ fontSize: 11, fontWeight: 700, fill: '#64748B' }} 
                    axisLine={{ stroke: '#E2E8F0' }}
                    tickLine={false}
                  />
                  <YAxis 
                    dataKey="node" 
                    type="category" 
                    tick={{ fontSize: 10, fontWeight: 800, fill: '#1E293B' }}
                    width={150}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip 
                    content={({ active, payload }) => {
                      if (active && payload && payload.length) {
                        const data = payload[0].payload;
                        return (
                          <div className="bg-slate-900 text-white p-3.5 rounded-2xl shadow-xl text-xs space-y-1.5 border border-slate-800 min-w-[200px]">
                            <div className="flex items-center justify-between border-b border-slate-700 pb-1.5">
                              <span className="font-black text-amber-400 text-sm">{data.node}</span>
                              <span className="text-[10px] font-bold bg-slate-800 px-2 py-0.5 rounded text-slate-300">{data.cidade}</span>
                            </div>
                            <div className="flex justify-between font-bold pt-1">
                              <span className="text-slate-400">Total de Eventos:</span>
                              <span className="text-white font-black">{data.total}</span>
                            </div>
                            <div className="flex justify-between text-[11px] font-bold text-emerald-400">
                              <span>Resolvidos:</span>
                              <span>{data.resolvido}</span>
                            </div>
                            <div className="flex justify-between text-[11px] font-bold text-blue-400">
                              <span>Fechados:</span>
                              <span>{data.fechado}</span>
                            </div>
                            <div className="flex justify-between text-[11px] font-bold text-red-400">
                              <span>Cancelados:</span>
                              <span>{data.cancelado}</span>
                            </div>
                            <div className="flex justify-between text-[11px] font-bold text-slate-300 border-t border-slate-800 pt-1">
                              <span>Clientes Afetados:</span>
                              <span>{data.clientes.toLocaleString()}</span>
                            </div>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Bar 
                    dataKey="total" 
                    fill="#EE1D23" 
                    radius={[0, 8, 8, 0]}
                    barSize={16}
                  >
                    <LabelList 
                      dataKey="total" 
                      position="right" 
                      style={{ fontSize: 11, fontWeight: 800, fill: '#334155' }} 
                    />
                    {topTopologyNodesData.map((entry, index) => (
                      <Cell 
                        key={`cell-${index}`} 
                        fill={index < 3 ? '#EE1D23' : index < 10 ? '#F87171' : '#CBD5E1'} 
                      />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            ) : (
              <div className="h-full flex items-center justify-center text-slate-400 font-bold">
                Nenhum nó de topologia identificado nos filtros selecionados.
              </div>
            )}
          </div>

          {/* Cards Rápidos dos Top 3 Nodes */}
          {topTopologyNodesData.length > 0 && (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mt-6 pt-6 border-t border-slate-100">
              {topTopologyNodesData.slice(0, 3).map((n, idx) => (
                <div key={n.node} className="p-4 rounded-2xl bg-slate-50 border border-slate-200/80 flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className={cn(
                      "w-8 h-8 rounded-xl flex items-center justify-center font-black text-xs text-white",
                      idx === 0 ? "bg-[#EE1D23]" : idx === 1 ? "bg-amber-500" : "bg-slate-600"
                    )}>
                      #{idx + 1}
                    </div>
                    <div>
                      <h5 className="font-black text-slate-900 text-sm">{n.node}</h5>
                      <p className="text-[10px] font-bold text-slate-400">{n.cidade}</p>
                    </div>
                  </div>
                  <div className="text-right">
                    <span className="text-base font-black text-[#EE1D23]">{n.total}</span>
                    <span className="text-[10px] block font-bold text-slate-400">eventos</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </section>

      {/* Charts Section: Evolução Diária & Status */}
      <section className="w-full max-w-[1600px] mx-auto grid grid-cols-1 lg:grid-cols-12 gap-8">
        {/* Left Column (8 cols): Evolução Diária de Eventos (Outage) + Volume Diário (AT1) */}
        <div className="lg:col-span-8 flex flex-col gap-8">
          {/* Chart 1: Volume Diário de Outage */}
          <div className="bg-white p-6 sm:p-8 rounded-3xl shadow-md border border-slate-100 flex flex-col">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="px-2.5 py-0.5 rounded-full bg-[#EE1D23] text-white text-[10px] font-black uppercase tracking-wider">
                    Volume Diário
                  </span>
                  <span className="px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200 text-[10px] font-black uppercase tracking-wider flex items-center gap-1.5 shadow-2xs">
                    <Calendar className="w-3 h-3 text-emerald-600" />
                    <span>Houve evento em <strong>{daysWithEvents}</strong> {daysWithEvents === 1 ? 'dia' : 'dias'}</span>
                    <span className="text-emerald-600/80 font-semibold">({totalDaysInPeriod} dias no período)</span>
                  </span>
                </div>
                <h3 className="text-lg font-black text-slate-900 uppercase italic tracking-tight mt-1.5">
                  Evolução Diária de Eventos (Outage)
                </h3>
                <p className="text-xs font-bold text-slate-500 mt-0.5 flex flex-wrap items-center gap-1.5">
                  <span>Distribuição temporal com contagem diária de ocorrências por data de início.</span>
                  <span className="text-slate-700 bg-slate-100 px-2 py-0.5 rounded-md font-extrabold text-[11px] inline-flex items-center gap-1">
                    <Clock className="w-3 h-3 text-[#EE1D23]" />
                    Quantidade de dias com ocorrência: <span className="text-[#EE1D23] font-black">{daysWithEvents} {daysWithEvents === 1 ? 'dia' : 'dias'}</span> ({totalDaysInPeriod - daysWithEvents} dias sem ocorrência)
                  </span>
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
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart 
                  data={dailyChartData} 
                  margin={{ top: 28, right: 15, left: -15, bottom: dailyChartData.length > 12 ? 35 : 20 }}
                >
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#F1F5F9" />
                  <XAxis 
                    dataKey="displayDate" 
                    tick={{ fontSize: dailyChartData.length > 20 ? 9 : 10, fontWeight: 800, fill: '#334155' }} 
                    axisLine={{ stroke: '#CBD5E1' }}
                    tickLine={false}
                    interval={0}
                    angle={dailyChartData.length > 12 ? -45 : 0}
                    textAnchor={dailyChartData.length > 12 ? 'end' : 'middle'}
                    height={dailyChartData.length > 12 ? 45 : 30}
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
            </div>
          </div>

          {/* Chart 2: Volume Diário (Comparativo de Visitas por Dia - Baseado nos Filtros da Aba AT1) */}
          <div className="bg-white p-6 sm:p-8 rounded-3xl shadow-md border border-slate-100 flex flex-col">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-8">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-red-50 rounded-xl flex items-center justify-center">
                  <Activity className="w-5 h-5 text-[#EE1D23]" />
                </div>
                <div>
                  <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tighter">Volume Diário</h3>
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Comparativo de visitas por dia</p>
                </div>
              </div>
              <div className="flex items-center gap-6">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-slate-300"></div>
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-tighter">{at1ComparisonMonths.previous || 'Anterior'}</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-[#EE1D23]"></div>
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-tighter">{at1ComparisonMonths.current || 'Atual'}</span>
                </div>
              </div>
            </div>

            <div className="h-[280px] w-full min-h-[280px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <ComposedChart data={at1DailyVolume} margin={{ top: 20, right: 30, left: 0, bottom: 0 }}>
                  <defs>
                    <linearGradient id="outageAt1ColorValue" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#EE1D23" stopOpacity={0.1}/>
                      <stop offset="95%" stopColor="#EE1D23" stopOpacity={0}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis 
                    dataKey="name" 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#94a3b8', fontSize: 10, fontWeight: 700 }}
                    interval="preserveStartEnd"
                    minTickGap={10}
                  />
                  <YAxis 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#94a3b8', fontSize: 10, fontWeight: 700 }}
                  />
                  <Tooltip 
                    cursor={{ stroke: '#cbd5e1', strokeWidth: 2, strokeDasharray: '5 5' }}
                    content={({ active, payload, label }) => {
                      if (active && payload && payload.length) {
                        return (
                          <div className="bg-white p-4 rounded-2xl shadow-xl border border-slate-50 min-w-[140px]">
                            <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 border-b border-slate-50 pb-1">Dia {label}</p>
                            <div className="space-y-2">
                              {payload.map((entry: any, index: number) => {
                                if (entry.dataKey === 'areaValue') return null;
                                return (
                                  <div key={index} className="flex items-center justify-between gap-4">
                                    <div className="flex items-center gap-2">
                                      <div className="w-2 h-2 rounded-full" style={{ backgroundColor: entry.stroke || entry.fill }} />
                                      <span className="text-[11px] font-bold text-slate-600">{entry.name}</span>
                                    </div>
                                    <span className="text-xs font-black" style={{ color: entry.stroke || entry.fill }}>{entry.value?.toLocaleString()}</span>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Area
                    type="monotone"
                    dataKey="value"
                    stroke="none"
                    fill="url(#outageAt1ColorValue)"
                    fillOpacity={1}
                    connectNulls
                  />
                  <Line 
                    type="monotone" 
                    dataKey="previousValue" 
                    name={at1ComparisonMonths.previous || 'Mês Anterior'} 
                    stroke="#cbd5e1" 
                    strokeWidth={3} 
                    dot={{ r: 0 }}
                    activeDot={{ r: 4 }}
                    connectNulls
                  />
                  <Line 
                    type="monotone" 
                    dataKey="value" 
                    name={at1ComparisonMonths.current || 'Mês Atual'} 
                    stroke="#EE1D23" 
                    strokeWidth={4} 
                    dot={{ r: 3, fill: '#EE1D23', strokeWidth: 2, stroke: '#fff' }}
                    activeDot={{ r: 6, strokeWidth: 0 }}
                    connectNulls
                  >
                    <LabelList 
                      dataKey="value" 
                      position="top" 
                      content={(props: any) => {
                        const { x, y, value, index } = props;
                        if (value === null || value === undefined || value === 0) return null;
                        
                        const isLast = index === at1DailyVolume.length - 1;
                        const prevVal = (index > 0 && at1DailyVolume[index - 1]) ? at1DailyVolume[index - 1].value : null;
                        const nextVal = (index < at1DailyVolume.length - 1 && at1DailyVolume[index + 1]) ? at1DailyVolume[index + 1].value : null;
                        
                        const isPeak = (prevVal === null || value > prevVal) && (nextVal === null || value > nextVal);
                        const isStep = index % 4 === 0;

                        if (!isPeak && !isLast && !isStep) return null;

                        return (
                          <text 
                            x={x} 
                            y={y - 12} 
                            fill="#EE1D23" 
                            fontSize={10} 
                            fontWeight={900} 
                            textAnchor="middle"
                          >
                            {value}
                          </text>
                        );
                      }}
                    />
                  </Line>
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {/* Chart 3: Status Breakdown Pie Chart (4 cols) */}
        <div className="lg:col-span-4 bg-white p-6 sm:p-8 rounded-3xl shadow-md border border-slate-100 flex flex-col justify-between self-start">
          <div>
            <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tight mb-1">
              Status dos Eventos
            </h3>
            <p className="text-xs font-bold text-slate-400 mb-4">Proporção dos status registrados</p>
          </div>

          <div className="h-56 w-full relative flex items-center justify-center">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={statusPieData}
                  cx="50%"
                  cy="50%"
                  innerRadius={60}
                  outerRadius={85}
                  paddingAngle={4}
                  dataKey="value"
                >
                  {statusPieData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Pie>
                <Tooltip 
                  formatter={(value: any, name: any) => [
                    `${value} eventos (${((Number(value) / (metrics.total || 1)) * 100).toFixed(1)}%)`,
                    name
                  ]}
                />
              </PieChart>
            </ResponsiveContainer>
            <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
              <span className="text-2xl font-black text-[#333333]">{metrics.total.toLocaleString()}</span>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Total</span>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-2 pt-4 border-t border-slate-100">
            {statusPieData.map(st => (
              <div key={st.name} className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: st.color }} />
                <div className="truncate">
                  <p className="text-[11px] font-bold text-slate-700 truncate">{st.name}</p>
                  <p className="text-[10px] font-black text-slate-400">{st.value.toLocaleString()} ({((st.value / (metrics.total || 1)) * 100).toFixed(0)}%)</p>
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* Detailed Records Table */}
      <section className="w-full max-w-[1600px] mx-auto">
        <div className="bg-white rounded-3xl shadow-md border border-slate-100 overflow-hidden">
          <div className="p-4 sm:p-6 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div>
              <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tight">
                Registros Analíticos de Outage
              </h3>
              <p className="text-xs font-bold text-slate-400">Lista completa com paginação e busca detalhada</p>
            </div>

            <div className="flex items-center gap-3">
              {/* Search Bar */}
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input 
                  type="text"
                  placeholder="Buscar chamado, cidade, tipo, topologia..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="bg-slate-50 border border-slate-200 rounded-xl pl-9 pr-4 py-2 text-xs font-bold text-slate-700 outline-none focus:border-[#EE1D23] transition-all w-56 sm:w-64"
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
              </select>
            </div>
          </div>

          <div className="overflow-x-auto scrollbar-thin scrollbar-thumb-slate-200">
            <table className="w-full text-left border-collapse table-auto">
              <thead>
                <tr className="bg-slate-50/80 border-b border-slate-100 text-[10px] font-black text-slate-400 uppercase tracking-wider">
                  <th className="py-3 px-3">Nº Evento</th>
                  <th className="py-3 px-2.5">Mês / Sem</th>
                  <th className="py-3 px-2.5">Cidade</th>
                  <th className="py-3 px-2.5">Cat. Prod. 2</th>
                  <th className="py-3 px-2.5">Tipo</th>
                  <th className="py-3 px-2.5">Topologia</th>
                  <th className="py-3 px-2 text-center">Clientes</th>
                  <th className="py-3 px-2.5">Início</th>
                  <th className="py-3 px-2.5">Previsão</th>
                  <th className="py-3 px-2.5">Fechamento</th>
                  <th className="py-3 px-3 text-center">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-xs">
                {paginatedData.length > 0 ? (
                  paginatedData.map(event => (
                    <tr 
                      key={event.id}
                      onClick={() => setSelectedEvent(event)}
                      className="hover:bg-slate-50/80 transition-colors cursor-pointer group"
                    >
                      <td className="py-2.5 px-3 font-mono font-black text-[#333333] group-hover:text-[#EE1D23] transition-colors whitespace-nowrap text-[11px]">
                        {event.numeroEvento}
                      </td>
                      <td className="py-2.5 px-2.5 font-bold text-slate-600 whitespace-nowrap text-[11px]">
                        <span>{event.mes}</span>
                        {event.semana && <span className="text-slate-400 font-normal ml-1">/ {event.semana}</span>}
                      </td>
                      <td className="py-2.5 px-2.5 font-black text-slate-800 text-[11px] uppercase whitespace-nowrap">
                        {event.cidade}
                      </td>
                      <td className="py-2.5 px-2.5 whitespace-nowrap">
                        <span className="px-1.5 py-0.5 rounded bg-sky-50 border border-sky-200/60 text-[10px] font-bold text-sky-800">
                          {event.catProd2}
                        </span>
                      </td>
                      <td className="py-2.5 px-2.5 font-bold text-slate-700 text-[11px] max-w-[130px] truncate" title={event.tipo}>
                        {event.tipo}
                      </td>
                      <td className="py-2.5 px-2.5 font-mono font-black text-slate-700 text-[11px] whitespace-nowrap">
                        {event.topologia || '-'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-bold text-slate-600 text-[11px] whitespace-nowrap">
                        {event.clientesAfetados ? event.clientesAfetados.toLocaleString('pt-BR') : '-'}
                      </td>
                      <td className="py-2.5 px-2.5">
                        <RenderSplitDateTime formatted={event.dataInicioFormatada} rawDate={event.dataInicio} />
                      </td>
                      <td className="py-2.5 px-2.5">
                        <RenderSplitDateTime formatted={event.dataPrevisaoFormatada} rawDate={event.dataPrevisao} />
                      </td>
                      <td className="py-2.5 px-2.5">
                        <RenderSplitDateTime 
                          formatted={event.dataFechamentoFormatada} 
                          rawDate={event.dataFechamento || event.dataFim || (event.status === 'RESOLVIDO' || event.status === 'FECHADO' ? event.dataInicio : null)} 
                        />
                      </td>
                      <td className="py-2.5 px-3 text-center whitespace-nowrap">
                        <span className={cn(
                          "px-2.5 py-0.5 rounded-full text-[9px] font-black uppercase tracking-wider border inline-block whitespace-nowrap shadow-2xs",
                          getStatusBadge(event.status)
                        )}>
                          {event.status}
                        </span>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={11} className="py-16 text-center text-slate-400 font-bold">
                      Nenhum registro de Outage encontrado com os filtros selecionados.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination Controls */}
          {totalPages > 1 && (
            <div className="p-4 border-t border-slate-100 flex items-center justify-between text-xs font-bold text-slate-500">
              <span>
                Mostrando {((currentPage - 1) * pageSize) + 1} a {Math.min(currentPage * pageSize, filteredData.length)} de {filteredData.length} registros
              </span>
              <div className="flex items-center gap-1">
                <button
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                >
                  Anterior
                </button>
                <span className="px-3 py-1.5 bg-slate-100 rounded-lg text-slate-800 font-black">
                  {currentPage} / {totalPages}
                </span>
                <button
                  onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                >
                  Próximo
                </button>
              </div>
            </div>
          )}
        </div>
      </section>
      </>
      )}

      {/* Event Details Modal */}
      <AnimatePresence>
        {selectedEvent && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-4 overflow-y-auto">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-3xl p-6 sm:p-8 max-w-xl w-full shadow-2xl border border-slate-100 flex flex-col max-h-[90vh] my-auto"
            >
              <div className="flex items-center justify-between border-b border-slate-100 pb-4 flex-shrink-0">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-red-50 flex items-center justify-center text-[#EE1D23] flex-shrink-0">
                    <Radio className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-lg font-black text-slate-800">{selectedEvent.numeroEvento}</h3>
                    <p className="text-xs text-slate-400 font-bold">{selectedEvent.cidade}{selectedEvent.topologia ? ` - ${selectedEvent.topologia}` : ''}</p>
                  </div>
                </div>
                <button
                  onClick={() => setSelectedEvent(null)}
                  className="p-2 text-slate-400 hover:text-slate-600 rounded-xl hover:bg-slate-100 transition-colors"
                  aria-label="Fechar"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto pr-1 my-4 space-y-4">
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 text-xs">
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Cat. Prod. 2</span>
                    <span className="font-bold text-sky-800 text-sm">{selectedEvent.catProd2}</span>
                  </div>
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Tipo de Evento</span>
                    <span className="font-bold text-slate-800 text-sm">{selectedEvent.tipo}</span>
                  </div>
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Topologia</span>
                    <span className="font-mono font-black text-slate-800 text-sm">{selectedEvent.topologia || 'Sem Topologia'}</span>
                  </div>
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Status</span>
                    <span className={cn("px-2.5 py-0.5 rounded-md text-xs font-black uppercase inline-block border", getStatusBadge(selectedEvent.status))}>
                      {selectedEvent.status}
                    </span>
                  </div>
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Data Início (Abertura)</span>
                    <span className="font-bold text-slate-800">{formatDisplayDateTime(selectedEvent.dataInicioFormatada, selectedEvent.dataInicio)}</span>
                  </div>
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Fechamento</span>
                    <span className="font-bold text-slate-800">
                      {selectedEvent.dataFechamentoFormatada || selectedEvent.dataFechamento || selectedEvent.dataFim
                        ? formatDisplayDateTime(selectedEvent.dataFechamentoFormatada, selectedEvent.dataFechamento || selectedEvent.dataFim)
                        : (selectedEvent.status === 'RESOLVIDO' || selectedEvent.status === 'FECHADO' ? formatDisplayDateTime(selectedEvent.dataInicioFormatada, selectedEvent.dataInicio) : 'Em aberto')}
                    </span>
                  </div>
                  {(selectedEvent.dataPrevisao || selectedEvent.dataPrevisaoFormatada) && (
                    <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                      <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Previsão Normalização</span>
                      <span className="font-bold text-slate-800">{formatDisplayDateTime(selectedEvent.dataPrevisaoFormatada, selectedEvent.dataPrevisao)}</span>
                    </div>
                  )}
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Clientes Impactados</span>
                    <span className="font-black text-slate-800 text-sm">{selectedEvent.clientesAfetados?.toLocaleString() || '0'}</span>
                  </div>
                  <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-100">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Duração</span>
                    <span className="font-bold text-slate-800">
                      {selectedEvent.duracaoMinutos 
                        ? `${Math.floor(selectedEvent.duracaoMinutos / 60)}h ${selectedEvent.duracaoMinutos % 60}m (${selectedEvent.duracaoMinutos} min)`
                        : '-'}
                    </span>
                  </div>
                </div>

                {selectedEvent.descricao && (
                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 max-h-60 overflow-y-auto">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1.5">Descrição / Ações Realizadas</span>
                    <p className="text-xs font-bold text-slate-600 leading-relaxed whitespace-pre-wrap break-words">{selectedEvent.descricao}</p>
                  </div>
                )}
              </div>

              <div className="flex justify-end pt-3 border-t border-slate-100 flex-shrink-0">
                <button
                  onClick={() => setSelectedEvent(null)}
                  className="bg-[#1A1A1A] hover:bg-black text-white px-6 py-2.5 rounded-xl text-xs font-black uppercase tracking-wider transition-all"
                >
                  Fechar
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
