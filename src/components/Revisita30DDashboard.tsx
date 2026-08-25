import React, { useState, useMemo, useRef, useEffect } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  FileSpreadsheet, 
  Filter, 
  X, 
  Building2, 
  CheckCircle2, 
  AlertTriangle, 
  RotateCcw, 
  Trash2, 
  Download, 
  Activity, 
  Calendar, 
  Search, 
  ChevronDown, 
  Layers, 
  Briefcase, 
  UserCheck, 
  FileCheck, 
  TrendingUp, 
  ShieldAlert, 
  AlertCircle, 
  Check, 
  Loader2, 
  RefreshCw,
  Users,
  Compass
} from 'lucide-react';
import { MultiFilterSelect } from './MultiFilterSelect';
import { motion, AnimatePresence } from 'motion/react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  Cell, 
  LineChart, 
  Line, 
  LabelList 
} from 'recharts';
import { cn, formatPercent, formatDecimal, formatInteger } from '../lib/utils';
import { getGithubRevisitaUrl, normalizeGithubRawUrl, fetchGithubFileArrayBuffer } from '../lib/githubSync';

export interface Revisita30DRow {
  contrato: string;
  municipio: string;
  tipoOs: string;
  empresa: string;
  unidadeNegocio: string;
  loginTecnico: string;
  codigoBaixa: string;
  qtdRevisitas: number; // 0 = Com Padrão, 1 = Sem Padrão
  dataBaixa: string;
}

// Function to clean baixa codes by stripping trailing/leading 0 ( ), (1), (0), ( ), etc.
export const cleanBaixaCode = (code: string | undefined | null): string => {
  if (!code) return 'NÃO INFORMADO';
  let cleaned = String(code).trim();
  // Strip trailing (1), (0), ( ), ( 0 ), ( 1 ), 0 ( ), etc.
  cleaned = cleaned.replace(/\s*\(\s*\d*\s*\)\s*$/g, '').trim();
  cleaned = cleaned.replace(/^0\s*\(\s*\)\s*/g, '').trim();
  cleaned = cleaned.replace(/\s*0\s*\(\s*\)\s*$/g, '').trim();
  cleaned = cleaned.replace(/\s*\(\s*\)\s*$/g, '').trim();
  return cleaned || 'NÃO INFORMADO';
};

// Generate realistic mock dataset matching exactly the user's Excel pivot table screenshot
const generateMockRevisita30DData = (): Revisita30DRow[] => {
  const result: Revisita30DRow[] = [];
  
  // Data matching the uploaded Excel print (Total: 7941 OS, 6846 Com Padrão (0), 1095 Sem Padrão (1) -> 13.79%):
  const citiesDistribution = [
    { name: 'Ananindeua', zero: 184, one: 26, un: 'UN PA/AP', empresa: 'TELEMONT' },
    { name: 'Belém', zero: 1969, one: 282, un: 'UN PA/AP', empresa: 'TELEMONT' },
    { name: 'Boa Vista', zero: 1, one: 0, un: 'UN AM/RR', empresa: 'I-SISTEMAS' },
    { name: 'Castanhal', zero: 78, one: 23, un: 'UN PA/AP', empresa: 'TELEMONT' },
    { name: 'Caxias', zero: 1, one: 0, un: 'UN MA/PI', empresa: 'I-SISTEMAS' },
    { name: 'Imperatriz', zero: 113, one: 15, un: 'UN MA/PI', empresa: 'I-SISTEMAS' },
    { name: 'Macapá', zero: 80, one: 23, un: 'UN PA/AP', empresa: 'TELEMONT' },
    { name: 'Manaus', zero: 2936, one: 491, un: 'UN AM/RR', empresa: 'TELEMONT' },
    { name: 'Marabá', zero: 19, one: 0, un: 'UN PA/AP', empresa: 'SAMP' },
    { name: 'Paragominas', zero: 19, one: 0, un: 'UN PA/AP', empresa: 'SAMP' },
    { name: 'Parauapebas', zero: 76, one: 19, un: 'UN PA/AP', empresa: 'SAMP' },
    { name: 'Santana', zero: 37, one: 5, un: 'UN PA/AP', empresa: 'TELEMONT' },
    { name: 'São Luís', zero: 1331, one: 211, un: 'UN MA/PI', empresa: 'I-SISTEMAS' },
    { name: 'Timon', zero: 2, one: 0, un: 'UN MA/PI', empresa: 'I-SISTEMAS' }
  ];

  const tiposOS = [
    'REPARO FTTH',
    'INSTALACAO FTTH',
    'MANUTENCAO PREVENTIVA',
    'MUDANCA DE ENDERECO',
    'TROCA DE TECNOLOGIA',
    'RETIRADA DE EQUIPAMENTO'
  ];

  const empresasList = ['TELEMONT', 'I-SISTEMAS', 'SAMP', 'ENGESET', 'COMFICA'];

  const codigosOfensores = [
    'REP01 - SEM CONEXAO DE SINAL ONT',
    'REP04 - FIBRA ROMPIDA NO DROP',
    'REP08 - POTENCIA OPTICA FORA DO PADRAO',
    'INS02 - INSTALACAO COM DEFEITO CONECTOR',
    'EQ01 - DEFEITO NO MODEM WI-FI',
    'EQ05 - FONTE DE ALIMENTACAO QUEIMADA',
    'REP12 - CONECTOR CONECTORIZADO MAL FIXADO',
    'REP15 - ATENUACAO ELEVADA NA CTO',
    'INS05 - CONFIGURACAO INCORRETA DE PPPOE',
    'MUD01 - PROBLEMA DE CABEAMENTO INTERNO',
    'REP19 - FUSAO COM PERDA EXCESSIVA',
    'EQ09 - FALHA DE PORTA LAN/ETHERNET',
    'REP22 - CABO DOBRADO / MICROCURVATURA',
    'REP25 - PORTA NA CTO DANIFICADA',
    'INS08 - IDENTIFICACAO INCORRETA DE FIBRA'
  ];

  const codigosNormais = [
    'OK01 - REPARO CONCLUIDO COM SUCESSO',
    'OK02 - INSTALACAO EFETUADA DENTRO DO PADRAO',
    'OK03 - SINAL NORMALIZADO NA CTO',
    'OK04 - TROCA DE EQUIPAMENTO EFETIVADA',
    'OK05 - CONECTOR REFITO E TESTADO'
  ];

  const tecLogins = [
    'TEC_MAN_0412', 'TEC_MAN_0891', 'TEC_MAN_1102', 'TEC_MAN_0234', 'TEC_MAN_1540',
    'TEC_BEL_0119', 'TEC_BEL_0342', 'TEC_BEL_0781', 'TEC_BEL_0955', 'TEC_BEL_1204',
    'TEC_SLZ_0451', 'TEC_SLZ_0882', 'TEC_SLZ_0129', 'TEC_SLZ_0664', 'TEC_SLZ_0901',
    'TEC_MCP_0311', 'TEC_MCP_0524', 'TEC_IMP_0219', 'TEC_PAR_0188', 'TEC_CAS_0072'
  ];

  let contractSeq = 208500000;

  citiesDistribution.forEach(({ name, zero, one, un, empresa }) => {
    // Generate 0 (Com Padrão - Sem revisita)
    for (let i = 0; i < zero; i++) {
      contractSeq++;
      const day = (i % 28) + 1;
      const tec = tecLogins[(i + name.length) % tecLogins.length];
      const emp = i % 4 === 0 ? empresasList[(i + 1) % empresasList.length] : empresa;
      const tipo = tiposOS[i % tiposOS.length];
      const cod = codigosNormais[i % codigosNormais.length];

      result.push({
        contrato: String(contractSeq),
        municipio: name,
        tipoOs: tipo,
        empresa: emp,
        unidadeNegocio: un,
        loginTecnico: tec,
        codigoBaixa: cod,
        qtdRevisitas: 0,
        dataBaixa: `${String(day).padStart(2, '0')}/08`
      });
    }

    // Generate 1 (Sem Padrão - Com revisita)
    for (let i = 0; i < one; i++) {
      contractSeq++;
      const day = (i % 28) + 1;
      const tec = tecLogins[(i * 3 + name.length) % tecLogins.length];
      const emp = i % 3 === 0 ? empresasList[(i + 2) % empresasList.length] : empresa;
      const tipo = tiposOS[i % 3]; // Revisit often in Reparo / Instalação
      const cod = codigosOfensores[i % codigosOfensores.length];

      result.push({
        contrato: String(contractSeq),
        municipio: name,
        tipoOs: tipo,
        empresa: emp,
        unidadeNegocio: un,
        loginTecnico: tec,
        codigoBaixa: cod,
        qtdRevisitas: 1,
        dataBaixa: `${String(day).padStart(2, '0')}/08`
      });
    }
  });

  return result;
};

const INITIAL_MOCK_REVISITA = generateMockRevisita30DData();

const COLORS_SERIES = ['#EE1D23', '#333333', '#475569', '#10B981', '#F59E0B', '#3B82F6', '#8B5CF6', '#EC4899'];

export default function Revisita30DDashboard() {
  const [data, setData] = useState<Revisita30DRow[]>([]);

  const [isLoading, setIsLoading] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [importProgress, setImportProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);

  // GitHub loader states
  const [githubUrl, setGithubUrl] = useState('');
  const [isGithubLoading, setIsGithubLoading] = useState(false);
  const [showGithubInput, setShowGithubInput] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Filters State
  const [filters, setFilters] = useState<{
    municipio: string[];
    tipoOs: string[];
    empresa: string[];
    unidadeNegocio: string[];
    padraoOs: string[]; // '0', '1'
    startDate: string;
    endDate: string;
  }>({
    municipio: [],
    tipoOs: [],
    empresa: [],
    unidadeNegocio: [],
    padraoOs: [],
    startDate: '',
    endDate: ''
  });

  // Table pagination & search state
  const [tableSearch, setTableSearch] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(25);
  const [sortField, setSortField] = useState<'municipio' | 'qtdRevisitas' | 'empresa' | 'unidadeNegocio' | 'tipoOs' | 'dataBaixa'>('dataBaixa');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');

  // UN chart state
  const [unidadeLimit, setUnidadeLimit] = useState<number>(15);
  const [unidadeMetric, setUnidadeMetric] = useState<'nota' | 'semPadrao'>('nota');
  const [unidadeSort, setUnidadeSort] = useState<'nota' | 'semPadrao' | 'volume'>('nota');
  const [unidadeLayout, setUnidadeLayout] = useState<'horizontal' | 'vertical'>('horizontal');

  // Tech chart state
  const [techLimit, setTechLimit] = useState<number>(15);
  const [techMetric, setTechMetric] = useState<'nota' | 'semPadrao'>('nota');
  const [techSort, setTechSort] = useState<'nota' | 'semPadrao' | 'volume'>('nota');

  // Clear any legacy localStorage data on mount to ensure fresh state
  useEffect(() => {
    try {
      localStorage.removeItem('REVISITA_30D_DATA_V3');
    } catch (err) {
      // ignore
    }
  }, []);

  // Generic Excel processor supporting exact column mappings (ESTR_MUNICIPIO, NM_EMPRESA_NEW, NM_UN_NEW, NR_CONTRATO, QTD_REVISITAS)
  const processExcelData = (ab: ArrayBuffer): boolean => {
    try {
      const wb = XLSX.read(ab, { type: 'array' });
      
      if (!wb.SheetNames || wb.SheetNames.length === 0) {
        setError('Nenhuma planilha encontrada no arquivo.');
        setIsLoading(false);
        return false;
      }

      // 1. Select the best worksheet by scoring headers across all sheets
      let targetSheetName = wb.SheetNames[0];
      let highestScore = -1;
      let headerRowIndex = 0;

      for (const sheetName of wb.SheetNames) {
        const wsTest = wb.Sheets[sheetName];
        if (!wsTest) continue;
        const matrix = XLSX.utils.sheet_to_json(wsTest, { header: 1, defval: '' }) as any[][];
        if (!matrix || matrix.length === 0) continue;

        for (let r = 0; r < Math.min(15, matrix.length); r++) {
          const rowArr = matrix[r];
          if (!Array.isArray(rowArr)) continue;
          
          const normalizedCols = rowArr.map(c => 
            String(c || '')
              .toUpperCase()
              .normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '')
              .replace(/[^A-Z0-9]/g, '')
          );

          let score = 0;
          if (normalizedCols.includes('ESTRMUNICIPIO')) score += 15;
          if (normalizedCols.includes('NMEMPRESANEW') || normalizedCols.includes('ANTNMEMPRESANEW')) score += 15;
          if (normalizedCols.includes('NMUNNEW')) score += 15;
          if (normalizedCols.includes('NRCONTRATO')) score += 15;
          if (normalizedCols.includes('QTDREVISITAS') || normalizedCols.includes('QTDREVISITA')) score += 15;
          if (normalizedCols.includes('WOTPATIVIDADE')) score += 8;
          if (normalizedCols.includes('WOLOGINTEC')) score += 8;
          if (normalizedCols.includes('CDBAIXAORDEMSERVICO')) score += 8;
          if (normalizedCols.includes('DTBAIXAORDEMSERVICO')) score += 8;

          // Also check for secondary variations
          if (normalizedCols.includes('MUNICIPIO') || normalizedCols.includes('CIDADE')) score += 6;
          if (normalizedCols.includes('EMPRESA') || normalizedCols.includes('PARCEIRO')) score += 6;
          if (normalizedCols.includes('CONTRATO') || normalizedCols.includes('OS')) score += 6;
          if (normalizedCols.includes('REVISITAS') || normalizedCols.includes('REVISITA')) score += 6;

          if (score > highestScore) {
            highestScore = score;
            targetSheetName = sheetName;
            headerRowIndex = r;
          }
        }
      }

      const ws = wb.Sheets[targetSheetName];
      const rawMatrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];

      if (!rawMatrix || rawMatrix.length <= headerRowIndex) {
        setError(`A planilha "${targetSheetName}" não contém dados válidos.`);
        setIsLoading(false);
        return false;
      }

      // Build header map from detected header row
      const headerRow = rawMatrix[headerRowIndex] as any[];
      const colMap: Record<string, number> = {};

      headerRow.forEach((colName, idx) => {
        if (colName !== undefined && colName !== null) {
          const rawStr = String(colName).trim();
          const cleanKey = rawStr
            .toUpperCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^A-Z0-9]/g, '');
          if (cleanKey) {
            colMap[cleanKey] = idx;
          }
        }
      });

      const findColIdx = (candidateList: string[]): number => {
        for (const cand of candidateList) {
          const cleanCand = cand
            .toUpperCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^A-Z0-9]/g, '');
          
          // Exact match
          if (colMap[cleanCand] !== undefined) {
            return colMap[cleanCand];
          }
        }

        // Secondary partial search
        for (const cand of candidateList) {
          const cleanCand = cand
            .toUpperCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^A-Z0-9]/g, '');
          
          const foundKey = Object.keys(colMap).find(k => k.includes(cleanCand) || cleanCand.includes(k));
          if (foundKey && colMap[foundKey] !== undefined) {
            return colMap[foundKey];
          }
        }

        return -1;
      };

      // Column Indices
      const idxMunicipio = findColIdx(['ESTR_MUNICIPIO', 'ESTRMUNICIPIO', 'ESTR_MUNICÍPIO', 'NM_MUNICIPIO', 'MUNICIPIO', 'MUNICÍPIO', 'CIDADE', 'ESTR_CIDADE']);
      const idxEmpresa = findColIdx(['NM_EMPRESA_NEW', 'NMEMPRESANEW', 'ANT_NM_EMPRESA_NEW', 'ANTNMEMPRESANEW', 'EMPRESA_NEW', 'NM_EMPRESA', 'NMEMPRESA', 'ANT_EMPRESA', 'EMPRESA', 'PARCEIRO', 'CONTRATADA']);
      const idxUn = findColIdx(['NM_UN_NEW', 'NMUNNEW', 'UNIDADE_NEGOCIO', 'UNIDADENEGOCIO', 'UNIDADE_NEGÓCIO', 'NM_UN', 'NMUN', 'UN_NEW', 'REGIONAL', 'UNIDADE']);
      const idxContrato = findColIdx(['NR_CONTRATO', 'NRCONTRATO', 'NUM_CONTRATO', 'NUMCONTRATO', 'NUMERO_CONTRATO', 'CONTRATO', 'NR_OS', 'NROS', 'NUM_OS', 'NUMOS', 'OS', 'WO_NUMBER', 'ID_OS']);
      const idxRevisitas = findColIdx(['QTD_REVISITAS', 'QTDREVISITAS', 'QTD_REVISITA', 'QTDREVISITA', 'QT_REVISITAS', 'REVISITAS', 'REVISITA', 'SEM_PADRAO', 'IS_REVISITA', 'FLAG_REVISITA']);
      const idxTipoOs = findColIdx(['WO_TP_ATIVIDADE', 'WOTPATIVIDADE', 'TP_ATIVIDADE', 'TIPO_OS', 'TIPOOS', 'TIPO_ATIVIDADE', 'WO_TIPO', 'TIPO', 'SERVICO', 'ATIVIDADE']);
      const idxLoginTec = findColIdx(['WO_LOGIN_TEC', 'WOLOGINTEC', 'LOGIN_TEC', 'LOGINTEC', 'LOGIN_TECNICO', 'NM_TECNICO', 'TECNICO', 'LOGIN', 'MATRICULA', 'WO_TEC']);
      const idxBaixa = findColIdx(['CD_BAIXA_ORDEM_SERVICO', 'CDBAIXAORDEMSERVICO', 'CODIGO_BAIXA', 'CODIGOBAIXA', 'CD_BAIXA', 'CDBAIXA', 'DS_BAIXA', 'MOTIVO_BAIXA', 'BAIXA']);
      const idxDataBaixa = findColIdx(['DT_BAIXA_ORDEM_SERVICO', 'DTBAIXAORDEMSERVICO', 'DT_BAIXA', 'DTBAIXA', 'DATA_BAIXA', 'DATA', 'DT_OS', 'FECHAMENTO']);

      // Format city name nicely (e.g. Belém, São Luís, Ananindeua)
      const formatCityName = (raw: string): string => {
        if (!raw) return 'NÃO INFORMADO';
        const str = raw.trim();
        const upper = str.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        
        if (upper === 'ANANINDEUA') return 'Ananindeua';
        if (upper === 'BELEM') return 'Belém';
        if (upper === 'BOA VISTA') return 'Boa Vista';
        if (upper === 'CASTANHAL') return 'Castanhal';
        if (upper === 'CAXIAS') return 'Caxias';
        if (upper === 'IMPERATRIZ') return 'Imperatriz';
        if (upper === 'MACAPA') return 'Macapá';
        if (upper === 'MANAUS') return 'Manaus';
        if (upper === 'MARABA') return 'Marabá';
        if (upper === 'PARAGOMINAS') return 'Paragominas';
        if (upper === 'PARAUAPEBAS') return 'Parauapebas';
        if (upper === 'SANTANA') return 'Santana';
        if (upper === 'SAO LUIS' || upper === 'SAO LUIZ') return 'São Luís';
        if (upper === 'TIMON') return 'Timon';
        
        // Capitalize words
        return str
          .toLowerCase()
          .split(' ')
          .map(w => w.charAt(0).toUpperCase() + w.slice(1))
          .join(' ');
      };

      const parsedData: Revisita30DRow[] = [];

      for (let r = headerRowIndex + 1; r < rawMatrix.length; r++) {
        const rowArr = rawMatrix[r];
        if (!rowArr || rowArr.length === 0) continue;

        // Extract raw fields using resolved indices
        const rawMun = idxMunicipio >= 0 ? String(rowArr[idxMunicipio] || '').trim() : '';
        const rawEmp = idxEmpresa >= 0 ? String(rowArr[idxEmpresa] || '').trim() : '';
        const rawUn = idxUn >= 0 ? String(rowArr[idxUn] || '').trim() : '';
        const rawCtr = idxContrato >= 0 ? String(rowArr[idxContrato] || '').trim() : '';
        const rawRev = idxRevisitas >= 0 ? rowArr[idxRevisitas] : '';
        const rawTipo = idxTipoOs >= 0 ? String(rowArr[idxTipoOs] || '').trim() : '';
        const rawTec = idxLoginTec >= 0 ? String(rowArr[idxLoginTec] || '').trim() : '';
        const rawBaixa = idxBaixa >= 0 ? String(rowArr[idxBaixa] || '').trim() : '';
        const rawData = idxDataBaixa >= 0 ? rowArr[idxDataBaixa] : '';

        // Ignore empty / summary footer lines
        if (!rawMun && !rawCtr && !rawEmp) continue;
        if (rawMun.toUpperCase().includes('TOTAL GERAL') || rawMun.toUpperCase().includes('TOTAL')) continue;

        // Parse QTD_REVISITAS (0 or 1)
        let qtdRevisitas = 0;
        if (rawRev !== undefined && rawRev !== null && rawRev !== '') {
          const num = Number(rawRev);
          if (!isNaN(num)) {
            qtdRevisitas = num >= 1 ? 1 : 0;
          } else {
            const strVal = String(rawRev).toUpperCase().trim();
            if (strVal === '1' || strVal === 'SIM' || strVal === 'S' || strVal.includes('SEM PADRAO') || strVal.includes('REVISITA')) {
              qtdRevisitas = 1;
            } else {
              qtdRevisitas = 0;
            }
          }
        }

        const municipio = formatCityName(rawMun);
        const empresa = rawEmp ? rawEmp.toUpperCase() : 'CLARO / PRÓPRIA';
        const unidadeNegocio = rawUn ? rawUn.toUpperCase() : 'UN NORTE';
        const contrato = rawCtr ? rawCtr : String(208500000 + r);
        const tipoOs = rawTipo ? rawTipo.toUpperCase() : 'REPARO FTTH';
        const loginTecnico = rawTec ? rawTec.toUpperCase() : 'TEC_SEM_LOGIN';
        const codigoBaixa = rawBaixa ? cleanBaixaCode(rawBaixa).toUpperCase() : 'BAIXA CONCLUÍDA';

        // Parse date
        let dataBaixa = `${String(((r - headerRowIndex) % 28) + 1).padStart(2, '0')}/08`;
        if (rawData) {
          if (rawData instanceof Date) {
            const d = String(rawData.getUTCDate()).padStart(2, '0');
            const m = String(rawData.getUTCMonth() + 1).padStart(2, '0');
            dataBaixa = `${d}/${m}`;
          } else if (typeof rawData === 'number') {
            const dateObj = new Date(Math.round((rawData - (25567 + 2)) * 86400 * 1000));
            const d = String(dateObj.getUTCDate()).padStart(2, '0');
            const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
            dataBaixa = `${d}/${m}`;
          } else {
            const strVal = String(rawData).trim();
            const dateMatch = strVal.match(/(\d{1,2})[\/\-](\d{1,2})/);
            if (dateMatch) {
              dataBaixa = `${dateMatch[1].padStart(2, '0')}/${dateMatch[2].padStart(2, '0')}`;
            }
          }
        }

        parsedData.push({
          contrato,
          municipio,
          tipoOs,
          empresa,
          unidadeNegocio,
          loginTecnico,
          codigoBaixa,
          qtdRevisitas,
          dataBaixa
        });
      }

      if (parsedData.length === 0) {
        setError('Não foi possível extrair dados válidos da planilha. Verifique se as colunas ESTR_MUNICIPIO, NM_EMPRESA_NEW, NM_UN_NEW, NR_CONTRATO e QTD_REVISITAS estão presentes.');
        setIsLoading(false);
        setIsImporting(false);
        setImportProgress(0);
        return false;
      }

      setData(parsedData);
      setError(null);
      setIsLoading(false);
      setImportProgress(100);
      setTimeout(() => {
        setIsImporting(false);
      }, 400);
      return true;
    } catch (err: any) {
      console.error('[Revisita30D] Error parsing file:', err);
      setError(err.message || 'Erro ao processar o arquivo Excel.');
      setIsLoading(false);
      setIsImporting(false);
      setImportProgress(0);
      return false;
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsImporting(true);
    setImportProgress(25);
    setIsLoading(true);
    setError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      setImportProgress(65);
      const buffer = evt.target?.result as ArrayBuffer;
      processExcelData(buffer);
    };
    reader.onerror = () => {
      setError('Erro ao ler o arquivo selecionado.');
      setIsLoading(false);
      setIsImporting(false);
      setImportProgress(0);
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop: (acceptedFiles) => {
      if (acceptedFiles.length === 0) return;
      setIsImporting(true);
      setImportProgress(25);
      setIsLoading(true);
      setError(null);
      const reader = new FileReader();
      reader.onload = (evt) => {
        setImportProgress(65);
        const buffer = evt.target?.result as ArrayBuffer;
        processExcelData(buffer);
      };
      reader.onerror = () => {
        setError('Erro ao ler arquivo arrastado.');
        setIsLoading(false);
        setIsImporting(false);
        setImportProgress(0);
      };
      reader.readAsArrayBuffer(acceptedFiles[0]);
    },
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    noClick: true
  });

  const handleGithubLoad = async (customUrl?: string) => {
    const targetUrl = customUrl || githubUrl.trim() || getGithubRevisitaUrl();
    if (!targetUrl) return;
    setIsImporting(true);
    setImportProgress(20);
    setIsGithubLoading(true);
    setError(null);

    try {
      setImportProgress(50);
      const arrayBuffer = await fetchGithubFileArrayBuffer(targetUrl);
      setImportProgress(75);
      const success = processExcelData(arrayBuffer);
      if (success) {
        setShowGithubInput(false);
      }
    } catch (err: any) {
      console.error('[Revisita30D] GitHub load error:', err);
      setError(err.message || 'Erro ao sincronizar com GitHub.');
      setIsImporting(false);
      setImportProgress(0);
    } finally {
      setIsGithubLoading(false);
    }
  };

  const handleRestoreDefault = () => {
    setData(INITIAL_MOCK_REVISITA);
    setFilters({
      municipio: [],
      tipoOs: [],
      empresa: [],
      unidadeNegocio: [],
      padraoOs: [],
      startDate: '',
      endDate: ''
    });
    setError(null);
  };

  const handleClearData = () => {
    setData([]);
    setFilters({
      municipio: [],
      tipoOs: [],
      empresa: [],
      unidadeNegocio: [],
      padraoOs: [],
      startDate: '',
      endDate: ''
    });
    setError(null);
  };

  const handleExportCSV = () => {
    if (filteredData.length === 0) return;
    const exportRows = filteredData.map(r => ({
      'CONTRATO / OS': r.contrato,
      'MUNICÍPIO': r.municipio,
      'UNIDADE DE NEGÓCIO': r.unidadeNegocio,
      'EMPRESA / PARCEIRO': r.empresa,
      'LOGIN TÉCNICO': r.loginTecnico,
      'TIPO OS': r.tipoOs,
      'CÓDIGO DE BAIXA': r.codigoBaixa,
      'QTD REVISITAS': r.qtdRevisitas,
      'PADRÃO': r.qtdRevisitas === 0 ? '0 - COM PADRÃO' : '1 - SEM PADRÃO (REVISITA)',
      'DATA BAIXA': r.dataBaixa
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Revisita 30D');
    XLSX.writeFile(wb, `Revisita_30D_Claro_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Helper to parse date strings for date range comparisons
  const parseRowDateVal = (dateStr: string) => {
    if (!dateStr) return null;
    const str = String(dateStr).trim();
    const ymd = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (ymd) {
      return new Date(parseInt(ymd[1]), parseInt(ymd[2]) - 1, parseInt(ymd[3]));
    }
    const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?/);
    if (dmy) {
      const day = parseInt(dmy[1]);
      const month = parseInt(dmy[2]) - 1;
      const year = dmy[3] ? parseInt(dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3]) : 2026;
      return new Date(year, month, day);
    }
    return null;
  };

  // Helper to check if a row matches start/end date filters
  const matchesDateRange = (dateStr: string) => {
    if (!filters.startDate && !filters.endDate) return true;
    const rowDate = parseRowDateVal(dateStr);
    if (!rowDate) return true;
    rowDate.setHours(0, 0, 0, 0);

    if (filters.startDate) {
      const [y, m, d] = filters.startDate.split('-').map(Number);
      const start = new Date(y, m - 1, d);
      start.setHours(0, 0, 0, 0);
      if (rowDate < start) return false;
    }
    if (filters.endDate) {
      const [y, m, d] = filters.endDate.split('-').map(Number);
      const end = new Date(y, m - 1, d);
      end.setHours(0, 0, 0, 0);
      if (rowDate > end) return false;
    }
    return true;
  };

  // Distinct filter options (cascading based on other active filters)
  const filterOptions = useMemo(() => {
    // 1. Municipios available considering other active filters
    const dataForMunicipios = data.filter(row => {
      if (filters.tipoOs.length > 0 && !filters.tipoOs.some(t => t.trim().toLowerCase() === (row.tipoOs || '').trim().toLowerCase())) return false;
      if (filters.empresa.length > 0 && !filters.empresa.some(e => e.trim().toLowerCase() === (row.empresa || '').trim().toLowerCase())) return false;
      if (filters.unidadeNegocio.length > 0 && !filters.unidadeNegocio.some(u => u.trim().toLowerCase() === (row.unidadeNegocio || '').trim().toLowerCase())) return false;
      if (filters.padraoOs.length > 0 && !filters.padraoOs.includes(String(row.qtdRevisitas))) return false;
      if (!matchesDateRange(row.dataBaixa)) return false;
      return true;
    });
    const municipios = Array.from(new Set(dataForMunicipios.map(d => d.municipio))).filter(Boolean).sort((a, b) => a.localeCompare(b, 'pt-BR'));

    // 2. Tipos OS available considering other active filters
    const dataForTipos = data.filter(row => {
      if (filters.municipio.length > 0 && !filters.municipio.some(m => m.trim().toLowerCase() === (row.municipio || '').trim().toLowerCase())) return false;
      if (filters.empresa.length > 0 && !filters.empresa.some(e => e.trim().toLowerCase() === (row.empresa || '').trim().toLowerCase())) return false;
      if (filters.unidadeNegocio.length > 0 && !filters.unidadeNegocio.some(u => u.trim().toLowerCase() === (row.unidadeNegocio || '').trim().toLowerCase())) return false;
      if (filters.padraoOs.length > 0 && !filters.padraoOs.includes(String(row.qtdRevisitas))) return false;
      if (!matchesDateRange(row.dataBaixa)) return false;
      return true;
    });
    const tiposOs = Array.from(new Set(dataForTipos.map(d => d.tipoOs))).filter(Boolean).sort((a, b) => a.localeCompare(b, 'pt-BR'));

    // 3. Empresas available considering other active filters
    const dataForEmpresas = data.filter(row => {
      if (filters.municipio.length > 0 && !filters.municipio.some(m => m.trim().toLowerCase() === (row.municipio || '').trim().toLowerCase())) return false;
      if (filters.tipoOs.length > 0 && !filters.tipoOs.some(t => t.trim().toLowerCase() === (row.tipoOs || '').trim().toLowerCase())) return false;
      if (filters.unidadeNegocio.length > 0 && !filters.unidadeNegocio.some(u => u.trim().toLowerCase() === (row.unidadeNegocio || '').trim().toLowerCase())) return false;
      if (filters.padraoOs.length > 0 && !filters.padraoOs.includes(String(row.qtdRevisitas))) return false;
      if (!matchesDateRange(row.dataBaixa)) return false;
      return true;
    });
    const empresas = Array.from(new Set(dataForEmpresas.map(d => d.empresa))).filter(Boolean).sort((a, b) => a.localeCompare(b, 'pt-BR'));

    // 4. Unidades available considering other active filters
    const dataForUnidades = data.filter(row => {
      if (filters.municipio.length > 0 && !filters.municipio.some(m => m.trim().toLowerCase() === (row.municipio || '').trim().toLowerCase())) return false;
      if (filters.tipoOs.length > 0 && !filters.tipoOs.some(t => t.trim().toLowerCase() === (row.tipoOs || '').trim().toLowerCase())) return false;
      if (filters.empresa.length > 0 && !filters.empresa.some(e => e.trim().toLowerCase() === (row.empresa || '').trim().toLowerCase())) return false;
      if (filters.padraoOs.length > 0 && !filters.padraoOs.includes(String(row.qtdRevisitas))) return false;
      if (!matchesDateRange(row.dataBaixa)) return false;
      return true;
    });
    const unidades = Array.from(new Set(dataForUnidades.map(d => d.unidadeNegocio))).filter(Boolean).sort((a, b) => a.localeCompare(b, 'pt-BR'));

    const padroes = ['0', '1'];

    return { municipios, tiposOs, empresas, unidades, padroes };
  }, [data, filters]);

  // Filtered dataset
  const filteredData = useMemo(() => {
    return data.filter(row => {
      if (filters.municipio.length > 0 && !filters.municipio.some(m => m.trim().toLowerCase() === (row.municipio || '').trim().toLowerCase())) return false;
      if (filters.tipoOs.length > 0 && !filters.tipoOs.some(t => t.trim().toLowerCase() === (row.tipoOs || '').trim().toLowerCase())) return false;
      if (filters.empresa.length > 0 && !filters.empresa.some(e => e.trim().toLowerCase() === (row.empresa || '').trim().toLowerCase())) return false;
      if (filters.unidadeNegocio.length > 0 && !filters.unidadeNegocio.some(u => u.trim().toLowerCase() === (row.unidadeNegocio || '').trim().toLowerCase())) return false;
      if (filters.padraoOs.length > 0 && !filters.padraoOs.includes(String(row.qtdRevisitas))) return false;
      if (!matchesDateRange(row.dataBaixa)) return false;
      return true;
    });
  }, [data, filters]);

  // Overall KPIs
  const kpis = useMemo(() => {
    const totalOS = filteredData.length;
    const comPadrao = filteredData.filter(d => d.qtdRevisitas === 0).length;
    const semPadrao = filteredData.filter(d => d.qtdRevisitas === 1).length;
    const notaRevisita = totalOS > 0 ? (semPadrao / totalOS) * 100 : 0;

    return {
      totalOS,
      comPadrao,
      semPadrao,
      notaRevisita
    };
  }, [filteredData]);

  // 1. Pivot Table (Cidades x 0 x 1 x Total x %) - ORDEM ALFABÉTICA
  const pivotTableData = useMemo(() => {
    const cityMap: Record<string, { zero: number; one: number }> = {};

    filteredData.forEach(row => {
      const city = row.municipio || 'NÃO INFORMADO';
      if (!cityMap[city]) {
        cityMap[city] = { zero: 0, one: 0 };
      }
      if (row.qtdRevisitas === 0) {
        cityMap[city].zero++;
      } else {
        cityMap[city].one++;
      }
    });

    const rows = Object.entries(cityMap).map(([city, counts]) => {
      const total = counts.zero + counts.one;
      const nota = total > 0 ? (counts.one / total) * 100 : 0;
      return {
        cidade: city,
        zero: counts.zero,
        one: counts.one,
        total,
        nota
      };
    });

    // Ordem alfabética para a tabela dinâmica
    return rows.sort((a, b) => a.cidade.localeCompare(b.cidade, 'pt-BR'));
  }, [filteredData]);

  // 2. REVISITA POR MUNICÍPIO (Chart data) - Classificado da MAIOR NOTA para a MENOR NOTA
  const chartMunicipioData = useMemo(() => {
    return [...pivotTableData]
      .filter(d => d.total > 0)
      .sort((a, b) => {
        if (b.nota !== a.nota) return b.nota - a.nota; // Classificado da maior nota para a menor
        return b.total - a.total;
      })
      .map(d => ({
        cidade: d.cidade,
        volume: d.total,
        semPadrao: d.one,
        comPadrao: d.zero,
        nota: Number(d.nota.toFixed(2))
      }));
  }, [pivotTableData]);

  // 3. EVOLUÇÃO DIÁRIA DA NOTA REVISITA
  const chartEvolucaoDiaria = useMemo(() => {
    const dayMap: Record<string, { total: number; semPadrao: number; comPadrao: number }> = {};

    filteredData.forEach(row => {
      const dt = row.dataBaixa || '01/08';
      if (!dayMap[dt]) {
        dayMap[dt] = { total: 0, semPadrao: 0, comPadrao: 0 };
      }
      dayMap[dt].total++;
      if (row.qtdRevisitas === 1) {
        dayMap[dt].semPadrao++;
      } else {
        dayMap[dt].comPadrao++;
      }
    });

    // Sort days chronologically
    const sortedDays = Object.keys(dayMap).sort((a, b) => {
      const [d1, m1] = a.split('/').map(Number);
      const [d2, m2] = b.split('/').map(Number);
      if (m1 !== m2) return m1 - m2;
      return d1 - d2;
    });

    return sortedDays.map(dt => {
      const item = dayMap[dt];
      const nota = item.total > 0 ? (item.semPadrao / item.total) * 100 : 0;
      return {
        data: dt,
        volume: item.total,
        semPadrao: item.semPadrao,
        comPadrao: item.comPadrao,
        nota: Number(nota.toFixed(2))
      };
    });
  }, [filteredData]);

  // 3.1. EVOLUÇÃO SEMANAL DA NOTA REVISITA (S1, S2, S3, S4, S5) - Igual ao AT1
  const chartEvolucaoSemanal = useMemo(() => {
    const weekMap: Record<string, { total: number; semPadrao: number; comPadrao: number }> = {
      'S1': { total: 0, semPadrao: 0, comPadrao: 0 },
      'S2': { total: 0, semPadrao: 0, comPadrao: 0 },
      'S3': { total: 0, semPadrao: 0, comPadrao: 0 },
      'S4': { total: 0, semPadrao: 0, comPadrao: 0 },
      'S5': { total: 0, semPadrao: 0, comPadrao: 0 }
    };

    filteredData.forEach(row => {
      const dateObj = parseRowDateVal(row.dataBaixa);
      const day = dateObj ? dateObj.getDate() : parseInt((row.dataBaixa || '').split('/')[0]) || 1;
      
      let weekKey = 'S5';
      if (day <= 7) weekKey = 'S1';
      else if (day <= 14) weekKey = 'S2';
      else if (day <= 21) weekKey = 'S3';
      else if (day <= 28) weekKey = 'S4';
      else weekKey = 'S5';

      weekMap[weekKey].total++;
      if (row.qtdRevisitas === 1) {
        weekMap[weekKey].semPadrao++;
      } else {
        weekMap[weekKey].comPadrao++;
      }
    });

    return Object.entries(weekMap)
      .map(([name, val]) => {
        const nota = val.total > 0 ? (val.semPadrao / val.total) * 100 : 0;
        return {
          name,
          volume: val.total,
          semPadrao: val.semPadrao,
          comPadrao: val.comPadrao,
          nota: Number(nota.toFixed(2))
        };
      })
      .filter(w => w.volume > 0 || ['S1', 'S2', 'S3'].includes(w.name));
  }, [filteredData]);

  // 4. REVISITA POR PARCEIRO (ANT_NM_EMPRESA_NEW)
  const chartEmpresaData = useMemo(() => {
    const map: Record<string, { total: number; semPadrao: number; comPadrao: number }> = {};

    filteredData.forEach(row => {
      const emp = row.empresa || 'OUTROS';
      if (!map[emp]) {
        map[emp] = { total: 0, semPadrao: 0, comPadrao: 0 };
      }
      map[emp].total++;
      if (row.qtdRevisitas === 1) {
        map[emp].semPadrao++;
      } else {
        map[emp].comPadrao++;
      }
    });

    return Object.entries(map)
      .map(([empresa, val]) => {
        const nota = val.total > 0 ? (val.semPadrao / val.total) * 100 : 0;
        return {
          empresa,
          volume: val.total,
          semPadrao: val.semPadrao,
          comPadrao: val.comPadrao,
          nota: Number(nota.toFixed(2))
        };
      })
      .sort((a, b) => {
        if (b.nota !== a.nota) return b.nota - a.nota;
        return b.volume - a.volume;
      });
  }, [filteredData]);

  // 5. REVISITA POR UNIDADE DE NEGÓCIO (NM_UN_NEW)
  const chartUnidadeData = useMemo(() => {
    const map: Record<string, { total: number; semPadrao: number; comPadrao: number }> = {};

    filteredData.forEach(row => {
      const un = row.unidadeNegocio || 'UN NORTE';
      if (!map[un]) {
        map[un] = { total: 0, semPadrao: 0, comPadrao: 0 };
      }
      map[un].total++;
      if (row.qtdRevisitas === 1) {
        map[un].semPadrao++;
      } else {
        map[un].comPadrao++;
      }
    });

    return Object.entries(map)
      .map(([unidade, val]) => {
        const nota = val.total > 0 ? (val.semPadrao / val.total) * 100 : 0;
        return {
          unidade,
          volume: val.total,
          semPadrao: val.semPadrao,
          comPadrao: val.comPadrao,
          nota: Number(nota.toFixed(2))
        };
      })
      .sort((a, b) => {
        if (unidadeSort === 'nota') {
          if (b.nota !== a.nota) return b.nota - a.nota;
          return b.semPadrao - a.semPadrao;
        }
        if (unidadeSort === 'volume') {
          return b.volume - a.volume;
        }
        // default: semPadrao (revisitas)
        if (b.semPadrao !== a.semPadrao) return b.semPadrao - a.semPadrao;
        return b.nota - a.nota;
      })
      .slice(0, unidadeLimit === 0 ? undefined : unidadeLimit);
  }, [filteredData, unidadeSort, unidadeLimit]);

  // 6. REVISITA POR TÉCNICO (WO_LOGIN_TEC) - Cálculo exato da Taxa de Revisita (%) e Reincidências
  const chartTecnicoData = useMemo(() => {
    const map: Record<string, { total: number; semPadrao: number; comPadrao: number; empresa: string; un: string }> = {};

    filteredData.forEach(row => {
      const tec = row.loginTecnico || 'TEC_NÃO_IDENTIFICADO';
      if (!map[tec]) {
        map[tec] = { total: 0, semPadrao: 0, comPadrao: 0, empresa: row.empresa, un: row.unidadeNegocio };
      }
      map[tec].total++;
      if (row.qtdRevisitas === 1) {
        map[tec].semPadrao++;
      } else {
        map[tec].comPadrao++;
      }
    });

    return Object.entries(map)
      .map(([tecnico, val]) => {
        // Cálculo da Nota de Revisita do Técnico = (Sem Padrão / Total de OS) * 100
        const nota = val.total > 0 ? (val.semPadrao / val.total) * 100 : 0;
        return {
          tecnico,
          empresa: val.empresa,
          unidade: val.un,
          volume: val.total,
          semPadrao: val.semPadrao,
          comPadrao: val.comPadrao,
          nota: Number(nota.toFixed(2))
        };
      })
      .sort((a, b) => {
        if (techSort === 'nota') {
          if (b.nota !== a.nota) return b.nota - a.nota;
          return b.semPadrao - a.semPadrao;
        }
        if (techSort === 'volume') {
          return b.volume - a.volume;
        }
        // default: semPadrao (revisitas)
        if (b.semPadrao !== a.semPadrao) return b.semPadrao - a.semPadrao;
        return b.nota - a.nota;
      })
      .slice(0, techLimit);
  }, [filteredData, techLimit, techSort]);

  // 7. TOP 15 DE CÓDIGO DE BAIXA OFENSOR NA REVISITA (CD_BAIXA_ORDEM_SERVICO)
  const chartOfensoresData = useMemo(() => {
    const map: Record<string, number> = {};

    // Filter only rows where qtdRevisitas === 1 (sem padrão / revisita)
    const revisitasRows = filteredData.filter(d => d.qtdRevisitas === 1);
    const totalRevisitas = revisitasRows.length;

    revisitasRows.forEach(row => {
      const rawCod = row.codigoBaixa || 'CÓDIGO NÃO INFORMADO';
      const cod = cleanBaixaCode(rawCod);
      map[cod] = (map[cod] || 0) + 1;
    });

    return Object.entries(map)
      .map(([codigo, count]) => {
        const perc = totalRevisitas > 0 ? (count / totalRevisitas) * 100 : 0;
        return {
          codigo,
          shortCode: codigo.length > 38 ? codigo.substring(0, 38) + '...' : codigo,
          volume: count,
          percentual: Number(perc.toFixed(2))
        };
      })
      .sort((a, b) => b.volume - a.volume)
      .slice(0, 15);
  }, [filteredData]);

  // Filtered detailed table with pagination
  const searchedTableData = useMemo(() => {
    let list = filteredData;
    if (tableSearch.trim()) {
      const q = tableSearch.toLowerCase().trim();
      list = list.filter(r => 
        r.contrato.toLowerCase().includes(q) ||
        r.municipio.toLowerCase().includes(q) ||
        r.empresa.toLowerCase().includes(q) ||
        r.unidadeNegocio.toLowerCase().includes(q) ||
        r.loginTecnico.toLowerCase().includes(q) ||
        r.codigoBaixa.toLowerCase().includes(q) ||
        r.tipoOs.toLowerCase().includes(q) ||
        r.dataBaixa.toLowerCase().includes(q)
      );
    }

    return [...list].sort((a, b) => {
      let valA: any = a[sortField];
      let valB: any = b[sortField];
      if (typeof valA === 'string') valA = valA.toLowerCase();
      if (typeof valB === 'string') valB = valB.toLowerCase();

      if (valA < valB) return sortOrder === 'asc' ? -1 : 1;
      if (valA > valB) return sortOrder === 'asc' ? 1 : -1;
      return 0;
    });
  }, [filteredData, tableSearch, sortField, sortOrder]);

  const totalPages = Math.ceil(searchedTableData.length / rowsPerPage) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return searchedTableData.slice(start, start + rowsPerPage);
  }, [searchedTableData, currentPage, rowsPerPage]);

  const activeFiltersCount = useMemo(() => {
    let count = 0;
    if (filters.municipio.length > 0) count++;
    if (filters.tipoOs.length > 0) count++;
    if (filters.empresa.length > 0) count++;
    if (filters.unidadeNegocio.length > 0) count++;
    if (filters.padraoOs.length > 0) count++;
    if (filters.startDate) count++;
    if (filters.endDate) count++;
    return count;
  }, [filters]);

  return (
    <div {...getRootProps()} className="space-y-6">
      <input {...getInputProps()} />

      {/* Loading Modal - Módulo Qualidade Revisita 30D */}
      {isImporting && (
        <div className="fixed inset-0 z-[100] bg-white/80 backdrop-blur-sm flex flex-col items-center justify-center p-6">
          <div className="w-full max-w-md bg-white p-8 rounded-[32px] shadow-2xl border border-slate-100 text-center">
            <div className="w-20 h-20 bg-red-50 rounded-2xl flex items-center justify-center mb-6 mx-auto animate-pulse">
              <RotateCcw className="w-10 h-10 text-[#EE1D23]" />
            </div>
            <h3 className="text-2xl font-black text-[#333333] uppercase italic tracking-tighter mb-2">IMPORTANDO BASE REVISITA 30D</h3>
            <p className="text-slate-500 font-bold mb-8 italic">Lendo contratos, reincidências e evolução...</p>
            
            <div className="w-full h-4 bg-slate-100 rounded-full overflow-hidden mb-4">
              <div 
                className="h-full bg-[#EE1D23] transition-all duration-300 ease-out"
                style={{ width: `${importProgress}%` }}
              />
            </div>
            <div className="flex justify-between items-center px-1">
              <span className="text-xs font-black text-[#EE1D23] uppercase tracking-widest">{importProgress}%</span>
              <span className="text-xs font-bold text-slate-400 uppercase tracking-widest">AGUARDE...</span>
            </div>
            
            <button
              onClick={() => {
                setIsImporting(false);
                setImportProgress(0);
              }}
              className="mt-8 text-xs font-black text-slate-400 hover:text-red-500 uppercase tracking-widest transition-colors cursor-pointer"
            >
              CANCELAR
            </button>
          </div>
        </div>
      )}

      {/* Corporate Claro Header */}
      <div className="bg-white rounded-3xl p-6 sm:p-8 border border-slate-200/80 shadow-xs relative overflow-hidden">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
          <div className="flex items-center gap-4">
            <div className="w-14 h-14 bg-red-50 rounded-2xl flex items-center justify-center text-[#EE1D23] border border-red-100 shrink-0">
              <RotateCcw className="w-7 h-7" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-black uppercase tracking-widest bg-red-50 text-[#EE1D23] px-2 py-0.5 rounded-md">
                  MÓDULO QUALIDADE
                </span>
                <span className="text-slate-400 text-xs font-bold">•</span>
                <span className="text-slate-500 text-xs font-bold">REINCIDÊNCIA TÉCNICA</span>
              </div>
              <h1 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight uppercase italic mt-0.5">
                REVISITA 30D
              </h1>
              <p className="text-xs font-medium text-slate-500 mt-1">
                Controle, Análise e Acompanhamento de Ordens de Serviço Revisitadas (Janela de 30 Dias)
              </p>
            </div>
          </div>

          {/* Action Buttons Toolbar */}
          <div className="flex flex-wrap items-center gap-2.5">
            <input 
              type="file" 
              ref={fileInputRef} 
              onChange={handleFileUpload} 
              accept=".xlsx, .xls" 
              className="hidden" 
            />

            {/* Sincronizar GitHub */}
            <button
              onClick={() => handleGithubLoad(getGithubRevisitaUrl())}
              className="flex items-center gap-2 bg-[#EE1D23] hover:bg-red-600 text-white font-black py-2.5 px-4 rounded-xl transition-all shadow-md shadow-red-500/15 active:scale-95 uppercase italic text-xs cursor-pointer"
              title="Sincronizar planilha REVISITA 30D com o repositório GitHub"
            >
              <Activity className="w-3.5 h-3.5" />
              <span>Sincronizar GitHub</span>
            </button>

            {/* Importar Excel */}
            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-800 font-black py-2.5 px-4 rounded-xl border border-slate-200 transition-all shadow-2xs active:scale-95 uppercase italic text-xs cursor-pointer"
            >
              <Upload className="w-3.5 h-3.5 text-[#EE1D23]" />
              <span>Importar Excel</span>
            </button>

            {/* Exportar */}
            <button
              onClick={handleExportCSV}
              disabled={filteredData.length === 0}
              className="flex items-center gap-2 bg-white hover:bg-slate-50 disabled:opacity-50 text-slate-800 font-black py-2.5 px-4 rounded-xl border border-slate-200 transition-all shadow-2xs active:scale-95 uppercase italic text-xs cursor-pointer"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span>Exportar</span>
            </button>

            {/* Restaurar */}
            <button
              onClick={handleRestoreDefault}
              title="Restaurar base de exemplo"
              className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all active:scale-95 text-xs cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Restaurar</span>
            </button>

            {/* Limpar */}
            {data.length > 0 && (
              <button
                onClick={handleClearData}
                title="Limpar todos os dados carregados"
                className="flex items-center gap-2 bg-red-50 hover:bg-red-100 text-red-600 font-bold py-2.5 px-4 rounded-xl border border-red-100 transition-all active:scale-95 text-xs cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Limpar</span>
              </button>
            )}
          </div>
        </div>

        {/* GitHub Drawer */}
        <AnimatePresence>
          {showGithubInput && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              className="mt-6 pt-6 border-t border-slate-100 overflow-hidden"
            >
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 flex flex-col md:flex-row gap-3">
                <div className="flex-1 relative">
                  <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                  <input
                    type="text"
                    placeholder="Cole o link do arquivo Excel de Revisita no GitHub (ex: https://github.com/usuario/repo/blob/main/revisita30d.xlsx)"
                    value={githubUrl}
                    onChange={(e) => setGithubUrl(e.target.value)}
                    className="w-full bg-white border border-slate-200 rounded-xl pl-10 pr-4 py-2.5 text-xs font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
                  />
                </div>
                <button
                  onClick={() => handleGithubLoad()}
                  disabled={!githubUrl.trim() || isGithubLoading}
                  className="bg-[#EE1D23] hover:bg-red-600 disabled:bg-slate-300 text-white font-black py-2.5 px-6 rounded-xl transition-all shadow-md shadow-red-500/20 active:scale-95 uppercase italic text-xs flex items-center justify-center gap-2 cursor-pointer"
                >
                  {isGithubLoading ? <Loader2 className="w-4 h-4 animate-spin" /> : <Check className="w-4 h-4" />}
                  Sincronizar Arquivo
                </button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Error Alert */}
        {error && (
          <div className="mt-4 p-3.5 bg-red-50 border border-red-200 rounded-2xl flex items-center gap-2 text-red-600 text-xs font-bold">
            <AlertCircle className="w-4 h-4 shrink-0" />
            <span>{error}</span>
          </div>
        )}
      </div>

      {/* Drag Active Overlay */}
      {isDragActive && (
        <div className="p-8 border-2 border-dashed border-[#EE1D23] bg-red-50/70 rounded-3xl text-center">
          <Upload className="w-10 h-10 text-[#EE1D23] mx-auto mb-2 animate-bounce" />
          <p className="text-sm font-black text-slate-900 uppercase italic">Solte o arquivo Excel (.xlsx, .xls) aqui para carregar</p>
        </div>
      )}

      {data.length === 0 ? (
        <div className="bg-white rounded-[32px] border border-slate-200/80 shadow-sm p-12 text-center flex flex-col items-center justify-center my-6">
          <div className="w-14 h-14 bg-red-50 rounded-2xl flex items-center justify-center mb-4">
            <FileSpreadsheet className="w-7 h-7 text-[#EE1D23]" />
          </div>
          <h3 className="text-xl font-black text-slate-800 uppercase italic tracking-tight mb-2">
            Nenhum Dado de Revisita 30D Carregado
          </h3>
          <p className="text-xs text-slate-400 max-w-md mb-6 leading-relaxed font-bold uppercase tracking-wider">
            Sincronize com o GitHub ou importe a planilha Excel (REVISITA_30D.xlsx) para visualizar os indicadores de repetição e retrabalho.
          </p>
          <div className="flex flex-wrap items-center justify-center gap-3">
            <button
              onClick={() => handleGithubLoad(getGithubRevisitaUrl())}
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
      {/* Filters Section */}
      <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs space-y-4">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Filter className="w-4 h-4 text-[#EE1D23]" />
            <h2 className="text-xs font-black uppercase tracking-wider text-slate-900">Filtros Dinâmicos</h2>
            {activeFiltersCount > 0 && (
              <span className="text-[10px] font-black bg-[#EE1D23] text-white px-2 py-0.5 rounded-full">
                {activeFiltersCount} ativo{activeFiltersCount > 1 ? 's' : ''}
              </span>
            )}
          </div>
          {activeFiltersCount > 0 && (
            <button
              onClick={() => setFilters({
                municipio: [],
                tipoOs: [],
                empresa: [],
                unidadeNegocio: [],
                padraoOs: [],
                startDate: '',
                endDate: ''
              })}
              className="text-[11px] font-bold text-red-600 hover:text-red-700 flex items-center gap-1 cursor-pointer transition-colors"
            >
              <X className="w-3 h-3" />
              Limpar Filtros
            </button>
          )}
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-7 gap-3.5 items-end">
          {/* MUNICÍPIO */}
          <MultiFilterSelect
            label="MUNICÍPIO"
            icon={<Compass className="w-3.5 h-3.5" />}
            value={filters.municipio}
            options={filterOptions.municipios}
            onChange={(val) => setFilters(prev => ({ ...prev, municipio: val }))}
            placeholder="Todos os Municípios"
          />

          {/* TIPO OS */}
          <MultiFilterSelect
            label="TIPO OS"
            icon={<FileCheck className="w-3.5 h-3.5" />}
            value={filters.tipoOs}
            options={filterOptions.tiposOs}
            onChange={(val) => setFilters(prev => ({ ...prev, tipoOs: val }))}
            placeholder="Todos os Tipos"
          />

          {/* EMPRESA */}
          <MultiFilterSelect
            label="EMPRESA"
            icon={<Building2 className="w-3.5 h-3.5" />}
            value={filters.empresa}
            options={filterOptions.empresas}
            onChange={(val) => setFilters(prev => ({ ...prev, empresa: val }))}
            placeholder="Todas as Empresas"
          />

          {/* UNIDADE DE NEGÓCIO */}
          <MultiFilterSelect
            label="UNIDADE DE NEGÓCIO"
            icon={<Briefcase className="w-3.5 h-3.5" />}
            value={filters.unidadeNegocio}
            options={filterOptions.unidades}
            onChange={(val) => setFilters(prev => ({ ...prev, unidadeNegocio: val }))}
            placeholder="Todas as UNs"
          />

          {/* PADRÃO OS */}
          <MultiFilterSelect
            label="PADRÃO OS"
            icon={<Layers className="w-3.5 h-3.5" />}
            value={filters.padraoOs}
            options={filterOptions.padroes}
            getOptionLabel={(opt) => opt === '0' ? '0 - Com Padrão (Sem Revisita)' : '1 - Sem Padrão (Revisita)'}
            onChange={(val) => setFilters(prev => ({ ...prev, padraoOs: val }))}
            placeholder="Todos os Padrões"
          />

          {/* INÍCIO */}
          <div className="flex flex-col gap-1.5">
            <label className="text-[10px] font-black uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
              <Calendar className="w-3 h-3 text-[#EE1D23]" />
              <span>INÍCIO</span>
            </label>
            <input 
              type="date" 
              value={filters.startDate}
              onChange={(e) => setFilters(prev => ({ ...prev, startDate: e.target.value }))}
              className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:ring-2 focus:ring-[#EE1D23] focus:border-transparent outline-none transition-all cursor-pointer"
            />
          </div>

          {/* FIM */}
          <div className="flex flex-col gap-1.5">
            <label className="text-[10px] font-black uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
              <Calendar className="w-3 h-3 text-[#EE1D23]" />
              <span>FIM</span>
            </label>
            <input 
              type="date" 
              value={filters.endDate}
              onChange={(e) => setFilters(prev => ({ ...prev, endDate: e.target.value }))}
              className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:ring-2 focus:ring-[#EE1D23] focus:border-transparent outline-none transition-all cursor-pointer"
            />
          </div>
        </div>
      </div>

      {/* KPI Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* VOLUME TOTAL DE OS */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs relative overflow-hidden">
          <div className="flex items-center justify-between">
            <span className="text-[10px] font-black uppercase tracking-widest text-slate-400">VOLUME TOTAL DE OS</span>
            <div className="w-8 h-8 rounded-xl bg-slate-100 flex items-center justify-center text-slate-700">
              <FileSpreadsheet className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-3">
            <p className="text-3xl font-black text-slate-900 tracking-tight">
              {formatInteger(kpis.totalOS)}
            </p>
            <p className="text-[11px] font-bold text-slate-400 mt-1 flex items-center gap-1">
              <span>Ordens analisadas na base</span>
            </p>
          </div>
        </div>

        {/* OS COM PADRÃO (0) */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs relative overflow-hidden">
          <div className="flex items-center justify-between">
            <span className="text-[10px] font-black uppercase tracking-widest text-emerald-600">OS COM PADRÃO (0)</span>
            <div className="w-8 h-8 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-600">
              <CheckCircle2 className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-3">
            <p className="text-3xl font-black text-emerald-600 tracking-tight">
              {formatInteger(kpis.comPadrao)}
            </p>
            <p className="text-[11px] font-bold text-slate-400 mt-1">
              {kpis.totalOS > 0 ? formatPercent((kpis.comPadrao / kpis.totalOS) * 100) : '0,00%'} do volume total
            </p>
          </div>
        </div>

        {/* OS SEM PADRÃO (1 - REVISITA) */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs relative overflow-hidden">
          <div className="flex items-center justify-between">
            <span className="text-[10px] font-black uppercase tracking-widest text-red-600">OS SEM PADRÃO (1)</span>
            <div className="w-8 h-8 rounded-xl bg-red-50 flex items-center justify-center text-[#EE1D23]">
              <AlertTriangle className="w-4 h-4" />
            </div>
          </div>
          <div className="mt-3">
            <p className="text-3xl font-black text-[#EE1D23] tracking-tight">
              {formatInteger(kpis.semPadrao)}
            </p>
            <p className="text-[11px] font-bold text-slate-400 mt-1">
              {kpis.totalOS > 0 ? formatPercent((kpis.semPadrao / kpis.totalOS) * 100) : '0,00%'} com revisita
            </p>
          </div>
        </div>

        {/* NOTA REVISITA */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs relative overflow-hidden">
          <div className="flex items-center justify-between">
            <span className="text-[10px] font-black uppercase tracking-widest text-slate-400">NOTA REVISITA 30D</span>
            <div className={cn(
              "w-8 h-8 rounded-xl flex items-center justify-center font-black text-xs",
              kpis.notaRevisita <= 10 ? "bg-emerald-50 text-emerald-600" :
              kpis.notaRevisita <= 15 ? "bg-amber-50 text-amber-600" : "bg-red-50 text-red-600"
            )}>
              %
            </div>
          </div>
          <div className="mt-3">
            <p className={cn(
              "text-3xl font-black tracking-tight",
              kpis.notaRevisita <= 10 ? "text-emerald-600" :
              kpis.notaRevisita <= 15 ? "text-amber-600" : "text-[#EE1D23]"
            )}>
              {formatPercent(kpis.notaRevisita)}
            </p>
            <p className="text-[11px] font-bold text-slate-400 mt-1">
              Total Sem Padrão / Total Geral
            </p>
          </div>
        </div>
      </div>

      {/* Pivot Table - Matching Excel Screenshot */}
      <div className="bg-white rounded-3xl border border-slate-200/80 shadow-xs overflow-hidden">
        <div className="p-6 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-red-50 text-[#EE1D23] flex items-center justify-center">
              <Layers className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                REVISITA 30D POR CIDADE
              </h3>
              <p className="text-xs text-slate-500 font-medium">
                Linhas: <span className="font-bold text-slate-700">ESTR_MUNICIPIO</span> | Colunas: <span className="font-bold text-slate-700">QTD_REVISITAS (0 e 1)</span> | Valores: <span className="font-bold text-slate-700">Contagem de NR_CONTRATO</span>
              </p>
            </div>
          </div>
          <span className="text-xs font-bold text-slate-500 bg-slate-100 px-3 py-1.5 rounded-xl self-start sm:self-auto">
            {pivotTableData.length} Municípios
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[700px]">
            <thead>
              <tr className="bg-slate-900 text-white text-[11px] font-black uppercase tracking-wider">
                <th className="py-3 px-4 border-b border-slate-800">RÓTULOS DE LINHA (ESTR_MUNICIPIO)</th>
                <th className="py-3 px-4 border-b border-slate-800 text-right bg-emerald-950/40 text-emerald-300">
                  0 (Com Padrão)
                </th>
                <th className="py-3 px-4 border-b border-slate-800 text-right bg-red-950/40 text-red-300">
                  1 (Sem Padrão)
                </th>
                <th className="py-3 px-4 border-b border-slate-800 text-right">Total Geral</th>
                <th className="py-3 px-4 border-b border-slate-800 text-right bg-slate-800 text-amber-300">% (Nota)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs font-semibold text-slate-700">
              {pivotTableData.map((row, idx) => (
                <tr key={row.cidade} className={idx % 2 === 0 ? 'bg-white hover:bg-slate-50/80' : 'bg-slate-50/40 hover:bg-slate-100/60'}>
                  <td className="py-2.5 px-4 font-bold text-slate-900 flex items-center gap-2">
                    <span className="w-2 h-2 rounded-full bg-slate-400 shrink-0" />
                    {row.cidade}
                  </td>
                  <td className="py-2.5 px-4 text-right font-medium text-emerald-700 bg-emerald-50/20">
                    {formatInteger(row.zero)}
                  </td>
                  <td className="py-2.5 px-4 text-right font-bold text-red-600 bg-red-50/20">
                    {formatInteger(row.one)}
                  </td>
                  <td className="py-2.5 px-4 text-right font-bold text-slate-900">
                    {formatInteger(row.total)}
                  </td>
                  <td className="py-2.5 px-4 text-right font-black text-slate-900 bg-slate-50/60">
                    <span className={cn(
                      "px-2 py-0.5 rounded-md text-[11px]",
                      row.nota <= 10 ? "bg-emerald-100/80 text-emerald-800" :
                      row.nota <= 15 ? "bg-amber-100/80 text-amber-800" : "bg-red-100/80 text-red-800"
                    )}>
                      {formatPercent(row.nota)}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="bg-slate-100 text-slate-900 text-xs font-black uppercase border-t-2 border-slate-300">
                <td className="py-3 px-4">Total Geral</td>
                <td className="py-3 px-4 text-right text-emerald-700 font-black">
                  {formatInteger(kpis.comPadrao)}
                </td>
                <td className="py-3 px-4 text-right text-red-600 font-black">
                  {formatInteger(kpis.semPadrao)}
                </td>
                <td className="py-3 px-4 text-right text-slate-900 font-black">
                  {formatInteger(kpis.totalOS)}
                </td>
                <td className="py-3 px-4 text-right font-black text-[#EE1D23] bg-red-50/50">
                  {formatPercent(kpis.notaRevisita)}
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* Gráficos Principais */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* 1. REVISITA POR MUNICÍPIO */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-50 text-[#EE1D23] flex items-center justify-center">
                <Compass className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                  REVISITA POR MUNICÍPIO
                </h3>
                <p className="text-xs text-slate-500 font-medium">Nota de Revisita (%) ordenada da maior para a menor</p>
              </div>
            </div>
            {filters.municipio.length > 0 && (
              <span className="text-[10px] font-black bg-red-50 text-[#EE1D23] border border-red-200 px-2.5 py-1 rounded-full">
                Filtrado: {filters.municipio.join(', ')}
              </span>
            )}
          </div>

          <div className="h-88 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartMunicipioData} margin={{ top: 25, right: 20, left: -10, bottom: 45 }} barCategoryGap="15%">
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
                <XAxis 
                  dataKey="cidade" 
                  angle={-35} 
                  textAnchor="end" 
                  interval={0} 
                  tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                  height={55}
                />
                <YAxis 
                  tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                  tickFormatter={(val) => `${val}%`}
                />
                <Tooltip
                  content={({ active, payload }) => {
                    if (active && payload && payload.length) {
                      const d = payload[0].payload;
                      return (
                        <div className="bg-slate-900 text-white p-3 rounded-xl shadow-xl text-xs space-y-1.5 border border-slate-700">
                          <p className="font-black text-amber-300 border-b border-slate-700 pb-1">Município: {d.cidade}</p>
                          <p className="flex justify-between gap-4">
                            <span className="text-slate-300 font-medium">Nota de Revisita:</span>
                            <span className="font-black text-white">{formatPercent(d.nota)}</span>
                          </p>
                          <p className="flex justify-between gap-4">
                            <span className="text-red-300 font-medium">Sem Padrão (Revisitas - 1):</span>
                            <span className="font-bold text-red-400">{formatInteger(d.semPadrao)} OS</span>
                          </p>
                          <p className="flex justify-between gap-4">
                            <span className="text-emerald-300 font-medium">Com Padrão (0):</span>
                            <span className="font-bold text-emerald-400">{formatInteger(d.comPadrao)} OS</span>
                          </p>
                          <p className="flex justify-between gap-4 border-t border-slate-800 pt-1">
                            <span className="text-slate-400 font-medium">Volume Total:</span>
                            <span className="font-bold text-slate-200">{formatInteger(d.volume)} OS</span>
                          </p>
                        </div>
                      );
                    }
                    return null;
                  }}
                />
                <Bar dataKey="nota" fill="#EE1D23" radius={[6, 6, 0, 0]} maxBarSize={45}>
                  <LabelList 
                    dataKey="nota" 
                    position="top" 
                    formatter={(val: number) => `${val.toFixed(1)}%`} 
                    style={{ fill: '#1E293B', fontSize: 10, fontWeight: 800 }} 
                  />
                  {chartMunicipioData.map((entry, index) => (
                    <Cell 
                      key={`cell-${index}`} 
                      fill={entry.nota <= 10 ? '#10B981' : entry.nota <= 15 ? '#F59E0B' : '#EE1D23'} 
                    />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 2. EVOLUÇÃO DIÁRIA DA NOTA REVISITA */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-50 text-[#EE1D23] flex items-center justify-center">
                <TrendingUp className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                  EVOLUÇÃO DIÁRIA DA NOTA REVISITA
                </h3>
                <p className="text-xs text-slate-500 font-medium">Tendência diária de revisita (%) ao longo do período</p>
              </div>
            </div>
          </div>

          <div className="h-88 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={chartEvolucaoDiaria} margin={{ top: 35, right: 25, left: -10, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
                <XAxis 
                  dataKey="data" 
                  tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                />
                <YAxis 
                  tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                  tickFormatter={(val) => `${val}%`}
                  domain={[0, 'auto']}
                />
                <Tooltip
                  content={({ active, payload }) => {
                    if (active && payload && payload.length) {
                      const d = payload[0].payload;
                      return (
                        <div className="bg-slate-900 text-white p-3 rounded-xl shadow-xl text-xs space-y-1.5 border border-slate-700">
                          <p className="font-black text-amber-300 border-b border-slate-700 pb-1">Data: {d.data}</p>
                          <p className="flex justify-between gap-4">
                            <span className="text-slate-300 font-medium">Nota de Revisita:</span>
                            <span className="font-black text-white">{formatPercent(d.nota)}</span>
                          </p>
                          <p className="flex justify-between gap-4">
                            <span className="text-red-300 font-medium">Sem Padrão (Revisitas - 1):</span>
                            <span className="font-bold text-red-400">{formatInteger(d.semPadrao)} OS</span>
                          </p>
                          <p className="flex justify-between gap-4">
                            <span className="text-emerald-300 font-medium">Com Padrão (0):</span>
                            <span className="font-bold text-emerald-400">{formatInteger(d.comPadrao)} OS</span>
                          </p>
                          <p className="flex justify-between gap-4 border-t border-slate-800 pt-1">
                            <span className="text-slate-400 font-medium">Volume Total:</span>
                            <span className="font-bold text-slate-200">{formatInteger(d.volume)} OS</span>
                          </p>
                        </div>
                      );
                    }
                    return null;
                  }}
                />
                <Line 
                  type="monotone" 
                  dataKey="nota" 
                  stroke="#EE1D23" 
                  strokeWidth={3.5} 
                  dot={{ r: 4.5, fill: '#EE1D23', stroke: '#FFFFFF', strokeWidth: 2 }}
                  activeDot={{ r: 7, fill: '#EE1D23', stroke: '#FFFFFF', strokeWidth: 2 }}
                >
                  <LabelList 
                    dataKey="nota" 
                    position="top" 
                    content={(props: any) => {
                      const { x, y, value, index } = props;
                      if (value === undefined || value === null) return null;
                      const isElevated = index % 2 !== 0;
                      const yOffset = isElevated ? -24 : -12;
                      const targetY = y + yOffset;
                      const formattedVal = `${Number(value).toFixed(1)}%`;
                      const textWidth = formattedVal.length * 6.5 + 8;

                      return (
                        <g key={`daily-lbl-${index}`}>
                          {isElevated && (
                            <line
                              x1={x}
                              y1={y - 4}
                              x2={x}
                              y2={targetY + 6}
                              stroke="#EF4444"
                              strokeWidth={1}
                              strokeDasharray="2 2"
                              opacity={0.5}
                            />
                          )}
                          <rect
                            x={x - textWidth / 2}
                            y={targetY - 9}
                            width={textWidth}
                            height={15}
                            rx={4}
                            fill="#FFFFFF"
                            stroke={value <= 10 ? '#10B981' : value <= 15 ? '#F59E0B' : '#EE1D23'}
                            strokeWidth={1.2}
                            style={{ filter: 'drop-shadow(0 1px 2px rgba(0,0,0,0.06))' }}
                          />
                          <text
                            x={x}
                            y={targetY + 2}
                            textAnchor="middle"
                            fill={value <= 10 ? '#065F46' : value <= 15 ? '#92400E' : '#991B1B'}
                            fontSize={9}
                            fontWeight={900}
                          >
                            {formattedVal}
                          </text>
                        </g>
                      );
                    }}
                  />
                </Line>
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 2.1. EVOLUÇÃO SEMANAL DA NOTA REVISITA (S1, S2, S3, S4, S5) */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-50 text-[#EE1D23] flex items-center justify-center">
                <Calendar className="w-5 h-5 text-[#EE1D23]" />
              </div>
              <div>
                <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                  EVOLUÇÃO SEMANAL DA NOTA (S1, S2, S3, S4, S5)
                </h3>
                <p className="text-xs text-slate-500 font-medium">Nota de Revisita 30D (%) agrupada por semana do mês</p>
              </div>
            </div>
            <div className="flex items-center gap-3">
              <span className="text-[10px] font-bold text-slate-400 bg-slate-50 border border-slate-200/60 px-2.5 py-1 rounded-lg">
                7 dias por semana
              </span>
            </div>
          </div>

          <div className="h-88 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartEvolucaoSemanal} margin={{ top: 25, right: 25, left: -10, bottom: 15 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
                <XAxis 
                  dataKey="name" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#333333', fontSize: 12, fontWeight: 900 }} 
                />
                <YAxis 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                  tickFormatter={(val) => `${val}%`}
                  domain={[0, 'auto']}
                />
                <Tooltip 
                  content={({ active, payload }) => {
                    if (active && payload && payload.length) {
                      const d = payload[0].payload;
                      return (
                        <div className="bg-slate-900 text-white p-3.5 rounded-xl shadow-xl text-xs space-y-1.5 border border-slate-700">
                          <p className="font-black text-amber-300 border-b border-slate-700 pb-1">Semana: {d.name}</p>
                          <p className="flex justify-between gap-4">
                            <span className="text-slate-300 font-medium">Nota de Revisita:</span>
                            <span className="font-black text-white">{formatPercent(d.nota)}</span>
                          </p>
                          <p className="flex justify-between gap-4">
                            <span className="text-red-300 font-medium">Sem Padrão (Revisitas - 1):</span>
                            <span className="font-bold text-red-400">{formatInteger(d.semPadrao)} OS</span>
                          </p>
                          <p className="flex justify-between gap-4">
                            <span className="text-emerald-300 font-medium">Com Padrão (0):</span>
                            <span className="font-bold text-emerald-400">{formatInteger(d.comPadrao)} OS</span>
                          </p>
                          <p className="flex justify-between gap-4 border-t border-slate-800 pt-1">
                            <span className="text-slate-400 font-medium">Volume Total na Semana:</span>
                            <span className="font-bold text-slate-200">{formatInteger(d.volume)} OS</span>
                          </p>
                        </div>
                      );
                    }
                    return null;
                  }}
                />
                <Bar 
                  dataKey="nota" 
                  fill="#EE1D23" 
                  radius={[8, 8, 0, 0]} 
                  barSize={48}
                  minPointSize={2}
                >
                  <LabelList 
                    dataKey="nota" 
                    position="top" 
                    formatter={(val: number) => `${val.toFixed(1)}%`}
                    style={{ fill: '#1E293B', fontSize: 11, fontWeight: 900 }} 
                  />
                  {chartEvolucaoSemanal.map((entry, index) => (
                    <Cell 
                      key={`week-cell-${index}`} 
                      fill={entry.nota <= 10 ? '#10B981' : entry.nota <= 15 ? '#F59E0B' : '#EE1D23'} 
                    />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 3. REVISITA POR PARCEIRO (ANT_NM_EMPRESA_NEW) */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-50 text-[#EE1D23] flex items-center justify-center">
                <Building2 className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                  REVISITA POR PARCEIRO
                </h3>
                <p className="text-xs text-slate-500 font-medium">Nota (%) e Volume por Empresa Parceira (ANT_NM_EMPRESA_NEW)</p>
              </div>
            </div>
          </div>

          <div className="h-88 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartEmpresaData} layout="vertical" margin={{ top: 10, right: 35, left: 20, bottom: 10 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#E2E8F0" />
                <XAxis 
                  type="number" 
                  tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                  tickFormatter={(val) => `${val}%`}
                />
                <YAxis 
                  type="category" 
                  dataKey="empresa" 
                  tick={{ fill: '#1E293B', fontSize: 10, fontWeight: 700 }}
                  width={110}
                />
                <Tooltip
                  content={({ active, payload }) => {
                    if (active && payload && payload.length) {
                      const d = payload[0].payload;
                      return (
                        <div className="bg-slate-900 text-white p-3 rounded-xl shadow-xl text-xs space-y-1.5 border border-slate-700">
                          <p className="font-black text-amber-300 border-b border-slate-700 pb-1">{d.empresa}</p>
                          <p className="flex justify-between gap-4">
                            <span className="text-slate-300 font-medium">Nota de Revisita:</span>
                            <span className="font-black text-white">{formatPercent(d.nota)}</span>
                          </p>
                          <p className="flex justify-between gap-4">
                            <span className="text-red-300 font-medium">Sem Padrão (Revisitas - 1):</span>
                            <span className="font-bold text-red-400">{formatInteger(d.semPadrao)} OS</span>
                          </p>
                          <p className="flex justify-between gap-4">
                            <span className="text-emerald-300 font-medium">Com Padrão (0):</span>
                            <span className="font-bold text-emerald-400">{formatInteger(d.comPadrao)} OS</span>
                          </p>
                          <p className="flex justify-between gap-4 border-t border-slate-800 pt-1">
                            <span className="text-slate-400 font-medium">Volume Total:</span>
                            <span className="font-bold text-slate-200">{formatInteger(d.volume)} OS</span>
                          </p>
                        </div>
                      );
                    }
                    return null;
                  }}
                />
                <Bar dataKey="nota" fill="#EE1D23" radius={[0, 6, 6, 0]} maxBarSize={24}>
                  <LabelList 
                    dataKey="nota" 
                    position="right" 
                    formatter={(val: number) => `${val.toFixed(1)}%`} 
                    style={{ fill: '#1E293B', fontSize: 10, fontWeight: 800 }} 
                  />
                  {chartEmpresaData.map((entry, index) => (
                    <Cell 
                      key={`cell-emp-${index}`} 
                      fill={entry.nota <= 10 ? '#10B981' : entry.nota <= 15 ? '#F59E0B' : '#EE1D23'} 
                    />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 4. REVISITA POR UNIDADE DE NEGÓCIO (NM_UN_NEW) */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-50 text-[#EE1D23] flex items-center justify-center shrink-0">
                <Briefcase className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                  REVISITA POR UNIDADE DE NEGÓCIO
                </h3>
                <p className="text-xs text-slate-500 font-medium">Nota (%) e Volume por UN Regional (NM_UN_NEW)</p>
              </div>
            </div>

            {/* Controles de métrica, ordenação e visualização */}
            <div className="flex flex-wrap items-center gap-2">
              <div className="flex bg-slate-100 p-0.5 rounded-xl border border-slate-200">
                <button
                  type="button"
                  onClick={() => setUnidadeMetric('nota')}
                  className={cn(
                    "px-2.5 py-1 text-[10px] font-black uppercase rounded-lg transition-all cursor-pointer",
                    unidadeMetric === 'nota' 
                      ? "bg-[#EE1D23] text-white shadow-xs" 
                      : "text-slate-600 hover:text-slate-900"
                  )}
                >
                  % Nota
                </button>
                <button
                  type="button"
                  onClick={() => setUnidadeMetric('semPadrao')}
                  className={cn(
                    "px-2.5 py-1 text-[10px] font-black uppercase rounded-lg transition-all cursor-pointer",
                    unidadeMetric === 'semPadrao' 
                      ? "bg-[#EE1D23] text-white shadow-xs" 
                      : "text-slate-600 hover:text-slate-900"
                  )}
                >
                  Qtd (1)
                </button>
              </div>

              <select
                value={unidadeSort}
                onChange={(e) => setUnidadeSort(e.target.value as any)}
                className="text-[11px] font-bold bg-slate-50 border border-slate-200 rounded-xl px-2 py-1 text-slate-700 outline-none cursor-pointer"
                title="Critério de ordenação"
              >
                <option value="nota">Maior Nota (%)</option>
                <option value="semPadrao">Maior Qtd Revisitas</option>
                <option value="volume">Maior Volume Geral</option>
              </select>

              <select
                value={unidadeLimit}
                onChange={(e) => setUnidadeLimit(Number(e.target.value))}
                className="text-[11px] font-bold bg-slate-50 border border-slate-200 rounded-xl px-2 py-1 text-slate-700 outline-none cursor-pointer"
              >
                <option value={10}>Top 10</option>
                <option value={15}>Top 15</option>
                <option value={0}>Todas</option>
              </select>

              <div className="flex bg-slate-100 p-0.5 rounded-xl border border-slate-200">
                <button
                  type="button"
                  onClick={() => setUnidadeLayout('horizontal')}
                  title="Barras Horizontais (Espaçamento otimizado)"
                  className={cn(
                    "px-2 py-1 text-[10px] font-black rounded-lg transition-all cursor-pointer",
                    unidadeLayout === 'horizontal' ? "bg-white text-slate-900 shadow-xs" : "text-slate-500"
                  )}
                >
                  Barras
                </button>
                <button
                  type="button"
                  onClick={() => setUnidadeLayout('vertical')}
                  title="Colunas Verticais"
                  className={cn(
                    "px-2 py-1 text-[10px] font-black rounded-lg transition-all cursor-pointer",
                    unidadeLayout === 'vertical' ? "bg-white text-slate-900 shadow-xs" : "text-slate-500"
                  )}
                >
                  Colunas
                </button>
              </div>
            </div>
          </div>

          <div className="h-88 w-full">
            <ResponsiveContainer width="100%" height="100%">
              {unidadeLayout === 'horizontal' ? (
                <BarChart data={chartUnidadeData} layout="vertical" margin={{ top: 10, right: 35, left: 15, bottom: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#E2E8F0" />
                  <XAxis 
                    type="number" 
                    tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                    tickFormatter={(val) => unidadeMetric === 'nota' ? `${val}%` : `${val}`}
                  />
                  <YAxis 
                    type="category" 
                    dataKey="unidade" 
                    tick={{ fill: '#1E293B', fontSize: 10, fontWeight: 700 }}
                    width={130}
                  />
                  <Tooltip
                    content={({ active, payload }) => {
                      if (active && payload && payload.length) {
                        const d = payload[0].payload;
                        return (
                          <div className="bg-slate-900 text-white p-3 rounded-xl shadow-xl text-xs space-y-1.5 border border-slate-700">
                            <p className="font-black text-amber-300 border-b border-slate-700 pb-1">UN: {d.unidade}</p>
                            <p className="flex justify-between gap-4">
                              <span className="text-slate-300 font-medium">Nota de Revisita:</span>
                              <span className="font-black text-white">{formatPercent(d.nota)}</span>
                            </p>
                            <p className="flex justify-between gap-4">
                              <span className="text-red-300 font-medium">Sem Padrão (Revisitas - 1):</span>
                              <span className="font-bold text-red-400">{formatInteger(d.semPadrao)} OS</span>
                            </p>
                            <p className="flex justify-between gap-4">
                              <span className="text-emerald-300 font-medium">Com Padrão (0):</span>
                              <span className="font-bold text-emerald-400">{formatInteger(d.comPadrao)} OS</span>
                            </p>
                            <p className="flex justify-between gap-4 border-t border-slate-800 pt-1">
                              <span className="text-slate-400 font-medium">Volume Total:</span>
                              <span className="font-bold text-slate-200">{formatInteger(d.volume)} OS</span>
                            </p>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  {unidadeMetric === 'nota' ? (
                    <Bar dataKey="nota" fill="#EE1D23" radius={[0, 6, 6, 0]} maxBarSize={22}>
                      <LabelList 
                        dataKey="nota" 
                        position="right" 
                        formatter={(val: number) => `${val.toFixed(1)}%`} 
                        style={{ fill: '#1E293B', fontSize: 10, fontWeight: 800 }} 
                      />
                      {chartUnidadeData.map((entry, index) => (
                        <Cell 
                          key={`cell-un-${index}`} 
                          fill={entry.nota <= 10 ? '#10B981' : entry.nota <= 15 ? '#F59E0B' : '#EE1D23'} 
                        />
                      ))}
                    </Bar>
                  ) : (
                    <Bar dataKey="semPadrao" fill="#EE1D23" radius={[0, 6, 6, 0]} maxBarSize={22}>
                      <LabelList 
                        dataKey="semPadrao" 
                        position="right" 
                        formatter={(val: number) => `${val}`} 
                        style={{ fill: '#EE1D23', fontSize: 10, fontWeight: 800 }} 
                      />
                    </Bar>
                  )}
                </BarChart>
              ) : (
                <BarChart data={chartUnidadeData} margin={{ top: 25, right: 20, left: -10, bottom: 45 }} barCategoryGap="20%">
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
                  <XAxis 
                    dataKey="unidade" 
                    angle={-40} 
                    textAnchor="end" 
                    interval={0} 
                    tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                    height={55}
                  />
                  <YAxis 
                    tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                    tickFormatter={(val) => unidadeMetric === 'nota' ? `${val}%` : `${val}`}
                  />
                  <Tooltip
                    content={({ active, payload }) => {
                      if (active && payload && payload.length) {
                        const d = payload[0].payload;
                        return (
                          <div className="bg-slate-900 text-white p-3 rounded-xl shadow-xl text-xs space-y-1.5 border border-slate-700">
                            <p className="font-black text-amber-300 border-b border-slate-700 pb-1">UN: {d.unidade}</p>
                            <p className="flex justify-between gap-4">
                              <span className="text-slate-300 font-medium">Nota de Revisita:</span>
                              <span className="font-black text-white">{formatPercent(d.nota)}</span>
                            </p>
                            <p className="flex justify-between gap-4">
                              <span className="text-red-300 font-medium">Sem Padrão (Revisitas - 1):</span>
                              <span className="font-bold text-red-400">{formatInteger(d.semPadrao)} OS</span>
                            </p>
                            <p className="flex justify-between gap-4">
                              <span className="text-emerald-300 font-medium">Com Padrão (0):</span>
                              <span className="font-bold text-emerald-400">{formatInteger(d.comPadrao)} OS</span>
                            </p>
                            <p className="flex justify-between gap-4 border-t border-slate-800 pt-1">
                              <span className="text-slate-400 font-medium">Volume Total:</span>
                              <span className="font-bold text-slate-200">{formatInteger(d.volume)} OS</span>
                            </p>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  {unidadeMetric === 'nota' ? (
                    <Bar dataKey="nota" fill="#333333" radius={[6, 6, 0, 0]} maxBarSize={30}>
                      <LabelList 
                        dataKey="nota" 
                        position="top" 
                        formatter={(val: number) => `${val.toFixed(1)}%`} 
                        style={{ fill: '#1E293B', fontSize: 10, fontWeight: 800 }} 
                      />
                      {chartUnidadeData.map((entry, index) => (
                        <Cell 
                          key={`cell-un-v-${index}`} 
                          fill={entry.nota <= 10 ? '#10B981' : entry.nota <= 15 ? '#F59E0B' : '#EE1D23'} 
                        />
                      ))}
                    </Bar>
                  ) : (
                    <Bar dataKey="semPadrao" fill="#EE1D23" radius={[6, 6, 0, 0]} maxBarSize={30}>
                      <LabelList 
                        dataKey="semPadrao" 
                        position="top" 
                        formatter={(val: number) => `${val}`} 
                        style={{ fill: '#EE1D23', fontSize: 10, fontWeight: 800 }} 
                      />
                    </Bar>
                  )}
                </BarChart>
              )}
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* Gráficos Secundários: TÉCNICO & TOP 15 OFENSORES */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* 5. REVISITA POR TÉCNICO (WO_LOGIN_TEC) */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-50 text-[#EE1D23] flex items-center justify-center shrink-0">
                <Users className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                  REVISITA POR TÉCNICO
                </h3>
                <p className="text-xs text-slate-500 font-medium">Cálculo de Reincidência e Nota (WO_LOGIN_TEC)</p>
              </div>
            </div>

            {/* Metric Switcher & Filters */}
            <div className="flex flex-wrap items-center gap-2">
              <div className="flex bg-slate-100 p-0.5 rounded-xl border border-slate-200">
                <button
                  type="button"
                  onClick={() => setTechMetric('nota')}
                  className={cn(
                    "px-2.5 py-1 text-[10px] font-black uppercase rounded-lg transition-all cursor-pointer",
                    techMetric === 'nota' 
                      ? "bg-[#EE1D23] text-white shadow-xs" 
                      : "text-slate-600 hover:text-slate-900"
                  )}
                >
                  % Nota
                </button>
                <button
                  type="button"
                  onClick={() => setTechMetric('semPadrao')}
                  className={cn(
                    "px-2.5 py-1 text-[10px] font-black uppercase rounded-lg transition-all cursor-pointer",
                    techMetric === 'semPadrao' 
                      ? "bg-[#EE1D23] text-white shadow-xs" 
                      : "text-slate-600 hover:text-slate-900"
                  )}
                >
                  Qtd (1)
                </button>
              </div>

              <select
                value={techSort}
                onChange={(e) => setTechSort(e.target.value as any)}
                className="text-[11px] font-bold bg-slate-50 border border-slate-200 rounded-xl px-2 py-1 text-slate-700 outline-none cursor-pointer"
                title="Critério de ordenação"
              >
                <option value="nota">Maior Nota (%)</option>
                <option value="semPadrao">Maior Qtd Revisitas</option>
                <option value="volume">Maior Volume Geral</option>
              </select>

              <select
                value={techLimit}
                onChange={(e) => setTechLimit(Number(e.target.value))}
                className="text-[11px] font-bold bg-slate-50 border border-slate-200 rounded-xl px-2 py-1 text-slate-700 outline-none cursor-pointer"
              >
                <option value={10}>Top 10</option>
                <option value={15}>Top 15</option>
                <option value={25}>Top 25</option>
              </select>
            </div>
          </div>

          <div className="h-96 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartTecnicoData} layout="vertical" margin={{ top: 10, right: 35, left: 20, bottom: 10 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#E2E8F0" />
                <XAxis 
                  type="number" 
                  tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                  tickFormatter={(val) => techMetric === 'nota' ? `${val}%` : `${val}`}
                />
                <YAxis 
                  type="category" 
                  dataKey="tecnico" 
                  tick={{ fill: '#1E293B', fontSize: 10, fontWeight: 700 }}
                  width={115}
                />
                <Tooltip
                  content={({ active, payload }) => {
                    if (active && payload && payload.length) {
                      const d = payload[0].payload;
                      return (
                        <div className="bg-slate-900 text-white p-3.5 rounded-xl shadow-xl text-xs space-y-1.5 border border-slate-700">
                          <p className="font-black text-amber-300 border-b border-slate-700 pb-1">Técnico: {d.tecnico}</p>
                          <p className="text-slate-300">Empresa: <span className="text-white font-bold">{d.empresa}</span></p>
                          <p className="text-slate-300">UN: <span className="text-white font-bold">{d.unidade}</span></p>
                          <div className="pt-1 border-t border-slate-800 space-y-1">
                            <p className="flex justify-between gap-4">
                              <span className="text-slate-300 font-medium">Nota de Revisita:</span>
                              <span className="font-black text-white">{formatPercent(d.nota)}</span>
                            </p>
                            <p className="flex justify-between gap-4">
                              <span className="text-red-300 font-medium">Sem Padrão (Revisitas - 1):</span>
                              <span className="font-bold text-red-400">{formatInteger(d.semPadrao)} OS</span>
                            </p>
                            <p className="flex justify-between gap-4">
                              <span className="text-emerald-300 font-medium">Com Padrão (0):</span>
                              <span className="font-bold text-emerald-400">{formatInteger(d.comPadrao)} OS</span>
                            </p>
                            <p className="flex justify-between gap-4">
                              <span className="text-slate-400 font-medium">Volume Total de OS:</span>
                              <span className="font-bold text-slate-200">{formatInteger(d.volume)} OS</span>
                            </p>
                          </div>
                        </div>
                      );
                    }
                    return null;
                  }}
                />
                {techMetric === 'nota' ? (
                  <Bar dataKey="nota" fill="#EE1D23" radius={[0, 6, 6, 0]}>
                    <LabelList 
                      dataKey="nota" 
                      position="right" 
                      formatter={(val: number) => `${val.toFixed(1)}%`} 
                      style={{ fill: '#1E293B', fontSize: 10, fontWeight: 800 }} 
                    />
                    {chartTecnicoData.map((entry, index) => (
                      <Cell 
                        key={`cell-tec-${index}`} 
                        fill={entry.nota <= 10 ? '#10B981' : entry.nota <= 15 ? '#F59E0B' : '#EE1D23'} 
                      />
                    ))}
                  </Bar>
                ) : (
                  <Bar dataKey="semPadrao" fill="#EE1D23" radius={[0, 6, 6, 0]}>
                    <LabelList 
                      dataKey="semPadrao" 
                      position="right" 
                      formatter={(val: number) => `${val}`} 
                      style={{ fill: '#EE1D23', fontSize: 10, fontWeight: 800 }} 
                    />
                  </Bar>
                )}
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 6. TOP 15 DE CÓDIGO DE BAIXA OFENSOR NA REVISITA */}
        <div className="bg-white rounded-3xl p-6 border border-slate-200/80 shadow-xs space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-50 text-[#EE1D23] flex items-center justify-center">
                <ShieldAlert className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                  TOP 15 CÓDIGOS DE BAIXA OFENSORES
                </h3>
                <p className="text-xs text-slate-500 font-medium">Motivos de baixa mais frequentes nas OS com revisita (CD_BAIXA_ORDEM_SERVICO)</p>
              </div>
            </div>
          </div>

          <div className="h-96 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartOfensoresData} layout="vertical" margin={{ top: 10, right: 40, left: 10, bottom: 10 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#E2E8F0" />
                <XAxis 
                  type="number" 
                  tick={{ fill: '#64748B', fontSize: 10, fontWeight: 700 }}
                />
                <YAxis 
                  type="category" 
                  dataKey="shortCode" 
                  tick={{ fill: '#1E293B', fontSize: 9, fontWeight: 700 }}
                  width={150}
                />
                <Tooltip
                  content={({ active, payload }) => {
                    if (active && payload && payload.length) {
                      const d = payload[0].payload;
                      return (
                        <div className="bg-slate-900 text-white p-3 rounded-xl shadow-xl text-xs space-y-1 max-w-xs">
                          <p className="font-black text-amber-300">{d.codigo}</p>
                          <p>Volume de Revisitas: <span className="font-black text-red-300">{formatInteger(d.volume)}</span></p>
                          <p>Participação nas Revisitas: <span className="font-bold text-white">{formatPercent(d.percentual)}</span></p>
                        </div>
                      );
                    }
                    return null;
                  }}
                />
                <Bar dataKey="volume" fill="#EE1D23" radius={[0, 6, 6, 0]}>
                  <LabelList 
                    dataKey="volume" 
                    position="right" 
                    formatter={(val: number) => `${val}`} 
                    style={{ fill: '#1E293B', fontSize: 10, fontWeight: 800 }} 
                  />
                  {chartOfensoresData.map((entry, index) => (
                    <Cell 
                      key={`cell-ofensor-${index}`} 
                      fill={COLORS_SERIES[index % COLORS_SERIES.length]} 
                    />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* Tabela Detalhada de Ordens de Serviço */}
      <div className="bg-white rounded-3xl border border-slate-200/80 shadow-xs overflow-hidden space-y-4">
        <div className="p-6 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-slate-100 text-slate-700 flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-black uppercase tracking-wide text-slate-900">
                REGISTROS DETALHADOS DE REVISITA
              </h3>
              <p className="text-xs text-slate-500 font-medium">
                Mostrando {searchedTableData.length} de {filteredData.length} registros filtrados
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
              <input
                type="text"
                placeholder="Pesquisar contrato, técnico, cidade..."
                value={tableSearch}
                onChange={(e) => {
                  setTableSearch(e.target.value);
                  setCurrentPage(1);
                }}
                className="bg-slate-50 border border-slate-200 rounded-xl pl-9 pr-3 py-2 text-xs font-bold text-slate-700 outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all w-60"
              />
            </div>
            <select
              value={rowsPerPage}
              onChange={(e) => {
                setRowsPerPage(Number(e.target.value));
                setCurrentPage(1);
              }}
              className="text-xs font-bold bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-slate-700 outline-none cursor-pointer"
            >
              <option value={10}>10 por pág.</option>
              <option value={25}>25 por pág.</option>
              <option value={50}>50 por pág.</option>
              <option value={100}>100 por pág.</option>
            </select>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[900px]">
            <thead>
              <tr className="bg-slate-900 text-white text-[11px] font-black uppercase tracking-wider">
                <th className="py-3 px-4 border-b border-slate-800">Contrato / OS</th>
                <th className="py-3 px-4 border-b border-slate-800">Município</th>
                <th className="py-3 px-4 border-b border-slate-800">UN</th>
                <th className="py-3 px-4 border-b border-slate-800">Empresa</th>
                <th className="py-3 px-4 border-b border-slate-800">Técnico</th>
                <th className="py-3 px-4 border-b border-slate-800">Tipo OS</th>
                <th className="py-3 px-4 border-b border-slate-800">Código de Baixa</th>
                <th className="py-3 px-4 border-b border-slate-800 text-center">Data</th>
                <th className="py-3 px-4 border-b border-slate-800 text-center">Padrão OS</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs font-semibold text-slate-700">
              {paginatedData.map((row, idx) => (
                <tr key={`${row.contrato}-${idx}`} className="hover:bg-slate-50 transition-colors">
                  <td className="py-2.5 px-4 font-mono font-bold text-slate-900">
                    {row.contrato}
                  </td>
                  <td className="py-2.5 px-4 font-bold text-slate-800">
                    {row.municipio}
                  </td>
                  <td className="py-2.5 px-4 text-slate-600">
                    {row.unidadeNegocio}
                  </td>
                  <td className="py-2.5 px-4 text-slate-600">
                    {row.empresa}
                  </td>
                  <td className="py-2.5 px-4 font-mono text-slate-700">
                    {row.loginTecnico}
                  </td>
                  <td className="py-2.5 px-4 text-slate-600">
                    {row.tipoOs}
                  </td>
                  <td className="py-2.5 px-4 text-slate-500 max-w-xs truncate" title={row.codigoBaixa}>
                    {row.codigoBaixa}
                  </td>
                  <td className="py-2.5 px-4 text-center font-mono text-slate-500">
                    {row.dataBaixa}
                  </td>
                  <td className="py-2.5 px-4 text-center">
                    {row.qtdRevisitas === 0 ? (
                      <span className="inline-flex items-center gap-1 bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded-full text-[10px] font-black">
                        <CheckCircle2 className="w-3 h-3" />
                        0 - COM PADRÃO
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 bg-red-50 text-red-600 border border-red-200 px-2 py-0.5 rounded-full text-[10px] font-black">
                        <AlertTriangle className="w-3 h-3" />
                        1 - SEM PADRÃO
                      </span>
                    )}
                  </td>
                </tr>
              ))}
              {paginatedData.length === 0 && (
                <tr>
                  <td colSpan={9} className="py-8 text-center text-slate-400 font-bold text-xs">
                    Nenhum registro encontrado para os filtros e busca aplicados.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination controls */}
        <div className="p-4 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs font-bold text-slate-500">
          <div>
            Página {currentPage} de {totalPages} ({searchedTableData.length} registros)
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
              disabled={currentPage === 1}
              className="px-3 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
            >
              Anterior
            </button>
            <span className="px-2">{currentPage}</span>
            <button
              onClick={() => setCurrentPage(p => Math.min(p + 1, totalPages))}
              disabled={currentPage === totalPages}
              className="px-3 py-1.5 rounded-xl border border-slate-200 hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
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
