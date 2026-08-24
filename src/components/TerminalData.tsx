import React, { useState, useMemo, useRef, useEffect } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  FileSpreadsheet, 
  Filter, 
  Search, 
  X, 
  BarChart3,
  PieChart as PieChartIcon,
  ChevronDown,
  Check,
  Database,
  MapPin,
  Wrench,
  Cpu,
  Activity,
  AlertCircle,
  Loader2
} from 'lucide-react';
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
  PieChart,
  Pie,
  Legend,
  LabelList
} from 'recharts';
import { cn } from '../lib/utils';

interface TerminalRow {
  NM_MUNICIPIO_BI: string;
  NM_MODELO: string;
  CD_BAIXA: string;
  GRUPO_BAIXA: string;
  NR_CONTRATO: string;
  NM_STATUS_ORDEM_SERVICO: string;
  TERMINAL: string;
  VOLUME: number;
  [key: string]: any;
}

const COLORS = ['#EE1D23', '#333333', '#666666', '#999999', '#CCCCCC', '#FF5555', '#555555'];

export default function TerminalData() {
  const [data, setData] = useState<TerminalRow[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  
  // Filters
  const [filters, setFilters] = useState<{
    cidade: string[];
    modelo: string[];
    status: string[];
    baixa: string[];
    grupoBaixa: string[];
    terminal: string[];
  }>({
    cidade: ['Todos'],
    modelo: ['Todos'],
    status: ['Todos'],
    baixa: ['Todos'],
    grupoBaixa: ['Todos'],
    terminal: ['Todos']
  });

  const [showFilters, setShowFilters] = useState({
    cidade: false,
    modelo: false,
    status: false,
    baixa: false,
    grupoBaixa: false,
    terminal: false
  });

  const filterRefs = {
    cidade: useRef<HTMLDivElement>(null),
    modelo: useRef<HTMLDivElement>(null),
    status: useRef<HTMLDivElement>(null),
    baixa: useRef<HTMLDivElement>(null),
    grupoBaixa: useRef<HTMLDivElement>(null),
    terminal: useRef<HTMLDivElement>(null)
  };

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      Object.entries(filterRefs).forEach(([key, ref]) => {
        if (ref.current && !ref.current.contains(event.target as Node)) {
          setShowFilters(prev => ({ ...prev, [key]: false }));
        }
      });
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const onDrop = async (acceptedFiles: File[]) => {
    setIsLoading(true);
    setError(null);
    try {
      const file = acceptedFiles[0];
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const ab = e.target?.result;
          const wb = XLSX.read(ab, { type: 'array' });
          
          // Try to find "analitico" sheet, otherwise use the first one
          const analiticoSheetName = wb.SheetNames.find(name => 
            name.toLowerCase().includes('analitico') || name.toLowerCase().includes('analítico')
          );
          const wsname = analiticoSheetName || wb.SheetNames[0];
          const ws = wb.Sheets[wsname];
          const jsonData = XLSX.utils.sheet_to_json(ws) as any[];

          // Map columns with flexibility
          const mappedData = jsonData.map(row => {
            const getValue = (keys: string[]) => {
              const rowKeys = Object.keys(row);
              const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
              const normalizedSearchKeys = keys.map(normalize);
              const foundKey = rowKeys.find(rk => normalizedSearchKeys.includes(normalize(rk)));
              return foundKey ? row[foundKey] : '';
            };

            const terminalValue = String(getValue(['TERMINAL', 'Terminal ID', 'ID Terminal', 'CD_TERMINAL', 'NM_TERMINAL', 'NOME_TERMINAL']) || 'N/A').trim();

            return {
              NM_MUNICIPIO_BI: String(getValue(['NM_MUNICIPIO_BI', 'Municipio', 'Cidade', 'NM_MUNICIPIO']) || 'N/A'),
              NM_MODELO: String(getValue(['NM_MODELO', 'Modelo', 'Equipamento', 'Modelo Equipamento']) || terminalValue || 'N/A'),
              CD_BAIXA: String(getValue(['CD_BAIXA', 'Codigo Baixa', 'Baixa', 'CD_BAIXA_OS']) || 'N/A'),
              GRUPO_BAIXA: String(getValue(['GRUPO BAIXA', 'GRUPO_BAIXA', 'GRUPO', 'DS_GRUPO_BAIXA', 'NM_GRUPO_BAIXA', 'GRUPO_ENCERRAMENTO', 'CATEGORIA_BAIXA']) || 'N/A'),
              NR_CONTRATO: String(getValue(['NR_CONTRATO', 'Contrato', 'Numero Contrato', 'CD_CONTRATO']) || 'N/A'),
              NM_STATUS_ORDEM_SERVICO: String(getValue(['NM_STATUS_ORDEM_SERVICO', 'Status', 'Status Ordem', 'Situacao']) || 'N/A'),
              TERMINAL: terminalValue,
              VOLUME: Number(getValue(['VOLUME', 'Quantidade', 'Qtd', 'Volume']) || 1),
              ...row
            };
          });

          setData(mappedData);
          setIsLoading(false);
        } catch (innerErr) {
          console.error('Erro ao processar Excel:', innerErr);
          setError('Erro ao processar os dados do arquivo Excel.');
          setIsLoading(false);
        }
      };
      reader.onerror = () => {
        setError('Erro ao ler o arquivo.');
        setIsLoading(false);
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      setError('Erro ao processar arquivo Excel');
      setIsLoading(false);
    }
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: false
  });

  const filteredData = useMemo(() => {
    return data.filter(item => {
      const matchCidade = filters.cidade.includes('Todos') || filters.cidade.includes(item.NM_MUNICIPIO_BI);
      const matchModelo = filters.modelo.includes('Todos') || filters.modelo.includes(item.NM_MODELO);
      const matchStatus = filters.status.includes('Todos') || filters.status.includes(item.NM_STATUS_ORDEM_SERVICO);
      const matchBaixa = filters.baixa.includes('Todos') || filters.baixa.includes(item.CD_BAIXA);
      const matchGrupo = filters.grupoBaixa.includes('Todos') || filters.grupoBaixa.includes(item.GRUPO_BAIXA);
      const matchTerminal = filters.terminal.includes('Todos') || filters.terminal.includes(item.TERMINAL);
      return matchCidade && matchModelo && matchStatus && matchBaixa && matchGrupo && matchTerminal;
    });
  }, [data, filters]);

  const cityOptions = useMemo(() => ['Todos', ...new Set(data.map(d => d.NM_MUNICIPIO_BI))].sort(), [data]);
  const modeloOptions = useMemo(() => ['Todos', ...new Set(data.map(d => d.NM_MODELO))].sort(), [data]);
  const statusOptions = useMemo(() => ['Todos', ...new Set(data.map(d => d.NM_STATUS_ORDEM_SERVICO))].sort(), [data]);
  const baixaOptions = useMemo(() => ['Todos', ...new Set(data.map(d => d.CD_BAIXA))].sort(), [data]);
  const grupoBaixaOptions = useMemo(() => ['Todos', ...new Set(data.map(d => d.GRUPO_BAIXA))].sort(), [data]);
  const terminalOptions = useMemo(() => ['Todos', ...new Set(data.map(d => d.TERMINAL))].sort(), [data]);

  const toggleFilter = (type: 'cidade' | 'modelo' | 'status' | 'baixa' | 'grupoBaixa' | 'terminal', value: string) => {
    setFilters(prev => {
      const current = prev[type];
      if (value === 'Todos') return { ...prev, [type]: ['Todos'] };
      
      let next = current.filter(v => v !== 'Todos');
      if (next.includes(value)) {
        next = next.filter(v => v !== value);
        if (next.length === 0) next = ['Todos'];
      } else {
        next = [...next, value];
      }
      return { ...prev, [type]: next };
    });
  };

  const setSingleFilter = (type: 'cidade' | 'modelo' | 'status' | 'baixa' | 'grupoBaixa' | 'terminal', value: string) => {
    setFilters(prev => {
      // If already only this one is selected, reset to Todos
      if (prev[type].length === 1 && prev[type][0] === value) {
        return { ...prev, [type]: ['Todos'] };
      }
      return { ...prev, [type]: [value] };
    });
  };

  const clearFilters = () => {
    setFilters({
      cidade: ['Todos'],
      modelo: ['Todos'],
      status: ['Todos'],
      baixa: ['Todos'],
      grupoBaixa: ['Todos'],
      terminal: ['Todos']
    });
  };

  const hasActiveFilters = useMemo(() => {
    return !filters.cidade.includes('Todos') || 
           !filters.modelo.includes('Todos') || 
           !filters.status.includes('Todos') ||
           !filters.baixa.includes('Todos') ||
           !filters.grupoBaixa.includes('Todos') ||
           !filters.terminal.includes('Todos');
  }, [filters]);

  const modeloChartData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => {
      counts[d.NM_MODELO] = (counts[d.NM_MODELO] || 0) + (d.VOLUME || 1);
    });
    return Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 15);
  }, [filteredData]);

  const baixaChartData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => {
      counts[d.CD_BAIXA] = (counts[d.CD_BAIXA] || 0) + (d.VOLUME || 1);
    });
    return Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 15);
  }, [filteredData]);

  const cidadeChartData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => {
      counts[d.NM_MUNICIPIO_BI] = (counts[d.NM_MUNICIPIO_BI] || 0) + (d.VOLUME || 1);
    });
    return Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 15);
  }, [filteredData]);

  const terminalChartData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => {
      if (d.TERMINAL && d.TERMINAL !== 'N/A') {
        counts[d.TERMINAL] = (counts[d.TERMINAL] || 0) + (d.VOLUME || 1);
      }
    });
    return Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 15);
  }, [filteredData]);

  return (
    <div className="max-w-7xl mx-auto pb-20 px-4">
      <div className="bg-white rounded-[32px] shadow-xl border border-slate-100">
        {/* Header */}
        <div className="p-8 border-b border-slate-100 bg-gradient-to-r from-white to-slate-50">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
            <div className="flex items-center gap-4">
              <div className="w-12 h-12 bg-red-50 rounded-2xl flex items-center justify-center text-[#EE1D23]">
                <Cpu className="w-6 h-6" />
              </div>
              <div>
                <h2 className="text-2xl font-black text-[#333333] uppercase italic tracking-tighter">GERADAS POR TERMINAL</h2>
                <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Análise de Equipamentos e Baixas</p>
              </div>
            </div>
            
            {data.length > 0 && (
              <div className="flex items-center gap-3">
                <div className="px-4 py-2 bg-slate-100 rounded-xl border border-slate-200">
                  <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block">Total Registros</span>
                  <span className="text-lg font-black text-slate-700 tracking-tighter">{data.length}</span>
                </div>
                <button 
                  onClick={() => setData([])}
                  className="p-3 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-xl transition-all"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            )}
          </div>
        </div>

        {data.length === 0 ? (
          <div className="p-12">
            <div 
              {...getRootProps()} 
              className={cn(
                "border-4 border-dashed rounded-[40px] p-20 transition-all cursor-pointer flex flex-col items-center justify-center gap-6",
                isDragActive ? "border-[#EE1D23] bg-red-50/50 scale-[0.99]" : "border-slate-100 hover:border-slate-200 hover:bg-slate-50/50",
                isLoading && "opacity-50 pointer-events-none"
              )}
            >
              <input {...getInputProps()} />
              <div className="w-24 h-24 bg-red-50 rounded-[32px] flex items-center justify-center text-[#EE1D23] shadow-inner">
                {isLoading ? (
                  <Loader2 className="w-10 h-10 animate-spin" />
                ) : (
                  <Upload className="w-10 h-10" />
                )}
              </div>
              <div className="text-center">
                <h3 className="text-2xl font-black text-slate-800 uppercase italic tracking-tighter mb-2">
                  {isLoading ? 'Processando...' : 'Importar Dados Terminal'}
                </h3>
                <p className="text-slate-400 font-bold uppercase text-xs tracking-widest">
                  {isLoading ? 'Aguarde enquanto processamos o arquivo' : 'Arraste seu arquivo Excel ou clique para selecionar'}
                </p>
              </div>
              {!isLoading && (
                <div className="flex items-center gap-4 mt-4">
                  <div className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-100 rounded-full shadow-sm">
                    <FileSpreadsheet className="w-4 h-4 text-emerald-500" />
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Suporta .xlsx, .xls</span>
                  </div>
                </div>
              )}
            </div>
            {error && (
              <div className="mt-6 p-4 bg-red-50 border border-red-100 rounded-2xl flex items-center gap-3 text-red-600">
                <AlertCircle className="w-5 h-5" />
                <p className="text-sm font-bold uppercase tracking-tight">{error}</p>
              </div>
            )}
          </div>
        ) : (
          <div className="flex flex-col">
            {/* Filters Bar */}
            <div className="p-6 border-b border-slate-100 bg-white flex flex-wrap items-center gap-4 relative z-[1000]">
              {hasActiveFilters && (
                <button 
                  onClick={clearFilters}
                  className="flex items-center gap-2 bg-red-50 text-[#EE1D23] border border-red-100 rounded-xl px-4 py-2 hover:bg-red-100 transition-all font-black uppercase italic text-[10px] tracking-widest"
                >
                  <X className="w-3 h-3" />
                  Limpar Filtros
                </button>
              )}
              
              {/* Cidade Filter */}
              <div className="relative" ref={filterRefs.cidade}>
                <button 
                  onClick={() => setShowFilters(prev => ({ ...prev, cidade: !prev.cidade }))}
                  className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 hover:bg-slate-100 transition-all"
                >
                  <MapPin className="w-4 h-4 text-slate-400" />
                  <div className="text-left">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block">Cidade</span>
                    <span className="text-xs font-bold text-slate-700">
                      {filters.cidade.includes('Todos') ? 'Todos' : `${filters.cidade.length} selecionados`}
                    </span>
                  </div>
                  <ChevronDown className="w-4 h-4 text-slate-400 ml-2" />
                </button>
                <AnimatePresence>
                  {showFilters.cidade && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="absolute top-full left-0 mt-2 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[1001] p-2 max-h-64 overflow-y-auto"
                    >
                      {cityOptions.map(option => (
                        <button
                          key={option}
                          onClick={() => toggleFilter('cidade', option)}
                          className={cn(
                            "w-full text-left px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center justify-between",
                            filters.cidade.includes(option) ? "bg-[#EE1D23] text-white" : "text-slate-600 hover:bg-slate-50"
                          )}
                        >
                          {option}
                          {filters.cidade.includes(option) && <Check className="w-3 h-3" />}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Modelo Filter */}
              <div className="relative" ref={filterRefs.modelo}>
                <button 
                  onClick={() => setShowFilters(prev => ({ ...prev, modelo: !prev.modelo }))}
                  className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 hover:bg-slate-100 transition-all"
                >
                  <Wrench className="w-4 h-4 text-slate-400" />
                  <div className="text-left">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block">Modelo Equipamento</span>
                    <span className="text-xs font-bold text-slate-700">
                      {filters.modelo.includes('Todos') ? 'Todos' : `${filters.modelo.length} selecionados`}
                    </span>
                  </div>
                  <ChevronDown className="w-4 h-4 text-slate-400 ml-2" />
                </button>
                <AnimatePresence>
                  {showFilters.modelo && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="absolute top-full left-0 mt-2 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[1001] p-2 max-h-64 overflow-y-auto"
                    >
                      {modeloOptions.map(option => (
                        <button
                          key={option}
                          onClick={() => toggleFilter('modelo', option)}
                          className={cn(
                            "w-full text-left px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center justify-between",
                            filters.modelo.includes(option) ? "bg-[#EE1D23] text-white" : "text-slate-600 hover:bg-slate-50"
                          )}
                        >
                          {option}
                          {filters.modelo.includes(option) && <Check className="w-3 h-3" />}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Status Filter */}
              <div className="relative" ref={filterRefs.status}>
                <button 
                  onClick={() => setShowFilters(prev => ({ ...prev, status: !prev.status }))}
                  className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 hover:bg-slate-100 transition-all"
                >
                  <Activity className="w-4 h-4 text-slate-400" />
                  <div className="text-left">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block">Status</span>
                    <span className="text-xs font-bold text-slate-700">
                      {filters.status.includes('Todos') ? 'Todos' : `${filters.status.length} selecionados`}
                    </span>
                  </div>
                  <ChevronDown className="w-4 h-4 text-slate-400 ml-2" />
                </button>
                <AnimatePresence>
                  {showFilters.status && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="absolute top-full left-0 mt-2 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[1001] p-2 max-h-64 overflow-y-auto"
                    >
                      {statusOptions.map(option => (
                        <button
                          key={option}
                          onClick={() => toggleFilter('status', option)}
                          className={cn(
                            "w-full text-left px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center justify-between",
                            filters.status.includes(option) ? "bg-[#EE1D23] text-white" : "text-slate-600 hover:bg-slate-50"
                          )}
                        >
                          {option}
                          {filters.status.includes(option) && <Check className="w-3 h-3" />}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Baixa Filter */}
              <div className="relative" ref={filterRefs.baixa}>
                <button 
                  onClick={() => setShowFilters(prev => ({ ...prev, baixa: !prev.baixa }))}
                  className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 hover:bg-slate-100 transition-all"
                >
                  <Database className="w-4 h-4 text-slate-400" />
                  <div className="text-left">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block">Código Baixa</span>
                    <span className="text-xs font-bold text-slate-700">
                      {filters.baixa.includes('Todos') ? 'Todos' : `${filters.baixa.length} selecionados`}
                    </span>
                  </div>
                  <ChevronDown className="w-4 h-4 text-slate-400 ml-2" />
                </button>
                <AnimatePresence>
                  {showFilters.baixa && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="absolute top-full left-0 mt-2 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[1001] p-2 max-h-64 overflow-y-auto"
                    >
                      {baixaOptions.map(option => (
                        <button
                          key={option}
                          onClick={() => toggleFilter('baixa', option)}
                          className={cn(
                            "w-full text-left px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center justify-between",
                            filters.baixa.includes(option) ? "bg-[#EE1D23] text-white" : "text-slate-600 hover:bg-slate-50"
                          )}
                        >
                          {option}
                          {filters.baixa.includes(option) && <Check className="w-3 h-3" />}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Terminal Filter */}
              <div className="relative" ref={filterRefs.terminal}>
                <button 
                  onClick={() => setShowFilters(prev => ({ ...prev, terminal: !prev.terminal }))}
                  className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 hover:bg-slate-100 transition-all"
                >
                  <Cpu className="w-4 h-4 text-slate-400" />
                  <div className="text-left">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block">Terminal</span>
                    <span className="text-xs font-bold text-slate-700">
                      {filters.terminal.includes('Todos') ? 'Todos' : `${filters.terminal.length} selecionados`}
                    </span>
                  </div>
                  <ChevronDown className="w-4 h-4 text-slate-400 ml-2" />
                </button>
                <AnimatePresence>
                  {showFilters.terminal && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="absolute top-full left-0 mt-2 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[1001] p-2 max-h-64 overflow-y-auto"
                    >
                      {terminalOptions.map(option => (
                        <button
                          key={option}
                          onClick={() => toggleFilter('terminal', option)}
                          className={cn(
                            "w-full text-left px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center justify-between",
                            filters.terminal.includes(option) ? "bg-[#EE1D23] text-white" : "text-slate-600 hover:bg-slate-50"
                          )}
                        >
                          <span className="truncate mr-2">{option}</span>
                          {filters.terminal.includes(option) && <Check className="w-3 h-3 flex-shrink-0" />}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Grupo de Baixa Filter */}
              <div className="relative" ref={filterRefs.grupoBaixa}>
                <button 
                  onClick={() => setShowFilters(prev => ({ ...prev, grupoBaixa: !prev.grupoBaixa }))}
                  className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 hover:bg-slate-100 transition-all"
                >
                  <Filter className="w-4 h-4 text-slate-400" />
                  <div className="text-left">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block">Grupo de Baixa</span>
                    <span className="text-xs font-bold text-slate-700">
                      {filters.grupoBaixa.includes('Todos') ? 'Todos' : `${filters.grupoBaixa.length} selecionados`}
                    </span>
                  </div>
                  <ChevronDown className="w-4 h-4 text-slate-400 ml-2" />
                </button>
                <AnimatePresence>
                  {showFilters.grupoBaixa && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="absolute top-full left-0 mt-2 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[1001] p-2 max-h-64 overflow-y-auto"
                    >
                      {grupoBaixaOptions.map(option => (
                        <button
                          key={option}
                          onClick={() => toggleFilter('grupoBaixa', option)}
                          className={cn(
                            "w-full text-left px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center justify-between",
                            filters.grupoBaixa.includes(option) ? "bg-[#EE1D23] text-white" : "text-slate-600 hover:bg-slate-50"
                          )}
                        >
                          <span className="truncate mr-2">{option}</span>
                          {filters.grupoBaixa.includes(option) && <Check className="w-3 h-3 flex-shrink-0" />}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            </div>

            {/* Charts Section */}
            <div className="p-8 grid grid-cols-1 lg:grid-cols-2 gap-8 bg-slate-50/50">
              {/* Terminal Chart - Full Width */}
              <div className="lg:col-span-2 bg-white p-8 rounded-[32px] border border-slate-100 shadow-sm">
                <div className="flex items-center gap-3 mb-8">
                  <div className="w-12 h-12 bg-red-50 rounded-2xl flex items-center justify-center text-[#EE1D23]">
                    <Database className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="text-lg font-black text-slate-800 uppercase italic tracking-tighter">Volume por Terminal</h3>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Top 15 Terminais por Volume de Atendimento</p>
                  </div>
                </div>
                <div className="h-[600px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={terminalChartData} layout="vertical" margin={{ left: 250, right: 80, top: 10, bottom: 10 }}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                      <XAxis type="number" hide />
                      <YAxis 
                        dataKey="name" 
                        type="category" 
                        width={240} 
                        tick={{ fontSize: 10, fontWeight: 800, fill: '#333333' }}
                        axisLine={true}
                        tickLine={true}
                        interval={0}
                      />
                      <Tooltip 
                        cursor={{ fill: '#f8fafc' }}
                        contentStyle={{ 
                          borderRadius: '16px', 
                          border: 'none', 
                          boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                          fontSize: '12px',
                          fontWeight: 'bold'
                        }}
                        formatter={(value: number) => [`${value} unidades`, 'Volume']}
                      />
                      <Bar 
                        dataKey="value" 
                        radius={[0, 8, 8, 0]}
                        onClick={(data) => data && setSingleFilter('terminal', data.name)}
                        className="cursor-pointer"
                        isAnimationActive={false}
                        barSize={30}
                      >
                        {terminalChartData.map((entry, index) => {
                          const isSelected = filters.terminal.includes(entry.name) && !filters.terminal.includes('Todos');
                          return (
                            <Cell 
                              key={`cell-${index}`} 
                              fill={isSelected ? "#D1191F" : "#EE1D23"} 
                              className="hover:opacity-80 transition-opacity"
                              stroke={isSelected ? "#333" : "none"}
                              strokeWidth={2}
                            />
                          );
                        })}
                        <LabelList dataKey="value" position="right" style={{ fontSize: '11px', fontWeight: 'bold', fill: '#EE1D23' }} offset={10} />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Cidade Chart */}
              <div className="bg-white p-6 rounded-[32px] border border-slate-100 shadow-sm">
                <div className="flex items-center gap-3 mb-6">
                  <div className="w-10 h-10 bg-slate-50 rounded-xl flex items-center justify-center text-slate-600">
                    <MapPin className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-sm font-black text-slate-800 uppercase italic tracking-tighter">Volume por Cidade</h3>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Top 15 Cidades</p>
                  </div>
                </div>
                <div className="h-[300px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={cidadeChartData} layout="vertical" margin={{ left: 100, right: 40 }}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                      <XAxis type="number" hide />
                      <YAxis 
                        dataKey="name" 
                        type="category" 
                        width={100} 
                        tick={{ fontSize: 10, fontWeight: 700, fill: '#64748b' }}
                        axisLine={false}
                        tickLine={false}
                      />
                      <Tooltip 
                        cursor={{ fill: '#f8fafc' }}
                        contentStyle={{ 
                          borderRadius: '16px', 
                          border: 'none', 
                          boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                          fontSize: '12px',
                          fontWeight: 'bold'
                        }}
                        formatter={(value: number) => [`${value} unidades`, 'Volume']}
                      />
                      <Bar 
                        dataKey="value" 
                        radius={[0, 8, 8, 0]}
                        onClick={(data) => data && setSingleFilter('cidade', data.name)}
                        className="cursor-pointer"
                        isAnimationActive={false}
                      >
                        {cidadeChartData.map((entry, index) => {
                          const isSelected = filters.cidade.includes(entry.name) && !filters.cidade.includes('Todos');
                          return (
                            <Cell 
                              key={`cell-${index}`} 
                              fill={isSelected ? "#D1191F" : "#EE1D23"} 
                              className="hover:opacity-80 transition-opacity"
                              stroke={isSelected ? "#333" : "none"}
                              strokeWidth={2}
                            />
                          );
                        })}
                        <LabelList dataKey="value" position="right" style={{ fontSize: '10px', fontWeight: 'bold', fill: '#64748b' }} />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Modelo Chart */}
              <div className="bg-white p-6 rounded-[32px] border border-slate-100 shadow-sm">
                <div className="flex items-center gap-3 mb-6">
                  <div className="w-10 h-10 bg-red-50 rounded-xl flex items-center justify-center text-[#EE1D23]">
                    <BarChart3 className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-sm font-black text-slate-800 uppercase italic tracking-tighter">Modelos de Equipamento</h3>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Volume por Modelo (Top 15)</p>
                  </div>
                </div>
                <div className="h-[300px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={modeloChartData} layout="vertical" margin={{ left: 100, right: 40 }}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                      <XAxis type="number" hide />
                      <YAxis 
                        dataKey="name" 
                        type="category" 
                        width={180} 
                        tick={{ fontSize: 10, fontWeight: 700, fill: '#64748b' }}
                        axisLine={false}
                        tickLine={false}
                      />
                      <Tooltip 
                        cursor={{ fill: '#f8fafc' }}
                        contentStyle={{ 
                          borderRadius: '16px', 
                          border: 'none', 
                          boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                          fontSize: '12px',
                          fontWeight: 'bold'
                        }}
                        formatter={(value: number) => [`${value} unidades`, 'Volume']}
                      />
                      <Bar 
                        dataKey="value" 
                        radius={[0, 8, 8, 0]}
                        onClick={(data) => data && setSingleFilter('modelo', data.name)}
                        className="cursor-pointer"
                        isAnimationActive={false}
                      >
                        {modeloChartData.map((entry, index) => {
                          const isSelected = filters.modelo.includes(entry.name) && !filters.modelo.includes('Todos');
                          return (
                            <Cell 
                              key={`cell-${index}`} 
                              fill={isSelected ? "#D1191F" : "#EE1D23"} 
                              className="hover:opacity-80 transition-opacity"
                              stroke={isSelected ? "#333" : "none"}
                              strokeWidth={2}
                            />
                          );
                        })}
                        <LabelList dataKey="value" position="right" style={{ fontSize: '10px', fontWeight: 'bold', fill: '#64748b' }} />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Baixa Chart */}
              <div className="bg-white p-6 rounded-[32px] border border-slate-100 shadow-sm">
                <div className="flex items-center gap-3 mb-6">
                  <div className="w-10 h-10 bg-slate-50 rounded-xl flex items-center justify-center text-slate-600">
                    <Database className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-sm font-black text-slate-800 uppercase italic tracking-tighter">Códigos de Baixa</h3>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Volume por Baixa (Top 15)</p>
                  </div>
                </div>
                <div className="h-[300px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={baixaChartData} margin={{ top: 20 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis 
                        dataKey="name" 
                        tick={{ fontSize: 10, fontWeight: 700, fill: '#64748b' }}
                        axisLine={false}
                        tickLine={false}
                      />
                      <YAxis hide />
                      <Tooltip 
                        cursor={{ fill: '#f8fafc' }}
                        contentStyle={{ 
                          borderRadius: '16px', 
                          border: 'none', 
                          boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                          fontSize: '12px',
                          fontWeight: 'bold'
                        }}
                        formatter={(value: number) => [`${value} unidades`, 'Volume']}
                      />
                      <Bar 
                        dataKey="value" 
                        radius={[8, 8, 0, 0]}
                        onClick={(data) => data && setSingleFilter('baixa', data.name)}
                        className="cursor-pointer"
                        isAnimationActive={false}
                      >
                        {baixaChartData.map((entry, index) => {
                          const isSelected = filters.baixa.includes(entry.name) && !filters.baixa.includes('Todos');
                          return (
                            <Cell 
                              key={`cell-${index}`} 
                              fill={isSelected ? "#D1191F" : "#EE1D23"} 
                              className="hover:opacity-80 transition-opacity"
                              stroke={isSelected ? "#333" : "none"}
                              strokeWidth={2}
                            />
                          );
                        })}
                        <LabelList dataKey="value" position="top" style={{ fontSize: '10px', fontWeight: 'bold', fill: '#64748b' }} />
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Data Table Preview */}
            <div className="p-8">
              <div className="flex items-center justify-between mb-6">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-slate-50 rounded-xl flex items-center justify-center text-slate-600">
                    <Search className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-sm font-black text-slate-800 uppercase italic tracking-tighter">Visualização de Dados</h3>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Mostrando {filteredData.length} registros filtrados</p>
                  </div>
                </div>
              </div>
              <div className="overflow-x-auto rounded-2xl border border-slate-100">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50">
                      <th className="p-4 text-[10px] font-black text-slate-400 uppercase tracking-widest border-b border-slate-100">Contrato</th>
                      <th className="p-4 text-[10px] font-black text-slate-400 uppercase tracking-widest border-b border-slate-100">Cidade</th>
                      <th className="p-4 text-[10px] font-black text-slate-400 uppercase tracking-widest border-b border-slate-100">Modelo</th>
                      <th className="p-4 text-[10px] font-black text-slate-400 uppercase tracking-widest border-b border-slate-100">Cód. Baixa</th>
                      <th className="p-4 text-[10px] font-black text-slate-400 uppercase tracking-widest border-b border-slate-100">Volume</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.slice(0, 50).map((row, idx) => (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        <td className="p-4 text-xs font-bold text-slate-600 border-b border-slate-50">{row.NR_CONTRATO}</td>
                        <td className="p-4 text-xs font-bold text-slate-600 border-b border-slate-50">{row.NM_MUNICIPIO_BI}</td>
                        <td className="p-4 text-xs font-bold text-slate-600 border-b border-slate-50">{row.NM_MODELO}</td>
                        <td className="p-4 text-xs font-bold text-slate-600 border-b border-slate-50">{row.CD_BAIXA}</td>
                        <td className="p-4 text-xs font-bold text-slate-600 border-b border-slate-50">{row.VOLUME}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {filteredData.length > 50 && (
                  <div className="p-4 text-center bg-slate-50/50">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">E mais {filteredData.length - 50} registros...</p>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
