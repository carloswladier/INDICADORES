import React, { useState, useMemo, useRef, useEffect, Component } from 'react';
import { 
  Filter, 
  BarChart3, 
  CheckCircle2, 
  XCircle, 
  Activity, 
  MapPin, 
  Calendar, 
  Cpu, 
  Calculator,
  ChevronDown,
  Upload,
  FileSpreadsheet,
  AlertCircle,
  Wrench,
  Signal,
  Search,
  ArrowUp,
  Plus,
  Save,
  Trash2 as Trash,
  Edit2,
  Check,
  Clock,
  AlertTriangle,
  Loader2,
  BookOpen,
  Star,
  Map as MapIcon,
  Database,
  Settings,
  Radio,
  RotateCcw,
  ChevronLeft,
  ChevronRight,
  PanelLeftClose,
  PanelLeft
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
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
  LabelList
} from 'recharts';
import * as XLSX from 'xlsx';
import { MOCK_DATA, VisitData, Technology, BaseCidadeData, MOCK_BASE_CIDADE } from './data';
import { cn, formatPercent, formatDecimal } from './lib/utils';
import { logApi, LogEntry as ApiLogEntry } from './services/api';
import { getDbConfig, setDbConfig } from './lib/database';
import { getGithubAt1Url, normalizeGithubRawUrl, fetchGithubFileArrayBuffer } from './lib/githubSync';
import AT5Dashboard from './components/AT5Dashboard';
import LOGDashboard from './components/LOGDashboard';
import OutageDashboard from './components/OutageDashboard';
import Revisita30DDashboard from './components/Revisita30DDashboard';

// Diário de Bordo Types
interface LogEntry {
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

// Diário de Bordo Component
const Logbook = ({ entries, isLoading, error, onRefresh, cities }: { 
  entries: LogEntry[], 
  isLoading: boolean, 
  error: string | null,
  onRefresh: () => void,
  cities: string[]
}) => {
  const [isAdding, setIsAdding] = useState(false);
  const [editingId, setEditingId] = useState<string | number | null>(null);
  const [deletingId, setDeletingId] = useState<string | number | null>(null);
  const [selectedEntry, setSelectedEntry] = useState<LogEntry | null>(null);
  const [dbConfigured, setDbConfigured] = useState(true);
  const [dbError, setDbError] = useState<string | null>(null);
  const [isCheckingDb, setIsCheckingDb] = useState(false);
  const [showDbModal, setShowDbModal] = useState(false);
  const [dbForm, setDbForm] = useState(() => getDbConfig());

  const checkDbStatus = async () => {
    setIsCheckingDb(true);
    try {
      console.log('[DEBUG] checkDbStatus checking Hostinger MySQL...');
      const status = await logApi.checkStatus();
      
      if (!status.configured) {
        setDbConfigured(false);
        setDbError(status.error || 'Banco de dados Hostinger (MySQL) não configurado.');
      } else if (status.status === 'error') {
        setDbConfigured(false);
        setDbError(status.details || status.error || 'Erro ao conectar ao banco de dados Hostinger.');
      } else {
        setDbConfigured(true);
        setDbError(null);
      }
    } catch (err: any) {
      setDbConfigured(false);
      setDbError(err.message || "Erro ao conectar ao banco de dados Hostinger.");
    } finally {
      setIsCheckingDb(false);
    }
  };

  const handleSaveDbConfig = async (e: React.FormEvent) => {
    e.preventDefault();
    setDbConfig(dbForm);
    setShowDbModal(false);
    await checkDbStatus();
    onRefresh();
  };

  useEffect(() => {
    checkDbStatus();
  }, []);

  const [logFilters, setLogFilters] = useState({
    startDate: '',
    endDate: '',
    cidades: ['Todos'] as string[]
  });
  const [showCityFilter, setShowCityFilter] = useState(false);
  const [showFormCityFilter, setShowFormCityFilter] = useState(false);
  const cityFilterRef = useRef<HTMLDivElement>(null);
  const formCityFilterRef = useRef<HTMLDivElement>(null);

  const getLocalDate = () => {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };

  const formatDateDisplay = (dateStr: string) => {
    if (!dateStr) return 'N/A';
    // Split YYYY-MM-DD and create date manually to avoid timezone shift
    const [year, month, day] = dateStr.split('-').map(Number);
    if (!year || !month || !day) return dateStr;
    return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
  };

  const [formData, setFormData] = useState<Omit<LogEntry, 'cidade'> & { cidade: string[] }>({
    data: getLocalDate(),
    cidade: cities.length > 1 ? [cities[1]] : ['São Paulo'],
    numero_chamado: '',
    incidente: '',
    descricao: '',
    status: 'Pendente',
    data_conclusao: null
  });

  // Filtering logic
  const filteredEntries = useMemo(() => {
    return entries.filter(entry => {
      // Date filter
      if (logFilters.startDate) {
        if (entry.data < logFilters.startDate) return false;
      }
      if (logFilters.endDate) {
        if (entry.data > logFilters.endDate) return false;
      }
      
      // City filter
      if (!logFilters.cidades.includes('Todos')) {
        const entryCities = (entry.cidade || '').split(',').map(c => c.trim());
        if (!entryCities.some(c => logFilters.cidades.includes(c))) return false;
      }
      
      return true;
    });
  }, [entries, logFilters]);

  const handleExport = () => {
    const dataToExport = filteredEntries.map(entry => ({
      'Data': formatDateDisplay(entry.data),
      'Cidade': entry.cidade,
      'Nº Chamado': entry.numero_chamado,
      'Incidente': entry.incidente,
      'Descrição': entry.descricao,
      'Status': entry.status,
      'Data Conclusão': entry.data_conclusao ? formatDateDisplay(entry.data_conclusao) : '-'
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Diário de Bordo");
    XLSX.writeFile(workbook, `Diario_de_Bordo_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleCityFilterToggle = (city: string) => {
    setLogFilters(prev => {
      if (city === 'Todos') return { ...prev, cidades: ['Todos'] };
      
      let newCidades = prev.cidades.filter(c => c !== 'Todos');
      if (newCidades.includes(city)) {
        newCidades = newCidades.filter(c => c !== city);
        if (newCidades.length === 0) newCidades = ['Todos'];
      } else {
        newCidades.push(city);
      }
      return { ...prev, cidades: newCidades };
    });
  };

  // Close city filters when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (cityFilterRef.current && !cityFilterRef.current.contains(event.target as Node)) {
        setShowCityFilter(false);
      }
      if (formCityFilterRef.current && !formCityFilterRef.current.contains(event.target as Node)) {
        setShowFormCityFilter(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);
  const [isSaving, setIsSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    setSaveError(null);
    console.log('Iniciando salvamento do registro no banco Hostinger...', formData);

    if (formData.cidade.length === 0) {
      setSaveError('Selecione pelo menos uma cidade.');
      return;
    }
    setIsSaving(true);
    
    try {
      // Prepare data for saving (handle empty strings for dates and join cities)
      // Exclude id and created_at from the data payload
      const { id, created_at, ...rest } = formData as any;
      const dataToSave = {
        ...rest,
        cidade: formData.cidade.join(', '),
        data_conclusao: (formData.data_conclusao === '' || formData.data_conclusao === null) ? null : formData.data_conclusao
      };

      console.log('Dados formatados para o banco Hostinger:', dataToSave);

      if (editingId) {
        console.log('Atualizando registro existente no banco Hostinger:', editingId);
        await logApi.update(editingId, dataToSave as any);
      } else {
        console.log('Inserindo novo registro no banco Hostinger...');
        await logApi.create(dataToSave as any);
      }
      
      console.log('Registro salvo com sucesso!');
      setIsAdding(false);
      setEditingId(null);
      setFormData({
        data: getLocalDate(),
        cidade: cities.length > 1 ? [cities[1]] : ['São Paulo'],
        numero_chamado: '',
        incidente: '',
        descricao: '',
        status: 'Pendente',
        data_conclusao: null
      });
      onRefresh();
      alert('Registro salvo com sucesso no banco de dados da Hostinger!');
    } catch (err: any) {
      console.error('Erro detalhado ao salvar log:', err);
      setSaveError(`Erro ao salvar: ${err.message || 'Verifique a configuração do banco de dados Hostinger.'}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: string | number) => {
    if (!id) return;
    try {
      await logApi.delete(id);
      setDeletingId(null);
      onRefresh();
    } catch (err: any) {
      console.error('Error deleting log:', err);
      alert(`Erro ao excluir registro: ${err.message || 'Erro desconhecido'}`);
    }
  };

  const handleEdit = (entry: LogEntry) => {
    setFormData({
      ...entry,
      cidade: (entry.cidade || '').split(',').map(c => c.trim())
    });
    setEditingId(entry.id || null);
    setIsAdding(true);
  };

  return (
    <div className="max-w-7xl mx-auto pb-20">
      <div className="bg-white rounded-[32px] shadow-xl border border-slate-100 overflow-hidden">
        <div className="p-8 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-6 bg-gradient-to-r from-white to-slate-50">
          <div className="flex flex-wrap items-center gap-4">
            {/* Date Range Filter */}
            <div className="flex items-center gap-2 bg-white border border-slate-200 rounded-xl px-3 py-2 shadow-sm">
              <Calendar className="w-4 h-4 text-slate-400" />
              <input 
                type="date" 
                value={logFilters.startDate}
                onChange={e => setLogFilters(prev => ({ ...prev, startDate: e.target.value }))}
                className="text-xs font-bold focus:outline-none bg-transparent"
              />
              <span className="text-slate-300">|</span>
              <input 
                type="date" 
                value={logFilters.endDate}
                onChange={e => setLogFilters(prev => ({ ...prev, endDate: e.target.value }))}
                className="text-xs font-bold focus:outline-none bg-transparent"
              />
            </div>

            {/* City Multi-select Filter */}
            <div className="relative" ref={cityFilterRef}>
              <button 
                onClick={() => setShowCityFilter(!showCityFilter)}
                className="flex items-center gap-2 bg-white border border-slate-200 rounded-xl px-4 py-2 shadow-sm hover:border-[#EE1D23] transition-all"
              >
                <MapPin className="w-4 h-4 text-slate-400" />
                <span className="text-xs font-bold text-slate-600">
                  {logFilters.cidades.includes('Todos') ? 'Todas as Cidades' : `${logFilters.cidades.length} Cidades`}
                </span>
                <ChevronDown className={cn("w-4 h-4 text-slate-400 transition-transform", showCityFilter && "rotate-180")} />
              </button>

              <AnimatePresence>
                {showCityFilter && (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: 10 }}
                    className="absolute top-full left-0 mt-2 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-50 p-4"
                  >
                    <div className="max-h-60 overflow-y-auto space-y-1">
                      {cities.map(city => (
                        <label key={city} className="flex items-center gap-3 p-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-colors">
                          <input 
                            type="checkbox" 
                            checked={logFilters.cidades.includes(city)}
                            onChange={() => handleCityFilterToggle(city)}
                            className="w-4 h-4 rounded border-slate-300 text-[#EE1D23] focus:ring-[#EE1D23]"
                          />
                          <span className="text-xs font-bold text-slate-600">{city}</span>
                        </label>
                      ))}
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>

            {/* Export Button */}
            <button
              onClick={handleExport}
              className="flex items-center gap-2 bg-emerald-50 text-emerald-600 border border-emerald-100 hover:bg-emerald-100 px-4 py-2 rounded-xl transition-all active:scale-95 text-xs font-black uppercase tracking-widest"
            >
              <FileSpreadsheet className="w-4 h-4" />
              Exportar
            </button>
          </div>

          <div className="flex items-center gap-3">
            {!dbConfigured && (
              <div className={cn(
                "border rounded-2xl p-3.5 flex items-center justify-between gap-3 max-w-xl shadow-sm",
                dbError ? "bg-red-50/90 border-red-200" : "bg-amber-50/90 border-amber-200"
              )}>
                <div className="flex items-center gap-2.5">
                  {dbError ? (
                    <XCircle className="w-5 h-5 text-red-600 shrink-0" />
                  ) : (
                    <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0" />
                  )}
                  <p className={cn(
                    "text-xs font-bold leading-tight",
                    dbError ? "text-red-800" : "text-amber-800"
                  )}>
                    {dbError || "Banco de Dados Hostinger (MySQL) não conectado."}
                  </p>
                </div>
                <button
                  onClick={checkDbStatus}
                  disabled={isCheckingDb}
                  title="Testar Conexão"
                  className="shrink-0 p-2 hover:bg-white/80 rounded-lg transition-colors disabled:opacity-50"
                >
                  <motion.div
                    animate={isCheckingDb ? { rotate: 360 } : {}}
                    transition={isCheckingDb ? { repeat: Infinity, duration: 1, ease: "linear" } : {}}
                  >
                    <Loader2 className={cn("w-4 h-4", dbError ? "text-red-600" : "text-amber-600")} />
                  </motion.div>
                </button>
              </div>
            )}

            <button
              onClick={() => {
                setIsAdding(!isAdding);
                if (!isAdding) {
                  setEditingId(null);
                  setFormData({
                    data: getLocalDate(),
                    cidade: cities.length > 1 ? [cities[1]] : ['São Paulo'],
                    numero_chamado: '',
                    incidente: '',
                    descricao: '',
                    status: 'Pendente',
                    data_conclusao: null
                  });
                }
              }}
              className="bg-[#EE1D23] hover:bg-[#D1191F] text-white font-black py-3 px-6 rounded-2xl transition-all shadow-lg shadow-red-500/30 active:scale-95 flex items-center gap-2 uppercase italic text-sm"
            >
              {isAdding ? <XCircle className="w-4 h-4" /> : <Plus className="w-4 h-4" />}
              {isAdding ? 'Cancelar' : 'Novo Registro'}
            </button>
          </div>
        </div>

        <AnimatePresence>
          {isAdding && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              className="overflow-hidden bg-slate-50 border-b border-slate-100"
            >
              <form onSubmit={handleSave} className="p-8 grid grid-cols-1 md:grid-cols-4 gap-6">
                <div className="space-y-2">
                  <label className="text-xs font-black text-slate-400 uppercase tracking-widest ml-1">Data</label>
                  <input
                    type="date"
                    required
                    value={formData.data}
                    onChange={e => setFormData({ ...formData, data: e.target.value })}
                    className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-black text-slate-400 uppercase tracking-widest ml-1">Cidades</label>
                  <div className="relative" ref={formCityFilterRef}>
                    <button 
                      type="button"
                      onClick={() => setShowFormCityFilter(!showFormCityFilter)}
                      className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all flex items-center justify-between"
                    >
                      <span className="truncate">
                        {formData.cidade.length === 0 ? 'Selecione as cidades' : 
                         formData.cidade.length === 1 ? formData.cidade[0] : 
                         `${formData.cidade.length} cidades selecionadas`}
                      </span>
                      <ChevronDown className={cn("w-4 h-4 text-slate-400 transition-transform", showFormCityFilter && "rotate-180")} />
                    </button>

                    <AnimatePresence>
                      {showFormCityFilter && (
                        <motion.div 
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: 1, y: 0 }}
                          exit={{ opacity: 0, y: 10 }}
                          className="absolute top-full left-0 mt-2 w-full bg-white border border-slate-100 rounded-2xl shadow-2xl z-50 p-4"
                        >
                          <div className="max-h-60 overflow-y-auto space-y-1">
                            {cities.filter(c => c !== 'Todos').map(city => (
                              <label key={city} className="flex items-center gap-3 p-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-colors">
                                <input 
                                  type="checkbox" 
                                  checked={formData.cidade.includes(city)}
                                  onChange={() => {
                                    const newCidades = formData.cidade.includes(city)
                                      ? formData.cidade.filter(c => c !== city)
                                      : [...formData.cidade, city];
                                    setFormData({ ...formData, cidade: newCidades });
                                  }}
                                  className="w-4 h-4 rounded border-slate-300 text-[#EE1D23] focus:ring-[#EE1D23]"
                                />
                                <span className="text-xs font-bold text-slate-600">{city}</span>
                              </label>
                            ))}
                          </div>
                        </motion.div>
                      )}
                    </AnimatePresence>
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-black text-slate-400 uppercase tracking-widest ml-1">Nº Chamado</label>
                  <input
                    type="text"
                    placeholder="Ex: 123456"
                    value={formData.numero_chamado}
                    onChange={e => setFormData({ ...formData, numero_chamado: e.target.value })}
                    className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-black text-slate-400 uppercase tracking-widest ml-1">Incidente</label>
                  <input
                    type="text"
                    placeholder="Ex: Queda de link"
                    value={formData.incidente}
                    onChange={e => setFormData({ ...formData, incidente: e.target.value })}
                    className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
                  />
                </div>
                <div className="md:col-span-2 space-y-2">
                  <label className="text-xs font-black text-slate-400 uppercase tracking-widest ml-1">Descrição</label>
                  <textarea
                    required
                    placeholder="Descreva o ocorrido..."
                    value={formData.descricao}
                    onChange={e => setFormData({ ...formData, descricao: e.target.value })}
                    className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all min-h-[100px]"
                  />
                </div>
                <div className="md:col-span-2 grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-xs font-black text-slate-400 uppercase tracking-widest ml-1">Status</label>
                    <select
                      value={formData.status}
                      onChange={e => setFormData({ ...formData, status: e.target.value as any })}
                      className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
                    >
                      <option value="Pendente">Pendente</option>
                      <option value="Em Andamento">Em Andamento</option>
                      <option value="Resolvido">Resolvido</option>
                    </select>
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-black text-slate-400 uppercase tracking-widest ml-1">Data Conclusão</label>
                    <input
                      type="date"
                      value={formData.data_conclusao || ''}
                      onChange={e => setFormData({ ...formData, data_conclusao: e.target.value || null })}
                      className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
                    />
                  </div>
                </div>
                <div className="md:col-span-4 flex flex-col md:flex-row items-center justify-between gap-4 pt-4 border-t border-slate-200">
                  {saveError && (
                    <div className="flex items-center gap-2 text-red-500 bg-red-50 px-4 py-2 rounded-xl border border-red-100 text-xs font-bold">
                      <AlertCircle className="w-4 h-4" />
                      {saveError}
                    </div>
                  )}
                  <div className="flex-1" />
                  <button
                    type="submit"
                    disabled={isSaving}
                    className="bg-[#333333] hover:bg-black text-white font-black py-4 px-10 rounded-2xl transition-all shadow-lg active:scale-95 flex items-center gap-3 uppercase italic text-sm disabled:opacity-50"
                  >
                    {isSaving ? <Loader2 className="w-4 h-4 animate-spin" /> : <Save className="w-4 h-4" />}
                    {editingId ? 'Atualizar Registro' : 'Gravar no Banco'}
                  </button>
                </div>
              </form>
            </motion.div>
          )}
        </AnimatePresence>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-100">
                <th className="px-8 py-4 text-xs font-black text-slate-400 uppercase tracking-widest">Data</th>
                <th className="px-8 py-4 text-xs font-black text-slate-400 uppercase tracking-widest">Cidade</th>
                <th className="px-8 py-4 text-xs font-black text-slate-400 uppercase tracking-widest">Nº Chamado</th>
                <th className="px-8 py-4 text-xs font-black text-slate-400 uppercase tracking-widest">Incidente</th>
                <th className="px-8 py-4 text-xs font-black text-slate-400 uppercase tracking-widest">Status</th>
                <th className="px-8 py-4 text-xs font-black text-slate-400 uppercase tracking-widest">Conclusão</th>
                <th className="px-8 py-4 text-xs font-black text-slate-400 uppercase tracking-widest text-right">Ações</th>
              </tr>
            </thead>
            <tbody>
              {isLoading ? (
                <tr>
                  <td colSpan={7} className="px-8 py-20 text-center">
                    <Loader2 className="w-10 h-10 text-[#EE1D23] animate-spin mx-auto mb-4" />
                    <p className="text-slate-400 font-bold uppercase tracking-widest text-xs">Carregando dados...</p>
                  </td>
                </tr>
              ) : filteredEntries.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-8 py-20 text-center">
                    <div className="w-16 h-16 bg-slate-50 rounded-2xl flex items-center justify-center mx-auto mb-4">
                      {error ? <AlertCircle className="w-8 h-8 text-amber-500" /> : <BookOpen className="w-8 h-8 text-slate-200" />}
                    </div>
                    <p className="text-slate-700 font-black uppercase tracking-widest text-xs mb-2">
                      {error ? 'Modo Offline / Sem Registros Remotos' : 'Nenhum registro encontrado'}
                    </p>
                    <p className="text-slate-400 text-xs font-bold max-w-md mx-auto mb-6">
                      {error ? error : 'Clique em "Novo Registro" para adicionar o primeiro apontamento no Diário de Bordo.'}
                    </p>
                    <div className="flex items-center justify-center gap-3">
                      <button 
                        onClick={onRefresh}
                        className="bg-slate-100 hover:bg-slate-200 text-slate-600 px-6 py-3 rounded-xl text-xs font-black uppercase tracking-widest transition-all active:scale-95"
                      >
                        Recarregar
                      </button>
                      <button 
                        onClick={() => {
                          setIsAdding(true);
                          setEditingId(null);
                        }}
                        className="bg-[#EE1D23] hover:bg-[#D1191F] text-white px-6 py-3 rounded-xl text-xs font-black uppercase tracking-widest transition-all active:scale-95 shadow-md shadow-red-500/20"
                      >
                        + Novo Registro
                      </button>
                    </div>
                  </td>
                </tr>
              ) : (
                filteredEntries.map(entry => (
                  <tr key={entry.id} className="border-b border-slate-50 hover:bg-slate-50/50 transition-colors group">
                    <td className="px-8 py-6">
                      <div 
                        className="flex items-center gap-3 cursor-pointer group/date"
                        onClick={() => setSelectedEntry(entry)}
                      >
                        <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center text-[#EE1D23] group-hover/date:bg-[#EE1D23] group-hover/date:text-white transition-colors">
                          <Calendar className="w-5 h-5" />
                        </div>
                        <span className="text-sm font-black text-[#333333] group-hover/date:text-[#EE1D23] transition-colors underline decoration-dotted underline-offset-4">{formatDateDisplay(entry.data)}</span>
                      </div>
                    </td>
                    <td className="px-8 py-6">
                      <div className="flex items-center gap-2">
                        <MapPin className="w-3 h-3 text-slate-400" />
                        <span className="text-xs font-bold text-slate-600">
                          {(entry.cidade || '').split(',').map(c => cleanCityName(c.trim())).join(', ')}
                        </span>
                      </div>
                    </td>
                    <td className="px-8 py-6">
                      <span className="text-xs font-black bg-slate-100 text-slate-600 px-3 py-1.5 rounded-lg uppercase tracking-wider">
                        #{entry.numero_chamado || '-'}
                      </span>
                    </td>
                    <td className="px-8 py-6">
                      <div className="max-w-xs">
                        <p className="text-sm font-black text-[#333333] mb-1">{entry.incidente}</p>
                        <p className="text-xs font-bold text-slate-400 line-clamp-1">{entry.descricao}</p>
                      </div>
                    </td>
                    <td className="px-8 py-6">
                      <span className={cn(
                        "text-[10px] font-black uppercase tracking-widest px-3 py-1.5 rounded-full flex items-center gap-1.5 w-fit",
                        entry.status === 'Resolvido' ? "bg-emerald-50 text-emerald-600 border border-emerald-100" :
                        entry.status === 'Em Andamento' ? "bg-amber-50 text-amber-600 border border-amber-100" :
                        "bg-red-50 text-[#EE1D23] border border-red-100"
                      )}>
                        {entry.status === 'Resolvido' ? <CheckCircle2 className="w-3 h-3" /> :
                         entry.status === 'Em Andamento' ? <Clock className="w-3 h-3" /> :
                         <AlertTriangle className="w-3 h-3" />}
                        {entry.status}
                      </span>
                    </td>
                    <td className="px-8 py-6">
                      <span className="text-xs font-bold text-slate-500 italic">
                        {entry.data_conclusao ? formatDateDisplay(entry.data_conclusao) : '-'}
                      </span>
                    </td>
                    <td className="px-8 py-6">
                      <div className="flex justify-end gap-2 transition-opacity">
                        <button
                          onClick={() => {
                            handleEdit(entry);
                          }}
                          title="Editar registro"
                          className="p-2 bg-white border border-slate-200 rounded-lg text-slate-400 hover:text-[#EE1D23] hover:border-[#EE1D23] transition-all active:scale-90"
                        >
                          <Edit2 className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => {
                            setDeletingId(entry.id!);
                          }}
                          title="Excluir registro"
                          className="p-2 bg-white border border-slate-200 rounded-lg text-slate-400 hover:text-red-600 hover:border-red-600 transition-all active:scale-90"
                        >
                          <Trash className="w-4 h-4" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Custom Delete Confirmation Modal */}
      <AnimatePresence>
        {deletingId && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white rounded-[32px] shadow-2xl border border-slate-100 p-8 max-w-sm w-full text-center"
            >
              <div className="w-20 h-20 bg-red-50 rounded-2xl flex items-center justify-center mb-6 mx-auto">
                <Trash className="w-10 h-10 text-[#EE1D23]" />
              </div>
              <h3 className="text-2xl font-black text-[#333333] uppercase italic tracking-tighter mb-2">Excluir Registro</h3>
              <p className="text-slate-500 font-bold mb-8 italic">Tem certeza que deseja excluir este registro? Esta ação não pode ser desfeita.</p>
              
              <div className="grid grid-cols-2 gap-4">
                <button
                  onClick={() => setDeletingId(null)}
                  className="bg-slate-100 hover:bg-slate-200 text-slate-500 font-black py-4 px-6 rounded-2xl transition-all active:scale-95 uppercase italic text-sm"
                >
                  Cancelar
                </button>
                <button
                  onClick={() => handleDelete(deletingId)}
                  className="bg-[#EE1D23] hover:bg-[#D1191F] text-white font-black py-4 px-6 rounded-2xl transition-all shadow-lg shadow-red-500/30 active:scale-95 uppercase italic text-sm"
                >
                  Excluir
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Detail Modal */}
      <AnimatePresence>
        {selectedEntry && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm overflow-y-auto">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white rounded-[32px] shadow-2xl border border-slate-100 p-6 sm:p-8 max-w-2xl w-full max-h-[90vh] flex flex-col my-auto"
            >
              <div className="flex items-center justify-between pb-4 border-b border-slate-100 flex-shrink-0">
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 bg-red-50 rounded-2xl flex items-center justify-center text-[#EE1D23] flex-shrink-0">
                    <BookOpen className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="text-2xl font-black text-[#333333] uppercase italic tracking-tighter">Detalhes do Registro</h3>
                    <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Informações Completas</p>
                  </div>
                </div>
                <button 
                  onClick={() => setSelectedEntry(null)} 
                  className="p-2 hover:bg-slate-100 rounded-xl transition-colors text-slate-400 hover:text-slate-600"
                  aria-label="Fechar modal"
                >
                  <XCircle className="w-6 h-6 text-slate-300 hover:text-[#EE1D23] transition-colors" />
                </button>
              </div>

              {/* Scrollable content body */}
              <div className="flex-1 overflow-y-auto pr-1 my-4 space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 sm:gap-6">
                  <div className="space-y-1">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Data do Registro</p>
                    <p className="text-sm font-black text-[#333333]">{formatDateDisplay(selectedEntry.data)}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Cidade</p>
                    <p className="text-sm font-black text-[#333333]">
                      {(selectedEntry.cidade || '').split(',').map(c => cleanCityName(c.trim())).join(', ')}
                    </p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Nº Chamado</p>
                    <p className="text-sm font-black text-[#333333]">#{selectedEntry.numero_chamado || 'N/A'}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Incidente</p>
                    <p className="text-sm font-black text-[#333333]">{selectedEntry.incidente}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Status</p>
                    <span className={cn(
                      "text-[10px] font-black uppercase tracking-widest px-3 py-1 rounded-full flex items-center gap-1.5 w-fit",
                      selectedEntry.status === 'Resolvido' ? "bg-emerald-50 text-emerald-600 border border-emerald-100" :
                      selectedEntry.status === 'Em Andamento' ? "bg-amber-50 text-amber-600 border border-amber-100" :
                      "bg-red-50 text-[#EE1D23] border border-red-100"
                    )}>
                      {selectedEntry.status}
                    </span>
                  </div>
                  <div className="md:col-span-2 space-y-1">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Descrição Completa</p>
                    <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 max-h-60 overflow-y-auto">
                      <p className="text-sm font-bold text-slate-600 leading-relaxed whitespace-pre-wrap break-words">{selectedEntry.descricao}</p>
                    </div>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Data de Conclusão</p>
                    <p className="text-sm font-black text-[#333333]">
                      {selectedEntry.data_conclusao ? formatDateDisplay(selectedEntry.data_conclusao) : 'Ainda não concluído'}
                    </p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Criado em</p>
                    <p className="text-xs font-bold text-slate-400">
                      {selectedEntry.created_at ? new Date(selectedEntry.created_at).toLocaleString('pt-BR') : 'N/A'}
                    </p>
                  </div>
                </div>
              </div>

              {/* Fixed Footer */}
              <div className="flex justify-end pt-4 border-t border-slate-100 flex-shrink-0">
                <button
                  onClick={() => setSelectedEntry(null)}
                  className="bg-[#333333] hover:bg-black text-white font-black py-3.5 px-10 rounded-2xl transition-all shadow-lg active:scale-95 uppercase italic text-sm"
                >
                  Fechar Detalhes
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

// Claro Colors
const CLARO_RED = '#EE1D23';
const CLARO_DARK = '#1A1A1A';
const CLARO_GRAY = '#F2F2F2';

// Helper to normalize strings for comparison (remove accents, uppercase, trim, remove non-alphanumeric)
const normalizeStr = (str: string) => {
  if (!str) return '';
  return str
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, ' ') // Replace non-alphanumeric with space
    .trim()
    .replace(/\s+/g, ' '); // Collapse multiple spaces
};

// Helper to clean city name (remove state suffix like " - PA")
const cleanCityName = (name: string) => {
  if (!name) return 'N/A';
  // Remove state suffix like " - PA", " - SP", etc. at the end of the string
  return name.split(/\s*-\s*[A-Z]{2}$/i)[0].trim();
};

// Helper to parse numbers from Excel (handles strings with dots/commas)
const parseExcelNumber = (val: any): number => {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const str = String(val).replace(/\s/g, '');
  // If it's something like "1.234,56" or "1,234.56"
  // This is tricky, but usually for base we expect integers or simple floats.
  // Let's try to remove thousands separators.
  // If there's both a dot and a comma, the last one is the decimal.
  const lastDot = str.lastIndexOf('.');
  const lastComma = str.lastIndexOf(',');
  
  if (lastDot > lastComma) {
    // Dot is decimal, remove commas
    return parseFloat(str.replace(/,/g, '')) || 0;
  } else if (lastComma > lastDot) {
    // Comma is decimal, remove dots, then replace comma with dot
    return parseFloat(str.replace(/\./g, '').replace(',', '.')) || 0;
  }
  
  return parseFloat(str) || 0;
};

// --- Error Boundary ---
class ErrorBoundary extends Component<{ children: React.ReactNode }, { hasError: boolean, error: any }> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: any) {
    return { hasError: true, error };
  }

  componentDidCatch(error: any, errorInfo: any) {
    console.error("ErrorBoundary caught an error", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6 text-center">
          <div className="w-20 h-20 bg-red-50 rounded-2xl flex items-center justify-center mb-6">
            <AlertCircle className="w-10 h-10 text-[#EE1D23]" />
          </div>
          <h2 className="text-3xl font-black text-[#333333] uppercase italic tracking-tighter mb-4">Ops! Algo deu errado</h2>
          <p className="text-slate-500 font-bold max-w-md mb-8">
            Ocorreu um erro inesperado ao renderizar o dashboard. Tente recarregar a página ou limpar os dados.
          </p>
          <button 
            onClick={() => window.location.reload()}
            className="bg-[#EE1D23] hover:bg-[#D1191F] text-white font-black py-4 px-8 rounded-2xl transition-all shadow-lg shadow-red-500/30 active:scale-95 uppercase italic"
          >
            Recarregar Página
          </button>
          {this.state.error && (
            <pre className="mt-8 p-4 bg-slate-100 rounded-xl text-left text-xs text-slate-500 max-w-2xl overflow-auto border border-slate-200">
              {this.state.error.toString()}
            </pre>
          )}
        </div>
      );
    }

    return this.props.children;
  }
}

// Constants
const MONTH_ORDER = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

export default function App() {
  const [baseData, setBaseData] = useState<VisitData[]>([]);
  const [baseCidadeData, setBaseCidadeData] = useState<BaseCidadeData[]>([]);
  const [importError, setImportError] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [importProgress, setImportProgress] = useState(0);
  const [showScrollTop, setShowScrollTop] = useState(false);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'outage' | 'at5' | 'log' | 'revisita30d' | 'logbook'>('dashboard');
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(() => {
    try {
      return localStorage.getItem('CLARO_SIDEBAR_COLLAPSED') === 'true';
    } catch {
      return false;
    }
  });
  const [logEntries, setLogEntries] = useState<LogEntry[]>([]);
  const [isLogLoading, setIsLogLoading] = useState(false);
  const [logFetchError, setLogFetchError] = useState<string | null>(null);
  const [githubUrl, setGithubUrl] = useState('');
  const [showGithubInput, setShowGithubInput] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Helper to get environment variables with localStorage fallback
  const getEnv = (key: string) => {
    const buildVal = (import.meta as any).env?.[key];
    if (buildVal && buildVal !== '') return buildVal;
    try {
      return localStorage.getItem(key) || '';
    } catch (e) {
      return '';
    }
  };

  // Fetch logs from Hostinger MySQL
  const fetchLogs = async () => {
    setIsLogLoading(true);
    setLogFetchError(null);
    try {
      console.log('[DEBUG] fetchLogs starting...');
      const data = await logApi.getAll();
      console.log('[DEBUG] fetchLogs success, count:', data?.length);
      setLogEntries(data || []);
    } catch (err: any) {
      console.error('[DEBUG] fetchLogs error:', err);
      setLogFetchError(err.message || String(err));
    } finally {
      setIsLogLoading(false);
    }
  };

  // Scroll to top effect
  React.useEffect(() => {
    const handleScroll = () => {
      setShowScrollTop(window.scrollY > 400);
    };
    window.addEventListener('scroll', handleScroll);
    return () => window.removeEventListener('scroll', handleScroll);
  }, []);

  // Fetch logs when logbook tab is active
  React.useEffect(() => {
    if (activeTab === 'logbook') {
      fetchLogs();
    }
  }, [activeTab]);

  const scrollToTop = () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  // Filters State
  const [filters, setFilters] = useState({
    mes: ['Todos'] as string[],
    semana: ['Todos'] as string[],
    cidade: ['Todos'] as string[],
    area: ['Todos'] as string[],
    expurgo: 'Todos',
    tecnologia: ['Todos'] as string[],
    grupoBaixa: ['Todos'] as string[],
    startDate: '',
    endDate: ''
  });
  const [selectedNodes, setSelectedNodes] = useState<string[]>([]);
  const [selectedBaixas, setSelectedBaixas] = useState<string[]>([]);

  // Comparison Window
  const comparisonMonths = useMemo(() => {
    if (baseData.length === 0) return { current: null, previous: null };
    
    const allMonthsInBase = [...new Set(baseData.map(d => d.mes))];
    const sortedAvailable = allMonthsInBase.sort((a, b) => MONTH_ORDER.indexOf(a) - MONTH_ORDER.indexOf(b));
    
    let current = filters.mes.includes('Todos') || filters.mes.length === 0
      ? sortedAvailable[sortedAvailable.length - 1] 
      : filters.mes[0];
    
    const currentIndex = MONTH_ORDER.indexOf(current);
    const previous = currentIndex > 0 ? MONTH_ORDER[currentIndex - 1] : null;
    
    return { current, previous };
  }, [baseData, filters.mes]);

  // Unique values for filters (Dependent Filters)
  const filterOptions = useMemo(() => {
    const getOptions = (key: keyof VisitData, currentFilters: typeof filters) => {
      const otherFiltersApplied = baseData.filter(item => {
        const matchMes = key === 'mes' || currentFilters.mes.includes('Todos') || currentFilters.mes.includes(item.mes);
        const matchCidade = key === 'cidade' || currentFilters.cidade.includes('Todos') || currentFilters.cidade.includes(item.cidade);
        const matchArea = key === 'area' || currentFilters.area.includes('Todos') || currentFilters.area.includes(item.area);
        const matchExpurgo = key === 'expurgo' || currentFilters.expurgo === 'Todos' || 
          (currentFilters.expurgo === 'Sim' ? item.expurgo : !item.expurgo);
        const matchTec = key === 'tecnologia' || currentFilters.tecnologia.includes('Todos') || currentFilters.tecnologia.includes(item.tecnologia);
        const matchGrupo = key === 'grupoBaixa' || currentFilters.grupoBaixa.includes('Todos') || currentFilters.grupoBaixa.includes(item.grupoBaixa);
        return matchMes && matchCidade && matchArea && matchExpurgo && matchTec && matchGrupo;
      });
      
      const unique = [...new Set(otherFiltersApplied.map(d => String(d[key])))];
      return ['Todos', ...(unique as string[]).sort((a, b) => a.localeCompare(b, 'pt-BR'))];
    };

    return {
      meses: getOptions('mes', filters),
      semanas: ['Todos', 'S1', 'S2', 'S3', 'S4', 'S5'],
      cidades: getOptions('cidade', filters),
      areas: getOptions('area', filters),
      expurgos: ['Todos', 'Sim', 'Não'],
      tecnologias: getOptions('tecnologia', filters),
      gruposBaixa: getOptions('grupoBaixa', filters)
    };
  }, [baseData, filters]);

  // Filtered Data (Base filters from dropdowns)
  const filteredData = useMemo(() => {
    return baseData.filter(item => {
      const matchMes = filters.mes.includes('Todos') || filters.mes.includes(item.mes);
      
      let matchSemana = filters.semana.includes('Todos');
      if (!matchSemana && item.fullDate instanceof Date && !isNaN(item.fullDate.getTime())) {
        const day = item.fullDate.getDate();
        let itemSemana = 'S5';
        if (day <= 7) itemSemana = 'S1';
        else if (day <= 14) itemSemana = 'S2';
        else if (day <= 21) itemSemana = 'S3';
        else if (day <= 28) itemSemana = 'S4';
        matchSemana = filters.semana.includes(itemSemana);
      }

      const matchCidade = filters.cidade.includes('Todos') || filters.cidade.includes(item.cidade);
      const matchArea = filters.area.includes('Todos') || filters.area.includes(item.area);
      const matchExpurgo = filters.expurgo === 'Todos' || 
        (filters.expurgo === 'Sim' ? item.expurgo : !item.expurgo);
      const matchTec = filters.tecnologia.includes('Todos') || filters.tecnologia.includes(item.tecnologia);
      const matchGrupo = filters.grupoBaixa.includes('Todos') || filters.grupoBaixa.includes(item.grupoBaixa);
      
      const itemDate = new Date(item.fullDate);
      itemDate.setHours(0, 0, 0, 0);
      
      let matchDate = true;
      if (filters.startDate) {
        const [y, m, d] = filters.startDate.split('-').map(Number);
        const start = new Date(y, m - 1, d);
        start.setHours(0, 0, 0, 0);
        matchDate = matchDate && itemDate >= start;
      }
      if (filters.endDate) {
        const [y, m, d] = filters.endDate.split('-').map(Number);
        const end = new Date(y, m - 1, d);
        end.setHours(0, 0, 0, 0);
        matchDate = matchDate && itemDate <= end;
      }
      
      return matchMes && matchSemana && matchCidade && matchArea && matchExpurgo && matchTec && matchGrupo && matchDate;
    });
  }, [filters, baseData]);

  // Comparison Data Filtered (All filters EXCEPT month)
  const momFilteredData = useMemo(() => {
    return baseData.filter(item => {
      let matchSemana = filters.semana.includes('Todos');
      if (!matchSemana && item.fullDate instanceof Date && !isNaN(item.fullDate.getTime())) {
        const day = item.fullDate.getDate();
        let itemSemana = 'S5';
        if (day <= 7) itemSemana = 'S1';
        else if (day <= 14) itemSemana = 'S2';
        else if (day <= 21) itemSemana = 'S3';
        else if (day <= 28) itemSemana = 'S4';
        matchSemana = filters.semana.includes(itemSemana);
      }

      const matchCidade = filters.cidade.includes('Todos') || filters.cidade.includes(item.cidade);
      const matchArea = filters.area.includes('Todos') || filters.area.includes(item.area);
      const matchExpurgo = filters.expurgo === 'Todos' || 
        (filters.expurgo === 'Sim' ? item.expurgo : !item.expurgo);
      const matchTec = filters.tecnologia.includes('Todos') || filters.tecnologia.includes(item.tecnologia);
      const matchGrupo = filters.grupoBaixa.includes('Todos') || filters.grupoBaixa.includes(item.grupoBaixa);
      
      const itemDate = new Date(item.fullDate);
      itemDate.setHours(0, 0, 0, 0);
      
      let matchDate = true;
      if (filters.startDate) {
        const [y, m, d] = filters.startDate.split('-').map(Number);
        const start = new Date(y, m - 1, d);
        start.setHours(0, 0, 0, 0);
        matchDate = matchDate && itemDate >= start;
      }
      if (filters.endDate) {
        const [y, m, d] = filters.endDate.split('-').map(Number);
        const end = new Date(y, m - 1, d);
        end.setHours(0, 0, 0, 0);
        matchDate = matchDate && itemDate <= end;
      }
      
      return matchSemana && matchCidade && matchArea && matchExpurgo && matchTec && matchGrupo && matchDate;
    });
  }, [filters, baseData]);

  // Data for Node Chart (Filtered by dropdowns + Selected Baixas)
  const nodeChartData = useMemo(() => {
    const { current, previous } = comparisonMonths;
    const data = momFilteredData.filter(d => d.mes === current || d.mes === previous);
    if (selectedBaixas.length === 0) return data;
    return data.filter(item => selectedBaixas.includes(item.cdBaixa));
  }, [momFilteredData, comparisonMonths, selectedBaixas]);

  // Data for Baixa Chart (Filtered by dropdowns + Selected Nodes)
  const baixaChartData = useMemo(() => {
    const { current, previous } = comparisonMonths;
    const data = momFilteredData.filter(d => d.mes === current || d.mes === previous);
    if (selectedNodes.length === 0) return data;
    return data.filter(item => selectedNodes.includes(item.node));
  }, [momFilteredData, comparisonMonths, selectedNodes]);

  // Final Filtered Data (For Metrics - includes all filters)
  const finalFilteredData = useMemo(() => {
    let data = filteredData;
    if (selectedNodes.length > 0) data = data.filter(item => selectedNodes.includes(item.node));
    if (selectedBaixas.length > 0) data = data.filter(item => selectedBaixas.includes(item.cdBaixa));
    return data;
  }, [filteredData, selectedNodes, selectedBaixas]);

  // Helper for metrics calculation
  const getMetricsSummary = useMemo(() => {
    return (data: VisitData[]) => {
      const total = data.length;
      const executada = data.filter(d => d.status === 'Executada').length;
      const cancelada = data.filter(d => d.status === 'Cancelada').length;
      
      const calculateAvg = (items: VisitData[]) => {
        const validNotes = items.filter(d => d.status === 'Executada' && d.notaAT1 > 0);
        if (validNotes.length === 0) return 0;
        return validNotes.reduce((acc, curr) => acc + curr.notaAT1, 0) / validNotes.length;
      };

      const calculateProjectionDetails = (items: VisitData[], techOverride?: Technology, globalDiasTrabOverride?: number) => {
        const osExec = items.filter(d => d.status === 'Executada').length;
        let diasTrab = 1;
        if (globalDiasTrabOverride) {
          diasTrab = globalDiasTrabOverride;
        } else {
          const uniqueDates = new Set(items.map(d => d.fullDate instanceof Date ? d.fullDate.toDateString() : ''));
          uniqueDates.delete('');
          diasTrab = uniqueDates.size || 1;
        }

        const referenceDate = (items.length > 0 && items[0].fullDate instanceof Date) ? items[0].fullDate : new Date();
        const diasMes = new Date(referenceDate.getFullYear(), referenceDate.getMonth() + 1, 0).getDate();

        let baseClientes = 0;
        const cityFilter = filters.cidade;
        const techFilter = techOverride || filters.tecnologia;
        const normCityFilters = cityFilter.map(c => normalizeStr(c));
        const isTechArray = Array.isArray(techFilter);
        const normTechFilters = isTechArray 
          ? (techFilter as string[]).map(t => normalizeStr(t))
          : [normalizeStr(techFilter as string)];

        const isTodos = cityFilter.includes('Todos');
        const isTechTodos = isTechArray ? (techFilter as string[]).includes('Todos') : (techFilter as string) === 'Todos';
        
        baseCidadeData.forEach(b => {
          const normBCidade = normalizeStr(b.cidade);
          const normBTech = normalizeStr(b.tecnologia);
          if (isTodos) {
            const isTotalOrRegional = normBCidade.includes('TOTAL') || normBCidade.includes('GERAL') || normBCidade.includes('SUM') ||
              ['NORTE', 'SUL', 'LESTE', 'OESTE', 'REGIONAL'].includes(normBCidade);
            if (isTotalOrRegional) return;
          }
          const cityMatch = isTodos || normCityFilters.some(cf => normBCidade === cf || normBCidade.includes(cf) || cf.includes(normBCidade));
          const techMatch = isTechTodos || normTechFilters.includes(normBTech);
          if (cityMatch && techMatch) baseClientes += b.base;
        });

        const projOsAT1 = (osExec / diasTrab) * diasMes;
        const at1 = baseClientes > 0 ? (projOsAT1 / baseClientes) * 100 : 0;
        const totalVisits = items.length;
        const projOsG1 = (totalVisits / diasTrab) * diasMes;
        const g1 = baseClientes > 0 ? (projOsG1 / baseClientes) * 100 : 0;

        return { osExec, totalVisits, diasTrab, projOsAT1, projOsG1, base: baseClientes, at1, g1, diasMes };
      };

      // Use data provided to the function (metrics or prevMetrics) to calculate progress
      // instead of using the global filters.mes, so that previous months don't get
      // affected by the current month's selection.
      const dayNumbers = data.map(d => d.fullDate instanceof Date ? d.fullDate.getDate() : 0).filter(d => d > 0);
      const currentMonthProgress = dayNumbers.length > 0 ? Math.max(...dayNumbers) : 1;

      const totalDetails = calculateProjectionDetails(data, undefined, currentMonthProgress);
      return {
        total, executada, cancelada,
        avgQualityTotal: calculateAvg(data),
        projAT1Total: totalDetails.at1,
        projG1Total: totalDetails.g1,
        details: totalDetails,
        techDetails: {
          HFC: calculateProjectionDetails(data.filter(d => d.tecnologia === 'HFC'), 'HFC', currentMonthProgress),
          GPON: calculateProjectionDetails(data.filter(d => d.tecnologia === 'GPON'), 'GPON', currentMonthProgress),
          HÍBRIDO: calculateProjectionDetails(data.filter(d => d.tecnologia === 'HÍBRIDO'), 'HÍBRIDO', currentMonthProgress),
          OUTROS: calculateProjectionDetails(data.filter(d => d.tecnologia === 'OUTROS'), 'OUTROS', currentMonthProgress)
        }
      };
    };
  }, [filters, baseCidadeData, baseData]);

  // Metrics Calculations (Uses finalFilteredData)
  const metrics = useMemo(() => {
    return getMetricsSummary(finalFilteredData);
  }, [finalFilteredData, getMetricsSummary]);

  // Previous Metrics for Comparison
  const prevMetrics = useMemo(() => {
    const { previous } = comparisonMonths;
    if (!previous) return null;
    const prevData = momFilteredData.filter(d => d.mes === previous);
    let data = prevData;
    if (selectedNodes.length > 0) data = data.filter(item => selectedNodes.includes(item.node));
    if (selectedBaixas.length > 0) data = data.filter(item => selectedBaixas.includes(item.cdBaixa));
    return getMetricsSummary(data);
  }, [momFilteredData, comparisonMonths, getMetricsSummary, selectedNodes, selectedBaixas]);

  const lastUpdateDate = useMemo(() => {
    if (baseData.length === 0) return null;
    const dates = baseData.map(d => d.fullDate instanceof Date ? d.fullDate.getTime() : 0).filter(t => t > 0);
    if (dates.length === 0) return null;
    return new Date(Math.max(...dates));
  }, [baseData]);

  const chartData = [
    { name: 'HFC', current: metrics.techDetails.HFC.at1, previous: prevMetrics?.techDetails.HFC.at1 || 0 },
    { name: 'GPON', current: metrics.techDetails.GPON.at1, previous: prevMetrics?.techDetails.GPON.at1 || 0 },
    { name: 'HÍBRIDO', current: metrics.techDetails.HÍBRIDO.at1, previous: prevMetrics?.techDetails.HÍBRIDO.at1 || 0 },
    { name: 'OUTROS', current: metrics.techDetails.OUTROS.at1, previous: prevMetrics?.techDetails.OUTROS.at1 || 0 },
  ];

  // Daily Volume Data
  const dailyVolume = useMemo(() => {
    const dayMap: Record<string, { current: number, previous: number }> = {};
    const { current: currentMonth, previous: previousMonth } = comparisonMonths;
    
    // For comparison, always show 1-31 days
    for (let d = 1; d <= 31; d++) {
      dayMap[String(d)] = { current: 0, previous: 0 };
    }

    // Fill with current data
    finalFilteredData.forEach(item => {
      if (item.mes === currentMonth && item.fullDate instanceof Date && !isNaN(item.fullDate.getTime())) {
        const dayKey = String(item.fullDate.getDate());
        if (dayMap[dayKey]) {
          dayMap[dayKey].current += 1;
        }
      }
    });

    // Fill with previous month data
    if (previousMonth) {
      momFilteredData.forEach(item => {
        if (item.mes === previousMonth && item.fullDate instanceof Date && !isNaN(item.fullDate.getTime())) {
          const dayKey = String(item.fullDate.getDate());
          if (dayMap[dayKey]) {
            dayMap[dayKey].previous += 1;
          }
        }
      });
    }

    // Identify the last day with data for the current month
    let lastDayWithData = 0;
    finalFilteredData.forEach(item => {
      if (item.mes === currentMonth && item.fullDate instanceof Date && !isNaN(item.fullDate.getTime())) {
        const d = item.fullDate.getDate();
        if (d > lastDayWithData) lastDayWithData = d;
      }
    });

    return Object.entries(dayMap)
      .map(([name, data]) => {
        const dayNum = Number(name);
        return { 
          name, 
          // If day is past current data point and value is 0, make it null to stop the line
          value: (dayNum > lastDayWithData && data.current === 0) ? null : data.current, 
          previousValue: data.previous 
        };
      })
      .sort((a, b) => Number(a.name) - Number(b.name))
      .filter(d => (d.value !== null && d.value >= 0) || d.previousValue > 0);
  }, [finalFilteredData, momFilteredData, comparisonMonths, filters]);

  // Weekly Volume Data
  const weeklyVolume = useMemo(() => {
    const weekMap: Record<string, number> = {
      'S1': 0,
      'S2': 0,
      'S3': 0,
      'S4': 0,
      'S5': 0
    };

    finalFilteredData.forEach(item => {
      if (item.fullDate instanceof Date && !isNaN(item.fullDate.getTime())) {
        const day = item.fullDate.getDate();
        if (day <= 7) weekMap['S1'] += 1;
        else if (day <= 14) weekMap['S2'] += 1;
        else if (day <= 21) weekMap['S3'] += 1;
        else if (day <= 28) weekMap['S4'] += 1;
        else weekMap['S5'] += 1;
      }
    });

    return Object.entries(weekMap).map(([name, value]) => ({ name, value }));
  }, [finalFilteredData]);

  // Top 15 Offending Nodes (Uses nodeChartData)
  const topNodes = useMemo(() => {
    const { current, previous } = comparisonMonths;
    const currentCounts: Record<string, number> = {};
    const prevCounts: Record<string, number> = {};

    nodeChartData.forEach(item => {
      if (item.node && item.node !== 'N/A') {
        if (item.mes === current) {
          currentCounts[item.node] = (currentCounts[item.node] || 0) + 1;
        } else if (item.mes === previous) {
          prevCounts[item.node] = (prevCounts[item.node] || 0) + 1;
        }
      }
    });

    const allNodeNames = [...new Set([...Object.keys(currentCounts), ...Object.keys(prevCounts)])];
    const allData = allNodeNames.map(name => ({
      name,
      current: currentCounts[name] || 0,
      previous: prevCounts[name] || 0
    })).sort((a, b) => b.current - a.current);

    let result = allData.slice(0, 15);
    
    // Ensure selected nodes are included
    selectedNodes.forEach(selNode => {
      if (!result.find(n => n.name === selNode)) {
        const found = allData.find(n => n.name === selNode);
        if (found) result.push(found);
      }
    });

    return result.sort((a, b) => b.current - a.current);
  }, [nodeChartData, selectedNodes, comparisonMonths]);

  // Top 20 Low Codes (CD_BAIXA) (Uses baixaChartData)
  const topBaixas = useMemo(() => {
    const { current, previous } = comparisonMonths;
    const currentCounts: Record<string, number> = {};
    const prevCounts: Record<string, number> = {};
    const descriptions: Record<string, string> = {};

    baixaChartData.forEach(item => {
      if (item.cdBaixa && item.cdBaixa !== 'N/A') {
        if (item.descriptionBaixa && item.descriptionBaixa !== 'N/A') {
          descriptions[item.cdBaixa] = item.descriptionBaixa;
        }
        if (item.mes === current) {
          currentCounts[item.cdBaixa] = (currentCounts[item.cdBaixa] || 0) + (item.volume || 1);
        } else if (item.mes === previous) {
          prevCounts[item.cdBaixa] = (prevCounts[item.cdBaixa] || 0) + (item.volume || 1);
        }
      }
    });

    const currentBase = metrics.details.base || 1;
    const previousBase = prevMetrics?.details.base || 1;
    const allBaixaNames = [...new Set([...Object.keys(currentCounts), ...Object.keys(prevCounts)])];

    const allData = allBaixaNames.map(name => {
      const currentVal = currentCounts[name] || 0;
      const previousVal = prevCounts[name] || 0;
      const impact = (currentVal / currentBase) * 100;
      const impactPrev = (previousVal / previousBase) * 100;
      const description = descriptions[name] || 'N/A';
      
      // Create a display name for the YAxis
      let displayName = name;
      if (description !== 'N/A') {
        // Truncate description for Y axis but keep code visible
        const truncatedDesc = description.length > 30 ? description.substring(0, 27) + '...' : description;
        displayName = `${name} - ${truncatedDesc}`;
      }
      
      return {
        name,
        displayName,
        description,
        current: currentVal,
        previous: previousVal,
        impact: Number(impact.toFixed(2)),
        impactPrev: Number(impactPrev.toFixed(2))
      };
    }).sort((a, b) => b.current - a.current);

    let result = allData.slice(0, 20);
    
    // Ensure selected baixas are included
    selectedBaixas.forEach(selBaixa => {
      if (!result.find(n => n.name === selBaixa)) {
        const found = allData.find(n => n.name === selBaixa);
        if (found) result.push(found);
      }
    });

    return result.sort((a, b) => b.current - a.current);
  }, [baixaChartData, selectedBaixas, comparisonMonths, metrics, prevMetrics]);

  // Volume por Área
  const areaData = useMemo(() => {
    const areaMap: Record<string, number> = {};
    finalFilteredData.forEach(item => {
      if (item.area && item.area !== 'N/A') {
        areaMap[item.area] = (areaMap[item.area] || 0) + (item.volume || 1);
      }
    });
    return Object.entries(areaMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
  }, [finalFilteredData]);

  // Volume por Terminal
  const terminalData = useMemo(() => {
    const { current, previous } = comparisonMonths;
    const currentMap: Record<string, number> = {};
    const prevMap: Record<string, number> = {};
    
    const data = momFilteredData.filter(d => d.mes === current || d.mes === previous);
    
    data.forEach(item => {
      if (item.terminal && item.terminal !== 'N/A') {
        if (item.mes === current) {
          currentMap[item.terminal] = (currentMap[item.terminal] || 0) + (item.volume || 1);
        } else if (item.mes === previous) {
          prevMap[item.terminal] = (prevMap[item.terminal] || 0) + (item.volume || 1);
        }
      }
    });
    
    const allTerminals = [...new Set([...Object.keys(currentMap), ...Object.keys(prevMap)])];
    return allTerminals.map(name => ({
      name,
      current: currentMap[name] || 0,
      previous: prevMap[name] || 0
    })).sort((a, b) => b.current - a.current).slice(0, 15);
  }, [momFilteredData, comparisonMonths]);

  // Volume por Grupo de Baixa
  const grupoBaixaData = useMemo(() => {
    const { current, previous } = comparisonMonths;
    const currentMap: Record<string, number> = {};
    const prevMap: Record<string, number> = {};
    
    // Get current and previous base for impact calculation
    const currentBase = metrics.details.base || 1;
    const previousBase = prevMetrics?.details.base || 1;
    
    const data = momFilteredData.filter(d => d.mes === current || d.mes === previous);
    
    data.forEach(item => {
      if (item.grupoBaixa && item.grupoBaixa !== 'N/A') {
        if (item.mes === current) {
          currentMap[item.grupoBaixa] = (currentMap[item.grupoBaixa] || 0) + (item.volume || 1);
        } else if (item.mes === previous) {
          prevMap[item.grupoBaixa] = (prevMap[item.grupoBaixa] || 0) + (item.volume || 1);
        }
      }
    });
    
    const allGrupos = [...new Set([...Object.keys(currentMap), ...Object.keys(prevMap)])];
    return allGrupos.map(name => {
      const currentVal = currentMap[name] || 0;
      const previousVal = prevMap[name] || 0;
      const impact = (currentVal / currentBase) * 100;
      const impactPrev = (previousVal / previousBase) * 100;
      
      return {
        name,
        current: currentVal,
        previous: previousVal,
        impact: Number(impact.toFixed(2)),
        impactPrev: Number(impactPrev.toFixed(2))
      };
    }).sort((a, b) => b.current - a.current);
  }, [momFilteredData, comparisonMonths, metrics, prevMetrics]);

  // Base Data Aggregations for Charts
  const baseMetrics = useMemo(() => {
    const cityMap: Record<string, number> = {};
    const techMap: Record<string, number> = { 'HFC': 0, 'GPON': 0, 'HÍBRIDO': 0, 'OUTROS': 0 };

    const normCityFilters = filters.cidade.map(c => normalizeStr(c));
    const normTechFilters = filters.tecnologia.map(t => normalizeStr(t));

    const isTodos = filters.cidade.includes('Todos');
    const isTechTodos = filters.tecnologia.includes('Todos');

    baseCidadeData.forEach(item => {
      const normCity = normalizeStr(item.cidade);
      const normTech = normalizeStr(item.tecnologia);

      // If "Todos", exclude total/regional rows from charts and aggregations
      if (isTodos) {
        const isTotalOrRegional = 
          normCity.includes('TOTAL') || 
          normCity.includes('GERAL') || 
          normCity.includes('SUM') ||
          normCity === 'NORTE' || 
          normCity === 'SUL' || 
          normCity === 'LESTE' || 
          normCity === 'OESTE' ||
          normCity === 'REGIONAL';
        
        if (isTotalOrRegional) return;
      }

      const matchCity = isTodos || normCityFilters.some(cf => 
        normCity === cf || normCity.includes(cf) || cf.includes(normCity)
      );
      
      const matchTech = isTechTodos || normTechFilters.includes(normTech);

      if (matchCity && matchTech) {
        cityMap[item.cidade] = (cityMap[item.cidade] || 0) + item.base;
        techMap[item.tecnologia] = (techMap[item.tecnologia] || 0) + item.base;
      }
    });

    const cityData = Object.entries(cityMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const totalBase = Object.values(techMap).reduce((acc, curr) => acc + curr, 0);
    const techData = Object.entries(techMap)
      .map(([name, value]) => ({ 
        name, 
        value,
        percentage: totalBase > 0 ? (value / totalBase) * 100 : 0
      }));

    if (totalBase > 0) {
      techData.push({
        name: 'TOTAL',
        value: totalBase,
        percentage: 100
      });
    }

    return { cityData, techData };
  }, [baseCidadeData, filters]);

  // AT1 by City Projection Data
  const at1ByCityData = useMemo(() => {
    const { previous } = comparisonMonths;
    
    // Prepare previous month filtered data
    const prevFilteredData = previous 
      ? momFilteredData.filter(d => 
          d.mes === previous && 
          (selectedNodes.length === 0 || selectedNodes.includes(d.node)) && 
          (selectedBaixas.length === 0 || selectedBaixas.includes(d.cdBaixa))
        )
      : [];

    const getCityStats = (visitData: VisitData[]) => {
      const cityMap: Record<string, { name: string, osExec: number, dates: Set<string>, base: number }> = {};
      
      const isTodos = filters.cidade.includes('Todos');
      const normCityFilters = filters.cidade.map(c => normalizeStr(c).replace(/\s+/g, ''));
      const normTechFilters = filters.tecnologia.map(t => normalizeStr(t));
      const isTechTodos = filters.tecnologia.includes('Todos');

      // 1. Populate cityMap from base data
      baseCidadeData.forEach(item => {
        const normCity = normalizeStr(item.cidade);
        const normCityKey = normCity.replace(/\s+/g, '');
        const normTech = normalizeStr(item.tecnologia);

        const matchCity = isTodos || normCityFilters.some(cf => 
          normCityKey === cf || normCityKey.includes(cf) || cf.includes(normCityKey)
        );
        
        const matchTech = isTechTodos || normTechFilters.includes(normTech);

        if (matchCity && matchTech) {
          // Exclude total/regional rows if "Todos" is selected
          if (isTodos) {
            const isTotalOrRegional = 
              normCityKey.includes('TOTAL') || 
              normCityKey.includes('GERAL') || 
              normCityKey.includes('SUM') ||
              normCityKey === 'NORTE' || 
              normCityKey === 'SUL' || 
              normCityKey === 'LESTE' || 
              normCityKey === 'OESTE' ||
              normCityKey === 'REGIONAL';
            if (isTotalOrRegional) return;
          }

          if (!cityMap[normCityKey]) {
            cityMap[normCityKey] = { name: item.cidade, osExec: 0, dates: new Set(), base: 0 };
          }
          cityMap[normCityKey].base += item.base;
        }
      });

      // 2. Add visit data
      const cityKeys = Object.keys(cityMap);
      visitData.forEach(item => {
        const normCityItem = normalizeStr(item.cidade).replace(/\s+/g, '');
        
        // Find the best matching key in cityMap
        let matchedKey = cityMap[normCityItem] ? normCityItem : cityKeys.find(key => 
          key === normCityItem || key.includes(normCityItem) || normCityItem.includes(key)
        );

        if (matchedKey) {
          if (normalizeStr(item.status) === normalizeStr('Executada')) {
            cityMap[matchedKey].osExec += (item.volume || 1);
          }
          if (item.fullDate instanceof Date && !isNaN(item.fullDate.getTime())) {
            cityMap[matchedKey].dates.add(item.fullDate.toDateString());
          }
        }
      });
      return cityMap;
    };

    const currentMap = getCityStats(finalFilteredData);
    const previousMap = getCityStats(prevFilteredData);

    const diasMes = metrics.details.diasMes || 30;
    const diasTrabGlobal = metrics.details.diasTrab || 1;

    // Days worked for previous month
    const prevDayNumbers = prevFilteredData.map(d => d.fullDate instanceof Date ? d.fullDate.getDate() : 0).filter(d => d > 0);
    const prevDiasTrab = prevDayNumbers.length > 0 ? Math.max(...prevDayNumbers) : 1;
    const prevReferenceDate = (prevFilteredData.length > 0 && prevFilteredData[0].fullDate instanceof Date) ? prevFilteredData[0].fullDate : new Date();
    const prevDiasMes = new Date(prevReferenceDate.getFullYear(), prevReferenceDate.getMonth() + 1, 0).getDate();

    return Object.keys(currentMap).map(key => {
      const curr = currentMap[key];
      const prev = previousMap[key];

      const projOsAT1 = (curr.osExec / diasTrabGlobal) * diasMes;
      const at1 = curr.base > 0 ? (projOsAT1 / curr.base) * 100 : 0;

      let prevAt1 = 0;
      if (prev && prev.base > 0) {
        const prevProj = (prev.osExec / prevDiasTrab) * prevDiasMes;
        prevAt1 = (prevProj / prev.base) * 100;
      }

      return { 
        name: curr.name, 
        current: Number(at1.toFixed(2)),
        previous: Number(prevAt1.toFixed(2)),
        value: Number(at1.toFixed(2)), 
        base: curr.base, 
        osExec: curr.osExec,
        osExecPrev: prev ? prev.osExec : 0
      };
    })
    .filter(item => item.current > 0 || item.previous > 0)
    .sort((a, b) => b.current - a.current);
  }, [finalFilteredData, momFilteredData, baseCidadeData, filters, metrics, comparisonMonths, selectedNodes, selectedBaixas]);

  const COLORS = [CLARO_RED, '#666666', '#999999'];

  // Excel Import Handler
  // Excel Import Handler
  const processExcelData = (bstr: any) => {
    try {
      const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
      
      // 1. Read Main Data
      const analiticoSheetName = wb.SheetNames.find(name => {
        const n = normalizeStr(name);
        return n.includes('ANALITICO') || n.includes('VISITAS') || n.includes('DADOS') || n.includes('GERAL');
      }) || wb.SheetNames.find(name => !normalizeStr(name).includes('BASE')) || wb.SheetNames[0];
      
      const ws = wb.Sheets[analiticoSheetName];
      if (!ws) {
        setIsImporting(false);
        setImportError(`Planilha "${analiticoSheetName}" não encontrada.`);
        return;
      }
      const rawRows = XLSX.utils.sheet_to_json(ws) as any[];

      if (rawRows.length === 0) {
        setIsImporting(false);
        setImportError("Arquivo vazio ou sem dados válidos.");
        return;
      }

      setImportProgress(5);

      // 2. Read BASE Data
      const baseCidadeSheetName = wb.SheetNames.find(name => {
        const n = normalizeStr(name);
        return n === 'BASE' || n.includes('BASE') || n.includes('REFERENCIA') || n.includes('METAS') || n.includes('CIDADE') || n.includes('CLIENTE') || n.includes('TOTAL');
      });
      
      let importedBaseCidade: BaseCidadeData[] = [];
      if (baseCidadeSheetName) {
        const wsBase = wb.Sheets[baseCidadeSheetName];
        const baseRawData = XLSX.utils.sheet_to_json(wsBase, { header: 1 }) as any[][];
        
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(baseRawData.length, 20); i++) {
          const row = baseRawData[i];
          if (row && row.some(cell => {
            const s = normalizeStr(String(cell || ''));
            return s === 'CIDADE' || s === 'MUNICIPIO' || s === 'LOCALIDADE' || s === 'BASE';
          })) {
            headerRowIndex = i;
            break;
          }
        }

        const startIdx = headerRowIndex === -1 ? 0 : headerRowIndex + 1;
        const headerRow = headerRowIndex === -1 ? [] : baseRawData[headerRowIndex];

        let colIdxCidade = 0;
        let colIdxHFC = -1;
        let colIdxGPON = -1;
        let colIdxHibrido = -1;
        let colIdxOutros = -1;
        let colIdxTotalBase = -1;

        if (headerRow.length > 0) {
          headerRow.forEach((cell, idx) => {
            const s = normalizeStr(String(cell || ''));
            if (s === 'CIDADE' || s === 'MUNICIPIO' || s === 'LOCALIDADE') colIdxCidade = idx;
            if (s.includes('HFC')) colIdxHFC = idx;
            if (s.includes('GPON') || s.includes('FIBRA') || s.includes('G PON')) colIdxGPON = idx;
            if (s.includes('HIBRID') || s.includes('HIBRIDO')) colIdxHibrido = idx;
            if (s.includes('OUTRO') || s.includes('OUTRA')) colIdxOutros = idx;
            if (s === 'BASE' || s === 'CLIENTES' || s === 'TOTAL' || s === 'BASE_CLIENTES') colIdxTotalBase = idx;
          });
        }

        // Fallbacks if columns not found by name
        if (colIdxHFC === -1 && colIdxTotalBase !== -1) colIdxHFC = colIdxTotalBase;
        if (colIdxHFC === -1) colIdxHFC = 2;
        if (colIdxGPON === -1) colIdxGPON = 3;
        if (colIdxHibrido === -1) colIdxHibrido = 4;
        if (colIdxOutros === -1) colIdxOutros = 5;

        importedBaseCidade = baseRawData.slice(startIdx).flatMap((row) => {
          if (!row || row.length <= colIdxCidade) return [];
          
          const rawCidade = String(row[colIdxCidade] || '').trim();
          if (!rawCidade || normalizeStr(rawCidade) === 'CIDADE' || normalizeStr(rawCidade) === 'MUNICIPIO') return [];
          
          const cidade = cleanCityName(rawCidade);
          
          const results = [];
          if (colIdxHFC !== -1) results.push({ cidade, tecnologia: 'HFC' as Technology, base: parseExcelNumber(row[colIdxHFC]) });
          if (colIdxGPON !== -1 && colIdxGPON !== colIdxHFC) results.push({ cidade, tecnologia: 'GPON' as Technology, base: parseExcelNumber(row[colIdxGPON]) });
          if (colIdxHibrido !== -1 && colIdxHibrido !== colIdxHFC) results.push({ cidade, tecnologia: 'HÍBRIDO' as Technology, base: parseExcelNumber(row[colIdxHibrido]) });
          if (colIdxOutros !== -1 && colIdxOutros !== colIdxHFC) results.push({ cidade, tecnologia: 'OUTROS' as Technology, base: parseExcelNumber(row[colIdxOutros]) });
          
          return results;
        });
      }

      // 3. Process Main Data in Chunks
      const mappedData: VisitData[] = [];
      const chunkSize = 1000;
      let currentIndex = 0;

      const processChunk = () => {
        try {
          const end = Math.min(currentIndex + chunkSize, rawRows.length);
          for (let i = currentIndex; i < end; i++) {
            const row = rawRows[i];
            
            let mesValue = 'N/A';
            let fullDateValue = new Date();
            const rawDate = row.DATA_NOTA || row.data_nota || row.Data || row.DATA || row.DATA_EXECUCAO || row.DATA_FIM || row.DT_EXEC;
            
            if (rawDate) {
              let date: Date;
              if (rawDate instanceof Date) {
                date = rawDate;
              } else if (typeof rawDate === 'number') {
                date = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
              } else {
                date = new Date(rawDate);
              }

              if (!isNaN(date.getTime())) {
                const year = date.getUTCFullYear();
                const month = date.getUTCMonth();
                const day = date.getUTCDate();
                fullDateValue = new Date(year, month, day);
                
                const months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
                mesValue = months[fullDateValue.getMonth()];
              }
            }

            let techValue = String(row.DSC_SEG_PRODUTO || row.tecnologia || row.TECNOLOGIA || row.PRODUTO || row.SEGMENTO || 'HFC').toUpperCase().trim();
            if (techValue.includes('HIBRID') || techValue.includes('HÍBRID')) techValue = 'HÍBRIDO';
            else if (techValue.includes('GPON') || techValue.includes('FIBRA') || techValue.includes('FTTH')) techValue = 'GPON';
            else if (techValue.includes('HFC') || techValue.includes('COAXIAL') || techValue.includes('CABO')) techValue = 'HFC';
            else techValue = 'HFC';

            const nota = Number(row.NOTA_AT1 || row.nota_at1 || row.NOTA || row.AT1 || row.notaAT1 || row.NOTA_FINAL || row.PONTUACAO || row.SCORE || 0);
            
            const rawStatus = String(row.NM_STATUS_OS || row.STATUS || row.status || row.SITUACAO || row.situação || row.DSC_STATUS || row.NM_SITUACAO || '').toUpperCase().trim();
            let statusValue: 'Executada' | 'Cancelada' = 'Cancelada';
            if (rawStatus.includes('EXEC') || rawStatus.includes('CONCLU') || rawStatus.includes('FINALIZ') || rawStatus.includes('ENCERRADA') || rawStatus.includes('REALIZADA') || rawStatus.includes('OK') || rawStatus.includes('FECHADA') || nota > 0) {
              statusValue = 'Executada';
            }

            const rawCidade = String(row.MUNICIPIO || row.municipio || row.cidade || row.CIDADE || row.LOCALIDADE || 'N/A').trim();
            const cidade = cleanCityName(rawCidade);

            const rawGrupoBaixa = String(
              row['GRUPO BAIXA'] || 
              row.GRUPO_BAIXA || 
              row.grupo_baixa || 
              row.GRUPO || 
              row.DS_GRUPO_BAIXA || 
              row.NM_GRUPO_BAIXA || 
              row.GRUPO_ENCERRAMENTO || 
              row.CATEGORIA_BAIXA || 
              'N/A'
            ).trim();

            const descBaixa = String(
              row['DESCRIÇÃO CÓDIGO'] || 
              row.DESCRICAO_CODIGO || 
              row.DESCRICAO_DO_CODIGO || 
              row.DESC_BAIXA || 
              row.DESCRIÇÃO_BAIXA || 
              row.DS_COD_BAIXA ||
              row.DSC_COD_BAIXA ||
              row.DS_BAIXA ||
              row.DSC_BAIXA ||
              row['DESCRIÇÃO'] ||
              row.DESCRICAO ||
              'N/A'
            ).trim();

            mappedData.push({
              id: `import-${i}`,
              mes: mesValue,
              fullDate: fullDateValue,
              cidade: cidade,
              area: String(row.AREA_DESPACHO || row.area_despacho || row.area || row.AREA || row.SETOR || 'N/A').trim(),
              expurgo: row.EXPURGO_AT1 === 'Sim' || row.EXPURGO_AT1 === 'S' || row.EXPURGO_AT1 === true || row.expurgo === true,
              tecnologia: techValue as Technology,
              status: statusValue,
              notaAT1: nota,
              node: String(row.CD_NODE || row.node || row.NODE || row.ESTACAO || 'N/A').trim(),
              cdBaixa: String(row.CD_BAIXA || row.cd_baixa || row.BAIXA || row.COD_BAIXA || 'N/A').trim(),
              descriptionBaixa: descBaixa,
              grupoBaixa: rawGrupoBaixa,
              terminal: String(row.TERMINAL || row.terminal || row.Terminal || row.CD_TERMINAL || row.NM_TERMINAL || row.NOME_TERMINAL || row.Terminal_ID || 'N/A').trim(),
              volume: Number(row.VOLUME || row.volume || row.Volume || row.QUANTIDADE || row.QTD || 1)
            });
          }

          currentIndex = end;
          const progress = Math.max(5, Math.min(99, Math.round((currentIndex / rawRows.length) * 100)));
          setImportProgress(progress);

          if (currentIndex < rawRows.length) {
            setTimeout(processChunk, 0);
          } else {
            setImportProgress(100);
            if (mappedData.length > 0) {
              setBaseData(mappedData);
              if (importedBaseCidade.length > 0) {
                setBaseCidadeData(importedBaseCidade);
              }
              setFilters({
                mes: ['Todos'],
                semana: ['Todos'],
                cidade: ['Todos'],
                area: ['Todos'],
                expurgo: 'Todos',
                tecnologia: ['Todos'],
                grupoBaixa: ['Todos'],
                startDate: '',
                endDate: ''
              });
            } else {
              setImportError('O arquivo parece estar vazio ou em formato inválido.');
            }
            setIsImporting(false);
          }
        } catch (err) {
          console.error('Erro ao processar chunk:', err);
          setImportError('Erro ao processar os dados do arquivo.');
          setIsImporting(false);
        }
      };

      processChunk();

    } catch (err) {
      setImportError('Erro ao processar o arquivo Excel. Verifique o formato.');
      setIsImporting(false);
      console.error(err);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsImporting(true);
    setImportProgress(0);
    setImportError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      processExcelData(evt.target?.result);
    };
    reader.onerror = () => {
      setImportError('Erro ao ler o arquivo.');
      setIsImporting(false);
    };
    reader.readAsBinaryString(file);
  };

  const handleGithubLoad = async (urlToLoad?: string | React.MouseEvent) => {
    const preConfiguredUrl = getGithubAt1Url();
    const targetUrl = (typeof urlToLoad === 'string' ? urlToLoad : null) || githubUrl || preConfiguredUrl;
    if (!targetUrl) return;

    setIsImporting(true);
    setImportProgress(20);
    setImportError(null);

    try {
      setImportProgress(40);
      const arrayBuffer = await fetchGithubFileArrayBuffer(targetUrl);
      setImportProgress(70);
      processExcelData(arrayBuffer);
      setShowGithubInput(false);
      setGithubUrl('');
    } catch (err: any) {
      setImportError(err.message || 'Erro ao carregar do GitHub.');
      setIsImporting(false);
    }
  };

  // Auto-load disabled on mount per user request. Data loads only on clicking 'Sincronizar GitHub'
  React.useEffect(() => {
    // Automatic loading disabled on mount; user loads manually via action buttons
  }, []);

  return (
    <ErrorBoundary>
      <div className="min-h-screen bg-[#F5F5F5] text-[#333333] font-sans flex flex-col md:flex-row">
        {/* Left Sidebar Navigation - Compact & Collapsible */}
        <aside className={cn(
          "bg-white border-b md:border-b-0 md:border-r border-slate-200 p-2.5 md:p-3 md:min-h-screen md:sticky md:top-0 flex flex-col justify-between shrink-0 shadow-xs z-30 transition-all duration-300 relative",
          isSidebarCollapsed ? "w-full md:w-16" : "w-full md:w-44 lg:w-48"
        )}>
          <div className="space-y-4">
            {/* Brand Header */}
            <div className={cn(
              "flex items-center justify-between gap-2 pb-2 border-b border-slate-100",
              isSidebarCollapsed ? "md:flex-col md:gap-2 md:pb-2.5 md:items-center" : ""
            )}>
              <div 
                onClick={() => {
                  if (isSidebarCollapsed) {
                    setIsSidebarCollapsed(false);
                    try { localStorage.setItem('CLARO_SIDEBAR_COLLAPSED', 'false'); } catch (e) {}
                  }
                }}
                className={cn(
                  "flex items-center gap-2.5 min-w-0",
                  isSidebarCollapsed ? "cursor-pointer" : ""
                )}
                title={isSidebarCollapsed ? "Clique para abrir menu" : undefined}
              >
                <div className="w-8 h-8 bg-[#EE1D23] rounded-xl flex items-center justify-center shadow-md shadow-red-500/25 relative overflow-hidden group shrink-0">
                  <Signal className="text-white/30 w-5 h-5 absolute -bottom-1 -right-1 rotate-12" />
                  <Wrench className="text-white w-4 h-4 relative z-10 drop-shadow-xs" />
                </div>
                {!isSidebarCollapsed && (
                  <div className="min-w-0">
                    <h1 className="text-sm font-black tracking-tight text-[#EE1D23] uppercase italic leading-none">CLARO</h1>
                    <p className="text-[9px] text-slate-500 font-bold uppercase tracking-wider mt-0.5 truncate">Indicadores</p>
                  </div>
                )}
              </div>

              {/* Collapse/Expand Toggle Button (Header) */}
              <button
                onClick={() => {
                  const nextState = !isSidebarCollapsed;
                  setIsSidebarCollapsed(nextState);
                  try {
                    localStorage.setItem('CLARO_SIDEBAR_COLLAPSED', String(nextState));
                  } catch (e) {}
                }}
                title={isSidebarCollapsed ? "Abrir menu lateral (Expandir)" : "Fechar menu lateral (Recolher)"}
                className={cn(
                  "hidden md:flex items-center justify-center rounded-lg transition-all cursor-pointer shrink-0",
                  isSidebarCollapsed 
                    ? "w-8 h-7 bg-red-50 text-[#EE1D23] hover:bg-red-100 border border-red-100 shadow-2xs" 
                    : "w-7 h-7 text-slate-400 hover:text-slate-700 hover:bg-slate-100"
                )}
              >
                {isSidebarCollapsed ? (
                  <ChevronRight className="w-4 h-4 stroke-[2.5]" />
                ) : (
                  <ChevronLeft className="w-4 h-4 stroke-[2.5]" />
                )}
              </button>
            </div>

            {/* Navigation Menu */}
            <div>
              {!isSidebarCollapsed && (
                <p className="text-[9px] font-black uppercase tracking-widest text-slate-400 mb-1.5 px-2">Painéis</p>
              )}
              <nav className="flex md:flex-col gap-1 overflow-x-auto md:overflow-x-visible pb-1 md:pb-0 scrollbar-none">
                {/* AT1 */}
                <button
                  onClick={() => setActiveTab('dashboard')}
                  title="AT1"
                  className={cn(
                    "flex items-center gap-2.5 px-2.5 py-2 rounded-xl font-black uppercase italic text-[11px] tracking-wide transition-all active:scale-95 whitespace-nowrap w-full text-left cursor-pointer",
                    isSidebarCollapsed ? "md:justify-center md:px-0" : "",
                    activeTab === 'dashboard'
                      ? "bg-[#EE1D23] text-white shadow-md shadow-red-500/25"
                      : "bg-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-100/80"
                  )}
                >
                  <BarChart3 className={cn("w-4 h-4 shrink-0", activeTab === 'dashboard' ? "text-white" : "text-[#EE1D23]")} />
                  {!isSidebarCollapsed && <span>AT1</span>}
                </button>

                {/* OUTAGE */}
                <button
                  onClick={() => setActiveTab('outage')}
                  title="OUTAGE"
                  className={cn(
                    "flex items-center gap-2.5 px-2.5 py-2 rounded-xl font-black uppercase italic text-[11px] tracking-wide transition-all active:scale-95 whitespace-nowrap w-full text-left cursor-pointer",
                    isSidebarCollapsed ? "md:justify-center md:px-0" : "",
                    activeTab === 'outage'
                      ? "bg-[#EE1D23] text-white shadow-md shadow-red-500/25"
                      : "bg-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-100/80"
                  )}
                >
                  <Radio className={cn("w-4 h-4 shrink-0", activeTab === 'outage' ? "text-white" : "text-[#EE1D23]")} />
                  {!isSidebarCollapsed && <span>OUTAGE</span>}
                </button>

                {/* AT5 */}
                <button
                  onClick={() => setActiveTab('at5')}
                  title="AT5"
                  className={cn(
                    "flex items-center gap-2.5 px-2.5 py-2 rounded-xl font-black uppercase italic text-[11px] tracking-wide transition-all active:scale-95 whitespace-nowrap w-full text-left cursor-pointer",
                    isSidebarCollapsed ? "md:justify-center md:px-0" : "",
                    activeTab === 'at5'
                      ? "bg-[#EE1D23] text-white shadow-md shadow-red-500/25"
                      : "bg-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-100/80"
                  )}
                >
                  <FileSpreadsheet className={cn("w-4 h-4 shrink-0", activeTab === 'at5' ? "text-white" : "text-[#EE1D23]")} />
                  {!isSidebarCollapsed && <span>AT5</span>}
                </button>

                {/* LOG */}
                <button
                  onClick={() => setActiveTab('log')}
                  title="LOG"
                  className={cn(
                    "flex items-center gap-2.5 px-2.5 py-2 rounded-xl font-black uppercase italic text-[11px] tracking-wide transition-all active:scale-95 whitespace-nowrap w-full text-left cursor-pointer",
                    isSidebarCollapsed ? "md:justify-center md:px-0" : "",
                    activeTab === 'log'
                      ? "bg-[#EE1D23] text-white shadow-md shadow-red-500/25"
                      : "bg-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-100/80"
                  )}
                >
                  <Activity className={cn("w-4 h-4 shrink-0", activeTab === 'log' ? "text-white" : "text-[#EE1D23]")} />
                  {!isSidebarCollapsed && <span>LOG</span>}
                </button>

                {/* REVISITA 30D */}
                <button
                  onClick={() => setActiveTab('revisita30d')}
                  title="REVISITA 30D"
                  className={cn(
                    "flex items-center gap-2.5 px-2.5 py-2 rounded-xl font-black uppercase italic text-[11px] tracking-wide transition-all active:scale-95 whitespace-nowrap w-full text-left cursor-pointer",
                    isSidebarCollapsed ? "md:justify-center md:px-0" : "",
                    activeTab === 'revisita30d'
                      ? "bg-[#EE1D23] text-white shadow-md shadow-red-500/25"
                      : "bg-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-100/80"
                  )}
                >
                  <RotateCcw className={cn("w-4 h-4 shrink-0", activeTab === 'revisita30d' ? "text-white" : "text-[#EE1D23]")} />
                  {!isSidebarCollapsed && <span className="truncate">REVISITA 30D</span>}
                </button>

                {/* DIÁRIO DE BORDO */}
                <button
                  onClick={() => setActiveTab('logbook')}
                  title="Diário de Bordo"
                  className={cn(
                    "flex items-center gap-2.5 px-2.5 py-2 rounded-xl font-black uppercase italic text-[11px] tracking-wide transition-all active:scale-95 whitespace-nowrap w-full text-left cursor-pointer",
                    isSidebarCollapsed ? "md:justify-center md:px-0" : "",
                    activeTab === 'logbook'
                      ? "bg-[#EE1D23] text-white shadow-md shadow-red-500/25"
                      : "bg-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-100/80"
                  )}
                >
                  <BookOpen className={cn("w-4 h-4 shrink-0", activeTab === 'logbook' ? "text-white" : "text-[#EE1D23]")} />
                  {!isSidebarCollapsed && <span className="truncate">Diário de Bordo</span>}
                </button>
              </nav>
            </div>
          </div>

          {/* Footer Info / Data status */}
          <div className={cn(
            "hidden md:flex flex-col gap-2 pt-3 border-t border-slate-100 mt-4",
            isSidebarCollapsed ? "items-center" : ""
          )}>
            {!isSidebarCollapsed ? (
              <div className="flex items-center gap-2 text-[10px] font-bold text-slate-500 bg-slate-50 px-2.5 py-2 rounded-xl border border-slate-200/60">
                <Calendar className="w-3 h-3 text-slate-400 shrink-0" />
                <span className="truncate">
                  {lastUpdateDate 
                    ? `Até ${lastUpdateDate.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' })}` 
                    : new Date().toLocaleDateString('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' })}
                </span>
              </div>
            ) : (
              <div 
                title={lastUpdateDate ? `Até ${lastUpdateDate.toLocaleDateString('pt-BR')}` : new Date().toLocaleDateString('pt-BR')} 
                className="w-8 h-8 rounded-lg bg-slate-50 border border-slate-200/60 flex items-center justify-center text-slate-400"
              >
                <Calendar className="w-3.5 h-3.5" />
              </div>
            )}
            
            {baseData.length > 0 && activeTab === 'dashboard' && (
              <button 
                onClick={() => {
                  setBaseData([]);
                  setBaseCidadeData([]);
                }}
                title="Limpar Dados AT1"
                className={cn(
                  "flex items-center justify-center gap-1.5 bg-slate-100 hover:bg-red-50 hover:text-red-600 text-slate-600 font-bold py-1.5 px-2 rounded-xl transition-all text-[11px] active:scale-95 cursor-pointer",
                  isSidebarCollapsed ? "w-8 h-8 px-0" : "w-full"
                )}
              >
                <Trash className="w-3.5 h-3.5 shrink-0" />
                {!isSidebarCollapsed && <span>Limpar AT1</span>}
              </button>
            )}
          </div>
        </aside>

        {/* Main Content Area */}
        <div className="flex-1 min-w-0 p-3 sm:p-5 lg:p-6 overflow-y-auto">
          <div className={activeTab === 'dashboard' ? 'block' : 'hidden'}>
            {/* Header */}
            <header className="max-w-7xl mx-auto mb-8">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                  <div className="flex items-center gap-4">
                    <div className="w-14 h-14 bg-[#EE1D23] rounded-2xl flex items-center justify-center shadow-xl shadow-red-500/20 relative overflow-hidden group">
                      <div className="absolute inset-0 bg-gradient-to-br from-white/20 to-transparent opacity-0 group-hover:opacity-100 transition-opacity" />
                      <Signal className="text-white/30 w-10 h-10 absolute -bottom-2 -right-2 rotate-12" />
                      <Wrench className="text-white w-7 h-7 relative z-10 drop-shadow-md" />
                    </div>
                    <div>
                      <h1 className="text-2xl sm:text-3xl font-black tracking-tighter text-[#EE1D23] uppercase italic leading-none">DASHBOARD INDICADORES DE MANUTENÇÃO</h1>
                      <p className="text-slate-500 font-bold text-xs mt-1 uppercase tracking-widest opacity-80">Manutenção & Performance Técnica</p>
                    </div>
                  </div>
                  
                  <div className="flex flex-wrap items-center gap-3">
                    {baseData.length > 0 && (
                      <>
                        <button
                          onClick={() => handleGithubLoad(getGithubAt1Url())}
                          className="flex items-center gap-2 bg-[#EE1D23] hover:bg-red-600 text-white font-black py-2 px-4 rounded-xl transition-all shadow-md shadow-red-500/15 active:scale-95 uppercase italic text-xs cursor-pointer"
                          title="Sincronizar planilha AT1 com GitHub"
                        >
                          <Activity className="w-3.5 h-3.5" />
                          <span>Sincronizar GitHub</span>
                        </button>
                        <button
                          onClick={() => fileInputRef.current?.click()}
                          className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-700 font-bold py-2 px-4 rounded-xl border border-slate-200 transition-all shadow-xs active:scale-95 uppercase italic text-xs cursor-pointer"
                          title="Importar planilha Excel"
                        >
                          <Upload className="w-3.5 h-3.5 text-[#EE1D23]" />
                          <span>Importar Excel</span>
                        </button>
                        <button 
                          onClick={() => {
                            setBaseData([]);
                            setBaseCidadeData([]);
                          }}
                          className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold py-2 px-4 rounded-xl transition-all shadow-xs active:scale-95 cursor-pointer text-xs uppercase"
                        >
                          <Trash className="w-3.5 h-3.5" />
                          <span>Limpar</span>
                        </button>
                      </>
                    )}
                    <div className="flex items-center gap-2 text-xs font-bold text-slate-500 bg-white px-4 py-2.5 rounded-full shadow-sm border border-slate-200">
                      <Calendar className="w-3.5 h-3.5" />
                      <span>
                        {lastUpdateDate 
                          ? `Atualizado até ${lastUpdateDate.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' })}` 
                          : new Date().toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' })}
                      </span>
                    </div>
                  </div>
                </div>

                {importError && (
                  <motion.div 
                    initial={{ opacity: 0, y: -10 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="mt-4 p-3 bg-red-50 border border-red-200 rounded-xl flex items-center gap-2 text-red-600 text-sm font-medium"
                  >
                    <AlertCircle className="w-4 h-4" />
                    {importError}
                  </motion.div>
                )}
              </header>

              {isImporting && (
        <div className="fixed inset-0 z-[100] bg-white/80 backdrop-blur-sm flex flex-col items-center justify-center p-6">
          <div className="w-full max-w-md bg-white p-8 rounded-[32px] shadow-2xl border border-slate-100 text-center">
            <div className="w-20 h-20 bg-red-50 rounded-2xl flex items-center justify-center mb-6 mx-auto animate-pulse">
              <FileSpreadsheet className="w-10 h-10 text-[#EE1D23]" />
            </div>
            <h3 className="text-2xl font-black text-[#333333] uppercase italic tracking-tighter mb-2">Importando Dados</h3>
            <p className="text-slate-500 font-bold mb-8 italic">Processando arquivo analítico...</p>
            
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
              Cancelar Importação
            </button>
          </div>
        </div>
      )}

      {baseData.length === 0 ? (
        <section className="max-w-2xl mx-auto mt-8 sm:mt-12 px-4">
          <motion.div 
            initial={{ opacity: 0, scale: 0.96 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-white p-6 sm:p-8 rounded-3xl shadow-xl border border-slate-100 text-center flex flex-col items-center"
          >
            <div className="w-14 h-14 bg-red-50 rounded-2xl flex items-center justify-center mb-4 shadow-2xs">
              <FileSpreadsheet className="w-7 h-7 text-[#EE1D23]" />
            </div>
            <h2 className="text-xl sm:text-2xl font-black text-[#333333] uppercase italic tracking-tight mb-2">
              Bem-vindo ao Dashboard Indicadores de Manutenção
            </h2>
            <p className="text-slate-500 font-bold text-xs max-w-md mb-6 leading-relaxed uppercase tracking-wide opacity-65">
              Clique no botão para sincronizar com o GitHub, importar uma planilha Excel ou carregar dados de exemplo.
            </p>

            <div className="flex flex-wrap items-center justify-center gap-3 mb-6">
              <button
                onClick={() => handleGithubLoad(getGithubAt1Url())}
                className="flex items-center gap-2 bg-[#EE1D23] hover:bg-red-600 text-white font-black py-2.5 px-5 rounded-xl transition-all shadow-md shadow-red-500/15 active:scale-95 uppercase italic text-xs"
              >
                <Activity className="w-3.5 h-3.5" />
                <span>Sincronizar GitHub</span>
              </button>

              <button
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-800 font-black py-2.5 px-5 rounded-xl border border-slate-200 transition-all shadow-2xs active:scale-95 uppercase italic text-xs"
              >
                <Upload className="w-3.5 h-3.5 text-[#EE1D23]" />
                <span>Importar Excel</span>
              </button>

              <button
                onClick={() => {
                  setBaseData(MOCK_DATA);
                  setBaseCidadeData(MOCK_BASE_CIDADE);
                }}
                className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2.5 px-4 rounded-xl transition-all shadow-2xs active:scale-95 uppercase italic text-xs"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-slate-500" />
                <span>Dados Exemplo</span>
              </button>
            </div>
            
            <div className="mt-4 grid grid-cols-1 sm:grid-cols-3 gap-4 w-full border-t border-slate-100 pt-5">
              <div className="flex flex-col items-center">
                <div className="w-8 h-8 bg-slate-50 rounded-lg flex items-center justify-center mb-1.5">
                  <BarChart3 className="w-4 h-4 text-slate-400" />
                </div>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Volume Diário</p>
              </div>
              <div className="flex flex-col items-center">
                <div className="w-8 h-8 bg-slate-50 rounded-lg flex items-center justify-center mb-1.5">
                  <Activity className="w-4 h-4 text-slate-400" />
                </div>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Performance G1</p>
              </div>
              <div className="flex flex-col items-center">
                <div className="w-8 h-8 bg-slate-50 rounded-lg flex items-center justify-center mb-1.5">
                  <MapPin className="w-4 h-4 text-slate-400" />
                </div>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Visão por Cidade</p>
              </div>
            </div>
          </motion.div>
        </section>
      ) : (
        <>
          {/* Filters Bar */}
          <section className="max-w-7xl mx-auto mb-8">
        <div className="bg-white p-6 rounded-2xl shadow-md border-t-4 border-[#EE1D23]">
          <div className="flex items-center gap-2 mb-6 text-[#333333] font-black uppercase italic tracking-tight">
            <Filter className="w-4 h-4 text-[#EE1D23]" />
            <h2>Filtros de Pesquisa</h2>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
            <MultiFilterSelect 
              label="Mês" 
              icon={<Calendar className="w-3.5 h-3.5" />}
              value={filters.mes}
              options={filterOptions.meses}
              onChange={(v) => setFilters(f => ({ ...f, mes: v }))}
            />
            <MultiFilterSelect 
              label="Semana" 
              icon={<Clock className="w-3.5 h-3.5" />}
              value={filters.semana}
              options={filterOptions.semanas}
              onChange={(v) => setFilters(f => ({ ...f, semana: v }))}
            />
            <MultiFilterSelect 
              label="Cidade" 
              icon={<MapPin className="w-3.5 h-3.5" />}
              value={filters.cidade}
              options={filterOptions.cidades}
              onChange={(v) => setFilters(f => ({ ...f, cidade: v }))}
            />
            <MultiFilterSelect 
              label="Área" 
              icon={<Activity className="w-3.5 h-3.5" />}
              value={filters.area}
              options={filterOptions.areas}
              onChange={(v) => setFilters(f => ({ ...f, area: v }))}
            />
            <FilterSelect 
              label="Expurgo" 
              icon={<Trash className="w-3.5 h-3.5" />}
              value={filters.expurgo}
              options={filterOptions.expurgos}
              onChange={(v) => setFilters(f => ({ ...f, expurgo: v }))}
            />
            <MultiFilterSelect 
              label="Tecnologia" 
              icon={<Cpu className="w-3.5 h-3.5" />}
              value={filters.tecnologia}
              options={filterOptions.tecnologias}
              onChange={(v) => setFilters(f => ({ ...f, tecnologia: v }))}
            />
            <MultiFilterSelect 
              label="Grupo de Baixa" 
              icon={<Filter className="w-3.5 h-3.5" />}
              value={filters.grupoBaixa}
              options={filterOptions.gruposBaixa}
              onChange={(v) => setFilters(f => ({ ...f, grupoBaixa: v }))}
            />
            <div className="flex flex-col gap-1.5">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Início</label>
              <input 
                type="date" 
                value={filters.startDate}
                onChange={(e) => setFilters(f => ({ ...f, startDate: e.target.value }))}
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-xs font-bold focus:ring-2 focus:ring-[#EE1D23] focus:border-transparent outline-none transition-all"
              />
            </div>
            <div className="flex flex-col gap-1.5">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Fim</label>
              <input 
                type="date" 
                value={filters.endDate}
                onChange={(e) => setFilters(f => ({ ...f, endDate: e.target.value }))}
                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 text-xs font-bold focus:ring-2 focus:ring-[#EE1D23] focus:border-transparent outline-none transition-all"
              />
            </div>
          </div>
        </div>
      </section>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto space-y-8">
        {/* Volume Metrics */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <MetricCard 
            title="Total de Visitas" 
            value={metrics.total} 
            icon={<BarChart3 />}
            color="dark"
          />
          <MetricCard 
            title="Visitas Executadas" 
            value={metrics.executada} 
            icon={<CheckCircle2 />}
            color="red"
            percentage={metrics.total > 0 ? (metrics.executada / metrics.total) * 100 : 0}
          />
          <MetricCard 
            title="Visitas Canceladas" 
            value={metrics.cancelada} 
            icon={<XCircle />}
            color="gray"
            percentage={metrics.total > 0 ? (metrics.cancelada / metrics.total) * 100 : 0}
          />
        </div>

        {/* G1 Scores */}
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-8 pb-12">
          {/* G1 Quality Gauge */}
          <div className="lg:col-span-1 bg-white p-8 rounded-3xl shadow-md border border-slate-100 flex flex-col items-center justify-center text-center">
            <h3 className="text-[#EE1D23] font-black uppercase italic mb-6 tracking-tight">Nota G1</h3>
            <div className="relative flex items-center justify-center">
              <svg className="w-44 h-44">
                <circle
                  className="text-slate-100"
                  strokeWidth="12"
                  stroke="currentColor"
                  fill="transparent"
                  r="75"
                  cx="88"
                  cy="88"
                />
                <circle
                  className="text-[#EE1D23] transition-all duration-1000 ease-out"
                  strokeWidth="12"
                  strokeDasharray={471}
                  strokeDashoffset={471 - (471 * Math.min(metrics.projG1Total, 100)) / 100}
                  strokeLinecap="round"
                  stroke="currentColor"
                  fill="transparent"
                  r="75"
                  cx="88"
                  cy="88"
                  transform="rotate(-90 88 88)"
                />
              </svg>
              <div className="absolute inset-0 flex flex-col items-center justify-center">
                <span className="text-4xl font-black text-[#333333]">{formatPercent(metrics.projG1Total)}</span>
                <span className="text-slate-400 font-bold text-[10px] uppercase tracking-widest">G1 Projetado</span>
                {prevMetrics && (
                  <div className="mt-1 flex flex-col items-center">
                    <span className="text-[10px] font-black text-slate-300 uppercase tracking-tighter">vs {formatPercent(prevMetrics.projG1Total)}</span>
                    <span className="text-[8px] font-bold text-slate-300 uppercase">{comparisonMonths.previous}</span>
                  </div>
                )}
              </div>
            </div>
            <div className="mt-6 w-full text-left space-y-1">
              <div className="flex justify-between text-[10px] font-bold">
                <span className="text-slate-400 uppercase">Média Qualidade:</span>
                <span className="text-[#333333]">{formatDecimal(metrics.avgQualityTotal)}</span>
              </div>
            </div>
          </div>

          {/* Projection Gauge */}
          <div className="lg:col-span-1 bg-white p-8 rounded-3xl shadow-md border border-slate-100 flex flex-col items-center justify-center text-center">
            <h3 className="text-slate-600 font-black uppercase italic mb-6 tracking-tight">AT1 Projetado (%)</h3>
            <div className="relative flex items-center justify-center">
              <svg className="w-44 h-44">
                <circle
                  className="text-slate-100"
                  strokeWidth="12"
                  stroke="currentColor"
                  fill="transparent"
                  r="75"
                  cx="88"
                  cy="88"
                />
                <circle
                  className="text-slate-600 transition-all duration-1000 ease-out"
                  strokeWidth="12"
                  strokeDasharray={471}
                  strokeDashoffset={471 - (471 * Math.min(metrics.projAT1Total, 100)) / 100}
                  strokeLinecap="round"
                  stroke="currentColor"
                  fill="transparent"
                  r="75"
                  cx="88"
                  cy="88"
                  transform="rotate(-90 88 88)"
                />
              </svg>
              <div className="absolute inset-0 flex flex-col items-center justify-center">
                <span className="text-4xl font-black text-[#333333]">{formatPercent(metrics.projAT1Total)}</span>
                <span className="text-slate-400 font-bold text-[10px] uppercase tracking-widest">Projetado AT1</span>
                {prevMetrics && (
                  <div className="mt-1 flex flex-col items-center">
                    <span className="text-[10px] font-black text-slate-400/50 uppercase tracking-tighter">vs {formatPercent(prevMetrics.projAT1Total)}</span>
                    <span className="text-[8px] font-bold text-slate-400/50 uppercase">{comparisonMonths.previous}</span>
                  </div>
                )}
              </div>
            </div>
            
            <div className="mt-6 w-full text-left space-y-1.5 bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div className="flex justify-between text-[9px] font-black uppercase tracking-tighter">
                <span className="text-slate-400">Dias Trab:</span>
                <span className="text-[#333333]">{metrics.details.diasTrab} / {metrics.details.diasMes}</span>
              </div>
              <div className="flex justify-between text-[9px] font-black uppercase tracking-tighter">
                <span className="text-slate-400">Proj OS:</span>
                <span className="text-[#333333]">{metrics.details.projOsAT1.toFixed(0)}</span>
              </div>
              <div className="flex justify-between text-[9px] font-black uppercase tracking-tighter">
                <span className="text-slate-400">Base:</span>
                <span className="text-[#333333]">{metrics.details.base.toLocaleString()}</span>
              </div>
            </div>
          </div>

          {/* Projection Summary Card - Explicitly requested values */}
          <div className="lg:col-span-2 bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex items-center gap-3 mb-6">
              <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center">
                <Calculator className="w-5 h-5 text-slate-600" />
              </div>
              <div className="flex flex-col">
                <h3 className="text-slate-600 font-black uppercase italic tracking-tight">Resumo de Projeção</h3>
                {baseCidadeData.length > 0 && (
                  <div className="flex items-center gap-1 text-[9px] font-bold text-emerald-600 uppercase">
                    <CheckCircle2 className="w-2.5 h-2.5" />
                    <span>
                      {`Base: ${Math.round(baseCidadeData.length / 3)} cidades importadas`}
                    </span>
                  </div>
                )}
                {metrics.details.base === 0 && !filters.cidade.includes('Todos') && (
                  <div className="flex items-center gap-1 text-[9px] font-bold text-red-500 uppercase mt-0.5">
                    <AlertCircle className="w-2.5 h-2.5" />
                    <span>Cidade não encontrada na base</span>
                  </div>
                )}
              </div>
            </div>
            
            <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">OS Executadas</p>
                <p className="text-2xl font-black text-[#333333]">{metrics.details.osExec}</p>
              </div>
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">Total Visitas</p>
                <p className="text-2xl font-black text-[#333333]">{metrics.details.totalVisits}</p>
              </div>
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">Dias Trabalhados</p>
                <p className="text-2xl font-black text-[#333333]">{metrics.details.diasTrab}</p>
              </div>
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">Dias do Mês</p>
                <p className="text-2xl font-black text-[#333333]">{metrics.details.diasMes}</p>
              </div>
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">Base de Clientes</p>
                <p className="text-2xl font-black text-[#333333]">{metrics.details.base.toLocaleString()}</p>
              </div>
              <div className="bg-[#EE1D23]/5 p-4 rounded-2xl border border-[#EE1D23]/10">
                <p className="text-[10px] font-black text-[#EE1D23] uppercase tracking-widest mb-1">AT1 Projetado</p>
                <p className="text-2xl font-black text-[#EE1D23]">{formatPercent(metrics.projAT1Total)}</p>
              </div>
              <div className="bg-[#EE1D23]/5 p-4 rounded-2xl border border-[#EE1D23]/10">
                <p className="text-[10px] font-black text-[#EE1D23] uppercase tracking-widest mb-1">G1 Projetado</p>
                <p className="text-2xl font-black text-[#EE1D23]">{formatPercent(metrics.projG1Total)}</p>
              </div>
            </div>
            
            <div className="mt-6 p-3 bg-slate-50 rounded-xl border border-slate-100">
              <p className="text-[9px] font-bold text-slate-500 italic">
                * Cálculo G1: ({metrics.details.totalVisits} / {metrics.details.diasTrab} * {metrics.details.diasMes}) / {metrics.details.base} * 100
              </p>
              <p className="text-[9px] font-bold text-slate-500 italic mt-1">
                * Cálculo AT1: ({metrics.details.osExec} / {metrics.details.diasTrab} * {metrics.details.diasMes}) / {metrics.details.base} * 100
              </p>
            </div>
          </div>

          {/* Daily Volume Chart */}
          <div className="lg:col-span-2 bg-white p-8 rounded-3xl shadow-md border border-slate-100">
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
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-tighter">{comparisonMonths.previous || 'Anterior'}</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-[#EE1D23]"></div>
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-tighter">{comparisonMonths.current || 'Atual'}</span>
                </div>
              </div>
            </div>
            <div className="h-[280px] w-full min-h-[280px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <ComposedChart data={dailyVolume} margin={{ top: 20, right: 30, left: 0, bottom: 0 }}>
                  <defs>
                    <linearGradient id="colorValue" x1="0" y1="0" x2="0" y2="1">
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
                    fill="url(#colorValue)"
                    fillOpacity={1}
                    connectNulls
                  />
                  <Line 
                    type="monotone" 
                    dataKey="previousValue" 
                    name={comparisonMonths.previous || 'Mês Anterior'} 
                    stroke="#cbd5e1" 
                    strokeWidth={3} 
                    dot={{ r: 0 }}
                    activeDot={{ r: 4 }}
                    connectNulls
                  />
                  <Line 
                    type="monotone" 
                    dataKey="value" 
                    name={comparisonMonths.current || 'Mês Atual'} 
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
                        
                        // Show labels only for local peaks or every few points
                        const isLast = index === dailyVolume.length - 1;
                        const prevVal = (index > 0 && dailyVolume[index - 1]) ? dailyVolume[index - 1].value : null;
                        const nextVal = (index < dailyVolume.length - 1 && dailyVolume[index + 1]) ? dailyVolume[index + 1].value : null;
                        
                        // Simple peak detection
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

          {/* Weekly Volume Chart */}
          <div className="bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex items-center gap-3 mb-8">
              <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center">
                <Calendar className="w-5 h-5 text-[#333333]" />
              </div>
              <div>
                <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tighter">Volume Semanal</h3>
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">S1 a S5 (7 dias cada)</p>
              </div>
            </div>
            <div className="h-[280px] w-full min-h-[280px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart data={weeklyVolume} margin={{ top: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis 
                    dataKey="name" 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 11, fontWeight: 800 }} 
                  />
                  <YAxis axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 10, fontWeight: 600 }} />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 'bold' }}
                  />
                  <Bar dataKey="value" fill="#EE1D23" radius={[8, 8, 0, 0]} barSize={40} minPointSize={2}>
                    <LabelList dataKey="value" position="top" style={{ fill: '#EE1D23', fontSize: 12, fontWeight: 900 }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* AT1 by Tech Chart */}
          <div className="lg:col-span-2 bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-10 gap-4">
              <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tight">AT1 por Tecnologia</h3>
              <div className="flex gap-4">
                {/* Previous month first */}
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-sm bg-[#cbd5e1]" />
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-tighter">{comparisonMonths.previous} (AT1)</span>
                </div>
                {/* Current month second */}
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-sm bg-[#EE1D23]" />
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-tighter">{comparisonMonths.current} (AT1)</span>
                </div>
              </div>
            </div>
            
            <div className="h-[280px] w-full min-h-[280px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart data={chartData} margin={{ top: 20, right: 30, left: 0, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis 
                    dataKey="name" 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 11, fontWeight: 800 }}
                    dy={12}
                  />
                  <YAxis 
                    domain={[0, 'auto']} 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#94a3b8', fontSize: 11, fontWeight: 600 }}
                  />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ 
                      borderRadius: '16px', 
                      border: 'none', 
                      boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1)',
                      padding: '16px',
                      fontWeight: 'bold'
                    }}
                    formatter={(val: number) => formatPercent(val)}
                  />
                  <Bar dataKey="previous" name={comparisonMonths.previous || 'Anterior'} fill="#cbd5e1" radius={[4, 4, 0, 0]} barSize={25}>
                    <LabelList dataKey="previous" position="top" formatter={(val: number) => formatPercent(val)} style={{ fontSize: 10, fontWeight: 800, fill: '#94a3b8' }} />
                  </Bar>
                  <Bar dataKey="current" name={comparisonMonths.current || 'Atual'} fill="#EE1D23" radius={[4, 4, 0, 0]} barSize={25}>
                    <LabelList dataKey="current" position="top" formatter={(val: number) => formatPercent(val)} style={{ fontSize: 10, fontWeight: 800, fill: '#EE1D23' }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Volume por Área Chart */}
          <div className="bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex items-center gap-3 mb-8">
              <div className="w-10 h-10 bg-red-50 rounded-xl flex items-center justify-center">
                <MapPin className="w-5 h-5 text-[#EE1D23]" />
              </div>
              <div>
                <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tighter">Volume por Área</h3>
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Distribuição por região</p>
              </div>
            </div>
            <div className="h-[400px] w-full min-h-[400px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart data={areaData} layout="vertical" margin={{ left: 40, right: 40 }}>
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                  <XAxis type="number" axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 10, fontWeight: 600 }} />
                  <YAxis 
                    dataKey="name" 
                    type="category"
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 10, fontWeight: 800 }} 
                    width={120}
                  />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 'bold' }}
                  />
                  <Bar dataKey="value" fill="#EE1D23" radius={[0, 8, 8, 0]} barSize={20}>
                    <LabelList dataKey="value" position="right" style={{ fill: '#EE1D23', fontSize: 11, fontWeight: 900 }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Volume por Terminal Chart */}
          <div className="lg:col-span-3 bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex items-center justify-between mb-8">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-red-50 rounded-xl flex items-center justify-center">
                  <Cpu className="w-5 h-5 text-[#EE1D23]" />
                </div>
                <div>
                  <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tighter">Volume por Terminal</h3>
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Top 15 Terminais por Volume de Atendimento</p>
                </div>
              </div>
              <div className="flex gap-4">
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-[#cbd5e1]" />
                  <span className="text-[9px] font-bold text-slate-500 uppercase">{comparisonMonths.previous}</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-[#EE1D23]" />
                  <span className="text-[9px] font-bold text-slate-500 uppercase">{comparisonMonths.current}</span>
                </div>
              </div>
            </div>
            <div className="h-[600px] w-full min-h-[600px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart data={terminalData} layout="vertical" margin={{ left: 250, right: 60, top: 10, bottom: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                  <XAxis type="number" axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 10, fontWeight: 600 }} />
                  <YAxis 
                    dataKey="name" 
                    type="category"
                    axisLine={true} 
                    tickLine={true} 
                    tick={{ fill: '#333333', fontSize: 10, fontWeight: 800 }} 
                    width={240}
                    interval={0}
                  />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 'bold' }}
                  />
                  <Bar dataKey="previous" name={comparisonMonths.previous || 'Anterior'} fill="#cbd5e1" radius={[0, 8, 8, 0]} barSize={15}>
                    <LabelList dataKey="previous" position="right" style={{ fill: '#94a3b8', fontSize: 10, fontWeight: 900 }} offset={10} />
                  </Bar>
                  <Bar dataKey="current" name={comparisonMonths.current || 'Atual'} fill="#EE1D23" radius={[0, 8, 8, 0]} barSize={15}>
                    <LabelList dataKey="current" position="right" style={{ fill: '#EE1D23', fontSize: 10, fontWeight: 900 }} offset={10} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Volume por Grupo de Baixa Chart */}
          <div className="lg:col-span-3 bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex items-center justify-between mb-8">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-red-50 rounded-xl flex items-center justify-center">
                  <Filter className="w-5 h-5 text-[#EE1D23]" />
                </div>
                <div>
                  <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tighter">Volume por Grupo de Baixa</h3>
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Distribuição por categoria de encerramento</p>
                </div>
              </div>
              <div className="flex gap-4">
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-[#cbd5e1]" />
                  <span className="text-[9px] font-bold text-slate-500 uppercase">{comparisonMonths.previous}</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-[#EE1D23]" />
                  <span className="text-[9px] font-bold text-slate-500 uppercase">{comparisonMonths.current}</span>
                </div>
              </div>
            </div>
            <div className="w-full" style={{ height: `${Math.max(600, grupoBaixaData.length * 38)}px` }}>
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart data={grupoBaixaData} layout="vertical" margin={{ left: 20, right: 130, top: 10, bottom: 10 }} barGap={4} barCategoryGap={12}>
                  <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                  <XAxis type="number" hide />
                  <YAxis 
                    dataKey="name" 
                    type="category"
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 11, fontWeight: 800 }} 
                    width={220}
                  />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 'bold' }}
                  />
                  <Bar dataKey="previous" name={comparisonMonths.previous || 'Anterior'} fill="#cbd5e1" radius={[0, 8, 8, 0]} barSize={12}>
                    <LabelList 
                      dataKey="previous" 
                      position="right" 
                      content={(props: any) => {
                        const { x, y, width, height, value, payload } = props;
                        if (value === undefined || value === null) return null;
                        
                        const data = payload && payload.payload ? payload.payload : payload;
                        const impactValue = (data && data.impactPrev !== undefined) ? formatPercent(data.impactPrev) : undefined;
                        const yCenter = y + (height ? height / 2 : 6);
                        
                        return (
                          <g>
                            <text 
                              x={x + width + 8} 
                              y={yCenter} 
                              dominantBaseline="central"
                              fill="#94a3b8" 
                              fontSize={10} 
                              fontWeight={900} 
                              textAnchor="start"
                            >
                              {typeof value === 'number' ? value.toLocaleString('pt-BR') : value}
                              {impactValue !== undefined && (
                                <tspan dx={6} fill="#94a3b8" fontSize={9} fontWeight={800}>
                                  ({impactValue})
                                </tspan>
                              )}
                            </text>
                          </g>
                        );
                      }}
                    />
                  </Bar>
                  <Bar dataKey="current" name={comparisonMonths.current || 'Atual'} fill="#EE1D23" radius={[0, 8, 8, 0]} barSize={12} minPointSize={2}>
                    <LabelList 
                      dataKey="current" 
                      position="right" 
                      content={(props: any) => {
                        const { x, y, width, height, value, payload } = props;
                        if (value === undefined || value === null) return null;
                        
                        const data = payload && payload.payload ? payload.payload : payload;
                        const impactValue = (data && data.impact !== undefined) ? formatPercent(data.impact) : undefined;
                        const yCenter = y + (height ? height / 2 : 6);
                        
                        return (
                          <g>
                            <text 
                              x={x + width + 8} 
                              y={yCenter} 
                              dominantBaseline="central"
                              fill="#EE1D23" 
                              fontSize={10} 
                              fontWeight={900} 
                              textAnchor="start"
                            >
                              {typeof value === 'number' ? value.toLocaleString('pt-BR') : value}
                              {impactValue !== undefined && (
                                <tspan dx={6} fill="#64748b" fontSize={9} fontWeight={800}>
                                  ({impactValue})
                                </tspan>
                              )}
                            </text>
                          </g>
                        );
                      }}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {/* Top 15 Nodes & Top 20 Low Codes */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-12">
          {/* Top 15 Nodes */}
          <div className="bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-8">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center">
                  <Cpu className="w-5 h-5 text-[#333333]" />
                </div>
                <div>
                  <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tighter">Top 15 Nodes Ofensores</h3>
                  <div className="flex gap-3 mt-1">
                    <div className="flex items-center gap-1">
                      <div className="w-2 h-2 rounded-full bg-[#cbd5e1]" />
                      <span className="text-[8px] font-bold text-slate-500 uppercase">{comparisonMonths.previous}</span>
                    </div>
                    <div className="flex items-center gap-1">
                      <div className="w-2 h-2 rounded-full bg-[#EE1D23]" />
                      <span className="text-[8px] font-bold text-slate-500 uppercase">{comparisonMonths.current}</span>
                    </div>
                  </div>
                </div>
              </div>
              <div className="w-full md:w-64">
                <MultiFilterSelect 
                  label="Filtrar Nodes"
                  icon={<Cpu className="w-3 h-3" />}
                  value={selectedNodes.length === 0 ? ['Todos'] : selectedNodes}
                  options={['Todos', ...[...new Set(nodeChartData.map(d => d.node))].sort()]}
                  onChange={(val) => setSelectedNodes(val.includes('Todos') ? [] : val)}
                />
              </div>
            </div>
            
            <div className="h-[550px] w-full min-h-[550px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart 
                  data={topNodes} 
                  layout="vertical"
                  margin={{ top: 5, right: 50, left: 40, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" horizontal={true} vertical={false} stroke="#f1f5f9" />
                  <XAxis type="number" hide />
                  <YAxis 
                    dataKey="name" 
                    type="category" 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 10, fontWeight: 800 }}
                    width={80}
                  />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ 
                      borderRadius: '16px', 
                      border: 'none', 
                      boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1)',
                      padding: '16px',
                      fontWeight: 'bold'
                    }}
                  />
                  <Bar dataKey="previous" name={comparisonMonths.previous || 'Anterior'} fill="#cbd5e1" radius={[0, 8, 8, 0]} barSize={12}>
                    <LabelList dataKey="previous" position="right" style={{ fill: '#94a3b8', fontSize: 10, fontWeight: 900 }} offset={10} />
                  </Bar>
                  <Bar 
                    dataKey="current" 
                    name={comparisonMonths.current || 'Atual'}
                    fill="#EE1D23" 
                    radius={[0, 8, 8, 0]} 
                    barSize={12}
                    label={{ 
                      position: 'right', 
                      fill: '#EE1D23', 
                      fontSize: 10, 
                      fontWeight: 900,
                      offset: 10
                    }}
                  />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Top 20 Low Codes */}
          <div className="bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-8">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center">
                  <Signal className="w-5 h-5 text-[#333333]" />
                </div>
                <div>
                  <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tighter">Top 20 Códigos de Baixa</h3>
                  <div className="flex gap-3 mt-1">
                    <div className="flex items-center gap-1">
                      <div className="w-2 h-2 rounded-full bg-[#cbd5e1]" />
                      <span className="text-[8px] font-bold text-slate-500 uppercase">{comparisonMonths.previous}</span>
                    </div>
                    <div className="flex items-center gap-1">
                      <div className="w-2 h-2 rounded-full bg-[#EE1D23]" />
                      <span className="text-[8px] font-bold text-slate-500 uppercase">{comparisonMonths.current}</span>
                    </div>
                  </div>
                </div>
              </div>
              <div className="w-full md:w-64">
                <MultiFilterSelect 
                  label="Filtrar Baixas"
                  icon={<Signal className="w-3 h-3" />}
                  value={selectedBaixas.length === 0 ? ['Todos'] : selectedBaixas}
                  options={['Todos', ...[...new Set(baixaChartData.map(d => d.cdBaixa))].sort()]}
                  onChange={(val) => setSelectedBaixas(val.includes('Todos') ? [] : val)}
                />
              </div>
            </div>
            
            <div className="h-[550px] w-full min-h-[550px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart 
                  data={topBaixas} 
                  layout="vertical"
                  margin={{ top: 5, right: 140, left: 10, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" horizontal={true} vertical={false} stroke="#f1f5f9" />
                  <XAxis type="number" hide />
                  <YAxis 
                    dataKey="displayName" 
                    type="category" 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 11, fontWeight: 800 }}
                    width={250}
                  />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    content={({ active, payload }: any) => {
                      if (active && payload && payload.length) {
                        const data = payload[0].payload;
                        if (!data) return null;
                        
                        return (
                          <div className="bg-white p-4 rounded-2xl shadow-2xl border border-slate-100 flex flex-col gap-2 min-w-[240px]">
                            <div className="text-xs font-black text-slate-400 uppercase tracking-widest">Código: {data.name}</div>
                            {data.description && data.description !== 'N/A' && (
                              <div className="text-[11px] font-bold text-[#EE1D23] italic leading-snug mb-2 whitespace-normal bg-red-50 p-2 rounded-lg border border-red-100">
                                {data.description}
                              </div>
                            )}
                            <div className="space-y-1.5 mt-1">
                              <div className="flex justify-between items-center gap-4">
                                <span className="text-[10px] uppercase font-bold text-slate-500">{comparisonMonths.current || 'Atual'}:</span>
                                <span className="text-sm font-black text-[#EE1D23]">{Number(data.current || 0).toLocaleString()} Visitas</span>
                              </div>
                              <div className="flex justify-between items-center gap-4">
                                <span className="text-[10px] uppercase font-bold text-slate-600">Impacto s/ Base:</span>
                                <span className="text-sm font-black text-slate-800 bg-slate-100 px-2 py-0.5 rounded">{formatPercent(data.impact || 0)}</span>
                              </div>
                              <div className="flex justify-between items-center gap-4 pt-1.5 border-t border-slate-100">
                                <span className="text-[10px] uppercase font-bold text-slate-400">{comparisonMonths.previous || 'Anterior'}:</span>
                                <div className="text-right">
                                  <div className="text-xs font-bold text-slate-400">{Number(data.previous || 0).toLocaleString()} Visitas</div>
                                  <div className="text-[9px] font-bold text-slate-400">Impacto: {formatPercent(data.impactPrev || 0)}</div>
                                </div>
                              </div>
                            </div>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Bar dataKey="previous" name={comparisonMonths.previous || 'Anterior'} fill="#cbd5e1" radius={[0, 8, 8, 0]} barSize={12}>
                    <LabelList 
                      dataKey="previous" 
                      position="right" 
                      content={(props: any) => {
                        const { x, y, width, height, value, payload } = props;
                        if (!value) return null;
                        const yCenter = y + (height ? height / 2 : 6);
                        
                        const data = payload && payload.payload ? payload.payload : payload;
                        if (!data) return (
                          <g>
                            <text x={x + width + 8} y={yCenter} dominantBaseline="central" fill="#94a3b8" fontSize={10} fontWeight={900} textAnchor="start">
                              {typeof value === 'number' ? value.toLocaleString('pt-BR') : value}
                            </text>
                          </g>
                        );

                        const impactValue = data.impactPrev !== undefined ? formatPercent(data.impactPrev) : undefined;
                        
                        return (
                          <g>
                            <text x={x + width + 8} y={yCenter} dominantBaseline="central" fill="#94a3b8" fontSize={10} fontWeight={900} textAnchor="start">
                              {typeof value === 'number' ? value.toLocaleString('pt-BR') : value}
                              {impactValue !== undefined && (
                                <tspan dx={6} fill="#94a3b8" fontSize={9} fontWeight={800}>
                                  ({impactValue})
                                </tspan>
                              )}
                            </text>
                          </g>
                        );
                      }}
                    />
                  </Bar>
                  <Bar 
                    dataKey="current" 
                    name={comparisonMonths.current || 'Atual'}
                    fill="#EE1D23" 
                    radius={[0, 8, 8, 0]} 
                    barSize={12}
                  >
                    <LabelList 
                      dataKey="current" 
                      position="right" 
                      content={(props: any) => {
                        const { x, y, width, height, value, payload } = props;
                        if (!value) return null;
                        const yCenter = y + (height ? height / 2 : 6);
                        
                        // Access nested data safely
                        const data = payload && payload.payload ? payload.payload : payload;
                        if (!data) return (
                          <g>
                            <text x={x + width + 8} y={yCenter} dominantBaseline="central" fill="#EE1D23" fontSize={10} fontWeight={900} textAnchor="start">
                              {typeof value === 'number' ? value.toLocaleString('pt-BR') : value}
                            </text>
                          </g>
                        );

                        const impactValue = data.impact !== undefined ? formatPercent(data.impact) : undefined;
                        
                        return (
                          <g>
                            <text x={x + width + 8} y={yCenter} dominantBaseline="central" fill="#EE1D23" fontSize={10} fontWeight={900} textAnchor="start">
                              {typeof value === 'number' ? value.toLocaleString('pt-BR') : value}
                              {impactValue !== undefined && (
                                <tspan dx={6} fill="#64748b" fontSize={9} fontWeight={800}>
                                  ({impactValue})
                                </tspan>
                              )}
                            </text>
                          </g>
                        );
                      }}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
        {/* Base de Clientes Charts */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-12">
          {/* Base por Cidade */}
          <div className="bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex items-center gap-3 mb-8">
              <div className="w-10 h-10 bg-red-50 rounded-xl flex items-center justify-center">
                <MapPin className="w-5 h-5 text-[#EE1D23]" />
              </div>
              <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tight">Base de Clientes por Cidade</h3>
            </div>
            <div className="h-[350px] w-full min-h-[350px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart 
                  data={baseMetrics.cityData} 
                  layout="vertical" 
                  margin={{ left: 40, right: 40 }}
                  onClick={(data) => {
                    if (data && data.activeLabel) {
                      const city = String(data.activeLabel);
                      setFilters(prev => ({
                        ...prev,
                        cidade: prev.cidade.includes(city) && prev.cidade.length === 1 
                          ? ['Todos'] 
                          : [city]
                      }));
                    }
                  }}
                >
                  <CartesianGrid strokeDasharray="3 3" horizontal={true} vertical={false} stroke="#f1f5f9" />
                  <XAxis type="number" hide />
                  <YAxis 
                    dataKey="name" 
                    type="category" 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 11, fontWeight: 800 }}
                    width={100}
                  />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 'bold' }}
                  />
                  <Bar 
                    dataKey="value" 
                    fill="#EE1D23" 
                    radius={[0, 8, 8, 0]} 
                    barSize={25} 
                    label={{ position: 'right', fill: '#666666', fontSize: 10, fontWeight: 800, offset: 10 }}
                    className="cursor-pointer"
                  >
                    {baseMetrics.cityData.map((entry, index) => (
                      <Cell 
                        key={`cell-${index}`} 
                        fill={filters.cidade.includes(entry.name) ? '#991B1B' : '#EE1D23'} 
                      />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <p className="text-[10px] text-slate-400 font-bold italic mt-4">* Clique nas barras para filtrar por cidade.</p>
          </div>

          {/* AT1 Projetado por Cidade */}
          <div className="bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-8">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-red-50 rounded-xl flex items-center justify-center">
                  <Calculator className="w-5 h-5 text-[#EE1D23]" />
                </div>
                <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tight">AT1 Projetado por Cidade (%)</h3>
              </div>
              <div className="flex items-center gap-6">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-slate-300"></div>
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-tighter">{comparisonMonths.previous || 'Anterior'} (AT1)</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-[#EE1D23]"></div>
                  <span className="text-[10px] text-slate-500 font-black uppercase tracking-tighter">{comparisonMonths.current || 'Atual'} (AT1)</span>
                </div>
              </div>
            </div>
            <div className="h-[800px] w-full min-h-[800px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart 
                  data={at1ByCityData} 
                  layout="vertical" 
                  margin={{ left: 40, right: 100, top: 10, bottom: 10 }}
                  barGap={4}
                >
                  <CartesianGrid strokeDasharray="3 3" horizontal={true} vertical={false} stroke="#f1f5f9" />
                  <XAxis type="number" hide />
                  <YAxis 
                    dataKey="name" 
                    type="category"
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 10, fontWeight: 800 }} 
                    width={120}
                  />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 'bold' }}
                    formatter={(val: number) => formatPercent(val)}
                  />
                  <Bar dataKey="previous" fill="#cbd5e1" radius={[0, 4, 4, 0]} barSize={12}>
                    <LabelList 
                      dataKey="previous" 
                      position="right" 
                      content={(props: any) => {
                        const { x, y, width, value, index } = props;
                        if (value === undefined || value === null) return null;
                        
                        // Use index to get reliable data from the source array
                        const entry = at1ByCityData[index];
                        const osExec = entry ? entry.osExecPrev : 0;
                        
                        return (
                          <g>
                            <text x={x + width + 8} y={y + 10} fill="#94a3b8" fontSize={10} fontWeight={900} textAnchor="start">
                              {formatPercent(value)}
                              <tspan dx={6} fill="#94a3b8" fontSize={9} fontWeight={800}>
                                ({Number(osExec).toLocaleString()})
                              </tspan>
                            </text>
                          </g>
                        );
                      }}
                    />
                  </Bar>
                  <Bar dataKey="current" fill="#EE1D23" radius={[0, 4, 4, 0]} barSize={12}>
                    <LabelList 
                      dataKey="current" 
                      position="right" 
                      content={(props: any) => {
                        const { x, y, width, value, index } = props;
                        if (value === undefined || value === null) return null;
                        
                        const entry = at1ByCityData[index];
                        const osExec = entry ? entry.osExec : 0;
                        
                        return (
                          <g>
                            <text x={x + width + 8} y={y + 10} fill="#EE1D23" fontSize={10} fontWeight={900} textAnchor="start">
                              {formatPercent(value)}
                              <tspan dx={6} fill="#64748b" fontSize={9} fontWeight={800}>
                                ({Number(osExec).toLocaleString()})
                              </tspan>
                            </text>
                          </g>
                        );
                      }}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <p className="text-[10px] text-slate-400 font-bold italic mt-4">* Projeção baseada em OS Executadas / Dias Trab * Dias Mês / Base.</p>
          </div>

          {/* Base por Tecnologia */}
          <div className="bg-white p-8 rounded-3xl shadow-md border border-slate-100">
            <div className="flex items-center gap-3 mb-8">
              <div className="w-10 h-10 bg-red-50 rounded-xl flex items-center justify-center">
                <Cpu className="w-5 h-5 text-[#EE1D23]" />
              </div>
              <h3 className="text-lg font-black text-[#333333] uppercase italic tracking-tight">Base de Clientes por Tecnologia</h3>
            </div>
            <div className="h-[350px] w-full min-h-[350px]">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart data={baseMetrics.techData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis 
                    dataKey="name" 
                    axisLine={false} 
                    tickLine={false} 
                    tick={{ fill: '#333333', fontSize: 11, fontWeight: 800 }} 
                  />
                  <YAxis axisLine={false} tickLine={false} tick={{ fill: '#666666', fontSize: 10, fontWeight: 600 }} />
                  <Tooltip 
                    cursor={{ fill: '#f8fafc' }}
                    contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 'bold' }}
                    formatter={(value: any, name: any, props: any) => {
                      const percentage = props.payload.percentage;
                      return [
                        `${Number(value).toLocaleString('pt-BR')} (${formatPercent(percentage)})`,
                        'Base'
                      ];
                    }}
                  />
                  <Bar dataKey="value" fill="#EE1D23" radius={[8, 8, 0, 0]} barSize={60}>
                    <LabelList 
                      dataKey="value" 
                      position="top" 
                      formatter={(val: number) => val.toLocaleString()}
                      style={{ fill: '#EE1D23', fontSize: 11, fontWeight: 900 }} 
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
        </main>
            </>
          )}
          </div>

          <div className={activeTab === 'outage' ? 'block' : 'hidden'}>
            <OutageDashboard 
              at1DailyVolume={dailyVolume}
              at1ComparisonMonths={comparisonMonths}
            />
          </div>

          <div className={activeTab === 'at5' ? 'block' : 'hidden'}>
            <AT5Dashboard />
          </div>

          <div className={activeTab === 'log' ? 'block' : 'hidden'}>
            <LOGDashboard />
          </div>

          <div className={activeTab === 'revisita30d' ? 'block' : 'hidden'}>
            <Revisita30DDashboard />
          </div>

          <div className={activeTab === 'logbook' ? 'block' : 'hidden'}>
            <Logbook 
              entries={logEntries} 
              isLoading={isLogLoading} 
              error={logFetchError}
              onRefresh={fetchLogs} 
              cities={filterOptions.cidades}
            />
          </div>
        </div>
        
        <AnimatePresence>
          {showScrollTop && (
            <motion.button
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 20 }}
              onClick={scrollToTop}
              className="fixed bottom-8 right-8 z-[60] w-14 h-14 bg-[#EE1D23] text-white rounded-2xl shadow-2xl shadow-red-500/40 flex items-center justify-center hover:bg-[#D1191F] transition-all active:scale-90 group"
            >
              <ArrowUp className="w-6 h-6 group-hover:-translate-y-1 transition-transform" />
            </motion.button>
          )}
        </AnimatePresence>
      </div>
    </ErrorBoundary>
  );
}

function MetricCard({ title, value, icon, color, percentage }: {  
  title: string; 
  value: number; 
  icon: React.ReactNode; 
  color: string;
  percentage?: number;
}) {
  const colorClasses: Record<string, string> = {
    red: "bg-[#EE1D23] text-white",
    dark: "bg-[#1A1A1A] text-white",
    gray: "bg-[#666666] text-white",
  };

  return (
    <motion.div 
      initial={{ opacity: 0, scale: 0.95 }}
      animate={{ opacity: 1, scale: 1 }}
      className="bg-white p-7 rounded-3xl shadow-md border border-slate-100 flex items-start justify-between relative overflow-hidden"
    >
      <div className="relative z-10">
        <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-1">{title}</p>
        <h4 className="text-4xl font-black text-[#333333] tracking-tighter">{value.toLocaleString()}</h4>
        {percentage !== undefined && (
          <div className="mt-4 flex items-center gap-3">
            <div className="w-28 h-2 bg-slate-100 rounded-full overflow-hidden">
              <div 
                className={cn("h-full transition-all duration-700 ease-out", 
                  color === 'red' ? 'bg-[#EE1D23]' : 'bg-[#666666]'
                )}
                style={{ width: `${percentage}%` }}
              />
            </div>
            <span className="text-xs font-black text-[#333333]">{percentage.toFixed(0)}%</span>
          </div>
        )}
      </div>
      <div className={cn("p-4 rounded-2xl shadow-lg", colorClasses[color])}>
        {React.cloneElement(icon as React.ReactElement<any>, { className: "w-7 h-7" })}
      </div>
    </motion.div>
  );
}

function MultiFilterSelect({ label, icon, value, options, onChange }: {
  label: string;
  icon: React.ReactNode;
  value: string[];
  options: string[];
  onChange: (val: string[]) => void;
}) {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const containerRef = useRef<HTMLDivElement>(null);

  // Close when clicking outside
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Reset search when opening/closing
  React.useEffect(() => {
    if (!isOpen) setSearchTerm('');
  }, [isOpen]);

  const toggleOption = (opt: string) => {
    if (opt === 'Todos') {
      onChange(['Todos']);
      return;
    }

    let newValue = [...value];
    if (newValue.includes('Todos')) {
      newValue = [opt];
    } else {
      if (newValue.includes(opt)) {
        newValue = newValue.filter(v => v !== opt);
        if (newValue.length === 0) newValue = ['Todos'];
      } else {
        newValue.push(opt);
      }
    }
    onChange(newValue);
  };

  const filteredOptions = options.filter(opt => 
    opt.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const displayValue = value.includes('Todos') ? 'Todos' : value.length === 1 ? value[0] : `${value.length} selecionados`;

  return (
    <div className="space-y-2" ref={containerRef}>
      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2 ml-1">
        <span className="text-[#EE1D23]">{icon}</span>
        {label}
      </label>
      <div className="relative">
        <button 
          type="button"
          onClick={() => setIsOpen(!isOpen)}
          className="w-full bg-slate-50 border-2 border-slate-100 text-[#333333] text-sm font-bold rounded-xl px-4 py-3 outline-none focus:ring-4 focus:ring-red-500/10 focus:border-[#EE1D23] transition-all cursor-pointer flex items-center justify-between"
        >
          <span className="truncate">{displayValue}</span>
          <ChevronDown className={cn("w-4 h-4 text-slate-400 transition-transform", isOpen && "rotate-180")} />
        </button>

        <AnimatePresence>
          {isOpen && (
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 10 }}
              className="absolute z-50 mt-2 w-full bg-white border border-slate-200 rounded-xl shadow-xl max-h-80 overflow-hidden flex flex-col"
            >
              <div className="p-2 border-bottom border-slate-100">
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                  <input 
                    type="text"
                    placeholder="Pesquisar..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg pl-9 pr-4 py-2 text-sm font-bold outline-none focus:border-[#EE1D23] transition-all"
                    autoFocus
                  />
                </div>
              </div>
              <div className="overflow-y-auto p-2 max-h-60">
                {filteredOptions.length > 0 ? (
                  filteredOptions.map(opt => {
                    const isSelected = value.includes(opt);
                    return (
                      <label 
                        key={opt} 
                        className={cn(
                          "flex items-center gap-3 px-3 py-2.5 rounded-lg cursor-pointer transition-all mb-1 last:mb-0",
                          isSelected ? "bg-red-50" : "hover:bg-slate-50"
                        )}
                      >
                        <div className={cn(
                          "w-5 h-5 rounded-md border-2 flex items-center justify-center transition-all shrink-0",
                          isSelected ? "bg-[#EE1D23] border-[#EE1D23]" : "bg-white border-slate-200"
                        )}>
                          {isSelected && <CheckCircle2 className="w-3.5 h-3.5 text-white" />}
                        </div>
                        <input 
                          type="checkbox" 
                          checked={isSelected}
                          onChange={() => toggleOption(opt)}
                          className="hidden"
                        />
                        <span className={cn(
                          "text-sm font-bold transition-colors",
                          isSelected ? "text-[#EE1D23]" : "text-[#333333]"
                        )}>
                          {opt}
                        </span>
                      </label>
                    );
                  })
                ) : (
                  <div className="p-4 text-center text-xs font-bold text-slate-400 uppercase italic">Nenhum resultado encontrado</div>
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
}

function FilterSelect({ label, icon, value, options, onChange }: {
  label: string;
  icon: React.ReactNode;
  value: string;
  options: string[];
  onChange: (val: string) => void;
}) {
  return (
    <div className="space-y-2">
      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2 ml-1">
        <span className="text-[#EE1D23]">{icon}</span>
        {label}
      </label>
      <div className="relative group">
        <select 
          value={value}
          onChange={(e) => onChange(e.target.value)}
          className="w-full appearance-none bg-slate-50 border-2 border-slate-100 text-[#333333] text-sm font-bold rounded-xl px-4 py-3 outline-none focus:ring-4 focus:ring-red-500/10 focus:border-[#EE1D23] transition-all cursor-pointer"
        >
          {options.map(opt => (
            <option key={opt} value={opt}>{opt}</option>
          ))}
        </select>
        <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none group-hover:text-[#EE1D23] transition-colors" />
      </div>
    </div>
  );
}
