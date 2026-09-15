import React, { useState, useMemo } from 'react';
import { 
  Search, 
  Download, 
  FileSpreadsheet, 
  ChevronLeft, 
  ChevronRight, 
  ChevronsLeft, 
  ChevronsRight, 
  ArrowUpDown, 
  ArrowUp, 
  ArrowDown, 
  CheckCircle2, 
  XCircle, 
  AlertTriangle,
  Filter,
  Layers
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { VisitData } from '../data';
import { cn } from '../lib/utils';

interface At1AnaliticoTableProps {
  data: VisitData[];
  totalDataCount: number;
}

type SortField = 'fullDate' | 'contrato' | 'cidade' | 'area' | 'tecnologia' | 'tipoOs' | 'status' | 'expurgo' | 'notaAT1';
type SortOrder = 'asc' | 'desc';

export const At1AnaliticoTable: React.FC<At1AnaliticoTableProps> = ({ data, totalDataCount }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<'Todos' | 'Executada' | 'Cancelada'>('Todos');
  const [expurgoFilter, setExpurgoFilter] = useState<'Todos' | 'Sem Expurgo' | 'Com Expurgo'>('Todos');
  const [techFilter, setTechFilter] = useState<'Todos' | 'GPON' | 'HFC' | 'HÍBRIDO' | 'OUTROS'>('Todos');
  
  const [sortField, setSortField] = useState<SortField>('fullDate');
  const [sortOrder, setSortOrder] = useState<SortOrder>('desc');
  
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState<number>(25);

  // Toggle sort direction or change column
  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortOrder(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('asc');
    }
    setCurrentPage(1);
  };

  // Filtered and Sorted list
  const filteredData = useMemo(() => {
    return data.filter(item => {
      // Status filter
      if (statusFilter !== 'Todos' && item.status !== statusFilter) return false;
      
      // Expurgo filter
      if (expurgoFilter === 'Sem Expurgo' && item.expurgo) return false;
      if (expurgoFilter === 'Com Expurgo' && !item.expurgo) return false;
      
      // Tech filter
      if (techFilter !== 'Todos' && item.tecnologia !== techFilter) return false;

      // Text search
      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase().trim();
        const contractStr = (item.contrato || item.id || '').toLowerCase();
        const cidadeStr = (item.cidade || '').toLowerCase();
        const areaStr = (item.area || '').toLowerCase();
        const techStr = (item.tecnologia || '').toLowerCase();
        const tipoOsStr = (item.tipoOs || '').toLowerCase();
        const nodeStr = (item.node || '').toLowerCase();
        const terminalStr = (item.terminal || '').toLowerCase();
        const cdBaixaStr = (item.cdBaixa || '').toLowerCase();
        const grupoBaixaStr = (item.grupoBaixa || '').toLowerCase();
        const descBaixaStr = (item.descriptionBaixa || '').toLowerCase();

        const match = 
          contractStr.includes(q) ||
          cidadeStr.includes(q) ||
          areaStr.includes(q) ||
          techStr.includes(q) ||
          tipoOsStr.includes(q) ||
          nodeStr.includes(q) ||
          terminalStr.includes(q) ||
          cdBaixaStr.includes(q) ||
          grupoBaixaStr.includes(q) ||
          descBaixaStr.includes(q);

        if (!match) return false;
      }

      return true;
    });
  }, [data, statusFilter, expurgoFilter, techFilter, searchTerm]);

  // Sorted list
  const sortedData = useMemo(() => {
    return [...filteredData].sort((a, b) => {
      let valA: any = a[sortField];
      let valB: any = b[sortField];

      if (sortField === 'contrato') {
        valA = a.contrato || a.id || '';
        valB = b.contrato || b.id || '';
      } else if (sortField === 'fullDate') {
        valA = a.fullDate ? new Date(a.fullDate).getTime() : 0;
        valB = b.fullDate ? new Date(b.fullDate).getTime() : 0;
      } else if (sortField === 'notaAT1') {
        valA = Number(a.notaAT1) || 0;
        valB = Number(b.notaAT1) || 0;
      } else if (typeof valA === 'string') {
        valA = valA.toLowerCase();
        valB = (valB || '').toLowerCase();
      }

      if (valA < valB) return sortOrder === 'asc' ? -1 : 1;
      if (valA > valB) return sortOrder === 'asc' ? 1 : -1;
      return 0;
    });
  }, [filteredData, sortField, sortOrder]);

  // Pagination calculation
  const totalPages = Math.max(1, Math.ceil(sortedData.length / pageSize));
  const currentSafePage = Math.min(currentPage, totalPages);
  const startIndex = (currentSafePage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, sortedData.length);
  const paginatedRows = useMemo(() => {
    return sortedData.slice(startIndex, endIndex);
  }, [sortedData, startIndex, endIndex]);

  // Export to Excel / CSV
  const handleExport = (format: 'xlsx' | 'csv') => {
    if (sortedData.length === 0) return;

    const rowsToExport = sortedData.map((row, idx) => ({
      'Item': idx + 1,
      'Data': row.fullDate ? new Date(row.fullDate).toLocaleDateString('pt-BR') : '',
      'Mês': row.mes || '',
      'Contrato / OS': row.contrato || row.id || '',
      'Cidade': row.cidade || '',
      'Área': row.area || '',
      'Tecnologia': row.tecnologia || '',
      'Tipo de OS': row.tipoOs || 'REPARO',
      'Status': row.status || '',
      'Expurgo': row.expurgo ? 'Sim' : 'Não',
      'Nota AT1': Number(row.notaAT1 || 0).toFixed(1),
      'Node': row.node || '',
      'Terminal': row.terminal || '',
      'Grupo Baixa': row.grupoBaixa || '',
      'Código Baixa': row.cdBaixa || '',
      'Descrição Baixa': row.descriptionBaixa || ''
    }));

    const ws = XLSX.utils.json_to_sheet(rowsToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Analitico_AT1');

    const fileName = `Analitico_AT1_${new Date().toISOString().slice(0, 10)}.${format}`;
    XLSX.writeFile(wb, fileName, { bookType: format });
  };

  // Helper for sort icon
  const renderSortIcon = (field: SortField) => {
    if (sortField !== field) {
      return <ArrowUpDown className="w-3 h-3 text-slate-300 opacity-60 group-hover:opacity-100" />;
    }
    return sortOrder === 'asc' ? (
      <ArrowUp className="w-3 h-3 text-[#EE1D23]" />
    ) : (
      <ArrowDown className="w-3 h-3 text-[#EE1D23]" />
    );
  };

  return (
    <section id="at1-analitico-view" className="w-full max-w-7xl mx-auto mt-10 mb-12">
      <div className="bg-white rounded-3xl shadow-lg border border-slate-100 overflow-hidden">
        
        {/* Header Section */}
        <div className="p-5 sm:p-7 border-b border-slate-100 bg-gradient-to-r from-slate-50/50 via-white to-red-50/20">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div>
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-red-50 flex items-center justify-center text-[#EE1D23] shadow-xs">
                  <Layers className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-lg sm:text-xl font-black uppercase italic tracking-tight text-slate-900">
                    Visão Analítica de Ordens de Serviço (AT1)
                  </h3>
                  <p className="text-xs font-bold text-slate-500 mt-0.5">
                    Detalhamento individual das visitas técnicas, notas AT1, contratos, expurgos e códigos de baixa.
                  </p>
                </div>
              </div>
            </div>

            {/* Quick Export Actions */}
            <div className="flex items-center gap-2.5">
              <button
                id="btn-export-excel-at1"
                onClick={() => handleExport('xlsx')}
                disabled={sortedData.length === 0}
                className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-black py-2 px-3.5 rounded-xl transition-all shadow-sm active:scale-95 text-xs uppercase italic cursor-pointer disabled:opacity-50"
                title="Exportar todos os registros filtrados para Excel"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Exportar Excel ({sortedData.length.toLocaleString()})</span>
              </button>

              <button
                id="btn-export-csv-at1"
                onClick={() => handleExport('csv')}
                disabled={sortedData.length === 0}
                className="flex items-center gap-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-2 px-3 rounded-xl transition-all active:scale-95 text-xs uppercase italic cursor-pointer disabled:opacity-50"
                title="Exportar formato CSV"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-slate-500" />
                <span>CSV</span>
              </button>
            </div>
          </div>

          {/* Quick Filters and Search Toolbar */}
          <div className="mt-5 grid grid-cols-1 md:grid-cols-12 gap-3 pt-4 border-t border-slate-100">
            {/* Search Input */}
            <div className="md:col-span-5 relative">
              <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
              <input
                id="input-search-at1-analitico"
                type="text"
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                placeholder="Buscar contrato, cidade, área, tipo de OS, node..."
                className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all"
              />
              {searchTerm && (
                <button 
                  onClick={() => setSearchTerm('')} 
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs font-black"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Status Filter */}
            <div className="md:col-span-2">
              <select
                id="select-status-at1-analitico"
                value={statusFilter}
                onChange={(e) => {
                  setStatusFilter(e.target.value as any);
                  setCurrentPage(1);
                }}
                className="w-full py-2 px-3 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-[#EE1D23] transition-all cursor-pointer"
              >
                <option value="Todos">Status: Todos</option>
                <option value="Executada">Executada</option>
                <option value="Cancelada">Cancelada</option>
              </select>
            </div>

            {/* Expurgo Filter */}
            <div className="md:col-span-2">
              <select
                id="select-expurgo-at1-analitico"
                value={expurgoFilter}
                onChange={(e) => {
                  setExpurgoFilter(e.target.value as any);
                  setCurrentPage(1);
                }}
                className="w-full py-2 px-3 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-[#EE1D23] transition-all cursor-pointer"
              >
                <option value="Todos">Expurgo: Todos</option>
                <option value="Sem Expurgo">Sem Expurgo</option>
                <option value="Com Expurgo">Com Expurgo</option>
              </select>
            </div>

            {/* Tecnologia Filter */}
            <div className="md:col-span-2">
              <select
                id="select-tech-at1-analitico"
                value={techFilter}
                onChange={(e) => {
                  setTechFilter(e.target.value as any);
                  setCurrentPage(1);
                }}
                className="w-full py-2 px-3 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-[#EE1D23] transition-all cursor-pointer"
              >
                <option value="Todos">Tecnologia: Todas</option>
                <option value="GPON">GPON</option>
                <option value="HFC">HFC</option>
                <option value="HÍBRIDO">HÍBRIDO</option>
                <option value="OUTROS">OUTROS</option>
              </select>
            </div>

            {/* Page Size */}
            <div className="md:col-span-1">
              <select
                id="select-pagesize-at1-analitico"
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                className="w-full py-2 px-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-[#EE1D23] transition-all cursor-pointer text-center"
              >
                <option value={10}>10</option>
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
              </select>
            </div>
          </div>

          {/* Counts and active filter info */}
          <div className="mt-3 flex flex-wrap items-center justify-between gap-2 text-xs font-semibold text-slate-500">
            <div>
              Exibindo <strong className="text-slate-800">{paginatedRows.length}</strong> de <strong className="text-slate-800">{sortedData.length.toLocaleString()}</strong> ordens filtradas 
              {totalDataCount > 0 && <span> (de um total de <strong className="text-slate-800">{totalDataCount.toLocaleString()}</strong>)</span>}.
            </div>
            {(statusFilter !== 'Todos' || expurgoFilter !== 'Todos' || techFilter !== 'Todos' || searchTerm) && (
              <button
                onClick={() => {
                  setSearchTerm('');
                  setStatusFilter('Todos');
                  setExpurgoFilter('Todos');
                  setTechFilter('Todos');
                  setCurrentPage(1);
                }}
                className="text-[11px] font-black text-[#EE1D23] hover:underline cursor-pointer flex items-center gap-1"
              >
                <Filter className="w-3 h-3" />
                <span>Limpar filtros da tabela</span>
              </button>
            )}
          </div>
        </div>

        {/* Table Body */}
        <div className="overflow-x-auto scrollbar-thin scrollbar-thumb-slate-200">
          <table className="w-full text-left border-collapse table-auto">
            <thead>
              <tr className="bg-slate-50/90 border-b border-slate-100 text-[10px] font-black text-slate-500 uppercase tracking-wider select-none">
                
                {/* Data */}
                <th 
                  onClick={() => handleSort('fullDate')}
                  className="py-3.5 px-3.5 cursor-pointer hover:bg-slate-100/80 transition-colors group"
                >
                  <div className="flex items-center gap-1">
                    <span>Data</span>
                    {renderSortIcon('fullDate')}
                  </div>
                </th>

                {/* Contrato / OS */}
                <th 
                  onClick={() => handleSort('contrato')}
                  className="py-3.5 px-3.5 cursor-pointer hover:bg-slate-100/80 transition-colors group"
                >
                  <div className="flex items-center gap-1">
                    <span>Contrato / OS</span>
                    {renderSortIcon('contrato')}
                  </div>
                </th>

                {/* Cidade */}
                <th 
                  onClick={() => handleSort('cidade')}
                  className="py-3.5 px-3 cursor-pointer hover:bg-slate-100/80 transition-colors group"
                >
                  <div className="flex items-center gap-1">
                    <span>Cidade</span>
                    {renderSortIcon('cidade')}
                  </div>
                </th>

                {/* Área */}
                <th 
                  onClick={() => handleSort('area')}
                  className="py-3.5 px-3 cursor-pointer hover:bg-slate-100/80 transition-colors group"
                >
                  <div className="flex items-center gap-1">
                    <span>Área</span>
                    {renderSortIcon('area')}
                  </div>
                </th>

                {/* Tecnologia */}
                <th 
                  onClick={() => handleSort('tecnologia')}
                  className="py-3.5 px-3 cursor-pointer hover:bg-slate-100/80 transition-colors group"
                >
                  <div className="flex items-center gap-1">
                    <span>Tecnologia</span>
                    {renderSortIcon('tecnologia')}
                  </div>
                </th>

                {/* Tipo de OS */}
                <th 
                  onClick={() => handleSort('tipoOs')}
                  className="py-3.5 px-3 cursor-pointer hover:bg-slate-100/80 transition-colors group"
                >
                  <div className="flex items-center gap-1">
                    <span>Tipo de OS</span>
                    {renderSortIcon('tipoOs')}
                  </div>
                </th>

                {/* Status */}
                <th 
                  onClick={() => handleSort('status')}
                  className="py-3.5 px-3 cursor-pointer hover:bg-slate-100/80 transition-colors group text-center"
                >
                  <div className="flex items-center justify-center gap-1">
                    <span>Status</span>
                    {renderSortIcon('status')}
                  </div>
                </th>

                {/* Expurgo */}
                <th 
                  onClick={() => handleSort('expurgo')}
                  className="py-3.5 px-2.5 cursor-pointer hover:bg-slate-100/80 transition-colors group text-center"
                >
                  <div className="flex items-center justify-center gap-1">
                    <span>Expurgo</span>
                    {renderSortIcon('expurgo')}
                  </div>
                </th>

                {/* Nota AT1 */}
                <th 
                  onClick={() => handleSort('notaAT1')}
                  className="py-3.5 px-3 cursor-pointer hover:bg-slate-100/80 transition-colors group text-center"
                >
                  <div className="flex items-center justify-center gap-1">
                    <span>Nota AT1</span>
                    {renderSortIcon('notaAT1')}
                  </div>
                </th>

                {/* Node / Terminal */}
                <th className="py-3.5 px-3">
                  <span>Node / Terminal</span>
                </th>

                {/* Grupo & Código de Baixa */}
                <th className="py-3.5 px-3.5">
                  <span>Grupo / Cód. Baixa</span>
                </th>
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-100 text-xs font-semibold text-slate-700">
              {paginatedRows.map((row, idx) => {
                const isExecuted = row.status === 'Executada';
                const formattedDate = row.fullDate 
                  ? new Date(row.fullDate).toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' })
                  : row.mes;

                // Tech color
                const techColors: Record<string, string> = {
                  'GPON': 'bg-emerald-50 text-emerald-700 border-emerald-200',
                  'HFC': 'bg-blue-50 text-blue-700 border-blue-200',
                  'HÍBRIDO': 'bg-purple-50 text-purple-700 border-purple-200',
                  'OUTROS': 'bg-slate-100 text-slate-700 border-slate-200'
                };

                // Tipo OS color
                const tipoOs = row.tipoOs || 'REPARO';
                const isReparo = tipoOs.includes('REPARO');
                const isInstalacao = tipoOs.includes('INSTAL') || tipoOs.includes('ATIV');
                const tipoOsColor = isInstalacao
                  ? 'bg-indigo-50 text-indigo-700 border-indigo-200'
                  : isReparo
                  ? 'bg-amber-50 text-amber-800 border-amber-200'
                  : 'bg-slate-100 text-slate-700 border-slate-200';

                return (
                  <tr 
                    key={`${row.id}-${idx}`}
                    className="hover:bg-red-50/20 transition-colors border-b border-slate-50 last:border-none"
                  >
                    {/* Data */}
                    <td className="py-3 px-3.5 whitespace-nowrap font-medium text-slate-600">
                      {formattedDate}
                    </td>

                    {/* Contrato / OS */}
                    <td className="py-3 px-3.5 whitespace-nowrap">
                      <span className="font-mono font-bold text-slate-900 bg-slate-100/90 px-2 py-0.5 rounded-md border border-slate-200 text-[11px]">
                        {row.contrato || row.id}
                      </span>
                    </td>

                    {/* Cidade */}
                    <td className="py-3 px-3 font-bold text-slate-800 whitespace-nowrap">
                      {row.cidade}
                    </td>

                    {/* Área */}
                    <td className="py-3 px-3 text-slate-600 whitespace-nowrap">
                      {row.area}
                    </td>

                    {/* Tecnologia */}
                    <td className="py-3 px-3 whitespace-nowrap">
                      <span className={cn(
                        "text-[10px] font-black uppercase px-2 py-0.5 rounded-full border tracking-wide",
                        techColors[row.tecnologia] || techColors['OUTROS']
                      )}>
                        {row.tecnologia}
                      </span>
                    </td>

                    {/* Tipo de OS */}
                    <td className="py-3 px-3 whitespace-nowrap">
                      <span className={cn(
                        "text-[10px] font-black uppercase px-2 py-0.5 rounded-md border tracking-wide",
                        tipoOsColor
                      )}>
                        {tipoOs}
                      </span>
                    </td>

                    {/* Status */}
                    <td className="py-3 px-3 text-center whitespace-nowrap">
                      {isExecuted ? (
                        <span className="inline-flex items-center gap-1 text-[10px] font-black uppercase bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded-full">
                          <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                          Executada
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 text-[10px] font-black uppercase bg-red-50 text-red-600 border border-red-200 px-2 py-0.5 rounded-full">
                          <XCircle className="w-3 h-3 text-red-500" />
                          Cancelada
                        </span>
                      )}
                    </td>

                    {/* Expurgo */}
                    <td className="py-3 px-2.5 text-center whitespace-nowrap">
                      {row.expurgo ? (
                        <span className="inline-flex items-center gap-1 text-[10px] font-black uppercase bg-red-100 text-red-700 px-2 py-0.5 rounded-md">
                          <AlertTriangle className="w-3 h-3 text-red-600" />
                          Sim
                        </span>
                      ) : (
                        <span className="text-[10px] font-bold text-slate-400 bg-slate-50 px-2 py-0.5 rounded-md">
                          Não
                        </span>
                      )}
                    </td>

                    {/* Nota AT1 */}
                    <td className="py-3 px-3 text-center whitespace-nowrap">
                      {row.notaAT1 > 0 ? (
                        <span className={cn(
                          "inline-block font-black text-xs px-2.5 py-0.5 rounded-lg border",
                          row.notaAT1 >= 8.5
                            ? "bg-emerald-50 text-emerald-700 border-emerald-200"
                            : row.notaAT1 >= 7.0
                            ? "bg-amber-50 text-amber-700 border-amber-200"
                            : "bg-red-50 text-red-600 border-red-200"
                        )}>
                          {Number(row.notaAT1).toFixed(1)}
                        </span>
                      ) : (
                        <span className="text-slate-300 font-bold">-</span>
                      )}
                    </td>

                    {/* Node / Terminal */}
                    <td className="py-3 px-3 whitespace-nowrap text-[11px]">
                      <div className="flex flex-col">
                        <span className="font-bold text-slate-800">{row.node || 'N/A'}</span>
                        <span className="text-slate-400 font-mono text-[10px]">{row.terminal || 'N/A'}</span>
                      </div>
                    </td>

                    {/* Grupo / Código Baixa */}
                    <td className="py-3 px-3.5 whitespace-nowrap text-[11px]">
                      <div className="flex flex-col">
                        <div className="flex items-center gap-1.5">
                          <span className="font-bold text-slate-800">{row.grupoBaixa || 'N/A'}</span>
                          {row.cdBaixa && row.cdBaixa !== 'N/A' && (
                            <span className="font-mono text-[10px] bg-slate-100 px-1 py-0.2 rounded text-slate-600 border border-slate-200">
                              {row.cdBaixa}
                            </span>
                          )}
                        </div>
                        {row.descriptionBaixa && row.descriptionBaixa !== 'N/A' && (
                          <span className="text-[10px] text-slate-400 truncate max-w-[200px]" title={row.descriptionBaixa}>
                            {row.descriptionBaixa}
                          </span>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}

              {paginatedRows.length === 0 && (
                <tr>
                  <td colSpan={11} className="py-12 text-center text-slate-400 font-bold text-xs">
                    Nenhum registro encontrado para os filtros e busca aplicados.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination Footer */}
        <div className="p-4 sm:p-5 border-t border-slate-100 bg-slate-50/50 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="text-xs font-semibold text-slate-500">
            Mostrando <strong className="text-slate-800">{sortedData.length > 0 ? startIndex + 1 : 0}</strong> a <strong className="text-slate-800">{endIndex}</strong> de <strong className="text-slate-800">{sortedData.length.toLocaleString()}</strong> registros
          </div>

          <div className="flex items-center gap-1.5 self-center sm:self-auto">
            {/* First */}
            <button
              id="btn-page-first"
              onClick={() => setCurrentPage(1)}
              disabled={currentSafePage <= 1}
              className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
              title="Primeira Página"
            >
              <ChevronsLeft className="w-4 h-4" />
            </button>

            {/* Prev */}
            <button
              id="btn-page-prev"
              onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
              disabled={currentSafePage <= 1}
              className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
              title="Página Anterior"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>

            <span className="px-3 text-xs font-black text-slate-700">
              Página {currentSafePage} de {totalPages}
            </span>

            {/* Next */}
            <button
              id="btn-page-next"
              onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
              disabled={currentSafePage >= totalPages}
              className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
              title="Próxima Página"
            >
              <ChevronRight className="w-4 h-4" />
            </button>

            {/* Last */}
            <button
              id="btn-page-last"
              onClick={() => setCurrentPage(totalPages)}
              disabled={currentSafePage >= totalPages}
              className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-30 disabled:pointer-events-none transition-all cursor-pointer"
              title="Última Página"
            >
              <ChevronsRight className="w-4 h-4" />
            </button>
          </div>
        </div>

      </div>
    </section>
  );
};
