import React, { useState, useMemo, useEffect } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import { 
  MapContainer, 
  TileLayer, 
  Marker,
  Popup, 
  useMap 
} from 'react-leaflet';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';
import MarkerClusterGroup from 'react-leaflet-cluster';
import { 
  Upload, 
  FileSpreadsheet, 
  Map as MapIcon, 
  Filter, 
  Search, 
  X, 
  AlertCircle,
  Table as TableIcon,
  ChevronDown,
  Activity,
  Signal,
  Database
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from '../lib/utils';

// Fix Leaflet marker icons
// @ts-ignore
delete L.Icon.Default.prototype._getIconUrl;
L.Icon.Default.mergeOptions({
  iconRetinaUrl: 'https://unpkg.com/leaflet@1.7.1/dist/images/marker-icon-2x.png',
  iconUrl: 'https://unpkg.com/leaflet@1.7.1/dist/images/marker-icon.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.7.1/dist/images/marker-shadow.png',
});

const createModemIcon = (color: string) => {
  return L.divIcon({
    html: `
      <div class="modem-marker-container" style="
        background-color: ${color}; 
        width: 32px; 
        height: 32px; 
        border-radius: 10px; 
        display: flex; 
        align-items: center; 
        justify-content: center; 
        border: 2px solid white; 
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        cursor: pointer;
      ">
        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <rect x="2" y="14" width="20" height="8" rx="2" ry="2"></rect>
          <path d="M6 18h.01"></path>
          <path d="M10 18h.01"></path>
          <path d="M14 18h.01"></path>
          <path d="M18 18h.01"></path>
          <path d="M2 14l3-9h14l3 9"></path>
        </svg>
      </div>
    `,
    className: 'custom-modem-icon',
    iconSize: [32, 32],
    iconAnchor: [16, 16],
    popupAnchor: [0, -16]
  });
};

export interface NodeData {
  cmtsRxPot: number;
  cmtsRxSNR: number;
  COD_NODE: string;
  NM_TIPO_EQUIPAMENTO: string;
  NM_MODELO: string;
  LOGRADOUDO_NUMERO: string;
  NOM_BAIRRO: string;
  geoLat: number;
  geoLng: number;
  hasCoords: boolean;
  codOperadora: string; // Código cidade
  cmMac: string;
  num_contrato: string;
  cmtsStreamModulation: string;
  sheetName: string;
}

interface NodeHeatmapProps {
  persistentData: NodeData[];
  setPersistentData: (data: NodeData[]) => void;
  persistentFilters: {
    codOperadora: string;
    codNode: string[];
    statusModem: string;
    snrRanges: string[];
  };
  setPersistentFilters: (filters: any) => void;
}

const createClusterCustomIcon = (cluster: any) => {
  const count = cluster.getChildCount();
  let color = '#334155'; // slate-700 (default/small)
  let bgColor = '#f1f5f9'; // slate-100

  if (count >= 100) {
    color = '#ffffff';
    bgColor = '#EE1D23'; // Red
  } else if (count >= 20) {
    color = '#ffffff';
    bgColor = '#3b82f6'; // blue-500
  }

  return L.divIcon({
    html: `
      <div style="
        background-color: ${bgColor};
        width: 44px;
        height: 44px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        border: 3px solid white;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      ">
        <span style="
          color: ${color};
          font-weight: 900;
          font-size: 16px;
          font-family: 'Inter', sans-serif;
          text-shadow: ${color === '#ffffff' ? '0 1px 2px rgba(0,0,0,0.2)' : 'none'};
        ">${count}</span>
      </div>
    `,
    className: 'custom-cluster-icon',
    iconSize: L.point(44, 44, true),
  });
};

const NodeHeatmap = ({ 
  persistentData: data, 
  setPersistentData: setData, 
  persistentFilters: filters, 
  setPersistentFilters: setFilters 
}: NodeHeatmapProps) => {
  const [isLoading, setIsLoading] = useState(false);
  const [loadingProgress, setLoadingProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<'map' | 'table'>('map');

  const [skippedCount, setSkippedCount] = useState<number>(0);

  const onDrop = (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    if (!file) return;

    setIsLoading(true);
    setLoadingProgress(0);
    setError(null);
    setSkippedCount(0);

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const dataBuffer = e.target?.result;
        if (!dataBuffer) throw new Error('Falha ao ler o arquivo');
        
        await new Promise(resolve => setTimeout(resolve, 100));
        setLoadingProgress(5);
        
        const workbook = XLSX.read(dataBuffer, { type: 'array' });
        setLoadingProgress(15);
        
        const sheetNames = workbook.SheetNames;
        let allData: NodeData[] = [];
        let totalRowsFound = 0;
        
        for (let i = 0; i < sheetNames.length; i++) {
          const sheetName = sheetNames[i];
          const sheet = workbook.Sheets[sheetName];
          if (sheet) {
            await new Promise(resolve => setTimeout(resolve, 50));
            
            // Get all rows as arrays to find the header
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
            if (rows.length === 0) continue;

            // Find header row by looking for known column names
            const headerKeys = [
              'geoLat', 'latitude', 'lat', 'COORD_X', 'X', 'LAT_GEO', 'GEO_LAT', 'COORDX', 'COORD X',
              'geoLng', 'longitude', 'lng', 'lon', 'COORD_Y', 'Y', 'LNG_GEO', 'GEO_LNG', 'COORDY', 'COORD Y',
              'cmtsRxPot', 'RX_POT', 'rxPot', 'POTENCIA', 'POTÊNCIA', 'RX_POWER', 'POWER', 'POT_RX',
              'cmtsRxSNR', 'RX_SNR', 'rxSNR', 'SNR', 'RELAÇÃO_SINAL_RUÍDO', 'SIGNAL_NOISE', 'SNR_RX', 'SNR_CMTS', 'VALOR_SNR', 'SNR_VALOR', 'SNR_DB',
              'COD_NODE', 'NODE', 'node', 'NÓ', 'CODIGO_NODE', 'CÓDIGO_NODE', 'CD_NODE', 'NOME_NODE', 'ID_NODE', 'NODE_ID', 'DESC_NODE', 'NOME_NÓ', 'CD_NÓ', 'NODE_NAME', 'NODE_CODE',
              'codOperadora', 'CIDADE', 'cidade', 'OPERADORA', 'MUNICÍPIO', 'MUNICIPIO', 'CITY', 'COD_OPERADORA', 'CD_OPERADORA'
            ].map(k => k.toLowerCase().trim());

            let headerRowIndex = -1;
            for (let r = 0; r < Math.min(rows.length, 20); r++) {
              const row = rows[r];
              if (!Array.isArray(row)) continue;
              const hasKey = row.some(cell => 
                cell && typeof cell === 'string' && headerKeys.includes(cell.toLowerCase().trim())
              );
              if (hasKey) {
                headerRowIndex = r;
                break;
              }
            }

            const actualHeaderRowIndex = headerRowIndex === -1 ? 0 : headerRowIndex;
            const jsonData = XLSX.utils.sheet_to_json(sheet, { range: actualHeaderRowIndex }) as any[];
            totalRowsFound += jsonData.length;

            const mappedData = jsonData.map((row: any, idx: number) => {
            const getValue = (keys: string[]) => {
              const rowKeys = Object.keys(row);
              
              // 1. Try exact match first
              const exactKey = keys.find(k => row[k] !== undefined);
              if (exactKey) return row[exactKey];

              // 2. Try normalized match (remove spaces, underscores, etc)
              const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
              const normalizedSearchKeys = keys.map(normalize);
              
              const foundKey = rowKeys.find(rk => normalizedSearchKeys.includes(normalize(rk)));
              return foundKey ? row[foundKey] : undefined;
            };

              const parseCoord = (val: any) => {
                if (val === undefined || val === null || val === '') return 0;
                if (typeof val === 'number') return val;
                const cleanStr = String(val).replace(/[^\d.,-]/g, '').replace(',', '.');
                const parsed = parseFloat(cleanStr);
                return isNaN(parsed) ? 0 : parsed;
              };

              const lat = parseCoord(getValue(['geoLat', 'latitude', 'lat', 'LATITUDE', 'LAT', 'COORD_X', 'X', 'LAT_GEO', 'GEO_LAT', 'COORDX', 'COORD X', 'LAT_GEO', 'LATITUDE_GEO']));
              const lng = parseCoord(getValue(['geoLng', 'longitude', 'lng', 'lon', 'LONGITUDE', 'LNG', 'LON', 'COORD_Y', 'Y', 'LNG_GEO', 'GEO_LNG', 'COORDY', 'COORD Y', 'LNG_GEO', 'LONGITUDE_GEO']));
              
              const jitter = (idx % 100) * 0.00001;
              
              return {
                cmtsRxPot: Number(parseCoord(getValue(['cmtsRxPot', 'RX_POT', 'rxPot', 'POTENCIA', 'POTÊNCIA', 'RX_POWER', 'POWER', 'POT_RX', 'POTENCIA_RX']))),
                cmtsRxSNR: Number(parseCoord(getValue(['cmtsRxSNR', 'RX_SNR', 'rxSNR', 'SNR', 'RELAÇÃO_SINAL_RUÍDO', 'SIGNAL_NOISE', 'SNR_RX', 'SNR_CMTS', 'VALOR_SNR', 'SNR_VALOR', 'SNR_DB']))),
                COD_NODE: String(getValue(['COD_NODE', 'NODE', 'node', 'NÓ', 'CODIGO_NODE', 'CÓDIGO_NODE', 'CD_NODE', 'NOME_NODE', 'ID_NODE', 'NODE_ID', 'DESC_NODE', 'NOME_NÓ', 'CD_NÓ', 'NODE_NAME', 'NODE_CODE']) || '').trim(),
                NM_TIPO_EQUIPAMENTO: String(getValue(['NM_TIPO_EQUIPAMENTO', 'TIPO', 'tipo', 'EQUIPAMENTO', 'TIPO_EQUIPAMENTO', 'NM_TIPO']) || '').trim(),
                NM_MODELO: String(getValue(['NM_MODELO', 'MODELO', 'modelo', 'MODEL', 'NM_MODELO_EQUIPAMENTO']) || '').trim(),
                LOGRADOUDO_NUMERO: String(getValue(['LOGRADOUDO_NUMERO', 'ENDERECO', 'endereco', 'END', 'LOGRADOURO', 'ENDEREÇO', 'NUMERO', 'NÚMERO', 'LOGRADOURO_NUMERO']) || '').trim(),
                NOM_BAIRRO: String(getValue(['NOM_BAIRRO', 'BAIRRO', 'bairro', 'NEIGHBORHOOD', 'NOME_BAIRRO']) || '').trim(),
                geoLat: lat !== 0 ? lat + jitter : 0,
                geoLng: lng !== 0 ? lng + jitter : 0,
                hasCoords: lat !== 0 && lng !== 0,
                codOperadora: String(getValue(['codOperadora', 'CIDADE', 'cidade', 'OPERADORA', 'MUNICÍPIO', 'MUNICIPIO', 'CITY', 'COD_OPERADORA', 'CD_OPERADORA']) || '').trim(),
                cmMac: String(getValue(['cmMac', 'MAC', 'mac', 'ENDEREÇO_MAC', 'ENDERECO_MAC', 'MAC_ADDRESS', 'MAC_ADDR', 'CM_MAC']) || '').trim(),
                num_contrato: String(getValue(['num_contrato', 'NUM_CONTRATO', 'CONTRATO', 'contrato', 'CLIENTE', 'CONTRATO_NUM', 'NUMERO_CONTRATO', 'NUM_CONTR']) || '').trim(),
                cmtsStreamModulation: String(getValue(['cmtsStreamModulation', 'STATUS', 'status', 'MODULATION', 'MODULAÇÃO', 'MODULACAO', 'STATE', 'STATUS_MODEM']) || '').trim(),
                sheetName: sheetName
              };
            });
            
            allData = [...allData, ...mappedData];
            setLoadingProgress(15 + ((i + 1) / sheetNames.length) * 85);
          }
        }

        const withCoords = allData.filter(d => d.hasCoords).length;
        setSkippedCount(allData.length - withCoords);

        if (allData.length === 0) {
          setError(`Nenhum dado encontrado no arquivo.`);
        } else {
          setLoadingProgress(100);
          setData(allData);
        }
      } catch (err) {
        console.error(err);
        setError('Erro ao processar o arquivo Excel.');
      } finally {
        setIsLoading(false);
        setLoadingProgress(0);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({ 
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: false
  });

  const [showNodeFilter, setShowNodeFilter] = useState(false);
  const [showSnrFilter, setShowSnrFilter] = useState(false);
  const nodeFilterRef = React.useRef<HTMLDivElement>(null);
  const snrFilterRef = React.useRef<HTMLDivElement>(null);

  // Close filters when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (nodeFilterRef.current && !nodeFilterRef.current.contains(event.target as Node)) {
        setShowNodeFilter(false);
      }
      if (snrFilterRef.current && !snrFilterRef.current.contains(event.target as Node)) {
        setShowSnrFilter(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const filteredData = useMemo(() => {
    return data.filter(item => {
      const matchCidade = filters.codOperadora === 'Todos' || item.codOperadora === filters.codOperadora;
      const matchNode = filters.codNode.includes('Todos') || filters.codNode.includes(item.COD_NODE);
      const matchStatus = filters.statusModem === 'Todos' || item.cmtsStreamModulation === filters.statusModem;
      
      let matchSnr = filters.snrRanges.includes('Todos');
      if (!matchSnr) {
        matchSnr = filters.snrRanges.some(range => {
          if (range === '< 30 dB') return item.cmtsRxSNR < 30;
          if (range === '30 - 35 dB') return item.cmtsRxSNR >= 30 && item.cmtsRxSNR < 35;
          if (range === '> 35 dB') return item.cmtsRxSNR >= 35;
          return false;
        });
      }

      return matchCidade && matchNode && matchStatus && matchSnr;
    });
  }, [data, filters]);

  const cityOptions = useMemo(() => {
    const cities = Array.from(new Set(data.map(item => item.codOperadora))).sort();
    return ['Todos', ...cities];
  }, [data]);

  const [nodeSearch, setNodeSearch] = useState('');

  const nodeOptions = useMemo(() => {
    // Filter nodes based on selected city only to keep the list stable when filtering by status
    const relevantData = data.filter(item => {
      return filters.codOperadora === 'Todos' || item.codOperadora === filters.codOperadora;
    });
    
    const nodes = Array.from(new Set(relevantData.map(item => item.COD_NODE))).filter(Boolean).sort();
    
    if (nodeSearch) {
      return nodes.filter(node => node.toLowerCase().includes(nodeSearch.toLowerCase()));
    }
    
    return nodes;
  }, [data, filters.codOperadora, nodeSearch]);

  const snrOptions = ['< 30 dB', '30 - 35 dB', '> 35 dB'];

  const statusOptions = useMemo(() => {
    const statuses = Array.from(new Set(data.map(item => item.cmtsStreamModulation))).filter(Boolean).sort();
    return ['Todos', ...statuses];
  }, [data]);

  // Reset node filter only if it's no longer valid in the new city options
  useEffect(() => {
    if (!filters.codNode.includes('Todos')) {
      const validNodes = filters.codNode.filter(node => nodeOptions.includes(node));
      if (validNodes.length === 0) {
        setFilters((prev: any) => ({ ...prev, codNode: ['Todos'] }));
      } else if (validNodes.length !== filters.codNode.length) {
        setFilters((prev: any) => ({ ...prev, codNode: validNodes }));
      }
    }
  }, [nodeOptions]);

  const exportToExcel = () => {
    const exportData = filteredData.map(item => ({
      'Node': item.COD_NODE,
      'MAC': item.cmMac,
      'Cidade': item.codOperadora,
      'Contrato': item.num_contrato,
      'Rx Pot (dBm)': item.cmtsRxPot,
      'Rx SNR (dB)': item.cmtsRxSNR,
      'Status Modem': item.cmtsStreamModulation,
      'Tipo': item.NM_TIPO_EQUIPAMENTO,
      'Modelo': item.NM_MODELO,
      'Bairro': item.NOM_BAIRRO,
      'Endereço': item.LOGRADOUDO_NUMERO,
      'Latitude': item.geoLat,
      'Longitude': item.geoLng,
      'Aba': item.sheetName
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Nodes_Filtrados");
    XLSX.writeFile(wb, `Export_Nodes_${new Date().toISOString().split('T')[0]}.xlsx`);
  };
  const mapCenter: [number, number] = useMemo(() => {
    const dataWithCoords = filteredData.filter(d => d.hasCoords);
    if (dataWithCoords.length > 0) {
      const avgLat = dataWithCoords.reduce((sum, item) => sum + item.geoLat, 0) / dataWithCoords.length;
      const avgLng = dataWithCoords.reduce((sum, item) => sum + item.geoLng, 0) / dataWithCoords.length;
      return [avgLat, avgLng];
    }
    return [-15.7801, -47.9292]; // Brasilia
  }, [filteredData]);

  const ChangeView = ({ center, filteredData }: { center: [number, number], filteredData: NodeData[] }) => {
    const map = useMap();
    useEffect(() => {
      const dataWithCoords = filteredData.filter(d => d.hasCoords);
      if (dataWithCoords.length > 0) {
        if (dataWithCoords.length === 1) {
          map.setView(center, 16);
        } else {
          const bounds = L.latLngBounds(dataWithCoords.map(item => [item.geoLat, item.geoLng]));
          map.fitBounds(bounds, { padding: [50, 50], maxZoom: 16 });
        }
      }
    }, [center, map, filteredData]);
    return null;
  };

  return (
    <div className="max-w-7xl mx-auto pb-20 px-4">
      <div className="bg-white rounded-[32px] shadow-xl border border-slate-100">
        {/* Header */}
        <div className="p-8 border-b border-slate-100 bg-gradient-to-r from-white to-slate-50">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
            <div className="flex items-center gap-4">
              <div className="w-12 h-12 bg-red-50 rounded-2xl flex items-center justify-center text-[#EE1D23]">
                <MapIcon className="w-6 h-6" />
              </div>
              <div>
                <h2 className="text-2xl font-black text-[#333333] uppercase italic tracking-tighter">MAPA DE CALOR NODE</h2>
                <div className="flex flex-col">
                  <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Análise Geográfica de Performance</p>
                  {data.length > 0 && (
                    <div className="flex items-center flex-wrap gap-2 mt-1">
                      <span className="text-[10px] font-black bg-slate-100 text-slate-600 px-2 py-0.5 rounded uppercase tracking-widest">
                        {data.length} Total
                      </span>
                      <span className="text-[10px] font-black bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded uppercase tracking-widest">
                        {data.length - skippedCount} Com Coordenadas
                      </span>
                      {skippedCount > 0 && (
                        <span className="text-[10px] font-black bg-amber-50 text-amber-600 px-2 py-0.5 rounded uppercase tracking-widest">
                          {skippedCount} Sem Coordenadas
                        </span>
                      )}
                      <span className="text-[10px] font-black bg-blue-50 text-blue-600 px-2 py-0.5 rounded uppercase tracking-widest">
                        {nodeOptions.length} Nodes Únicos
                      </span>
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-3">
              <button
                onClick={() => setViewMode('map')}
                className={cn(
                  "px-4 py-2 rounded-xl text-xs font-black uppercase tracking-widest transition-all flex items-center gap-2",
                  viewMode === 'map' ? "bg-[#EE1D23] text-white shadow-lg shadow-red-500/20" : "bg-slate-100 text-slate-400 hover:bg-slate-200"
                )}
              >
                <MapIcon className="w-4 h-4" />
                Mapa
              </button>
              <button
                onClick={() => setViewMode('table')}
                className={cn(
                  "px-4 py-2 rounded-xl text-xs font-black uppercase tracking-widest transition-all flex items-center gap-2",
                  viewMode === 'table' ? "bg-[#EE1D23] text-white shadow-lg shadow-red-500/20" : "bg-slate-100 text-slate-400 hover:bg-slate-200"
                )}
              >
                <TableIcon className="w-4 h-4" />
                Tabela
              </button>
            </div>
          </div>
        </div>

        {/* Upload Area */}
        {data.length === 0 ? (
          <div className="p-12">
            <div 
              {...getRootProps()} 
              className={cn(
                "border-4 border-dashed rounded-[32px] p-16 text-center transition-all cursor-pointer",
                isDragActive ? "border-[#EE1D23] bg-red-50" : "border-slate-100 hover:border-slate-200 bg-slate-50"
              )}
            >
              <input {...getInputProps()} />
              <div className="w-24 h-24 bg-white rounded-3xl shadow-lg flex items-center justify-center mx-auto mb-6 text-[#EE1D23]">
                <Upload className="w-10 h-10" />
              </div>
              <h3 className="text-2xl font-black text-[#333333] uppercase italic tracking-tighter mb-2">
                {isDragActive ? "Solte o arquivo aqui" : "Importar Dados dos Nodes"}
              </h3>
              <p className="text-slate-400 font-bold mb-8">
                Arraste o arquivo Excel ou clique para selecionar.<br/>
                O sistema lerá <span className="text-slate-600">todas as abas</span> do arquivo.
              </p>
              {isLoading && (
                <div className="mt-8 max-w-md mx-auto">
                  <div className="flex items-center justify-between mb-2">
                    <div className="flex items-center gap-3 text-[#EE1D23] font-black uppercase italic text-xs">
                      <Activity className="w-4 h-4 animate-spin" />
                      Processando Dados...
                    </div>
                    <span className="text-xs font-black text-[#EE1D23]">{Math.round(loadingProgress)}%</span>
                  </div>
                  <div className="w-full h-3 bg-slate-100 rounded-full overflow-hidden border border-slate-200">
                    <motion.div 
                      className="h-full bg-[#EE1D23]"
                      initial={{ width: 0 }}
                      animate={{ width: `${loadingProgress}%` }}
                      transition={{ duration: 0.3 }}
                    />
                  </div>
                </div>
              )}
              {error && (
                <div className="flex items-center justify-center gap-2 text-red-500 font-bold bg-red-50 p-4 rounded-2xl border border-red-100 max-w-md mx-auto">
                  <AlertCircle className="w-5 h-5" />
                  {error}
                </div>
              )}
            </div>
          </div>
        ) : (
          <div className="flex flex-col h-[700px]">
            {/* Filters Bar */}
            <div className="p-6 border-b border-slate-100 bg-white flex flex-wrap items-center gap-4 relative z-[1000]">
              <div className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2">
                <Filter className="w-4 h-4 text-slate-400" />
                <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Código Cidade:</span>
                <select 
                  value={filters.codOperadora}
                  onChange={e => setFilters((prev: any) => ({ ...prev, codOperadora: e.target.value }))}
                  className="bg-transparent text-xs font-bold focus:outline-none cursor-pointer"
                >
                  {cityOptions.map(city => (
                    <option key={city} value={city}>{city}</option>
                  ))}
                </select>
              </div>

              {/* Multi-select Node Filter */}
              <div className="relative" ref={nodeFilterRef}>
                <button 
                  onClick={() => setShowNodeFilter(!showNodeFilter)}
                  className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 hover:border-[#EE1D23] transition-all min-w-[160px]"
                >
                  <Search className="w-4 h-4 text-slate-400" />
                  <div className="flex flex-col items-start text-left">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Node:</span>
                    <span className="text-xs font-bold truncate max-w-[120px]">
                      {filters.codNode.includes('Todos') ? 'Todos' : 
                       filters.codNode.length === 1 ? filters.codNode[0] : 
                       `${filters.codNode.length} selecionados`}
                    </span>
                  </div>
                  <ChevronDown className={cn("w-4 h-4 text-slate-400 ml-auto transition-transform", showNodeFilter && "rotate-180")} />
                </button>

                <AnimatePresence>
                  {showNodeFilter && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="absolute top-full left-0 mt-2 w-64 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[1001] p-4"
                    >
                      <div className="mb-3">
                        <div className="relative">
                          <Search className="w-3 h-3 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                          <input
                            type="text"
                            placeholder="Buscar node..."
                            value={nodeSearch}
                            onChange={(e) => setNodeSearch(e.target.value)}
                            className="w-full pl-8 pr-3 py-2 bg-slate-50 border border-slate-100 rounded-lg text-xs font-bold focus:outline-none focus:border-[#EE1D23] transition-all"
                          />
                        </div>
                      </div>

                      <div className="max-h-60 overflow-y-auto space-y-1">
                        <label className="flex items-center gap-3 p-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-colors">
                          <input 
                            type="checkbox" 
                            checked={filters.codNode.includes('Todos')}
                            onChange={() => setFilters((prev: any) => ({ ...prev, codNode: ['Todos'] }))}
                            className="w-4 h-4 rounded border-slate-300 text-[#EE1D23] focus:ring-[#EE1D23]"
                          />
                          <span className="text-xs font-bold text-slate-600">Todos</span>
                        </label>
                        <div className="h-px bg-slate-100 my-1" />
                        {nodeOptions.length > 0 ? (
                          nodeOptions.map(node => (
                            <label key={node} className="flex items-center gap-3 p-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-colors">
                              <input 
                                type="checkbox" 
                                checked={filters.codNode.includes(node)}
                                onChange={() => {
                                  setFilters((prev: any) => {
                                    let newNodes = prev.codNode.filter((n: string) => n !== 'Todos');
                                    if (newNodes.includes(node)) {
                                      newNodes = newNodes.filter((n: string) => n !== node);
                                      if (newNodes.length === 0) newNodes = ['Todos'];
                                    } else {
                                      newNodes.push(node);
                                    }
                                    return { ...prev, codNode: newNodes };
                                  });
                                }}
                                className="w-4 h-4 rounded border-slate-300 text-[#EE1D23] focus:ring-[#EE1D23]"
                              />
                              <span className="text-xs font-bold text-slate-600">{node}</span>
                            </label>
                          ))
                        ) : (
                          <div className="p-4 text-center">
                            <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Nenhum node encontrado</p>
                          </div>
                        )}
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Multi-select SNR Filter */}
              <div className="relative" ref={snrFilterRef}>
                <button 
                  onClick={() => setShowSnrFilter(!showSnrFilter)}
                  className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 hover:border-[#EE1D23] transition-all min-w-[140px]"
                >
                  <Signal className="w-4 h-4 text-slate-400" />
                  <div className="flex flex-col items-start text-left">
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">SNR:</span>
                    <span className="text-xs font-bold truncate max-w-[100px]">
                      {filters.snrRanges.includes('Todos') ? 'Todos' : 
                       filters.snrRanges.length === 1 ? filters.snrRanges[0] : 
                       `${filters.snrRanges.length} selecionados`}
                    </span>
                  </div>
                  <ChevronDown className={cn("w-4 h-4 text-slate-400 ml-auto transition-transform", showSnrFilter && "rotate-180")} />
                </button>

                <AnimatePresence>
                  {showSnrFilter && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="absolute top-full left-0 mt-2 w-56 bg-white border border-slate-100 rounded-2xl shadow-2xl z-[1001] p-4"
                    >
                      <div className="space-y-1">
                        <label className="flex items-center gap-3 p-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-colors">
                          <input 
                            type="checkbox" 
                            checked={filters.snrRanges.includes('Todos')}
                            onChange={() => setFilters((prev: any) => ({ ...prev, snrRanges: ['Todos'] }))}
                            className="w-4 h-4 rounded border-slate-300 text-[#EE1D23] focus:ring-[#EE1D23]"
                          />
                          <span className="text-xs font-bold text-slate-600">Todos</span>
                        </label>
                        <div className="h-px bg-slate-100 my-1" />
                        {snrOptions.map(range => (
                          <label key={range} className="flex items-center gap-3 p-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-colors">
                            <input 
                              type="checkbox" 
                              checked={filters.snrRanges.includes(range)}
                              onChange={() => {
                                setFilters((prev: any) => {
                                  let newRanges = prev.snrRanges.filter((r: string) => r !== 'Todos');
                                  if (newRanges.includes(range)) {
                                    newRanges = newRanges.filter((r: string) => r !== range);
                                    if (newRanges.length === 0) newRanges = ['Todos'];
                                  } else {
                                    newRanges.push(range);
                                  }
                                  return { ...prev, snrRanges: newRanges };
                                });
                              }}
                              className="w-4 h-4 rounded border-slate-300 text-[#EE1D23] focus:ring-[#EE1D23]"
                            />
                            <span className="text-xs font-bold text-slate-600">{range}</span>
                          </label>
                        ))}
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              <div className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-2">
                <Activity className="w-4 h-4 text-slate-400" />
                <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Status Modem:</span>
                <select 
                  value={filters.statusModem}
                  onChange={e => setFilters((prev: any) => ({ ...prev, statusModem: e.target.value }))}
                  className="bg-transparent text-xs font-bold focus:outline-none cursor-pointer"
                >
                  {statusOptions.map(status => (
                    <option key={status} value={status}>{status}</option>
                  ))}
                </select>
              </div>

              <div className="flex-1" />

              <div className="flex items-center gap-6">
                <div className="text-right">
                  <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Total de Modens</p>
                  <p className="text-lg font-black text-[#333333] tracking-tighter">{filteredData.length}</p>
                </div>
                <button
                  onClick={() => {
                    setData([]);
                    setFilters({ codOperadora: 'Todos', codNode: 'Todos', statusModem: 'Todos' });
                  }}
                  className="p-3 bg-slate-100 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-xl transition-all"
                  title="Limpar dados"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            {/* Content Area */}
            <div className="flex-1 relative">
              {viewMode === 'map' ? (
                <div className="h-full w-full z-10">
                  <MapContainer 
                    center={mapCenter} 
                    zoom={13} 
                    style={{ height: '100%', width: '100%' }}
                    scrollWheelZoom={true}
                    preferCanvas={true}
                  >
                    <TileLayer
                      attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                      url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
                    />
                    <ChangeView center={mapCenter} filteredData={filteredData} />
                    <MarkerClusterGroup
                      spiderfyOnMaxZoom={true}
                      showCoverageOnHover={false}
                      zoomToBoundsOnClick={true}
                      maxClusterRadius={50}
                      spiderfyDistanceMultiplier={1.5}
                      iconCreateFunction={createClusterCustomIcon}
                    >
                      {filteredData.filter(d => d.hasCoords).map((item, idx) => (
                        <Marker 
                          key={`${item.COD_NODE}-${idx}`} 
                          position={[item.geoLat, item.geoLng]}
                          icon={createModemIcon(
                            item.cmtsRxSNR < 30 ? '#ef4444' : 
                            item.cmtsRxSNR < 35 ? '#f97316' :
                            '#10b981'
                          )}
                        >
                          <Popup className="custom-popup">
                              <div className="p-2 min-w-[200px]">
                                <div className="flex items-center gap-2 mb-2 border-b border-slate-100 pb-2">
                                  <Database className="w-4 h-4 text-[#EE1D23]" />
                                  <span className="text-sm font-black text-[#333333] uppercase italic">Node: {item.COD_NODE}</span>
                                </div>
                                <div className="grid grid-cols-2 gap-3 mb-3">
                                  <div className="bg-slate-50 p-2 rounded-lg">
                                    <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest">Rx Pot</p>
                                    <div className="flex items-center gap-1">
                                      <Activity className="w-3 h-3 text-slate-400" />
                                      <span className="text-xs font-black text-[#333333]">{item.cmtsRxPot} dBm</span>
                                    </div>
                                  </div>
                                  <div className="bg-slate-50 p-2 rounded-lg">
                                    <p className="text-[8px] font-black text-slate-400 uppercase tracking-widest">Rx SNR</p>
                                    <div className="flex items-center gap-1">
                                      <Signal className={cn("w-3 h-3", item.cmtsRxSNR < 30 ? "text-red-500" : "text-emerald-500")} />
                                      <span className="text-xs font-black text-[#333333]">{item.cmtsRxSNR} dB</span>
                                    </div>
                                  </div>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-[9px] font-bold text-slate-500"><span className="font-black uppercase text-[8px] text-slate-400 mr-1">Contrato:</span> {item.num_contrato}</p>
                                  <p className="text-[9px] font-bold text-slate-500"><span className="font-black uppercase text-[8px] text-slate-400 mr-1">MAC:</span> {item.cmMac}</p>
                                  <p className="text-[9px] font-bold text-slate-500"><span className="font-black uppercase text-[8px] text-slate-400 mr-1">Status:</span> {item.cmtsStreamModulation}</p>
                                  <p className="text-[9px] font-bold text-slate-500"><span className="font-black uppercase text-[8px] text-slate-400 mr-1">TIPO:</span> {item.NM_TIPO_EQUIPAMENTO}</p>
                                  <p className="text-[9px] font-bold text-slate-500"><span className="font-black uppercase text-[8px] text-slate-400 mr-1">MODELO:</span> {item.NM_MODELO}</p>
                                  <p className="text-[9px] font-bold text-slate-500"><span className="font-black uppercase text-[8px] text-slate-400 mr-1">BAIRRO:</span> {item.NOM_BAIRRO}</p>
                                  <p className="text-[9px] font-bold text-slate-500"><span className="font-black uppercase text-[8px] text-slate-400 mr-1">END:</span> {item.LOGRADOUDO_NUMERO}</p>
                                </div>
                              </div>
                            </Popup>
                        </Marker>
                      ))}
                    </MarkerClusterGroup>
                  </MapContainer>
                </div>
              ) : (
                <div className="h-full flex flex-col">
                  <div className="p-4 border-b border-slate-100 bg-white flex justify-end">
                    <button
                      onClick={exportToExcel}
                      className="flex items-center gap-2 px-4 py-2 bg-emerald-500 text-white rounded-xl text-xs font-black uppercase tracking-widest hover:bg-emerald-600 transition-all shadow-lg shadow-emerald-500/20"
                    >
                      <FileSpreadsheet className="w-4 h-4" />
                      Exportar Excel
                    </button>
                  </div>
                  <div className="flex-1 overflow-auto">
                    <table className="w-full text-left border-collapse">
                      <thead className="sticky top-0 bg-slate-50 z-20">
                        <tr className="border-b border-slate-100">
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Node</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">MAC</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Cidade</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Contrato</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-center">Rx Pot</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-center">Rx SNR</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Status</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Tipo</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Bairro</th>
                          <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Aba</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredData.slice(0, 1000).map((item, idx) => (
                          <tr key={idx} className="border-b border-slate-50 hover:bg-slate-50 transition-colors">
                            <td className="px-6 py-4 text-xs font-black text-[#333333]">{item.COD_NODE}</td>
                            <td className="px-6 py-4 text-xs font-bold text-slate-500">{item.cmMac}</td>
                            <td className="px-6 py-4 text-xs font-bold text-slate-600">{item.codOperadora}</td>
                            <td className="px-6 py-4 text-xs font-bold text-slate-500">{item.num_contrato}</td>
                            <td className="px-6 py-4 text-xs font-black text-center text-slate-600">{item.cmtsRxPot}</td>
                            <td className="px-6 py-4 text-center">
                              <span className={cn(
                                "text-[10px] font-black px-2 py-1 rounded-lg",
                                item.cmtsRxSNR < 30 ? "bg-red-50 text-red-600" : "bg-emerald-50 text-emerald-600"
                              )}>
                                {item.cmtsRxSNR} dB
                              </span>
                            </td>
                            <td className="px-6 py-4 text-[10px] font-bold text-slate-500">{item.cmtsStreamModulation}</td>
                            <td className="px-6 py-4 text-xs font-bold text-slate-500">{item.NM_TIPO_EQUIPAMENTO}</td>
                            <td className="px-6 py-4 text-xs font-bold text-slate-500">{item.NOM_BAIRRO}</td>
                            <td className="px-6 py-4">
                              <span className="text-[9px] font-black bg-slate-100 text-slate-400 px-2 py-1 rounded uppercase tracking-widest">
                                {item.sheetName}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {filteredData.length > 1000 && (
                      <div className="p-8 text-center bg-slate-50 border-t border-slate-100">
                        <p className="text-sm font-bold text-slate-400 italic">
                          Mostrando os primeiros 1000 modens de {filteredData.length}. 
                          Use os filtros ou o mapa para ver todos.
                        </p>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default NodeHeatmap;
