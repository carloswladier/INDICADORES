import React from 'react';
import { RefreshCw, FileSpreadsheet, Layers, CheckCircle2 } from 'lucide-react';

interface FileLoadingOverlayProps {
  isOpen: boolean;
  title?: string;
  subtitle?: string;
  currentStep?: string;
  progress?: number;
}

export const FileLoadingOverlay: React.FC<FileLoadingOverlayProps> = ({
  isOpen,
  title = 'Carregando e Processando Arquivo...',
  subtitle = 'Mapeando colunas, processando planilhas e correlacionando indicadores',
  currentStep,
  progress
}) => {
  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-950/70 backdrop-blur-md animate-fade-in p-4"
      id="file-loading-screen"
    >
      <div className="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-md w-full p-8 flex flex-col items-center text-center relative overflow-hidden">
        {/* Glow effect */}
        <div className="absolute -top-12 -left-12 w-36 h-36 bg-[#EE1D23]/10 rounded-full blur-2xl pointer-events-none" />
        <div className="absolute -bottom-12 -right-12 w-36 h-36 bg-red-500/10 rounded-full blur-2xl pointer-events-none" />

        {/* Central animated icon */}
        <div className="relative mb-6">
          <div className="w-20 h-20 rounded-2xl bg-red-50 flex items-center justify-center text-[#EE1D23] shadow-inner relative">
            <FileSpreadsheet className="w-10 h-10 text-[#EE1D23] animate-pulse" />
            <div className="absolute -bottom-1 -right-1 w-7 h-7 rounded-full bg-[#EE1D23] text-white flex items-center justify-center shadow-md">
              <RefreshCw className="w-4 h-4 animate-spin" />
            </div>
          </div>
        </div>

        {/* Title and Subtitle */}
        <span className="px-3 py-1 rounded-full bg-red-50 text-[#EE1D23] text-[10px] font-black uppercase tracking-widest mb-2 inline-block">
          Processamento de Dados
        </span>
        <h3 className="text-xl font-black text-slate-900 uppercase italic tracking-tight mb-2">
          {title}
        </h3>
        <p className="text-xs text-slate-500 font-bold max-w-xs leading-relaxed mb-6">
          {subtitle}
        </p>

        {/* Animated Progress Bar */}
        <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden mb-4 relative">
          <div
            className="h-full bg-gradient-to-r from-[#EE1D23] via-red-500 to-[#EE1D23] rounded-full transition-all duration-300 animate-[pulse_1.5s_infinite]"
            style={{ width: progress !== undefined ? `${Math.min(100, Math.max(15, progress))}%` : '85%' }}
          />
        </div>

        {/* Step Indicator */}
        <div className="w-full bg-slate-50 border border-slate-100 rounded-2xl p-3 flex items-center justify-between text-left">
          <div className="flex items-center gap-2.5">
            <div className="w-6 h-6 rounded-lg bg-red-100/70 text-[#EE1D23] flex items-center justify-center shrink-0">
              <Layers className="w-3.5 h-3.5" />
            </div>
            <div>
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-wider block">Etapa Atual</span>
              <span className="text-xs font-bold text-slate-700 block truncate max-w-[220px]">
                {currentStep || 'Mapeando colunas e correlacionando datas (Início x Abertura Solic)...'}
              </span>
            </div>
          </div>
          <RefreshCw className="w-3.5 h-3.5 text-[#EE1D23] animate-spin shrink-0" />
        </div>

        <p className="text-[10px] text-slate-400 font-medium mt-4">
          Por favor, aguarde. Os dados estão sendo estruturados com precisão.
        </p>
      </div>
    </div>
  );
};
