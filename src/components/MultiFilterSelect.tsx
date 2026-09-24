import React, { useState, useRef, useEffect } from 'react';
import { ChevronDown, Search, Check } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import { cn } from '../lib/utils';

export interface MultiFilterSelectProps {
  label: string;
  icon?: React.ReactNode;
  value: string[];
  options: string[];
  onChange: (value: string[]) => void;
  getOptionLabel?: (opt: string) => string;
  placeholder?: string;
  className?: string;
  optionCounts?: Record<string, number>;
  dropdownClassName?: string;
  dropdownAlign?: 'left' | 'right';
}

export function MultiFilterSelect({
  label,
  icon,
  value,
  options,
  onChange,
  getOptionLabel,
  placeholder = 'Pesquisar...',
  className,
  optionCounts,
  dropdownClassName,
  dropdownAlign = 'left',
}: MultiFilterSelectProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  useEffect(() => {
    if (!isOpen) setSearchTerm('');
  }, [isOpen]);

  const allOptions = options.includes('Todos') ? options : ['Todos', ...options];

  const isTodosSelected = value.length === 0 || value.includes('Todos');

  const toggleOption = (opt: string) => {
    if (opt === 'Todos') {
      onChange([]);
      return;
    }

    let newValue = value.filter(v => v !== 'Todos');

    if (newValue.includes(opt)) {
      newValue = newValue.filter(v => v !== opt);
    } else {
      newValue.push(opt);
    }

    onChange(newValue);
  };

  const filteredOptions = allOptions.filter(opt => {
    if (opt === 'Todos') return true;
    const labelText = getOptionLabel ? getOptionLabel(opt) : opt;
    return labelText.toLowerCase().includes(searchTerm.toLowerCase()) || opt.toLowerCase().includes(searchTerm.toLowerCase());
  });

  const displayValue = isTodosSelected
    ? 'Todos'
    : value.length === 1
      ? (getOptionLabel ? getOptionLabel(value[0]) : value[0])
      : `${value.length} selecionados`;

  const MAX_OPTIONS = 150;
  const visibleOptions = filteredOptions.slice(0, MAX_OPTIONS);

  return (
    <div className={cn("relative space-y-1.5", className)} ref={containerRef}>
      {label && (
        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-1.5 ml-1 h-5 whitespace-nowrap overflow-hidden text-ellipsis">
          {icon && <span className="text-[#EE1D23] shrink-0">{icon}</span>}
          <span className="truncate">{label}</span>
        </label>
      )}
      <div className="relative">
        <button
          type="button"
          onClick={() => setIsOpen(!isOpen)}
          title={typeof displayValue === 'string' ? displayValue : undefined}
          className="w-full h-10 bg-slate-50 border border-slate-200/80 hover:bg-slate-100/60 text-[#333333] text-xs font-bold rounded-xl px-3 outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all cursor-pointer flex items-center justify-between shadow-2xs"
        >
          <span className="truncate pr-1 text-left flex-1 min-w-0">{displayValue}</span>
          <ChevronDown className={cn("w-4 h-4 text-slate-400 transition-transform shrink-0 ml-1", isOpen && "rotate-180")} />
        </button>

        <AnimatePresence>
          {isOpen && (
            <motion.div
              initial={{ opacity: 0, y: 6, scale: 0.98 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 6, scale: 0.98 }}
              transition={{ duration: 0.15 }}
              className={cn(
                "absolute z-50 mt-1.5 min-w-[280px] sm:min-w-[320px] max-w-[92vw] bg-white border border-slate-200 rounded-2xl shadow-2xl overflow-hidden flex flex-col",
                dropdownAlign === 'right' ? 'right-0' : 'left-0',
                dropdownClassName
              )}
            >
              <div className="p-2.5 border-b border-slate-100 bg-white sticky top-0 z-10">
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                  <input
                    type="text"
                    placeholder={placeholder}
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl pl-8 pr-3 py-1.5 text-xs font-bold text-slate-700 outline-none focus:border-[#EE1D23] focus:bg-white transition-all"
                    autoFocus
                  />
                </div>
              </div>

              <div className="overflow-y-auto p-1.5 max-h-64 space-y-0.5 scrollbar-thin">
                {visibleOptions.length > 0 ? (
                  visibleOptions.map((opt) => {
                    const isSelected = opt === 'Todos' ? isTodosSelected : (!isTodosSelected && value.includes(opt));
                    const optLabel = opt === 'Todos' ? 'Todos' : (getOptionLabel ? getOptionLabel(opt) : opt);

                    return (
                      <div
                        key={opt}
                        onClick={() => toggleOption(opt)}
                        className={cn(
                          "flex items-center gap-2.5 px-3 py-2 rounded-xl cursor-pointer transition-all select-none hover:bg-slate-50",
                          isSelected ? "bg-red-50/70 text-[#EE1D23]" : "text-slate-700"
                        )}
                        title={optLabel}
                      >
                        <div
                          className={cn(
                            "w-4 h-4 rounded-md border flex items-center justify-center transition-all shrink-0",
                            isSelected
                              ? "bg-[#EE1D23] border-[#EE1D23] text-white shadow-xs"
                              : "bg-white border-slate-300"
                          )}
                        >
                          {isSelected && <Check className="w-3 h-3 stroke-[3]" />}
                        </div>
                        <span className={cn("text-xs font-bold truncate flex-1 min-w-0 pr-2", isSelected ? "text-[#EE1D23]" : "text-slate-700")}>
                          {optLabel}
                        </span>
                        {optionCounts && opt !== 'Todos' && optionCounts[opt] !== undefined && (
                          <span className={cn(
                            "ml-auto text-[10px] font-black px-2 py-0.5 rounded-md shrink-0 transition-colors tabular-nums",
                            isSelected ? "bg-red-100 text-[#EE1D23]" : "bg-slate-100 text-slate-500"
                          )}>
                            {optionCounts[opt].toLocaleString()}
                          </span>
                        )}
                      </div>
                    );
                  })
                ) : (
                  <div className="p-4 text-center text-xs font-bold text-slate-400 uppercase italic">
                    Nenhum resultado
                  </div>
                )}
                {filteredOptions.length > MAX_OPTIONS && (
                  <div className="p-2 text-center text-[10px] font-bold text-slate-400 bg-slate-50 rounded-lg mt-1 border border-slate-100">
                    Mostrando 150 de {filteredOptions.length.toLocaleString()} opções. Digite para refinar a busca.
                  </div>
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
}
