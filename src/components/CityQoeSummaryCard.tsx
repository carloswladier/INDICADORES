import React from 'react';
import { Calculator, Info, CheckCircle2, AlertTriangle, AlertCircle } from 'lucide-react';
import { CityQoeCalculation } from '../data/qoeGponData';
import { cn } from '../lib/utils';

export interface CityQoeSummaryCardProps {
  selectedCity: string;
  onSelectCity: (city: string) => void;
  availableCities: string[];
  calculation: CityQoeCalculation;
}

export const CityQoeSummaryCard: React.FC<CityQoeSummaryCardProps> = ({
  selectedCity,
  onSelectCity,
  availableCities,
  calculation,
}) => {
  const {
    cidade,
    totalModems,
    cronicoOffline,
    cronicoImpactado,
    cronicoEstressado,
    impactado,
    estressado,
    qoe,
    color,
    statusLabel,
  } = calculation;

  // Conforme o print do usuário: o cálculo leva APENAS em consideração Impactado e Estressado
  const penaltyPoints = impactado * 2.5 + estressado * 1.0;
  const penaltyPct = totalModems > 0 ? (penaltyPoints / totalModems) * 100 : 0;

  return (
    <section className="bg-white rounded-3xl p-6 shadow-sm border border-slate-100 overflow-hidden">
      {/* Header bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-slate-100">
        <div className="flex items-center gap-2.5">
          <div className="w-9 h-9 rounded-2xl bg-red-50 border border-red-100 flex items-center justify-center text-[#EE1D23]">
            <Calculator className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h3 className="text-base font-black uppercase tracking-tight text-slate-900">
                Memória de Cálculo QoE da Cidade
              </h3>
              <span className="px-2 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wide bg-emerald-50 text-emerald-700 border border-emerald-200">
                Apenas Impactado e Estressado
              </span>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Cálculo baseado estritamente na proporção de terminais <strong>Impactados</strong> e <strong>Estressados</strong> sobre o total de modems reportados.
            </p>
          </div>
        </div>

        {/* City selector dropdown */}
        <div className="flex items-center gap-2 self-start sm:self-auto">
          <label htmlFor="city-qoe-select" className="text-xs font-bold text-slate-500 whitespace-nowrap">
            Cidade Selecionada:
          </label>
          <select
            id="city-qoe-select"
            value={cidade || selectedCity}
            onChange={(e) => onSelectCity(e.target.value)}
            className="bg-slate-50 border border-slate-200 text-xs font-black text-slate-800 rounded-xl px-3 py-2 outline-none focus:ring-2 focus:ring-red-500/20 focus:border-[#EE1D23] transition-all cursor-pointer shadow-sm"
          >
            {availableCities.map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>
        </div>
      </div>

      {/* Excel Reproduction Layout matching the user's uploaded print */}
      <div className="mt-6 font-sans">
        {/* Top summary row: TOTAL DE MODEMS + Threshold Color Scale */}
        <div className="flex flex-wrap items-center justify-between gap-6 pb-6 border-b border-slate-100">
          {/* Total de Modems Box (reproducing print header) */}
          <div className="flex items-center gap-4">
            <span className="text-xs font-black uppercase tracking-wider text-slate-800">
              TOTAL DE MODEMS
            </span>
            <div className="text-center">
              <div className="text-[10px] font-black uppercase tracking-widest text-slate-500 mb-0.5">
                REPORTES
              </div>
              <div className="bg-slate-200/80 border-2 border-slate-800 px-6 py-1.5 rounded text-lg font-black text-slate-900 font-mono min-w-[120px] shadow-inner">
                {totalModems.toLocaleString()}
              </div>
            </div>
          </div>

          {/* Threshold Colors Scale (reproducing print right side) */}
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2">
              <span className="text-xs font-black text-slate-800 font-mono">QoE &gt;= 80</span>
              <div className="w-16 h-6 bg-[#86EFAC] border border-slate-700 rounded-sm shadow-sm" title="Meta Atingida (>= 80)"></div>
            </div>
            <div className="w-16 h-6 bg-[#FDE047] border border-slate-700 rounded-sm shadow-sm" title="Atenção (40 a 79)"></div>
            <div className="flex items-center gap-2">
              <span className="text-xs font-black text-slate-800 font-mono">QoE &lt; 40</span>
              <div className="w-16 h-6 bg-[#EF4444] border border-slate-700 rounded-sm shadow-sm" title="Crítico (< 40)"></div>
            </div>
          </div>
        </div>

        {/* The Table Layout exactly as formatted in the print */}
        <div className="mt-6 overflow-x-auto">
          <table className="w-full border-collapse text-xs">
            <thead>
              <tr>
                <th className="text-left font-black uppercase tracking-wider text-slate-800 py-2.5 pr-4 w-52 sm:w-60">
                  STATUS DOS MODEMS
                </th>
                <th className="text-center font-black uppercase tracking-wider text-slate-800 py-2.5 px-4 w-36">
                  QUANTIDADE
                </th>
                <th className="text-left font-black uppercase tracking-wider text-slate-500 py-2.5 px-4">
                  CRITÉRIO OPERACIONAL
                </th>
              </tr>
            </thead>
            <tbody className="space-y-2">
              {/* Crônico + Offline */}
              <tr className="border-t border-slate-100">
                <td className="py-2 pr-4">
                  <div className="border border-slate-800 rounded px-3 py-1.5 font-bold text-slate-800 bg-white shadow-xs">
                    Crônico + Offline
                  </div>
                </td>
                <td className="py-2 px-4 text-center">
                  <div className="bg-slate-200/80 border border-slate-800 rounded py-1.5 font-mono font-black text-slate-900 shadow-inner">
                    {cronicoOffline}
                  </div>
                </td>
                <td className="py-2 px-4 text-slate-600 text-xs">
                  Em Node Summary, clicar em Offline e contabilizar os modems com status Chronic
                </td>
              </tr>

              {/* Crônico + Impactado */}
              <tr>
                <td className="py-2 pr-4">
                  <div className="border border-slate-800 rounded px-3 py-1.5 font-bold text-slate-800 bg-white shadow-xs">
                    Crônico + Impactado
                  </div>
                </td>
                <td className="py-2 px-4 text-center">
                  <div className="bg-slate-200/80 border border-slate-800 rounded py-1.5 font-mono font-black text-slate-900 shadow-inner">
                    {cronicoImpactado}
                  </div>
                </td>
                <td className="py-2 px-4 text-slate-600 text-xs">
                  Em Node Summary, clicar em Chronic e contabilizar os modems com status Impacted
                </td>
              </tr>

              {/* Crônico + Estressado */}
              <tr>
                <td className="py-2 pr-4">
                  <div className="border border-slate-800 rounded px-3 py-1.5 font-bold text-slate-800 bg-white shadow-xs">
                    Crônico + Estressado
                  </div>
                </td>
                <td className="py-2 px-4 text-center">
                  <div className="bg-slate-200/80 border border-slate-800 rounded py-1.5 font-mono font-black text-slate-900 shadow-inner">
                    {cronicoEstressado}
                  </div>
                </td>
                <td className="py-2 px-4 text-slate-600 text-xs">
                  Em Node Summary, clicar em Chronic e contabilizar os modems com status Stressed
                </td>
              </tr>

              {/* Impactado (Participa do Cálculo) */}
              <tr className="bg-orange-50/40 rounded-lg">
                <td className="py-2 pr-4">
                  <div className="border-2 border-orange-500 rounded px-3 py-1.5 font-black text-orange-950 bg-white shadow-xs flex items-center justify-between">
                    <span>Impactado</span>
                    <span className="text-[10px] bg-orange-100 text-orange-800 px-1.5 py-0.2 rounded font-bold">
                      Peso 2.5
                    </span>
                  </div>
                </td>
                <td className="py-2 px-4 text-center">
                  <div className="bg-slate-200/90 border-2 border-orange-500 rounded py-1.5 font-mono font-black text-orange-950 shadow-inner text-sm">
                    {impactado}
                  </div>
                </td>
                <td className="py-2 px-4 text-orange-900 font-medium text-xs">
                  Em Node Summary, contabilizar os modems com status Impacted &bull;{' '}
                  <span className="text-orange-700 font-bold">{(impactado * 2.5).toFixed(1)} pontos ponderados</span>
                </td>
              </tr>

              {/* Estressado (Participa do Cálculo) */}
              <tr className="bg-amber-50/40 rounded-lg">
                <td className="py-2 pr-4">
                  <div className="border-2 border-amber-500 rounded px-3 py-1.5 font-black text-amber-950 bg-white shadow-xs flex items-center justify-between">
                    <span>Estressado</span>
                    <span className="text-[10px] bg-amber-100 text-amber-800 px-1.5 py-0.2 rounded font-bold">
                      Peso 1.0
                    </span>
                  </div>
                </td>
                <td className="py-2 px-4 text-center">
                  <div className="bg-slate-200/90 border-2 border-amber-500 rounded py-1.5 font-mono font-black text-amber-950 shadow-inner text-sm">
                    {estressado}
                  </div>
                </td>
                <td className="py-2 px-4 text-amber-900 font-medium text-xs">
                  Em Node Summary, contabilizar os modems com status Stressed &bull;{' '}
                  <span className="text-amber-700 font-bold">{(estressado * 1.0).toFixed(1)} pontos ponderados</span>
                </td>
              </tr>

              {/* Linha Final: QoE (Resultado exato do print) */}
              <tr className="border-t-2 border-slate-300">
                <td className="py-3 pr-4">
                  <div className="border-2 border-slate-900 rounded px-3 py-2 font-black text-base uppercase tracking-wider text-slate-900 bg-white shadow-xs">
                    QoE
                  </div>
                </td>
                <td className="py-3 px-4 text-center">
                  <div
                    className={cn(
                      'border-2 border-slate-900 rounded py-2 font-mono font-black text-xl shadow-md transition-colors',
                      color === 'green'
                        ? 'bg-[#86EFAC] text-emerald-950'
                        : color === 'yellow'
                        ? 'bg-[#FDE047] text-amber-950'
                        : 'bg-[#EF4444] text-white'
                    )}
                  >
                    {qoe}
                  </div>
                </td>
                <td className="py-3 px-4 text-slate-700 font-bold text-xs">
                  <div className="flex items-center gap-2">
                    <span className={cn(
                      'px-2 py-0.5 rounded text-[11px] font-black uppercase tracking-wide',
                      color === 'green' ? 'bg-emerald-100 text-emerald-800' :
                      color === 'yellow' ? 'bg-amber-100 text-amber-800' :
                      'bg-red-100 text-red-800'
                    )}>
                      {statusLabel}
                    </span>
                    <span>
                      Dedução: -{penaltyPct.toFixed(2)}% ({penaltyPoints.toFixed(1)} pts sobre {totalModems.toLocaleString()} modems)
                    </span>
                  </div>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* Footer calculation notice */}
        <div className="mt-4 p-3.5 rounded-2xl bg-slate-50 border border-slate-200/80 flex items-start gap-2.5 text-xs text-slate-600">
          <Info className="w-4 h-4 text-slate-500 mt-0.5 shrink-0" />
          <div>
            <strong>Regra Aplicada:</strong> O cálculo da nota QoE GPON leva em consideração <strong>apenas</strong> os modems <strong>Impactados</strong> e <strong>Estressados</strong>. Os registros de Crônicos são monitorados para conformidade mas não penalizam a nota da cidade.
            <div className="mt-1 font-mono text-[11px] text-slate-700 bg-white px-2 py-1 rounded border border-slate-200 inline-block">
              QoE = 100 - [ (Impactados × 2.5 + Estressados × 1.0) / Total de Modems ] × 100
            </div>
            {totalModems === 1250 && impactado === 63 && estressado === 49 && (
              <span className="ml-2 font-bold text-emerald-700">
                &rarr; 100 - [ (63 × 2.5 + 49 × 1.0) / 1250 ] × 100 = 83,48% &rarr; Nota 83
              </span>
            )}
          </div>
        </div>
      </div>
    </section>
  );
};
