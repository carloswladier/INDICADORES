export interface QoeGponRow {
  id: string | number;
  contrato: string;
  mes: string;
  cidade: string;
  olt: string;
  topologia: string;
  nmModelo?: string;
  status: 'OK' | 'IMPACTED' | 'OFFLINE' | 'STRESSED' | string;
  temperatura: number | null;
  rxOnuCliente: number | null;
  impacted: 'YES' | 'NO' | string;
  stressed: 'YES' | 'NO' | string;
  cronico: 'YES' | 'NO' | string;
  logradouro?: string;
  logradouroNumero?: string;
  bairro?: string;
  numero?: string;
  cep?: string;
  slotPon?: string;
  onuId?: string;
  serial?: string;
  txOlt?: number | null;
  txOnu?: number | null;
  cliente?: string;
  dataHora?: string;
  raw?: Record<string, any>;
}

// Mapeamento obrigatório das iniciais da coluna OLT para os nomes completos das cidades
export const OLT_CITY_MAP: Record<string, string> = {
  'AIU': 'ANANINDEUA',
  'BLM': 'BELÉM',
  'CAH': 'CASTANHAL',
  'CXS': 'CAXIAS',
  'ITZ': 'IMPERATRIZ',
  'MBA': 'MARABÁ',
  'MNS': 'MANAUS',
  'MPA': 'MACAPÁ',
  'PGN': 'PARAGOMINAS',
  'PUP': 'PARAUAPEBAS',
  'SLS': 'SÃO LUÍS',
  'SQA': 'SANTANA',
  'TMN': 'TIMON',
};

export const CITIES_LIST = [
  'ANANINDEUA',
  'BELÉM',
  'CASTANHAL',
  'CAXIAS',
  'IMPERATRIZ',
  'MARABÁ',
  'MANAUS',
  'MACAPÁ',
  'PARAGOMINAS',
  'PARAUAPEBAS',
  'SÃO LUÍS',
  'SANTANA',
  'TIMON',
];

/**
 * Converte as iniciais da OLT (ou token) para a cidade correspondente
 */
export function getCityFromOlt(oltName: string, fallbackCity?: string): string {
  if (!oltName && fallbackCity) {
    return normalizeCityName(fallbackCity);
  }
  if (!oltName) return 'NÃO INFORMADO';

  const clean = oltName.trim().toUpperCase();
  
  // 1. Testa as primeiras 3 letras
  const prefix3 = clean.substring(0, 3);
  if (OLT_CITY_MAP[prefix3]) {
    return OLT_CITY_MAP[prefix3];
  }

  // 2. Testa delimitadores comuns como BLM-OLT-01, OLT_BLM_02, GPON-AIU-01
  for (const [code, fullCity] of Object.entries(OLT_CITY_MAP)) {
    const regex = new RegExp(`(^|[-_\\s./])${code}([-_\\s./]|\\d|$)`, 'i');
    if (regex.test(clean) || clean.includes(code)) {
      return fullCity;
    }
  }

  // 3. Fallback para a cidade informada no registro caso exista
  if (fallbackCity && fallbackCity.trim()) {
    return normalizeCityName(fallbackCity);
  }

  return clean.substring(0, 3);
}

function normalizeCityName(city: string): string {
  const c = city.trim().toUpperCase();
  for (const fullCity of CITIES_LIST) {
    if (c === fullCity || c.includes(fullCity) || fullCity.includes(c)) {
      return fullCity;
    }
  }
  // Remove acentos para comparação
  const norm = c.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  for (const fullCity of CITIES_LIST) {
    const normFull = fullCity.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (norm === normFull) return fullCity;
  }
  return c;
}

/**
 * Classificação da faixa de RX_ONU_CLIENTE conforme solicitação:
 * - -8 até -24,99: Verde
 * - -25 até -26,99: Laranja
 * - >= -27 e <= -27,99: Amarelo
 * - >= -28 (ou <= -28 em termos numéricos de dBm): Vermelho
 */
export type RxRangeCategory = 
  | '-8 até -24,99'
  | '-25 até -26,99'
  | '>=-27 e <= -27,99'
  | '>= -28';

export function getRxClassification(rx: number | null | undefined): {
  range: RxRangeCategory | 'Outros';
  color: 'green' | 'orange' | 'yellow' | 'red' | 'gray';
  label: string;
} {
  if (rx === null || rx === undefined || isNaN(rx)) {
    return { range: 'Outros', color: 'gray', label: 'Sem Leitura' };
  }

  const val = rx > 0 ? -rx : rx; // Normaliza se vier positivo por engano

  // Faixa 1: -8 até -24,99 (Verde)
  if (val >= -24.99 && val <= -8.0) {
    return { range: '-8 até -24,99', color: 'green', label: '-8 a -24.99 dBm (Normal)' };
  }

  // Faixa 2: -25 até -26,99 (Laranja)
  if (val >= -26.99 && val <= -25.0) {
    return { range: '-25 até -26,99', color: 'orange', label: '-25 a -26.99 dBm (Alerta)' };
  }

  // Faixa 3: >= -27 e <= -27,99 (Amarelo)
  if (val >= -27.99 && val <= -27.0) {
    return { range: '>=-27 e <= -27,99', color: 'yellow', label: '-27 a -27.99 dBm (Atenção)' };
  }

  // Faixa 4: >= -28 (ou pior, <= -28 dBm, Vermelho)
  if (val < -27.99 || val <= -28.0) {
    return { range: '>= -28', color: 'red', label: '≤ -28 dBm (Crítico)' };
  }

  // Para valores entre -24.99 e -25.00
  if (val < -24.99 && val > -25.0) {
    return { range: '-25 até -26,99', color: 'orange', label: '-25 a -26.99 dBm (Alerta)' };
  }

  // Para valores entre -26.99 e -27.00
  if (val < -26.99 && val > -27.0) {
    return { range: '>=-27 e <= -27,99', color: 'yellow', label: '-27 a -27.99 dBm (Atenção)' };
  }

  // Se for maior que -8 (potência muito alta/saturação)
  if (val > -8.0) {
    return { range: '-8 até -24,99', color: 'green', label: 'Alta Potência' };
  }

  return { range: '>= -28', color: 'red', label: 'Crítico' };
}

/**
 * Estrutura da Memória de Cálculo do QoE da Cidade (Padrão Node Summary)
 */
export interface CityQoeCalculation {
  cidade: string;
  totalModems: number;
  cronicoOffline: number;
  cronicoImpactado: number;
  cronicoEstressado: number;
  impactado: number;
  estressado: number;
  qoe: number; // 0 a 100
  color: 'green' | 'yellow' | 'red';
  statusLabel: string;
}

/**
 * Calcula a Nota QoE da cidade conforme a lógica do Node Summary:
 * Total de Modems (Reportes)
 * Para o cálculo da nota QoE, leva APENAS em consideração:
 * - Impactado: peso 2.5
 * - Estressado: peso 1.0
 * (Crônicos constam na tabela para rastreabilidade/auditoria, mas NÃO entram no cálculo da nota)
 * 
 * Faixas de Conformidade:
 * - QoE >= 80: Verde (#10B981)
 * - 40 <= QoE < 80: Amarelo (#EAB308)
 * - QoE < 40: Vermelho (#EE1D23)
 * 
 * Exemplo do Print (Ananindeua):
 * Total Modems = 1250 | Impactado = 63 | Estressado = 49
 * Penalidade = (63*2.5 + 49*1.0) / 1250 * 100 = 206.5 / 12.5 = 16.52%
 * QoE = 100 - 16.52% = 83.48% => Round: 83 (Verde)
 */
export function calculateCityQoe(counts: {
  cidade?: string;
  totalModems: number;
  cronicoOffline: number;
  cronicoImpactado: number;
  cronicoEstressado: number;
  impactado: number;
  estressado: number;
}): CityQoeCalculation {
  const {
    cidade = '',
    totalModems,
    cronicoOffline,
    cronicoImpactado,
    cronicoEstressado,
    impactado,
    estressado,
  } = counts;

  if (totalModems <= 0) {
    return {
      cidade,
      totalModems: 0,
      cronicoOffline: 0,
      cronicoImpactado: 0,
      cronicoEstressado: 0,
      impactado: 0,
      estressado: 0,
      qoe: 100,
      color: 'green',
      statusLabel: 'QoE >= 80 (Meta Atingida)',
    };
  }

  // O cálculo da nota QoE GPON leva APENAS em consideração Impactado e Estressado conforme o print
  const penaltyPoints = (impactado * 2.5) + (estressado * 1.0);

  const penaltyPct = (penaltyPoints / totalModems) * 100;
  const rawQoe = Math.max(0, Math.min(100, 100 - penaltyPct));
  const qoe = Math.round(rawQoe);

  let color: 'green' | 'yellow' | 'red' = 'green';
  let statusLabel = 'QoE >= 80 (Meta Atingida)';

  if (qoe >= 80) {
    color = 'green';
    statusLabel = 'QoE >= 80';
  } else if (qoe >= 40) {
    color = 'yellow';
    statusLabel = '40 <= QoE < 80';
  } else {
    color = 'red';
    statusLabel = 'QoE < 40';
  }

  return {
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
  };
}

/**
 * Gera dados de exemplo realistas para o primeiro carregamento da aba
 * Inclui o volume real do terminal de Ananindeua (1250 modems: 63 Impactados, 49 Estressados, QoE = 83)
 */
export function generateSampleQoeGponData(): QoeGponRow[] {
  const rows: QoeGponRow[] = [];
  const currentMonth = 'Setembro';

  let idCounter = 100000;

  // 1. Geração específica e precisa para ANANINDEUA conforme a solicitação do usuário:
  // Total = 1250, Impactado = 63, Estressado = 49, Crônicos = 0, OK = 1138
  const ananindeuaTotal = 1250;
  const ananindeuaImpacted = 63;
  const ananindeuaStressed = 49;
  const ananindeuaOk = ananindeuaTotal - ananindeuaImpacted - ananindeuaStressed; // 1138

  const ananindeuaLogradouros = [
    'TV DE BREVES',
    'Av. Mário Covas',
    'Rodovia BR-316',
    'Tv. We 16 (Cidade Nova)',
    'Av. Três Corações',
    'Av. Dom Vicente Zico',
    'Rua Cláudio Sanders',
    'Av. Independência',
    'Passagem São Pedro',
    'Rua da Providência',
    'Estrada do Icuí-Guajará',
    'Conjunto Guajará I',
    'Av. Arterial 18',
    'Av. Arterial 5A',
    'Rua Leopoldo Teixeira',
    'Passagem Santo Antônio',
  ];

  const brevesNumbers = ['1182', '333', '752', '746', '882', '1633', '1190', '1118', '759', '763', '392', '790', '1303', '1174', '1494', '951', '837', '842', '845', '1111', '58', '375', '1115', '704'];
  const topologyCodes = [
    'AGL.AA.021.00.M.010',
    'CNV.AA.001.00.M.010',
    'AGL.AA.001.00.M.010',
    'AGL.AA.001.00.M.020',
    'CQR.AA.002.00.M.040',
    'CQR.AA.004.04.010',
    'CQR.AA.003.01.040',
    'CQR.AA.003.01.020',
    'GPON',
    'XGS-PON',
    'EPON',
  ];

  const equipmentModels = [
    'HG8245H',
    'EG8145V5',
    'ZXHN F670L',
    'ZXHN F680',
    'AN5506-04-F',
    'HG8145V5',
    'G-140W-ME',
    'ZXHN F660',
    'HG8245Q2',
    'AN5506-02-B',
    'EG8245W5-6T',
    'G-2425G-A'
  ];

  for (let i = 0; i < ananindeuaTotal; i++) {
    idCounter++;
    const olt = (i % 2 === 0) ? 'AIU-OLT-01' : 'AIU-OLT-02';
    const slot = `0/${(i % 16) + 1}/${(i % 8) + 1}`;
    const onuIndex = (i % 64) + 1;
    const contratoNum = 7000000 + idCounter;
    const isBreves = (i % 4 === 0);
    const logradouro = isBreves ? 'TV DE BREVES' : ananindeuaLogradouros[i % ananindeuaLogradouros.length];
    const numero = isBreves ? brevesNumbers[i % brevesNumbers.length] : String((i % 450) + 12);
    const logradouroNumero = `${logradouro}, ${numero}`;
    const topologia = topologyCodes[i % topologyCodes.length];
    const nmModelo = equipmentModels[i % equipmentModels.length];

    let status: 'OK' | 'IMPACTED' | 'OFFLINE' | 'STRESSED' = 'OK';
    let impacted: 'YES' | 'NO' = 'NO';
    let stressed: 'YES' | 'NO' = 'NO';
    const cronico: 'YES' | 'NO' = 'NO';
    let rx = parseFloat((-19.0 - (i % 5) * 0.9).toFixed(2));
    let temperatura = parseFloat((42.0 + (i % 8) * 1.2).toFixed(1));

    if (i < ananindeuaImpacted) {
      status = 'IMPACTED';
      impacted = 'YES';
      rx = parseFloat((-27.3 - (i % 10) * 0.2).toFixed(2));
      temperatura = (i % 5 === 0) ? 53.4 : 48.2;
    } else if (i < ananindeuaImpacted + ananindeuaStressed) {
      status = 'STRESSED';
      stressed = 'YES';
      rx = parseFloat((-25.5 - (i % 10) * 0.15).toFixed(2));
      temperatura = (i % 6 === 0) ? 52.8 : 46.5;
    } else {
      status = 'OK';
      if (i % 15 === 0) {
        temperatura = 52.5; // Alguns acima de 52 mesmo com status OK
      }
    }

    const txOlt = 2.85;
    const txOnu = 2.10;

    rows.push({
      id: `QOE-${idCounter}`,
      contrato: String(contratoNum),
      mes: currentMonth,
      cidade: 'ANANINDEUA',
      olt,
      topologia,
      nmModelo,
      status,
      temperatura,
      rxOnuCliente: rx,
      impacted,
      stressed,
      cronico,
      logradouro,
      logradouroNumero,
      bairro: 'Cidade Nova',
      numero,
      cep: '67130-000',
      slotPon: slot,
      onuId: `ONU-${onuIndex}`,
      serial: `HWTC${idCounter.toString(16).toUpperCase().padStart(8, '0')}`,
      txOlt,
      txOnu,
      cliente: `Cliente ${contratoNum}`,
      dataHora: `2026-09-${String((i % 28) + 1).padStart(2, '0')} 10:${String(i % 60).padStart(2, '0')}`,
      raw: {
        CONTRATO: String(contratoNum),
        MES: currentMonth,
        CIDADE: 'ANANINDEUA',
        OLT: olt,
        TOPOLOGIA: topologia,
        NM_MODELO: nmModelo,
        STATUS: status,
        LOGRADOURO: logradouro,
        LOGRADOURO_NUMERO: logradouroNumero,
        BAIRRO: 'Cidade Nova',
        TEMPERATURA: temperatura,
        RX_ONU_CLIENTE: rx,
        IMPACTED: impacted,
        STRESSED: stressed,
        CRONICO: cronico,
        SLOT_PON: slot,
        ONU_ID: `ONU-${onuIndex}`,
        SERIAL: `HWTC${idCounter.toString(16).toUpperCase().padStart(8, '0')}`,
        TX_OLT: txOlt,
        TX_ONU: txOnu,
      },
    });
  }

  // 2. Geração para as demais cidades da regional Norte
  const otherCities = [
    { code: 'BLM', city: 'BELÉM', olts: ['BLM-OLT-01', 'BLM-OLT-02', 'BLM-OLT-03'], count: 180, imp: 8, est: 9, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'MNS', city: 'MANAUS', olts: ['MNS-OLT-01', 'MNS-OLT-02'], count: 160, imp: 5, est: 7, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'SLS', city: 'SÃO LUÍS', olts: ['SLS-OLT-01', 'SLS-OLT-02'], count: 140, imp: 7, est: 8, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'CAH', city: 'CASTANHAL', olts: ['CAH-OLT-01'], count: 90, imp: 2, est: 3, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'MPA', city: 'MACAPÁ', olts: ['MPA-OLT-01'], count: 95, imp: 6, est: 5, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'ITZ', city: 'IMPERATRIZ', olts: ['ITZ-OLT-01'], count: 85, imp: 4, est: 4, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'MBA', city: 'MARABÁ', olts: ['MBA-OLT-01'], count: 75, imp: 5, est: 4, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'PGN', city: 'PARAGOMINAS', olts: ['PGN-OLT-01'], count: 60, imp: 2, est: 2, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'PUP', city: 'PARAUAPEBAS', olts: ['PUP-OLT-01'], count: 65, imp: 4, est: 5, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'CXS', city: 'CAXIAS', olts: ['CXS-OLT-01'], count: 50, imp: 3, est: 2, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'SQA', city: 'SANTANA', olts: ['SQA-OLT-01'], count: 45, imp: 2, est: 2, croOff: 0, croImp: 0, croEst: 0 },
    { code: 'TMN', city: 'TIMON', olts: ['TMN-OLT-01'], count: 40, imp: 1, est: 1, croOff: 0, croImp: 0, croEst: 0 },
  ];

  const commonLogradouros = [
    'Av. Principal',
    'Rua das Flores',
    'Av. Brasil',
    'Rua São José',
    'Av. Getúlio Vargas',
    'Rua 15 de Novembro',
    'Av. Santos Dumont',
    'Rua Amazonas',
    'Av. Central',
    'Rua Santa Luzia',
  ];

  for (const cfg of otherCities) {
    for (let i = 0; i < cfg.count; i++) {
      idCounter++;
      const olt = cfg.olts[i % cfg.olts.length];
      const slot = `0/${(i % 16) + 1}/${(i % 8) + 1}`;
      const onuIndex = (i % 64) + 1;
      const contratoNum = 7000000 + idCounter;
      const logradouro = commonLogradouros[i % commonLogradouros.length];
      const topologia = i % 8 === 0 ? 'XGS-PON' : i % 20 === 0 ? 'EPON' : 'GPON';

      let status: 'OK' | 'IMPACTED' | 'OFFLINE' | 'STRESSED' = 'OK';
      let impacted: 'YES' | 'NO' = 'NO';
      let stressed: 'YES' | 'NO' = 'NO';
      const cronico: 'YES' | 'NO' = 'NO';
      let rx = parseFloat((-19.5 - (i % 6) * 0.8).toFixed(2));
      let temperatura = parseFloat((41.0 + (i % 10) * 1.1).toFixed(1));

      if (i < cfg.imp) {
        status = 'IMPACTED';
        impacted = 'YES';
        rx = parseFloat((-27.2 - (i % 4) * 0.3).toFixed(2));
      } else if (i < cfg.imp + cfg.est) {
        status = 'STRESSED';
        stressed = 'YES';
        rx = parseFloat((-25.4 - (i % 5) * 0.2).toFixed(2));
      }

      if (temperatura > 51.5 && Math.random() < 0.2) {
        temperatura = parseFloat((52.1 + Math.random() * 4).toFixed(1));
      }

      const txOlt = 2.75;
      const txOnu = 2.05;

      const numero = String((i % 300) + 10);
      const logradouroNumero = `${logradouro}, ${numero}`;
      const nmModelo = equipmentModels[(i + cfg.count) % equipmentModels.length];

      rows.push({
        id: `QOE-${idCounter}`,
        contrato: String(contratoNum),
        mes: currentMonth,
        cidade: cfg.city,
        olt,
        topologia,
        nmModelo,
        status,
        temperatura,
        rxOnuCliente: rx,
        impacted,
        stressed,
        cronico,
        logradouro,
        logradouroNumero,
        bairro: 'Centro',
        numero,
        cep: '66000-000',
        slotPon: slot,
        onuId: `ONU-${onuIndex}`,
        serial: `HWTC${idCounter.toString(16).toUpperCase().padStart(8, '0')}`,
        txOlt,
        txOnu,
        cliente: `Cliente ${contratoNum}`,
        dataHora: `2026-09-${String((i % 28) + 1).padStart(2, '0')} 10:${String(i % 60).padStart(2, '0')}`,
        raw: {
          CONTRATO: String(contratoNum),
          MES: currentMonth,
          CIDADE: cfg.city,
          OLT: olt,
          TOPOLOGIA: topologia,
          NM_MODELO: nmModelo,
          STATUS: status,
          LOGRADOURO: logradouro,
          LOGRADOURO_NUMERO: logradouroNumero,
          BAIRRO: 'Centro',
          TEMPERATURA: temperatura,
          RX_ONU_CLIENTE: rx,
          IMPACTED: impacted,
          STRESSED: stressed,
          CRONICO: cronico,
          SLOT_PON: slot,
          ONU_ID: `ONU-${onuIndex}`,
          SERIAL: `HWTC${idCounter.toString(16).toUpperCase().padStart(8, '0')}`,
          TX_OLT: txOlt,
          TX_ONU: txOnu,
        },
      });
    }
  }

  return rows;
}

