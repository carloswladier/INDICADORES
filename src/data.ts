export type Technology = 'HFC' | 'GPON' | 'HÍBRIDO' | 'OUTROS';

export interface VisitData {
  id: string;
  contrato?: string;
  mes: string;
  fullDate: Date;
  cidade: string;
  area: string;
  expurgo: boolean;
  tecnologia: Technology;
  status: 'Executada' | 'Cancelada';
  notaAT1: number;
  node: string;
  cdBaixa: string;
  descriptionBaixa?: string;
  grupoBaixa: string;
  terminal: string;
  volume: number;
  tipoOs?: string;
}

export interface BaseCidadeData {
  cidade: string;
  tecnologia: Technology;
  base: number;
}

export const MOCK_DATA: VisitData[] = [
  { id: '1', contrato: '003/0129841', mes: 'Janeiro', fullDate: new Date(2024, 0, 15), cidade: 'São Paulo', area: 'Norte', expurgo: false, tecnologia: 'HFC', status: 'Executada', notaAT1: 8.5, node: 'SP01', cdBaixa: 'B01', descriptionBaixa: 'REPARO CABO / CONECTOR', grupoBaixa: 'TÉCNICO', terminal: 'TERM-01', volume: 1, tipoOs: 'REPARO' },
  { id: '2', contrato: '003/0145892', mes: 'Janeiro', fullDate: new Date(2024, 0, 16), cidade: 'Rio de Janeiro', area: 'Sul', expurgo: false, tecnologia: 'GPON', status: 'Executada', notaAT1: 9.2, node: 'RJ05', cdBaixa: 'B02', descriptionBaixa: 'INSTALAÇÃO FTTH CONCLUÍDA', grupoBaixa: 'CLIENTE', terminal: 'TERM-02', volume: 1, tipoOs: 'INSTALAÇÃO' },
  { id: '3', contrato: '003/0198234', mes: 'Fevereiro', fullDate: new Date(2024, 1, 10), cidade: 'São Paulo', area: 'Leste', expurgo: true, tecnologia: 'HÍBRIDO', status: 'Cancelada', notaAT1: 0, node: 'SP03', cdBaixa: 'B03', descriptionBaixa: 'CLIENTE AUSENTE', grupoBaixa: 'INFRA', terminal: 'TERM-03', volume: 1, tipoOs: 'REPARO' },
  { id: '4', contrato: '003/0201478', mes: 'Fevereiro', fullDate: new Date(2024, 1, 11), cidade: 'Belo Horizonte', area: 'Oeste', expurgo: false, tecnologia: 'GPON', status: 'Executada', notaAT1: 7.8, node: 'BH12', cdBaixa: 'B01', descriptionBaixa: 'MUDANÇA DE PONTO', grupoBaixa: 'TÉCNICO', terminal: 'TERM-04', volume: 1, tipoOs: 'MUDANÇA DE ENDEREÇO' },
  { id: '5', contrato: '003/0215981', mes: 'Março', fullDate: new Date(2024, 2, 5), cidade: 'São Paulo', area: 'Sul', expurgo: false, tecnologia: 'HFC', status: 'Executada', notaAT1: 8.9, node: 'SP01', cdBaixa: 'B05', descriptionBaixa: 'SUBSTITUIÇÃO DE MODEM', grupoBaixa: 'TÉCNICO', terminal: 'TERM-01', volume: 1, tipoOs: 'REPARO' },
  { id: '6', contrato: '003/0228490', mes: 'Março', fullDate: new Date(2024, 2, 6), cidade: 'Rio de Janeiro', area: 'Norte', expurgo: false, tecnologia: 'GPON', status: 'Executada', notaAT1: 9.5, node: 'RJ02', cdBaixa: 'B01', descriptionBaixa: 'INSTALAÇÃO GPON FIBRA', grupoBaixa: 'TÉCNICO', terminal: 'TERM-05', volume: 1, tipoOs: 'INSTALAÇÃO' },
  { id: '7', contrato: '003/0239102', mes: 'Janeiro', fullDate: new Date(2024, 0, 20), cidade: 'Belo Horizonte', area: 'Leste', expurgo: false, tecnologia: 'HÍBRIDO', status: 'Executada', notaAT1: 8.2, node: 'BH08', cdBaixa: 'B02', descriptionBaixa: 'TRANSFERÊNCIA DE TITULARIDADE', grupoBaixa: 'CLIENTE', terminal: 'TERM-06', volume: 1, tipoOs: 'TRANSFERÊNCIA' },
  { id: '8', contrato: '003/0248173', mes: 'Fevereiro', fullDate: new Date(2024, 1, 25), cidade: 'São Paulo', area: 'Sul', expurgo: false, tecnologia: 'HFC', status: 'Executada', notaAT1: 9.0, node: 'SP01', cdBaixa: 'B01', descriptionBaixa: 'NIVELAMENTO DE SINAL RF', grupoBaixa: 'TÉCNICO', terminal: 'TERM-01', volume: 1, tipoOs: 'REPARO' },
  { id: '9', contrato: '003/0259810', mes: 'Março', fullDate: new Date(2024, 2, 15), cidade: 'Rio de Janeiro', area: 'Oeste', expurgo: true, tecnologia: 'GPON', status: 'Cancelada', notaAT1: 0, node: 'RJ05', cdBaixa: 'B04', descriptionBaixa: 'CANCELAMENTO A PEDIDO DO CLIENTE', grupoBaixa: 'INFRA', terminal: 'TERM-02', volume: 1, tipoOs: 'DESCONEXÃO' },
  { id: '10', contrato: '003/0268491', mes: 'Janeiro', fullDate: new Date(2024, 0, 5), cidade: 'São Paulo', area: 'Norte', expurgo: false, tecnologia: 'GPON', status: 'Executada', notaAT1: 8.7, node: 'SP02', cdBaixa: 'B01', descriptionBaixa: 'ATIVAÇÃO ONU GPON', grupoBaixa: 'TÉCNICO', terminal: 'TERM-07', volume: 1, tipoOs: 'INSTALAÇÃO' },
  // Add more data to make it look real
  ...Array.from({ length: 40 }).map((_, i) => {
    const monthIdx = Math.floor(Math.random() * 3);
    const day = Math.floor(Math.random() * 28) + 1;
    const tipos = ['REPARO', 'INSTALAÇÃO', 'MUDANÇA DE ENDEREÇO', 'TRANSFERÊNCIA', 'DESCONEXÃO'];
    const descricoes = [
      'REPARO FIBRA ÓPTICA',
      'TROCA DE DROP ÓPTICO',
      'SUBSTITUIÇÃO DE FONTE',
      'REPARO CONECTOR ÓPTICO',
      'CONFIGURAÇÃO DE WI-FI',
      'INSTALAÇÃO COMPLETA'
    ];
    return {
      id: `extra-${i}`,
      contrato: `003/0${270000 + i * 137}`,
      mes: ['Janeiro', 'Fevereiro', 'Março'][monthIdx],
      fullDate: new Date(2024, monthIdx, day),
      cidade: ['São Paulo', 'Rio de Janeiro', 'Belo Horizonte', 'Curitiba'][Math.floor(Math.random() * 4)],
      area: ['Norte', 'Sul', 'Leste', 'Oeste'][Math.floor(Math.random() * 4)],
      expurgo: Math.random() > 0.8,
      tecnologia: (['HFC', 'GPON', 'HÍBRIDO'] as Technology[])[Math.floor(Math.random() * 3)],
      status: Math.random() > 0.2 ? 'Executada' : 'Cancelada' as 'Executada' | 'Cancelada',
      notaAT1: Math.random() > 0.2 ? Number((Math.random() * 3 + 7).toFixed(1)) : 0,
      node: `NODE-${Math.floor(Math.random() * 20)}`,
      cdBaixa: `B-${Math.floor(Math.random() * 30)}`,
      descriptionBaixa: descricoes[i % descricoes.length],
      grupoBaixa: ['TÉCNICO', 'CLIENTE', 'INFRA', 'SISTEMA'][Math.floor(Math.random() * 4)],
      terminal: `TERM-${Math.floor(Math.random() * 50)}`,
      volume: 1,
      tipoOs: tipos[i % tipos.length]
    };
  })
];

export const MOCK_BASE_CIDADE: BaseCidadeData[] = [
  { cidade: 'São Paulo', tecnologia: 'HFC', base: 5000 },
  { cidade: 'São Paulo', tecnologia: 'GPON', base: 3000 },
  { cidade: 'São Paulo', tecnologia: 'HÍBRIDO', base: 1000 },
  { cidade: 'Rio de Janeiro', tecnologia: 'HFC', base: 4000 },
  { cidade: 'Rio de Janeiro', tecnologia: 'GPON', base: 2500 },
  { cidade: 'Belo Horizonte', tecnologia: 'HFC', base: 3000 },
  { cidade: 'Curitiba', tecnologia: 'HFC', base: 2000 },
];
