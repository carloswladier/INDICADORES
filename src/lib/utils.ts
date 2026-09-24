import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function formatPercent(val: number | string | null | undefined, decimals = 2): string {
  if (val === null || val === undefined || val === '') return `0,${'0'.repeat(decimals)}%`;
  const num = Number(val);
  if (isNaN(num)) return `0,${'0'.repeat(decimals)}%`;
  return `${num.toFixed(decimals).replace('.', ',')}%`;
}

export function formatInteger(val: number | string | null | undefined): string {
  if (val === null || val === undefined || val === '') return '0';
  const num = Number(val);
  if (isNaN(num)) return '0';
  return Math.round(num).toLocaleString('pt-BR');
}

export function formatDecimal(val: number | string | null | undefined, decimals = 2): string {
  if (val === null || val === undefined || val === '') return `0,${'0'.repeat(decimals)}`;
  const num = Number(val);
  if (isNaN(num)) return `0,${'0'.repeat(decimals)}`;
  return num.toFixed(decimals).replace('.', ',');
}

