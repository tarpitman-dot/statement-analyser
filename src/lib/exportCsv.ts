import type { Transaction } from './types';
export function toCsv(rows:Record<string,unknown>[], headers?:string[]){const cols=headers??Object.keys(rows[0]??{}); const esc=(v:unknown)=>`"${String(v??'').replace(/"/g,'""')}"`; return '\uFEFF'+[cols.join(','),...rows.map(r=>cols.map(c=>esc(r[c])).join(','))].join('\n');}
export function downloadCsv(name:string,csv:string){const blob=new Blob([csv],{type:'text/csv;charset=utf-8'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; a.click(); URL.revokeObjectURL(a.href);}
export function normalisedRows(rows:Transaction[]){return rows.map(({originalRow:_,...r})=>r as unknown as Record<string,unknown>);}
