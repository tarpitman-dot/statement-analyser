import Decimal from 'decimal.js';
export const D = (v:unknown)=>{ if(v===null||v===undefined||v==='') return new Decimal(0); const s=String(v).trim().replace(/[£$€,]/g,'').replace(/^\((.*)\)$/,'-$1'); const n=new Decimal(s||0); return n.isFinite()?n:new Decimal(0); };
export const fmtGbp=(v:Decimal.Value)=>{ const d=new Decimal(v).toDecimalPlaces(2); const abs=d.abs().toNumber().toLocaleString('en-GB',{minimumFractionDigits:2,maximumFractionDigits:2}); return `${d.isNegative()?'-':''}£${abs}`; };
export const fmtMoney=(v:Decimal.Value)=>fmtGbp(v);
export const fmtDec=(v:Decimal.Value,places=8)=>new Decimal(v).toDecimalPlaces(places).toString();
export const fmtInt=(v:Decimal.Value)=>new Decimal(v).toDecimalPlaces(0).toNumber().toLocaleString();
export const fmtRate=(v:string)=>v?`${D(v).mul(100).toDecimalPlaces(4).toString().replace(/\.0+$/,'')}%`:'';
export const text = (v:unknown)=> v===null||v===undefined?'':String(v).trim();
