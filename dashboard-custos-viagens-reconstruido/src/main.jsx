import React, {useMemo, useState} from 'react'
import {createRoot} from 'react-dom/client'
import * as XLSX from 'xlsx'
import {BarChart,Bar,XAxis,YAxis,Tooltip,ResponsiveContainer,CartesianGrid} from 'recharts'
import {Upload,Search,RefreshCw,FileSpreadsheet,Download} from 'lucide-react'
import './styles.css'

const demo=[
 {OS:'8979',Data:'04/09/2026',Categoria:'Passagem aérea',Origem:'Belém',Destino:'Macapá',Custo:1280},
 {OS:'8988',Data:'04/09/2026',Categoria:'Hospedagem',Origem:'Santarém',Destino:'Macapá',Custo:2450},
 {OS:'9001',Data:'08/09/2026',Categoria:'Locação',Origem:'São Paulo',Destino:'Rio de Janeiro',Custo:1850}
]
function money(v){return new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(Number(v)||0)}
function normalize(rows){
 return rows.map((r,i)=>{
   const keys=Object.keys(r); const find=(names)=>keys.find(k=>names.some(n=>k.toLowerCase().normalize('NFD').replace(/[^a-z]/g,'').includes(n)))
   const costKey=find(['custo','valor','total','preco']);
   return {OS:r.OS||r.os||r['Nº OS']||r['OS Nº']||`LINHA-${i+1}`,Data:r.Data||r.data||r['Data viagem']||'',Categoria:r.Categoria||r.categoria||'Outros',Origem:r.Origem||r.origem||'',Destino:r.Destino||r.destino||'',Custo: Number(String(costKey?r[costKey]:0).replace(/\./g,'').replace(',','.').replace(/[^0-9.-]/g,''))||0}
 })
}
function App(){
 const [rows,setRows]=useState(demo),[q,setQ]=useState(''),[cat,setCat]=useState('Todas'),[file,setFile]=useState('')
 const cats=['Todas',...new Set(rows.map(r=>r.Categoria))]
 const filtered=useMemo(()=>rows.filter(r=>(cat==='Todas'||r.Categoria===cat)&&Object.values(r).join(' ').toLowerCase().includes(q.toLowerCase())),[rows,q,cat])
 const total=filtered.reduce((s,r)=>s+r.Custo,0), avg=filtered.length?total/filtered.length:0
 const chart=Object.entries(filtered.reduce((a,r)=>(a[r.Categoria]=(a[r.Categoria]||0)+r.Custo,a),{})).map(([name,total])=>({name,total}))
 function importFile(e){
   const f=e.target.files?.[0]; if(!f)return; setFile(f.name)
   const reader=new FileReader(); reader.onload=x=>{const wb=XLSX.read(x.target.result,{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]]; setRows(normalize(XLSX.utils.sheet_to_json(ws,{defval:''})))}; reader.readAsArrayBuffer(f)
 }
 function exportCsv(){
   const csv=[['OS','Data','Categoria','Origem','Destino','Custo'],...filtered.map(r=>[r.OS,r.Data,r.Categoria,r.Origem,r.Destino,r.Custo])].map(r=>r.join(';')).join('\n')
   const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download='analise-custos-viagens.csv';a.click()
 }
 return <div className="app">
  <header><div><span className="eyebrow">GESTÃO DE VIAGENS</span><h1>Central de Custos</h1><p>Consulta, consolidação e análise de despesas por OS.</p></div>
  <label className="upload"><Upload size={18}/> Importar Excel<input type="file" accept=".xlsx,.xls,.xlsb,.csv" onChange={importFile}/></label></header>
  <section className="toolbar"><div className="search"><Search size={18}/><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Pesquisar OS, cidade, categoria..."/></div><select value={cat} onChange={e=>setCat(e.target.value)}>{cats.map(c=><option key={c}>{c}</option>)}</select><button onClick={()=>{setQ('');setCat('Todas')}}><RefreshCw size={16}/> Limpar</button><button onClick={exportCsv}><Download size={16}/> Exportar</button></section>
  {file&&<div className="file"><FileSpreadsheet size={17}/> {file} <span>{rows.length} registros carregados</span></div>}
  <section className="cards"><div><small>CUSTO TOTAL</small><strong>{money(total)}</strong></div><div><small>REGISTROS</small><strong>{filtered.length}</strong></div><div><small>CUSTO MÉDIO</small><strong>{money(avg)}</strong></div><div><small>CATEGORIAS</small><strong>{Math.max(0,cats.length-1)}</strong></div></section>
  <section className="grid"><div className="panel chart"><div className="panel-title"><h2>Custos por categoria</h2></div><ResponsiveContainer width="100%" height={280}><BarChart data={chart}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="name"/><YAxis/><Tooltip formatter={v=>money(v)}/><Bar dataKey="total"/></BarChart></ResponsiveContainer></div>
  <div className="panel"><div className="panel-title"><h2>Resumo</h2></div><p className="big">{money(total)}</p><p className="muted">Valor consolidado após aplicação dos filtros.</p><div className="rules">• Importação de XLSX/XLS/XLSB<br/>• Busca por múltiplos campos<br/>• Filtro por categoria<br/>• Exportação dos resultados</div></div></section>
  <section className="panel"><div className="panel-title"><h2>Detalhamento</h2><span>{filtered.length} linhas</span></div><div className="table-wrap"><table><thead><tr>{['OS','Data','Categoria','Origem','Destino','Custo'].map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{filtered.map((r,i)=><tr key={i}><td>{r.OS}</td><td>{r.Data}</td><td>{r.Categoria}</td><td>{r.Origem}</td><td>{r.Destino}</td><td>{money(r.Custo)}</td></tr>)}</tbody></table></div></section>
  <footer>Dashboard Custos de Viagens • versão reconstruída para evolução e versionamento no GitHub</footer>
 </div>
}
createRoot(document.getElementById('root')).render(<App/>)
