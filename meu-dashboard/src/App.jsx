import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, AlertCircle } from 'lucide-react';
import { 
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, 
  PieChart, Pie, Cell, Legend 
} from 'recharts';

const COLORS = ['#06b6d4', '#f59e0b', '#ef4444', '#10b981', '#6366f1', '#ec4899', '#8b5cf6'];

export default function App() {
  const [data, setData] = useState([]);
  const [fileName, setFileName] = useState("base padrão (planilha enviada)");

  // Filtros
  const [selectedUnidade, setSelectedUnidade] = useState("Todos");
  const [selectedStatus, setSelectedStatus] = useState("Todos");
  const [selectedPagamento, setSelectedPagamento] = useState("Todos");

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const workbook = XLSX.read(bstr, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet);
      
      setData(jsonData);
    };

    reader.readAsBinaryString(file);
  };

  // Opções dos Filtros
  const unidades = useMemo(() => ["Todos", ...new Set(data.map(d => (d["UNIDADE"] || "").trim()).filter(Boolean))], [data]);
  const statuses = useMemo(() => ["Todos", ...new Set(data.map(d => (d["Status"] || d["STATUS"] || "").trim()).filter(Boolean))], [data]);
  const pagamentos = useMemo(() => ["Todos", ...new Set(data.map(d => (d["PAGAMENTO"] || "").trim()).filter(Boolean))], [data]);

  // Filtragem dos Dados
  const filteredData = useMemo(() => {
    return data.filter(item => {
      const u = (item["UNIDADE"] || "").trim();
      const s = (item["Status"] || item["STATUS"] || "").trim();
      const p = (item["PAGAMENTO"] || "").trim();

      const matchU = selectedUnidade === "Todos" || u === selectedUnidade;
      const matchS = selectedStatus === "Todos" || s === selectedStatus;
      const matchP = selectedPagamento === "Todos" || p === selectedPagamento;

      return matchU && matchS && matchP;
    });
  }, [data, selectedUnidade, selectedStatus, selectedPagamento]);

  // Totalizadores Globais da Base
  const totalGeralNotas = data.length || 1;
  const totalGeralMontante = useMemo(() => {
    return data.reduce((acc, row) => acc + (parseFloat(row["MONTANTE"]) || 0), 0) || 1;
  }, [data]);

  // KPIs Principais
  const totalRegistros = filteredData.length;
  const pctNotasFiltradas = ((totalRegistros / totalGeralNotas) * 100).toFixed(1);

  const montanteBloqueado = useMemo(() => {
    return filteredData.reduce((acc, row) => acc + (parseFloat(row["MONTANTE"]) || 0), 0);
  }, [filteredData]);
  const pctMontanteFiltrado = ((montanteBloqueado / totalGeralMontante) * 100).toFixed(1);

  const totalDiferenca = useMemo(() => {
    return filteredData.reduce((acc, row) => acc + (parseFloat(row["TOTAL DA DIFERENÇA"]) || 0), 0);
  }, [filteredData]);

  const pctMediaDivergencia = useMemo(() => {
    if (!filteredData.length) return 0;
    const sum = filteredData.reduce((acc, row) => acc + (parseFloat(row["% DIF"]) || 0), 0);
    return (sum / filteredData.length) * 100;
  }, [filteredData]);

  const agingMedio = useMemo(() => {
    if (!filteredData.length) return 0;
    const sum = filteredData.reduce((acc, row) => acc + (parseFloat(row["DIAS "]) || parseFloat(row["DIAS"]) || 0), 0);
    return sum / filteredData.length;
  }, [filteredData]);

  // 1. TOP FORNECEDORES
  const topFornecedoresData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      let forn = "Outros";
      Object.keys(row).forEach(k => {
        if (k.includes("FORNECEDOR")) forn = row[k];
      });
      forn = String(forn).trim();
      const montante = parseFloat(row["MONTANTE"]) || 0;
      map[forn] = (map[forn] || 0) + montante;
    });
    return Object.keys(map)
      .map(key => ({ 
        name: key.length > 22 ? key.substring(0, 22) + "..." : key, 
        montante: map[key],
        pct: montanteBloqueado ? ((map[key] / montanteBloqueado) * 100).toFixed(1) : 0
      }))
      .sort((a, b) => b.montante - a.montante)
      .slice(0, 10);
  }, [filteredData, montanteBloqueado]);

  // 2. NEGOCIADORES
  const negociadorTableData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      const neg = (row["NEGOCIADOR"] || "Não Atribuído").trim();
      const montante = parseFloat(row["MONTANTE"]) || 0;
      const dif = parseFloat(row["TOTAL DA DIFERENÇA"]) || 0;
      if (!map[neg]) map[neg] = { notas: 0, montante: 0, dif: 0 };
      map[neg].notas += 1;
      map[neg].montante += montante;
      map[neg].dif += dif;
    });
    return Object.keys(map)
      .map(key => ({ 
        name: key, 
        ...map[key],
        pctNotas: totalRegistros ? ((map[key].notas / totalRegistros) * 100).toFixed(1) : 0,
        pctMontante: montanteBloqueado ? ((map[key].montante / montanteBloqueado) * 100).toFixed(1) : 0
      }))
      .sort((a, b) => b.montante - a.montante);
  }, [filteredData, totalRegistros, montanteBloqueado]);

  // 3. RESPONSÁVEIS
  const responsavelTableData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      const resp = (row["RESPONSÁVEL"] || "Não Atribuído").trim();
      const montante = parseFloat(row["MONTANTE"]) || 0;
      const dif = parseFloat(row["TOTAL DA DIFERENÇA"]) || 0;
      if (!map[resp]) map[resp] = { notas: 0, montante: 0, dif: 0 };
      map[resp].notas += 1;
      map[resp].montante += montante;
      map[resp].dif += dif;
    });
    return Object.keys(map)
      .map(key => ({ 
        name: key, 
        ...map[key],
        pctNotas: totalRegistros ? ((map[key].notas / totalRegistros) * 100).toFixed(1) : 0,
        pctMontante: montanteBloqueado ? ((map[key].montante / montanteBloqueado) * 100).toFixed(1) : 0
      }))
      .sort((a, b) => b.montante - a.montante);
  }, [filteredData, totalRegistros, montanteBloqueado]);

  // 4. TABELA TIPO E MOTIVO
  const divergenciasTipoMotivoData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      let tipo = "Não Informado";
      let motivo = "Não Informado";

      Object.keys(row).forEach(k => {
        if (k.toUpperCase().includes("TIPO")) tipo = row[k];
        if (k.toUpperCase().includes("MOTIVO")) motivo = row[k];
      });

      const key = `${String(tipo).trim()} - ${String(motivo).trim()}`;
      const montante = parseFloat(row["MONTANTE"]) || 0;

      if (!map[key]) map[key] = { tipo: String(tipo).trim(), motivo: String(motivo).trim(), notas: 0, montante: 0 };
      map[key].notas += 1;
      map[key].montante += montante;
    });

    return Object.values(map)
      .map(item => ({
        ...item,
        pctNotas: totalRegistros ? ((item.notas / totalRegistros) * 100).toFixed(1) : 0,
        pctMontante: montanteBloqueado ? ((item.montante / montanteBloqueado) * 100).toFixed(1) : 0
      }))
      .sort((a, b) => b.montante - a.montante);
  }, [filteredData, totalRegistros, montanteBloqueado]);

  // 5. PLANTAS OPERACIONAIS
  const plantaChartData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      const un = (row["UNIDADE"] || "Outros").trim().replace('Indústria - ', '');
      const montante = parseFloat(row["MONTANTE"]) || 0;
      if (!map[un]) map[un] = { notas: 0, montante: 0 };
      map[un].notas += 1;
      map[un].montante += montante;
    });
    return Object.keys(map).map(key => ({
      unidade: key,
      Notas: map[key].notas,
      pctNotas: totalRegistros ? ((map[key].notas / totalRegistros) * 100).toFixed(1) : 0,
      Montante: Math.round(map[key].montante),
      pctMontante: montanteBloqueado ? ((map[key].montante / montanteBloqueado) * 100).toFixed(1) : 0
    })).sort((a, b) => b.Montante - a.Montante);
  }, [filteredData, totalRegistros, montanteBloqueado]);

  const formatBRL = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(val);

  return (
    <div className="min-h-screen p-6 bg-[#0b0f19] text-white font-sans">
      {/* HEADER */}
      <header className="flex flex-wrap justify-between items-center mb-6 pb-4 border-b border-gray-800 gap-4">
        <div>
          <span className="text-[10px] text-cyan-400 font-semibold tracking-wider uppercase">PROCUREMENT · NOTAS BLOQUEADAS</span>
          <h1 className="text-xl font-bold mt-0.5">Dashboard de Divergências</h1>
          <p className="text-xs text-gray-400 flex items-center gap-1.5 mt-0.5">
            <FileSpreadsheet size={14} />
            {fileName}
          </p>
        </div>

        <label className="flex items-center gap-2 bg-cyan-500 hover:bg-cyan-600 text-slate-950 text-xs font-semibold px-3 py-1.5 rounded cursor-pointer transition">
          <Upload size={14} />
          Atualizar base (.xlsx)
          <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleFileUpload} />
        </label>
      </header>

      {/* FILTROS */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-6 bg-[#131927] p-3 rounded-lg border border-gray-800 text-xs">
        <div>
          <label className="block text-gray-400 mb-1 font-medium">Unidade:</label>
          <select value={selectedUnidade} onChange={e => setSelectedUnidade(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded px-2 py-1 text-white">
            {unidades.map(u => <option key={u} value={u}>{u}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-gray-400 mb-1 font-medium">Status:</label>
          <select value={selectedStatus} onChange={e => setSelectedStatus(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded px-2 py-1 text-white">
            {statuses.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-gray-400 mb-1 font-medium">Pagamento:</label>
          <select value={selectedPagamento} onChange={e => setSelectedPagamento(e.target.value)} className="w-full bg-[#1f2937] border border-gray-700 rounded px-2 py-1 text-white">
            {pagamentos.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </div>
      </div>

      {/* KPIS NOVO (COM TOTAL DE NOTAS E PORCENTAGENS) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 mb-6">
        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">TOTAL DE NOTAS</span>
          <p className="text-2xl font-bold text-white mt-1">{totalRegistros}</p>
          <span className="text-[10px] text-cyan-400 font-medium">{pctNotasFiltradas}% do total</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">MONTANTE BLOQUEADO</span>
          <p className="text-2xl font-bold text-emerald-400 mt-1">{formatBRL(montanteBloqueado)}</p>
          <span className="text-[10px] text-emerald-400/80 font-medium">{pctMontanteFiltrado}% do total</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">TOTAL DA DIFERENÇA</span>
          <p className="text-2xl font-bold text-cyan-400 mt-1">{formatBRL(totalDiferenca)}</p>
          <span className="text-[10px] text-gray-500">
            {montanteBloqueado ? ((totalDiferenca / montanteBloqueado) * 100).toFixed(1) : 0}% do montante
          </span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">% MÉDIO DIVERGÊNCIA</span>
          <p className="text-2xl font-bold text-purple-400 mt-1">{pctMediaDivergencia.toFixed(1)}%</p>
          <span className="text-[10px] text-gray-500">Média ponderada % DIF</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">AGING MÉDIO BLOQUEIO</span>
          <p className="text-2xl font-bold text-amber-400 mt-1">{agingMedio.toFixed(1)} dias</p>
          <span className="text-[10px] text-gray-500">Dias decorridos em média</span>
        </div>
      </div>

      {data.length > 0 ? (
        <>
          {/* 1. RANKING FORNECEDORES PRIMEIRO NO LAYOUT */}
          <div className="bg-[#131927] p-4 rounded-lg border border-gray-800 mb-6">
            <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">RANKING DE FORNECEDORES OFENSORES (TOP 10 MONTANTE E %)</h2>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={topFornecedoresData} layout="vertical">
                  <XAxis type="number" stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                  <YAxis type="category" dataKey="name" stroke="#6b7280" width={170} tick={{ fontSize: 10, fill: '#9ca3af' }} />
                  <Tooltip 
                    formatter={(value, name, props) => [`${formatBRL(value)} (${props.payload.pct}%)`, 'Montante']} 
                    contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} 
                  />
                  <Bar dataKey="montante" fill="#06b6d4" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* 2. NEGOCIADORES E RESPONSÁVEIS */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6 text-xs">
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">DESEMPENHO POR NEGOCIADOR / COMPRADOR</h2>
              <div className="overflow-x-auto max-h-64">
                <table className="w-full text-left border-collapse">
                  <thead className="sticky top-0 bg-[#131927]">
                    <tr className="border-b border-gray-800 text-gray-400 text-[10px] uppercase">
                      <th className="pb-2">Nome</th>
                      <th className="pb-2 text-center">Notas (% Tot)</th>
                      <th className="pb-2 text-right">Montante (% Tot)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800/50">
                    {negociadorTableData.map((row, i) => (
                      <tr key={i} className="hover:bg-[#1a2234]">
                        <td className="py-2 text-gray-200 font-medium">{row.name}</td>
                        <td className="py-2 text-center text-gray-300">
                          {row.notas} <span className="text-[10px] text-cyan-400">({row.pctNotas}%)</span>
                        </td>
                        <td className="py-2 text-right text-emerald-400 font-medium">
                          {formatBRL(row.montante)} <span className="text-[10px] text-emerald-300">({row.pctMontante}%)</span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">DESEMPENHO POR RESPONSÁVEL</h2>
              <div className="overflow-x-auto max-h-64">
                <table className="w-full text-left border-collapse">
                  <thead className="sticky top-0 bg-[#131927]">
                    <tr className="border-b border-gray-800 text-gray-400 text-[10px] uppercase">
                      <th className="pb-2">Nome</th>
                      <th className="pb-2 text-center">Notas (% Tot)</th>
                      <th className="pb-2 text-right">Montante (% Tot)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800/50">
                    {responsavelTableData.map((row, i) => (
                      <tr key={i} className="hover:bg-[#1a2234]">
                        <td className="py-2 text-gray-200 font-medium">{row.name}</td>
                        <td className="py-2 text-center text-gray-300">
                          {row.notas} <span className="text-[10px] text-cyan-400">({row.pctNotas}%)</span>
                        </td>
                        <td className="py-2 text-right text-emerald-400 font-medium">
                          {formatBRL(row.montante)} <span className="text-[10px] text-emerald-300">({row.pctMontante}%)</span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          {/* 3. TABELA DIVERGÊNCIAS POR TIPO E MOTIVO */}
          <div className="bg-[#131927] p-4 rounded-lg border border-gray-800 mb-6 text-xs">
            <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">TABELA DE DIVERGÊNCIAS POR TIPO E MOTIVO</h2>
            <div className="overflow-x-auto max-h-64">
              <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-[#131927]">
                  <tr className="border-b border-gray-800 text-gray-400 text-[10px] uppercase">
                    <th className="pb-2">Tipo</th>
                    <th className="pb-2">Motivo</th>
                    <th className="pb-2 text-center">Qtd. Notas</th>
                    <th className="pb-2 text-center">% Qtd.</th>
                    <th className="pb-2 text-right">Montante (R$)</th>
                    <th className="pb-2 text-right">% Montante</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-800/50">
                  {divergenciasTipoMotivoData.map((row, i) => (
                    <tr key={i} className="hover:bg-[#1a2234]">
                      <td className="py-2 text-cyan-400 font-semibold">{row.tipo}</td>
                      <td className="py-2 text-gray-300">{row.motivo}</td>
                      <td className="py-2 text-center text-gray-200 font-medium">{row.notas}</td>
                      <td className="py-2 text-center text-cyan-400 font-medium">{row.pctNotas}%</td>
                      <td className="py-2 text-right text-emerald-400 font-medium">{formatBRL(row.montante)}</td>
                      <td className="py-2 text-right text-emerald-300 font-medium">{row.pctMontante}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* 4. PLANTA OPERACIONAL DIVIDIDA EM DOIS GRÁFICOS */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-cyan-400 uppercase tracking-wider">INCONSISTÊNCIAS POR PLANTA — QTD. NOTAS (% DO TOTAL)</h2>
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={plantaChartData}>
                    <XAxis dataKey="unidade" stroke="#6b7280" tick={{ fontSize: 9, fill: '#9ca3af' }} interval={0} angle={-25} textAnchor="end" />
                    <YAxis stroke="#06b6d4" tick={{ fontSize: 10, fill: '#06b6d4' }} />
                    <Tooltip 
                      formatter={(value, name, props) => [`${value} notas (${props.payload.pctNotas}%)`, 'Notas']}
                      contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} 
                    />
                    <Bar dataKey="Notas" fill="#06b6d4" radius={[2, 2, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-amber-400 uppercase tracking-wider">INCONSISTÊNCIAS POR PLANTA — MONTANTE R$ (% DO TOTAL)</h2>
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={plantaChartData}>
                    <XAxis dataKey="unidade" stroke="#6b7280" tick={{ fontSize: 9, fill: '#9ca3af' }} interval={0} angle={-25} textAnchor="end" />
                    <YAxis stroke="#f59e0b" tick={{ fontSize: 10, fill: '#f59e0b' }} />
                    <Tooltip 
                      formatter={(value, name, props) => [`${formatBRL(value)} (${props.payload.pctMontante}%)`, 'Montante']}
                      contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} 
                    />
                    <Bar dataKey="Montante" fill="#f59e0b" radius={[2, 2, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        </>
      ) : (
        <div className="flex flex-col items-center justify-center py-24 bg-[#131927] rounded-lg border border-gray-800 text-gray-500 gap-3">
          <AlertCircle size={40} className="text-gray-600" />
          <p className="text-xs">Carregue a planilha no botão superior para carregar os dados.</p>
        </div>
      )}
    </div>
  );
}