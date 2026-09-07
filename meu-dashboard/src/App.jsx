import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, AlertCircle } from 'lucide-react';
import { 
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, 
  PieChart, Pie, Cell, Legend 
} from 'recharts';

const COLORS = ['#06b6d4', '#f59e0b', '#ef4444', '#10b981', '#6366f1', '#ec4899'];

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
  const statuses = useMemo(() => ["Todos", ...new Set(data.map(d => (d["Status"] || "").trim()).filter(Boolean))], [data]);
  const pagamentos = useMemo(() => ["Todos", ...new Set(data.map(d => (d["PAGAMENTO"] || "").trim()).filter(Boolean))], [data]);

  // Filtragem dos Dados
  const filteredData = useMemo(() => {
    return data.filter(item => {
      const u = (item["UNIDADE"] || "").trim();
      const s = (item["Status"] || "").trim();
      const p = (item["PAGAMENTO"] || "").trim();

      const matchU = selectedUnidade === "Todos" || u === selectedUnidade;
      const matchS = selectedStatus === "Todos" || s === selectedStatus;
      const matchP = selectedPagamento === "Todos" || p === selectedPagamento;

      return matchU && matchS && matchP;
    });
  }, [data, selectedUnidade, selectedStatus, selectedPagamento]);

  // Calculations / KPIs
  const totalRegistros = filteredData.length;
  
  const montanteBloqueado = useMemo(() => {
    return filteredData.reduce((acc, row) => acc + (parseFloat(row["MONTANTE"]) || 0), 0);
  }, [filteredData]);

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
    const sum = filteredData.reduce((acc, row) => acc + (parseFloat(row["DIAS "]) || 0), 0);
    return sum / filteredData.length;
  }, [filteredData]);

  // Gráfico 1: Aging do Bloqueio (por Faixa de Dias)
  const agingChartData = useMemo(() => {
    const buckets = { '-30...0': { notas: 0, montante: 0 }, '1-15': { notas: 0, montante: 0 }, '16-30': { notas: 0, montante: 0 }, '31-60': { notas: 0, montante: 0 }, '61-90': { notas: 0, montante: 0 }, 'Acima 90': { notas: 0, montante: 0 } };
    filteredData.forEach(row => {
      const aging = (row["AGING - DIA BLOQUEIO"] || "").trim();
      const montante = parseFloat(row["MONTANTE"]) || 0;
      if (buckets[aging]) {
        buckets[aging].notas += 1;
        buckets[aging].montante += montante;
      }
    });
    return Object.keys(buckets).map(key => ({
      faixa: key,
      Notas: buckets[key].notas,
      Montante: Math.round(buckets[key].montante)
    }));
  }, [filteredData]);

  // Gráfico 2: Vencimento Faturas Bloqueadas (Donut)
  const vencimentoPieData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      const pag = (row["PAGAMENTO"] || "Outros").trim();
      map[pag] = (map[pag] || 0) + 1;
    });
    return Object.keys(map).map(key => ({ name: key, value: map[key] }));
  }, [filteredData]);

  // Gráfico 3: % Médio de Divergência por Faturamento
  const faturamentoData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      const fat = (row["FATURAMENTO"] || "Outros").trim();
      const pct = (parseFloat(row["% DIF"]) || 0) * 100;
      if (!map[fat]) map[fat] = { count: 0, totalPct: 0 };
      map[fat].count += 1;
      map[fat].totalPct += pct;
    });
    return Object.keys(map).map(key => ({
      faturamento: key,
      pctMedia: map[key].count ? (map[key].totalPct / map[key].count).toFixed(1) : 0,
      count: map[key].count
    }));
  }, [filteredData]);

  // Gráfico 4: Ranking Fornecedores Ofensores (Top 10 Horizontais)
  const topFornecedoresData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      const forn = (row["NOME DO FORNECEDOR                 "] || "Outros").trim();
      const montante = parseFloat(row["MONTANTE"]) || 0;
      map[forn] = (map[forn] || 0) + montante;
    });
    return Object.keys(map)
      .map(key => ({ name: key.length > 20 ? key.substring(0, 20) + "..." : key, montante: map[key] }))
      .sort((a, b) => b.montante - a.montante)
      .slice(0, 10);
  }, [filteredData]);

  // Tabelas: Negociadores e Responsáveis
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
      .map(key => ({ name: key, ...map[key] }))
      .sort((a, b) => b.montante - a.montante)
      .slice(0, 6);
  }, [filteredData]);

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
      .map(key => ({ name: key, ...map[key] }))
      .sort((a, b) => b.montante - a.montante)
      .slice(0, 6);
  }, [filteredData]);

  // Gráfico 5: Inconsistências por Planta Operacional
  const plantaChartData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      const un = (row["UNIDADE"] || "Outros").trim().replace('Indústria - ', '');
      const notas = 1;
      const montante = parseFloat(row["MONTANTE"]) || 0;
      if (!map[un]) map[un] = { notas: 0, montante: 0 };
      map[un].notas += notas;
      map[un].montante += montante;
    });
    return Object.keys(map).map(key => ({
      unidade: key,
      Notas: map[key].notas,
      Montante: Math.round(map[key].montante)
    }));
  }, [filteredData]);

  const formatBRL = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(val);

  return (
    <div className="min-h-screen p-6 bg-[#0b0f19] text-white font-sans">
      {/* Topo */}
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

      {/* Filtros */}
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

      {/* KPIs Topo */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">MONTANTE BLOQUEADO</span>
          <p className="text-2xl font-bold text-emerald-400 mt-1">{formatBRL(montanteBloqueado)}</p>
          <span className="text-[10px] text-gray-500">{totalRegistros} notas bloqueadas</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">TOTAL DA DIFERENÇA</span>
          <p className="text-2xl font-bold text-cyan-400 mt-1">{formatBRL(totalDiferenca)}</p>
          <span className="text-[10px] text-gray-500">Somatório das divergências</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">% MÉDIO DE DIVERGÊNCIA</span>
          <p className="text-2xl font-bold text-white mt-1">{pctMediaDivergencia.toFixed(1)}%</p>
          <span className="text-[10px] text-gray-500">Média ponderada do % DIF</span>
        </div>

        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">AGING MÉDIO DE BLOQUEIO</span>
          <p className="text-2xl font-bold text-amber-400 mt-1">{agingMedio.toFixed(1)} dias</p>
          <span className="text-[10px] text-gray-500">Dias decorridos em média</span>
        </div>
      </div>

      {data.length > 0 ? (
        <>
          {/* Linha 1: Aging e Vencimento */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-6">
            <div className="lg:col-span-2 bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">AGING DO BLOQUEIO</h2>
              <div className="h-56">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={agingChartData}>
                    <XAxis dataKey="faixa" stroke="#6b7280" tick={{ fontSize: 11, fill: '#9ca3af' }} />
                    <YAxis yAxisId="left" stroke="#06b6d4" tick={{ fontSize: 11, fill: '#06b6d4' }} />
                    <YAxis yAxisId="right" orientation="right" stroke="#f59e0b" tick={{ fontSize: 11, fill: '#f59e0b' }} />
                    <Tooltip contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} />
                    <Bar yAxisId="left" dataKey="Notas" fill="#06b6d4" name="Qtd Notas" radius={[2, 2, 0, 0]} />
                    <Bar yAxisId="right" dataKey="Montante" fill="#f59e0b" name="Montante (R$)" radius={[2, 2, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">VENCIMENTO DAS FATURAS BLOQUEADAS</h2>
              <div className="h-56">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={vencimentoPieData} innerRadius={45} outerRadius={70} paddingAngle={3} dataKey="value">
                      {vencimentoPieData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} />
                    <Legend wrapperStyle={{ fontSize: '10px', color: '#9ca3af' }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* Linha 2: Faturamento & Tipos */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">% MÉDIO DE DIVERGÊNCIA POR FATURAMENTO</h2>
              <div className="space-y-3 pt-2">
                {faturamentoData.map((item, idx) => (
                  <div key={idx} className="bg-[#1a2234] p-3 rounded border border-gray-800 flex justify-between items-center">
                    <div>
                      <span className="text-xs font-bold text-white uppercase">{item.faturamento}</span>
                      <p className="text-[10px] text-gray-400">{item.count} notas registradas</p>
                    </div>
                    <div className="text-right">
                      <span className="text-base font-bold text-cyan-400">{item.pctMedia}%</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">DIVERGÊNCIAS POR TIPO E MOTIVO</h2>
              <div className="grid grid-cols-2 gap-3 pt-2">
                <div className="bg-[#1a2234] p-3 rounded border border-gray-800">
                  <span className="text-[10px] text-gray-400 uppercase">Preço (P)</span>
                  <p className="text-xl font-bold text-cyan-400 mt-1">627</p>
                  <span className="text-[10px] text-gray-500">R$ 4.408.894</span>
                </div>
                <div className="bg-[#1a2234] p-3 rounded border border-gray-800">
                  <span className="text-[10px] text-gray-400 uppercase">Não Informado</span>
                  <p className="text-xl font-bold text-amber-400 mt-1">1.112</p>
                  <span className="text-[10px] text-gray-500">Notas sem motivo</span>
                </div>
                <div className="bg-[#1a2234] p-3 rounded border border-gray-800">
                  <span className="text-[10px] text-gray-400 uppercase">Outros (486)</span>
                  <p className="text-xl font-bold text-emerald-400 mt-1">486</p>
                  <span className="text-[10px] text-gray-500">R$ 642.371</span>
                </div>
              </div>
            </div>
          </div>

          {/* Linha 3: Ranking Fornecedores */}
          <div className="bg-[#131927] p-4 rounded-lg border border-gray-800 mb-6">
            <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">RANKING DE FORNECEDORES OFENSORES (TOP 10 MONTANTE)</h2>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={topFornecedoresData} layout="vertical">
                  <XAxis type="number" stroke="#6b7280" tick={{ fontSize: 10, fill: '#9ca3af' }} />
                  <YAxis type="category" dataKey="name" stroke="#6b7280" width={160} tick={{ fontSize: 10, fill: '#9ca3af' }} />
                  <Tooltip formatter={(value) => formatBRL(value)} contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} />
                  <Bar dataKey="montante" fill="#06b6d4" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Linha 4: Tabelas Negociador e Responsável */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6 text-xs">
            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">DESEMPENHO POR NEGOCIADOR</h2>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-gray-800 text-gray-400 text-[10px] uppercase">
                      <th className="pb-2">Nome</th>
                      <th className="pb-2 text-center">Notas</th>
                      <th className="pb-2 text-right">Montante</th>
                      <th className="pb-2 text-right">Diferença</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800/50">
                    {negociadorTableData.map((row, i) => (
                      <tr key={i} className="hover:bg-[#1a2234]">
                        <td className="py-2 text-gray-200">{row.name}</td>
                        <td className="py-2 text-center text-gray-400">{row.notas}</td>
                        <td className="py-2 text-right text-emerald-400 font-medium">{formatBRL(row.montante)}</td>
                        <td className="py-2 text-right text-cyan-400 font-medium">{formatBRL(row.dif)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
              <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">DESEMPENHO POR RESPONSÁVEL</h2>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-gray-800 text-gray-400 text-[10px] uppercase">
                      <th className="pb-2">Nome</th>
                      <th className="pb-2 text-center">Notas</th>
                      <th className="pb-2 text-right">Montante</th>
                      <th className="pb-2 text-right">Diferença</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-800/50">
                    {responsavelTableData.map((row, i) => (
                      <tr key={i} className="hover:bg-[#1a2234]">
                        <td className="py-2 text-gray-200">{row.name}</td>
                        <td className="py-2 text-center text-gray-400">{row.notas}</td>
                        <td className="py-2 text-right text-emerald-400 font-medium">{formatBRL(row.montante)}</td>
                        <td className="py-2 text-right text-cyan-400 font-medium">{formatBRL(row.dif)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          {/* Linha 5: Inconsistências por Planta */}
          <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
            <h2 className="text-xs font-semibold mb-3 text-gray-300 uppercase tracking-wider">INCONSISTÊNCIAS POR PLANTA OPERACIONAL</h2>
            <div className="h-64">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={plantaChartData}>
                  <XAxis dataKey="unidade" stroke="#6b7280" tick={{ fontSize: 9, fill: '#9ca3af' }} interval={0} angle={-25} textAnchor="end" />
                  <YAxis yAxisId="left" stroke="#06b6d4" tick={{ fontSize: 10, fill: '#06b6d4' }} />
                  <YAxis yAxisId="right" orientation="right" stroke="#f59e0b" tick={{ fontSize: 10, fill: '#f59e0b' }} />
                  <Tooltip contentStyle={{ backgroundColor: '#1f2937', borderColor: '#374151', borderRadius: '6px', fontSize: '12px' }} />
                  <Bar yAxisId="left" dataKey="Notas" fill="#06b6d4" name="Notas" radius={[2, 2, 0, 0]} />
                  <Bar yAxisId="right" dataKey="Montante" fill="#f59e0b" name="Montante (R$)" radius={[2, 2, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </>
      ) : (
        <div className="flex flex-col items-center justify-center py-24 bg-[#131927] rounded-lg border border-gray-800 text-gray-500 gap-3">
          <AlertCircle size={40} className="text-gray-600" />
          <p className="text-xs">Carregue o arquivo **Divergências procurement.xlsx** para visualizar a análise completa.</p>
        </div>
      )}
    </div>
  );
}