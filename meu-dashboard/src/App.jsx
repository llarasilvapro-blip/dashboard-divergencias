import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, AlertCircle } from 'lucide-react';
import { 
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, 
  PieChart, Pie, Cell, Legend 
} from 'recharts';

const COLORS = ['#06b6d4', '#f59e0b', '#ef4444', '#10b981', '#6366f1', '#ec4899', '#8b5cf6', '#14b8a6'];

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

  // Totalizadores Globais da Base (para % comparativo)
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

  // 1. TOP 10 FORNECEDORES
  const topFornecedoresData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      let forn = "Outros";
      Object.keys(row).forEach(k => {
        if (k.toUpperCase().includes("FORNECEDOR")) forn = row[k];
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

  // 2. NEGOCIADORES / COMPRADORES
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

  // 3. MATRIZ DE CRITICIDADE (AGING x MONTANTE) - EM SUBSTITUIÇÃO AO RESPONSÁVEL
  const matrizCriticidadeData = useMemo(() => {
    let critico = { notas: 0, montante: 0 };
    let atencao = { notas: 0, montante: 0 };
    let normal = { notas: 0, montante: 0 };

    filteredData.forEach(row => {
      const montante = parseFloat(row["MONTANTE"]) || 0;
      const dias = parseFloat(row["DIAS "]) || parseFloat(row["DIAS"]) || 0;

      if (dias > 60 || montante >= 100000) {
        critico.notas += 1;
        critico.montante += montante;
      } else if (dias >= 31 || montante >= 20000) {
        atencao.notas += 1;
        atencao.montante += montante;
      } else {
        normal.notas += 1;
        normal.montante += montante;
      }
    });

    return [
      {
        nivel: "Crítico",
        descricao: "Aging > 60 dias ou > R$ 100k",
        corBadge: "bg-red-500/10 text-red-400 border-red-500/30",
        corTexto: "text-red-400",
        ...critico,
        pctNotas: totalRegistros ? ((critico.notas / totalRegistros) * 100).toFixed(1) : 0,
        pctMontante: montanteBloqueado ? ((critico.montante / montanteBloqueado) * 100).toFixed(1) : 0
      },
      {
        nivel: "Atenção",
        descricao: "Aging 31-60 dias ou R$ 20k-100k",
        corBadge: "bg-amber-500/10 text-amber-400 border-amber-500/30",
        corTexto: "text-amber-400",
        ...atencao,
        pctNotas: totalRegistros ? ((atencao.notas / totalRegistros) * 100).toFixed(1) : 0,
        pctMontante: montanteBloqueado ? ((atencao.montante / montanteBloqueado) * 100).toFixed(1) : 0
      },
      {
        nivel: "Normal",
        descricao: "Aging ≤ 30 dias e < R$ 20k",
        corBadge: "bg-emerald-500/10 text-emerald-400 border-emerald-500/30",
        corTexto: "text-emerald-400",
        ...normal,
        pctNotas: totalRegistros ? ((normal.notas / totalRegistros) * 100).toFixed(1) : 0,
        pctMontante: montanteBloqueado ? ((normal.montante / montanteBloqueado) * 100).toFixed(1) : 0
      }
    ];
  }, [filteredData, totalRegistros, montanteBloqueado]);

  // 4. TABELA DE DIVERGÊNCIAS POR TIPO E MOTIVO
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

  // 5. AGING DO BLOQUEIO POR FAIXAS
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
      Montante: Math.round(buckets[key].montante),
      pctNotas: totalRegistros ? ((buckets[key].notas / totalRegistros) * 100).toFixed(1) : 0,
      pctMontante: montanteBloqueado ? ((buckets[key].montante / montanteBloqueado) * 100).toFixed(1) : 0
    }));
  }, [filteredData, totalRegistros, montanteBloqueado]);

  // 6. VENCIMENTO DAS FATURAS (DONUT)
  const vencimentoPieData = useMemo(() => {
    const map = {};
    filteredData.forEach(row => {
      const pag = (row["PAGAMENTO"] || "Outros").trim();
      map[pag] = (map[pag] || 0) + 1;
    });
    return Object.keys(map).map(key => ({ 
      name: key, 
      value: map[key],
      pct: totalRegistros ? ((map[key] / totalRegistros) * 100).toFixed(1) : 0
    }));
  }, [filteredData, totalRegistros]);

  // 7. INCONSISTÊNCIAS POR PLANTA OPERACIONAL (Qtd vs Montante)
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

      {/* KPIS PRINCIPAIS (COM TOTAL DE NOTAS E %) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 mb-6">
        <div className="bg-[#131927] p-4 rounded-lg border border-gray-800">
          <span className="text-[10px] text-gray-400 uppercase font-semibold">TOTAL DE NOTAS</span>
          <p className="text-2xl font-bold text-white mt-1">{totalRegistros}</p>
          <span className="text