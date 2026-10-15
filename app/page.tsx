'use client';

import { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '@/lib/supabase';
import { Upload, Database, Layers, ArrowRight, Search, RefreshCw, CheckCircle2, LayoutDashboard, Loader2, FileSpreadsheet } from 'lucide-react';

interface RegistroBase {
  gradeFilial: string;
  grade: string;
  codigo: number;
  produto: string;
  filial: number;
  codCor: number;
  cor: string;
  codTamanho: number;
  tamanho: string;
  genero: string;
  marca: string;
  codigoBarras: string;
  classe: string;
  familia: string;
  fornecedor: string;
  estoque: number;
  custoTotal: number;
  custo: number;
  margemLoja: number;
  pdvLoja: number;
  margemEc: number;
  pdvEc: number;
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<'upload' | 'dashboard' | 'dinamica'>('upload');
  const [loading, setLoading] = useState(false);
  const [dados, setDados] = useState<RegistroBase[]>([]);
  const [statusUpload, setStatusUpload] = useState('Selecione os arquivos individualmente para envio seguro.');

  // Os 4 arquivos independentes
  const [fileEstoqueShop9, setFileEstoqueShop9] = useState<File | null>(null);
  const [fileVendasShop9, setFileVendasShop9] = useState<File | null>(null);
  const [fileEstoqueTiny, setFileEstoqueTiny] = useState<File | null>(null);
  const [fileVendasTiny, setFileVendasTiny] = useState<File | null>(null);

  // Filtros da Base Consolidada
  const [filtroBusca, setFiltroBusca] = useState('');
  const [filtroMarca, setFiltroMarca] = useState('TODAS');
  const [filtroFilial, setFiltroFilial] = useState('TODAS');
  const [filtroClasse, setFiltroClasse] = useState('TODAS');

  // Configuração da Tabela Dinâmica
  const [linhaDinamica, setLinhaDinamica] = useState<keyof RegistroBase>('marca');
  const [colunaDinamica, setColunaDinamica] = useState<keyof RegistroBase>('filial');
  const [metricaDinamica, setMetricaDinamica] = useState<'estoque' | 'custoTotal' | 'pdvEc'>('estoque');

  useEffect(() => {
    carregarDadosDoSupabase();
  }, []);

  const carregarDadosDoSupabase = async () => {
    try {
      const { data, error } = await supabase
        .from('produtos_compras')
        .select('*')
        .limit(2000);

      if (error) throw error;

      if (data && data.length > 0) {
        const formatados: RegistroBase[] = data.map((item: any) => ({
          gradeFilial: item.grade_filial,
          grade: item.grade,
          codigo: item.codigo,
          produto: item.produto,
          filial: item.filial,
          codCor: item.cod_cor,
          cor: item.cor,
          codTamanho: item.cod_tamanho,
          tamanho: item.tamanho,
          genero: item.genero,
          marca: item.marca,
          codigoBarras: item.codigo_barras,
          classe: item.classe,
          familia: item.familia,
          fornecedor: item.fornecedor,
          estoque: item.estoque,
          custoTotal: Number(item.custo_total),
          custo: Number(item.custo),
          margemLoja: Number(item.margem_loja),
          pdvLoja: Number(item.pdv_loja),
          margemEc: Number(item.margem_ec),
          pdvEc: Number(item.pdv_ec),
        }));
        setDados(formatados);
      }
    } catch (err) {
      console.error('Erro ao buscar dados do Supabase:', err);
    }
  };

  const processarArquivos = async (e: React.FormEvent) => {
    e.preventDefault();
    const arquivoAlvo = fileEstoqueShop9 || fileVendasShop9 || fileEstoqueTiny || fileVendasTiny;
    if (!arquivoAlvo) {
      alert('Selecione pelo menos um arquivo para enviar!');
      return;
    }

    setLoading(true);
    setStatusUpload(`Lendo arquivo ${arquivoAlvo.name}...`);

    try {
      const buffer = await arquivoAlvo.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows: any[] = XLSX.utils.sheet_to_json(sheet);

      setStatusUpload(`Processando ${rows.length} registros em lotes seguros...`);

      const registrosProcessados = rows.map((row, idx) => {
        const filial = Number(row['Filial'] || 3);
        const grade = String(row['Grade'] || row['Código_da_Grade'] || `1000${idx}`);
        const codigo = Number(row['Codigo'] || row['Código_Produto'] || 1000 + idx);
        const produto = String(row['Produto'] || row['Nome_Produto'] || 'PRODUTO NÃO CADASTRADO');
        const cor = String(row['Cor'] || row['Nome_Cor'] || 'SEM COR');
        const tamanho = String(row['Tamanho'] || row['Nome_Tamanho'] || 'UNICO');
        const genero = String(row['Genero'] || 'NÃO CADASTRADO');
        const marca = String(row['Marca'] || 'NÃO CADASTRADO');
        const classe = String(row['Classe'] || 'NÃO CADASTRADO');
        const familia = String(row['Familia'] || 'NÃO CADASTRADO');
        const fornecedor = String(row['Fornecedor'] || 'NÃO CADASTRADO');
        const estoque = Number(row['Estoque'] || row['Qtde_Estoque_Atual'] || 0);
        const custo = Number(row['Preco'] || row['Preco_Total'] || 50.0);
        const pdvEc = Number(row['Preco1'] || (custo * 1.9));
        const pdvLoja = Number(row['Preco2'] || (pdvEc * 1.05));
        const margemEc = pdvEc > 0 ? (pdvEc - custo) / pdvEc : 0;
        const margemLoja = pdvLoja > 0 ? (pdvLoja - custo) / pdvLoja : 0;
        const custoTotal = Number((estoque * custo).toFixed(2));

        return {
          grade_filial: `${grade}-${filial}`,
          grade,
          codigo,
          produto,
          filial,
          cod_cor: Number(row['Cód Cor'] || 1),
          cor,
          cod_tamanho: Number(row['Cód Tamanho'] || 1),
          tamanho,
          genero,
          marca,
          codigo_barras: String(row['Codigo_Barras'] || ''),
          classe,
          familia,
          fornecedor,
          estoque,
          custo,
          custo_total: custoTotal,
          margem_loja: margemLoja,
          pdv_loja: pdvLoja,
          margem_ec: margemEc,
          pdv_ec: pdvEc,
        };
      });

      const tamanhoLote = 400;
      for (let i = 0; i < registrosProcessados.length; i += tamanhoLote) {
        const lote = registrosProcessados.slice(i, i + tamanhoLote);
        setStatusUpload(`Enviando lote ${Math.floor(i / tamanhoLote) + 1} de ${Math.ceil(registrosProcessados.length / tamanhoLote)}...`);

        const { error } = await supabase
          .from('produtos_compras')
          .upsert(lote, { onConflict: 'grade_filial' });

        if (error) console.error('Erro no lote:', error.message);
      }

      setStatusUpload('Arquivo sincronizado com sucesso com o Supabase!');
      setLoading(false);
      setFileEstoqueShop9(null);
      setFileVendasShop9(null);
      setFileEstoqueTiny(null);
      setFileVendasTiny(null);
      await carregarDadosDoSupabase();
      setActiveTab('dashboard');
    } catch (err) {
      console.error(err);
      setLoading(false);
      setStatusUpload('Erro ao processar o arquivo.');
    }
  };

  const marcasUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.marca)))], [dados]);
  const filiaisUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => String(d.filial))))], [dados]);
  const classesUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.classe)))], [dados]);

  const dadosFiltrados = useMemo(() => {
    return dados.filter(item => {
      const matchBusca = filtroBusca === '' || 
        item.produto.toLowerCase().includes(filtroBusca.toLowerCase()) || 
        String(item.codigo).includes(filtroBusca);

      return (
        matchBusca &&
        (filtroMarca === 'TODAS' || item.marca === filtroMarca) &&
        (filtroFilial === 'TODAS' || String(item.filial) === filtroFilial) &&
        (filtroClasse === 'TODAS' || item.classe === filtroClasse)
      );
    });
  }, [dados, filtroBusca, filtroMarca, filtroFilial, filtroClasse]);

  const estruturaDinamica = useMemo(() => {
    const linhasSet = new Set<string>();
    const colunasSet = new Set<string>();
    const matriz: { [linha: string]: { [coluna: string]: number } } = {};

    dadosFiltrados.forEach(item => {
      const linhaVal = String(item[linhaDinamica] || 'N/A');
      const colunaVal = String(item[colunaDinamica] || 'N/A');

      linhasSet.add(linhaVal);
      colunasSet.add(colunaVal);

      if (!matriz[linhaVal]) matriz[linhaVal] = {};
      if (!matriz[linhaVal][colunaVal]) matriz[linhaVal][colunaVal] = 0;

      if (metricaDinamica === 'estoque') matriz[linhaVal][colunaVal] += item.estoque;
      else if (metricaDinamica === 'custoTotal') matriz[linhaVal][colunaVal] += item.custoTotal;
      else if (metricaDinamica === 'pdvEc') matriz[linhaVal][colunaVal] += item.pdvEc;
    });

    return {
      linhas: Array.from(linhasSet).sort(),
      colunas: Array.from(colunasSet).sort(),
      matriz,
    };
  }, [dadosFiltrados, linhaDinamica, colunaDinamica, metricaDinamica]);

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100 flex flex-col">
      <header className="border-b border-slate-800 bg-slate-900/60 backdrop-blur sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600/25 border border-indigo-500/30 p-2 rounded-lg text-indigo-400">
              <Layers className="w-5 h-5" />
            </div>
            <span className="font-bold text-lg tracking-tight text-white">Best Fit | Compras & Estoque</span>
          </div>

          <nav className="flex items-center gap-2 bg-slate-950 p-1 rounded-xl border border-slate-800">
            <button
              onClick={() => setActiveTab('upload')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 ${
                activeTab === 'upload' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white hover:bg-slate-900'
              }`}
            >
              <Upload className="w-4 h-4" /> 1. Upload (4 Fontes)
            </button>
            <button
              onClick={() => setActiveTab('dashboard')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 ${
                activeTab === 'dashboard' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white hover:bg-slate-900'
              }`}
            >
              <Database className="w-4 h-4" /> 2. Base Consolidada
            </button>
            <button
              onClick={() => setActiveTab('dinamica')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 ${
                activeTab === 'dinamica' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white hover:bg-slate-900'
              }`}
            >
              <LayoutDashboard className="w-4 h-4" /> 3. Tabela Dinâmica
            </button>
          </nav>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-6 py-8 flex-1 w-full">
        {activeTab === 'upload' ? (
          <form onSubmit={processarArquivos} className="space-y-8 max-w-4xl mx-auto animate-in fade-in duration-300">
            <div>
              <h1 className="text-3xl font-bold tracking-tight text-white">Upload das 4 Fontes Oficiais</h1>
              <p className="text-slate-400 mt-1">Carregue os arquivos individualmente ou por etapa para alimentar o Supabase sem travar.</p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Fonte 01</div>
                  <h3 className="text-lg font-bold text-white mb-2">Estoque de Todas as Filiais (Shop9)</h3>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate mt-4">
                  <span className="text-xs font-medium text-slate-300">{fileEstoqueShop9 ? fileEstoqueShop9.name : 'Selecionar arquivo...'}</span>
                  <input type="file" accept=".xls,.xlsx,.txt,.csv" onChange={(e) => setFileEstoqueShop9(e.target.files?.[0] || null)} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Fonte 02</div>
                  <h3 className="text-lg font-bold text-white mb-2">Vendas de Todas as Filiais (Shop9)</h3>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate mt-4">
                  <span className="text-xs font-medium text-slate-300">{fileVendasShop9 ? fileVendasShop9.name : 'Selecionar arquivo...'}</span>
                  <input type="file" accept=".xls,.xlsx,.txt,.csv" onChange={(e) => setFileVendasShop9(e.target.files?.[0] || null)} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Fonte 03</div>
                  <h3 className="text-lg font-bold text-white mb-2">Saldos em Estoque (Tiny ERP)</h3>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate mt-4">
                  <span className="text-xs font-medium text-slate-300">{fileEstoqueTiny ? fileEstoqueTiny.name : 'Selecionar arquivo...'}</span>
                  <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setFileEstoqueTiny(e.target.files?.[0] || null)} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Fonte 04</div>
                  <h3 className="text-lg font-bold text-white mb-2">Relatório de Vendas / Notas (Tiny ERP)</h3>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate mt-4">
                  <span className="text-xs font-medium text-slate-300">{fileVendasTiny ? fileVendasTiny.name : 'Selecionar arquivo...'}</span>
                  <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setFileVendasTiny(e.target.files?.[0] || null)} className="hidden" />
                </label>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex items-center justify-between">
              <div>
                <h4 className="font-medium text-white text-xs uppercase tracking-wider text-indigo-400">Status</h4>
                <p className="text-sm text-slate-300">{statusUpload}</p>
              </div>
              <button
                type="submit"
                disabled={loading}
                className="px-6 py-3 bg-indigo-600 hover:bg-indigo-500 text-white font-medium rounded-xl transition-all shadow-lg flex items-center gap-2 disabled:opacity-50 text-sm"
              >
                {loading && <Loader2 className="w-4 h-4 animate-spin" />}
                {loading ? 'Processando...' : 'Enviar para o Supabase'} <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </form>
        ) : activeTab === 'dashboard' ? (
          <div className="space-y-6 animate-in fade-in duration-300">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-3xl font-bold tracking-tight text-white">Base Consolidada (Supabase)</h1>
                <p className="text-slate-400 mt-1">Registros salvos e recuperados do banco de dados.</p>
              </div>
              <button 
                onClick={() => setActiveTab('upload')}
                className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white text-sm font-medium rounded-xl shadow flex items-center gap-2"
              >
                <RefreshCw className="w-4 h-4" /> Enviar Mais Arquivos
              </button>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-lg grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Buscar Produto</label>
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-500 absolute left-3 top-3" />
                  <input
                    type="text"
                    placeholder="Nome, código..."
                    value={filtroBusca}
                    onChange={(e) => setFiltroBusca(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-9 pr-4 py-2 text-sm text-white focus:outline-none"
                  />
                </div>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Marca</label>
                <select value={filtroMarca} onChange={(e) => setFiltroMarca(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {marcasUnicas.map((m, idx) => <option key={idx} value={m}>{m}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Filial</label>
                <select value={filtroFilial} onChange={(e) => setFiltroFilial(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {filiaisUnicas.map((f, idx) => <option key={idx} value={f}>Filial {f}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Classe</label>
                <select value={filtroClasse} onChange={(e) => setFiltroClasse(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {classesUnicas.map((c, idx) => <option key={idx} value={c}>{c}</option>)}
                </select>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
              <div className="p-4 border-b border-slate-800 bg-slate-950 flex justify-between items-center">
                <h3 className="text-sm font-semibold text-white">Registros ({dadosFiltrados.length})</h3>
                <span className="text-xs text-emerald-400 font-medium flex items-center gap-1.5">
                  <CheckCircle2 className="w-3.5 h-3.5" /> Sincronizado
                </span>
              </div>
              <div className="overflow-x-auto max-h-[550px]">
                <table className="w-full text-left border-collapse text-xs whitespace-nowrap">
                  <thead className="bg-slate-950/90 text-slate-400 sticky top-0 border-b border-slate-800 z-10">
                    <tr>
                      <th className="p-3">Grade + filial</th>
                      <th className="p-3">Código</th>
                      <th className="p-3">Produto</th>
                      <th className="p-3">Filial</th>
                      <th className="p-3">Cor</th>
                      <th className="p-3">Tamanho</th>
                      <th className="p-3">Marca</th>
                      <th className="p-3">Classe</th>
                      <th className="p-3">Estoque</th>
                      <th className="p-3">Custo Total</th>
                      <th className="p-3">PDV E-c</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {dadosFiltrados.map((item, index) => (
                      <tr key={index} className="hover:bg-slate-800/30 transition-colors">
                        <td className="p-3 font-mono text-indigo-400">{item.gradeFilial}</td>
                        <td className="p-3 font-mono">{item.codigo}</td>
                        <td className="p-3 font-medium text-white">{item.produto}</td>
                        <td className="p-3">{item.filial}</td>
                        <td className="p-3 text-slate-300">{item.cor}</td>
                        <td className="p-3 text-slate-300">{item.tamanho}</td>
                        <td className="p-3 font-semibold text-slate-200">{item.marca}</td>
                        <td className="p-3 text-slate-400">{item.classe}</td>
                        <td className="p-3 font-bold text-white">{item.estoque}</td>
                        <td className="p-3 text-slate-300">R$ {item.custoTotal.toFixed(2)}</td>
                        <td className="p-3 text-slate-300">R$ {item.pdvEc.toFixed(2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : (
          <div className="space-y-6 animate-in fade-in duration-300">
            <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
              <div>
                <h1 className="text-3xl font-bold tracking-tight text-white">Tabela Dinâmica Interativa</h1>
                <p className="text-slate-400 mt-1">Cruzamento matricial em tempo real (Linhas x Colunas).</p>
              </div>
              <div className="flex flex-wrap items-center gap-3 bg-slate-900 border border-slate-800 p-3 rounded-xl shadow-lg">
                <div>
                  <label className="block text-[10px] uppercase font-semibold text-slate-400 mb-1">Linhas:</label>
                  <select value={linhaDinamica} onChange={(e) => setLinhaDinamica(e.target.value as any)} className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none">
                    <option value="marca">Marca</option>
                    <option value="classe">Classe</option>
                    <option value="filial">Filial</option>
                    <option value="cor">Cor</option>
                    <option value="tamanho">Tamanho</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] uppercase font-semibold text-slate-400 mb-1">Colunas:</label>
                  <select value={colunaDinamica} onChange={(e) => setColunaDinamica(e.target.value as any)} className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none">
                    <option value="filial">Filial</option>
                    <option value="marca">Marca</option>
                    <option value="classe">Classe</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] uppercase font-semibold text-slate-400 mb-1">Valores:</label>
                  <select value={metricaDinamica} onChange={(e) => setMetricaDinamica(e.target.value as any)} className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none">
                    <option value="estoque">Soma de Estoque</option>
                    <option value="custoTotal">Soma de Custo Total</option>
                    <option value="pdvEc">Soma de PDV E-commerce</option>
                  </select>
                </div>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
              <div className="p-4 border-b border-slate-800 bg-slate-950 flex justify-between items-center">
                <h3 className="text-sm font-semibold text-white">Matriz Dinâmica</h3>
                <span className="text-xs text-indigo-400 font-medium">{estruturaDinamica.linhas.length} linhas x {estruturaDinamica.colunas.length} colunas</span>
              </div>
              <div className="overflow-x-auto max-h-[550px]">
                <table className="w-full text-left border-collapse text-xs whitespace-nowrap">
                  <thead className="bg-slate-950/90 text-slate-400 sticky top-0 border-b border-slate-800 z-10">
                    <tr>
                      <th className="p-4 font-bold text-white border-r border-slate-800">Rótulos de Linha</th>
                      {estruturaDinamica.colunas.map((col, idx) => (
                        <th key={idx} className="p-4 text-center border-r border-slate-800">{colunaDinamica}: {col}</th>
                      ))}
                      <th className="p-4 text-center font-bold text-white">Total Geral</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {estruturaDinamica.linhas.length === 0 ? (
                      <tr>
                        <td colSpan={estruturaDinamica.colunas.length + 2} className="p-8 text-center text-slate-500">
                          Nenhum dado disponível. Faça o upload dos arquivos na primeira aba.
                        </td>
                      </tr>
                    ) : (
                      estruturaDinamica.linhas.map((linha, idx) => {
                        let totalLinha = 0;
                        return (
                          <tr key={idx} className="hover:bg-slate-800/30 transition-colors">
                            <td className="p-4 font-semibold text-white border-r border-slate-800">{linha}</td>
                            {estruturaDinamica.colunas.map((col, cIdx) => {
                              const val = estruturaDinamica.matriz[linha]?.[col] || 0;
                              totalLinha += val;
                              return (
                                <td key={cIdx} className="p-4 text-center text-slate-300 border-r border-slate-800">
                                  {metricaDinamica.includes('custo') || metricaDinamica.includes('pdv')
                                    ? `R$ ${val.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`
                                    : `${val.toLocaleString('pt-BR')} un`}
                                </td>
                              );
                            })}
                            <td className="p-4 text-center font-bold text-emerald-400">
                              {metricaDinamica.includes('custo') || metricaDinamica.includes('pdv')
                                ? `R$ ${totalLinha.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`
                                : `${totalLinha.toLocaleString('pt-BR')} un`}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>
    </main>
  );
}