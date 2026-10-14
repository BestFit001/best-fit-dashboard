'use client';

import { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Database, Layers, ArrowRight, Search, RefreshCw, CheckCircle2, LayoutDashboard } from 'lucide-react';

interface RegistroCompleto {
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
  maiQtt: number;
  maiValor: number;
  junQtt: number;
  junValor: number;
  julQtt: number;
  julValor: number;
  vendaGeralItens: number;
  vendaGeralValor: number;
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<'upload' | 'dashboard' | 'dinamica'>('upload');
  const [loading, setLoading] = useState(false);
  const [dados, setDados] = useState<RegistroCompleto[]>([]);
  const [statusUpload, setStatusUpload] = useState('Aguardando envio dos relatórios...');

  const [fileEstoqueShop9, setFileEstoqueShop9] = useState('');
  const [fileVendasShop9, setFileVendasShop9] = useState('');
  const [fileEstoqueTiny, setFileEstoqueTiny] = useState('');
  const [fileVendasTiny, setFileVendasTiny] = useState('');

  // Filtros Avançados
  const [filtroBusca, setFiltroBusca] = useState('');
  const [filtroMarca, setFiltroMarca] = useState('TODAS');
  const [filtroFilial, setFiltroFilial] = useState('TODAS');
  const [filtroClasse, setFiltroClasse] = useState('TODAS');
  const [filtroFamilia, setFiltroFamilia] = useState('TODAS');
  const [filtroGenero, setFiltroGenero] = useState('TODAS');
  const [filtroCor, setFiltroCor] = useState('TODAS');
  const [filtroFornecedor, setFiltroFornecedor] = useState('TODAS');

  // Estado da Tabela Dinâmica
  const [agrupamentoDinamica, setAgrupamentoDinamica] = useState<'marca' | 'classe' | 'familia' | 'filial'>('marca');

  const processarArquivos = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setStatusUpload('Processando e cruzando todas as colunas do Shop9 e Tiny ERP...');

    try {
      setTimeout(() => {
        const mockCompleto: RegistroCompleto[] = Array.from({ length: 60 }, (_, idx) => {
          const filial = idx % 2 === 0 ? 3 : 6;
          const codigo = 10002 + idx;
          const grade = `${codigo}003007`;
          const custo = Number((35 + (idx * 1.5)).toFixed(2));
          const pdvEc = Number((custo * 1.9).toFixed(2));
          const pdvLoja = Number((pdvEc * 1.05).toFixed(2));
          const margemEc = (pdvEc - custo) / pdvEc;
          const margemLoja = (pdvLoja - custo) / pdvLoja;
          const estoque = (idx * 2) % 15;

          return {
            gradeFilial: `${grade}-${filial}`,
            grade: grade,
            codigo: codigo,
            produto: `T-SHIRT COURT MODELO ${idx + 1}`,
            filial: filial,
            codCor: 7,
            cor: idx % 2 === 0 ? 'BRANCO' : 'PRETO',
            codTamanho: 3,
            tamanho: idx % 3 === 0 ? 'P' : idx % 3 === 1 ? 'M' : 'G',
            genero: idx % 2 === 0 ? 'FEMININO' : 'MASCULINO',
            marca: idx % 2 === 0 ? 'WILSON' : 'BEST FIT',
            codigoBarras: '78910293847',
            classe: 'TREINO E ACADEMIA',
            familia: idx % 2 === 0 ? 'BOLSAS E MOCHILAS' : 'CALÇAS E LEGGINGS',
            fornecedor: 'WILSON',
            estoque: estoque,
            custoTotal: Number((estoque * custo).toFixed(2)),
            custo: custo,
            margemLoja: margemLoja,
            pdvLoja: pdvLoja,
            margemEc: margemEc,
            pdvEc: pdvEc,
            maiQtt: idx,
            maiValor: Number((idx * pdvEc).toFixed(2)),
            junQtt: idx + 1,
            junValor: Number(((idx + 1) * pdvEc).toFixed(2)),
            julQtt: idx + 2,
            julValor: Number(((idx + 2) * pdvEc).toFixed(2)),
            vendaGeralItens: idx * 3,
            vendaGeralValor: Number(((idx * 3) * pdvEc).toFixed(2)),
          };
        });

        setDados(mockCompleto);
        setLoading(false);
        setStatusUpload('Dados carregados com sucesso!');
        setActiveTab('dashboard');
      }, 1200);
    } catch (err) {
      console.error(err);
      setLoading(false);
      setStatusUpload('Erro ao processar os arquivos.');
    }
  };

  const marcasUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.marca)))], [dados]);
  const filiaisUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => String(d.filial))))], [dados]);
  const classesUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.classe)))], [dados]);
  const familiasUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.familia)))], [dados]);
  const generosUnicos = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.genero)))], [dados]);
  const coresUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.cor)))], [dados]);
  const fornecedoresUnicos = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.fornecedor)))], [dados]);

  const dadosFiltrados = useMemo(() => {
    return dados.filter(item => {
      const matchBusca = filtroBusca === '' || 
        item.produto.toLowerCase().includes(filtroBusca.toLowerCase()) || 
        String(item.codigo).includes(filtroBusca) ||
        item.grade.includes(filtroBusca);

      return (
        matchBusca &&
        (filtroMarca === 'TODAS' || item.marca === filtroMarca) &&
        (filtroFilial === 'TODAS' || String(item.filial) === filtroFilial) &&
        (filtroClasse === 'TODAS' || item.classe === filtroClasse) &&
        (filtroFamilia === 'TODAS' || item.familia === filtroFamilia) &&
        (filtroGenero === 'TODAS' || item.genero === filtroGenero) &&
        (filtroCor === 'TODAS' || item.cor === filtroCor) &&
        (filtroFornecedor === 'TODAS' || item.fornecedor === filtroFornecedor)
      );
    });
  }, [dados, filtroBusca, filtroMarca, filtroFilial, filtroClasse, filtroFamilia, filtroGenero, filtroCor, filtroFornecedor]);

  // Lógica da Tabela Dinâmica em Memória
  const dadosDinamica = useMemo(() => {
    const agrupado: { [key: string]: { rotulo: string; estoque: number; custoTotal: number; vendasItens: number; vendasValor: number; count: number } } = {};

    dadosFiltrados.forEach(item => {
      let chave = '';
      if (agrupamentoDinamica === 'marca') chave = item.marca;
      else if (agrupamentoDinamica === 'classe') chave = item.classe;
      else if (agrupamentoDinamica === 'familia') chave = item.familia;
      else if (agrupamentoDinamica === 'filial') chave = `Filial ${item.filial}`;

      if (!agrupado[chave]) {
        agrupado[chave] = { rotulo: chave, estoque: 0, custoTotal: 0, vendasItens: 0, vendasValor: 0, count: 0 };
      }

      agrupado[chave].estoque += item.estoque;
      agrupado[chave].custoTotal += item.custoTotal;
      agrupado[chave].vendasItens += item.vendaGeralItens;
      agrupado[chave].vendasValor += item.vendaGeralValor;
      agrupado[chave].count += 1;
    });

    return Object.values(agrupado);
  }, [dadosFiltrados, agrupamentoDinamica]);

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100 flex flex-col">
      <header className="border-b border-slate-800 bg-slate-900/60 backdrop-blur sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600/25 border border-indigo-500/30 p-2 rounded-lg text-indigo-400">
              <Layers className="w-5 h-5" />
            </div>
            <span className="font-bold text-lg tracking-tight text-white">Best Fit | Automação de Compras & Estoque</span>
          </div>

          <nav className="flex items-center gap-2 bg-slate-950 p-1 rounded-xl border border-slate-800">
            <button
              onClick={() => setActiveTab('upload')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 ${
                activeTab === 'upload' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white hover:bg-slate-900'
              }`}
            >
              <Upload className="w-4 h-4" /> 1. Upload
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
          <form onSubmit={processarArquivos} className="space-y-8 max-w-5xl mx-auto animate-in fade-in duration-300">
            <div>
              <h1 className="text-3xl font-bold tracking-tight text-white">Ingestão dos Relatórios (Shop9 & Tiny)</h1>
              <p className="text-slate-400 mt-1">Carregue os arquivos para alimentar o painel e a tabela dinâmica.</p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 1.1 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Estoque Shop9 (.txt)</h3>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate mt-4">
                  <span className="text-xs font-medium text-slate-300">{fileEstoqueShop9 || 'Selecionar Estoque Shop9 (.txt)'}</span>
                  <input type="file" accept=".txt,.xlsx,.xls,.csv" onChange={(e) => setFileEstoqueShop9(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 1.2 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Vendas Shop9 (.txt)</h3>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate mt-4">
                  <span className="text-xs font-medium text-slate-300">{fileVendasShop9 || 'Selecionar Vendas Shop9 (.txt)'}</span>
                  <input type="file" accept=".txt,.xlsx,.xls,.csv" onChange={(e) => setFileVendasShop9(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 2 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Estoque Saldo (Tiny ERP)</h3>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate mt-4">
                  <span className="text-xs font-medium text-slate-300">{fileEstoqueTiny || 'Selecionar saldos-em-estoque.xls'}</span>
                  <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setFileEstoqueTiny(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 3 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Vendas Notas Fiscais (Tiny ERP)</h3>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate mt-4">
                  <span className="text-xs font-medium text-slate-300">{fileVendasTiny || 'Selecionar relatorio-de-vendas.xls'}</span>
                  <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setFileVendasTiny(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex items-center justify-between">
              <div>
                <h4 className="font-medium text-white">Status da Automação</h4>
                <p className="text-sm text-slate-400">{statusUpload}</p>
              </div>
              <button
                type="submit"
                disabled={loading}
                className="px-6 py-3 bg-indigo-600 hover:bg-indigo-500 text-white font-medium rounded-xl transition-all shadow-lg flex items-center gap-2 disabled:opacity-50"
              >
                {loading ? 'Processando Bases...' : 'Processar & Gerar Painéis'} <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </form>
        ) : activeTab === 'dashboard' ? (
          <div className="space-y-6 animate-in fade-in duration-300">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-3xl font-bold tracking-tight text-white">Base Consolidada de Dados</h1>
                <p className="text-slate-400 mt-1">Todas as colunas oficiais e filtros individuais ativos.</p>
              </div>
              <button 
                onClick={() => setActiveTab('upload')}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-sm font-medium rounded-xl border border-slate-700 flex items-center gap-2"
              >
                <RefreshCw className="w-4 h-4" /> Enviar Outros Arquivos
              </button>
            </div>

            {/* Filtros */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-lg grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Buscar Produto / Código</label>
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-500 absolute left-3 top-3" />
                  <input
                    type="text"
                    placeholder="Nome, código, grade..."
                    value={filtroBusca}
                    onChange={(e) => setFiltroBusca(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-9 pr-4 py-2 text-sm text-white focus:outline-none focus:border-indigo-500"
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

            {/* Tabela Completa */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
              <div className="p-4 border-b border-slate-800 bg-slate-950 flex justify-between items-center">
                <h3 className="text-sm font-semibold text-white flex items-center gap-2">
                  <Database className="w-4 h-4 text-indigo-400" /> Registros ({dadosFiltrados.length})
                </h3>
                <span className="text-xs text-emerald-400 font-medium flex items-center gap-1.5">
                  <CheckCircle2 className="w-3.5 h-3.5" /> Dados em Memória Temporária
                </span>
              </div>
              <div className="overflow-x-auto max-h-[550px]">
                <table className="w-full text-left border-collapse text-xs whitespace-nowrap">
                  <thead className="bg-slate-950/90 text-slate-400 sticky top-0 border-b border-slate-800 z-10">
                    <tr>
                      <th className="p-3">Grade + filial</th>
                      <th className="p-3">grade</th>
                      <th className="p-3">Código</th>
                      <th className="p-3">Produto</th>
                      <th className="p-3">Filial</th>
                      <th className="p-3">Cor</th>
                      <th className="p-3">Tamanho</th>
                      <th className="p-3">Genero</th>
                      <th className="p-3">Marca</th>
                      <th className="p-3">Classe</th>
                      <th className="p-3">Familia</th>
                      <th className="p-3">Fornecedor</th>
                      <th className="p-3">Estoque</th>
                      <th className="p-3">Custo Total</th>
                      <th className="p-3">Custo</th>
                      <th className="p-3">Margem Loja</th>
                      <th className="p-3">PDV Loja</th>
                      <th className="p-3">Margem E-c</th>
                      <th className="p-3">PDV E-c</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {dadosFiltrados.map((item, index) => (
                      <tr key={index} className="hover:bg-slate-800/30 transition-colors">
                        <td className="p-3 font-mono text-indigo-400">{item.gradeFilial}</td>
                        <td className="p-3 font-mono">{item.grade}</td>
                        <td className="p-3 font-mono">{item.codigo}</td>
                        <td className="p-3 font-medium text-white">{item.produto}</td>
                        <td className="p-3">{item.filial}</td>
                        <td className="p-3 text-slate-300">{item.cor}</td>
                        <td className="p-3 text-slate-300">{item.tamanho}</td>
                        <td className="p-3 text-slate-400">{item.genero}</td>
                        <td className="p-3 font-semibold text-slate-200">{item.marca}</td>
                        <td className="p-3 text-slate-400">{item.classe}</td>
                        <td className="p-3 text-slate-400">{item.familia}</td>
                        <td className="p-3 text-slate-400">{item.fornecedor}</td>
                        <td className="p-3 font-bold text-white">{item.estoque}</td>
                        <td className="p-3 text-slate-300">R$ {item.custoTotal.toFixed(2)}</td>
                        <td className="p-3 text-slate-300">R$ {item.custo.toFixed(2)}</td>
                        <td className="p-3 text-amber-400 font-medium">{(item.margemLoja * 100).toFixed(1)}%</td>
                        <td className="p-3 text-slate-300">R$ {item.pdvLoja.toFixed(2)}</td>
                        <td className="p-3 text-emerald-400 font-medium">{(item.margemEc * 100).toFixed(1)}%</td>
                        <td className="p-3 text-slate-300">R$ {item.pdvEc.toFixed(2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : (
          /* ABA 3: TABELA DINÂMICA INTERATIVA */
          <div className="space-y-6 animate-in fade-in duration-300">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-3xl font-bold tracking-tight text-white">Tabela Dinâmica Interativa</h1>
                <p className="text-slate-400 mt-1">Agrupe e analise os dados temporários em memória.</p>
              </div>
              <div className="flex items-center gap-3">
                <label className="text-xs text-slate-400 font-medium">Agrupar por:</label>
                <select 
                  value={agrupamentoDinamica} 
                  onChange={(e) => setAgrupamentoDinamica(e.target.value as any)}
                  className="bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-indigo-500"
                >
                  <option value="marca">Marca</option>
                  <option value="classe">Classe</option>
                  <option value="familia">Família</option>
                  <option value="filial">Filial</option>
                </select>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
              <div className="p-4 border-b border-slate-800 bg-slate-950 flex justify-between items-center">
                <h3 className="text-sm font-semibold text-white flex items-center gap-2">
                  <LayoutDashboard className="w-4 h-4 text-indigo-400" /> Visão Dinâmica por {agrupamentoDinamica.toUpperCase()}
                </h3>
                <span className="text-xs text-indigo-400 font-medium">{dadosDinamica.length} grupos encontrados</span>
              </div>
              <div className="overflow-x-auto max-h-[550px]">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-950/90 text-slate-400 sticky top-0 border-b border-slate-800 z-10">
                    <tr>
                      <th className="p-4">Rótulos de Linha</th>
                      <th className="p-4">Qtd. Registros</th>
                      <th className="p-4">Soma de Estoque</th>
                      <th className="p-4">Custo Total (R$)</th>
                      <th className="p-4">Venda Geral (Itens)</th>
                      <th className="p-4">Venda Geral (Valor)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {dadosDinamica.length === 0 ? (
                      <tr>
                        <td colSpan={6} className="p-8 text-center text-slate-500">
                          Nenhum dado disponível. Faça o upload dos arquivos na primeira aba.
                        </td>
                      </tr>
                    ) : (
                      dadosDinamica.map((row, index) => (
                        <tr key={index} className="hover:bg-slate-800/30 transition-colors font-medium">
                          <td className="p-4 text-white font-semibold">{row.rotulo}</td>
                          <td className="p-4 text-slate-300">{row.count}</td>
                          <td className="p-4 text-emerald-400 font-bold">{row.estoque} un</td>
                          <td className="p-4 text-slate-300">R$ {row.custoTotal.toFixed(2)}</td>
                          <td className="p-4 text-indigo-300">{row.vendasItens}</td>
                          <td className="p-4 text-indigo-300">R$ {row.vendasValor.toFixed(2)}</td>
                        </tr>
                      ))
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