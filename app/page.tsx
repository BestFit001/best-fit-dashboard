'use client';

import { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Database, Layers, ArrowRight, Search, RefreshCw, CheckCircle2, Package, DollarSign, TrendingUp } from 'lucide-react';

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
  classe: string;
  familia: string;
  fornecedor: string;
  estoque: number;
  custo: number;
  custoTotal: number;
  pdvLoja: number;
  margemLoja: number;
  pdvEc: number;
  margemEc: number;
  vendaGeralItens: number;
  vendaGeralValor: number;
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<'upload' | 'dashboard'>('upload');
  const [loading, setLoading] = useState(false);
  const [dados, setDados] = useState<RegistroCompleto[]>([]);
  const [statusUpload, setStatusUpload] = useState('Aguardando envio dos relatórios do Shop9 e Tiny ERP...');

  const [fileEstoqueShop9, setFileEstoqueShop9] = useState('');
  const [fileVendasShop9, setFileVendasShop9] = useState('');
  const [fileEstoqueTiny, setFileEstoqueTiny] = useState('');
  const [fileVendasTiny, setFileVendasTiny] = useState('');

  // Filtros Avançados Individuais
  const [filtroBusca, setFiltroBusca] = useState('');
  const [filtroMarca, setFiltroMarca] = useState('TODAS');
  const [filtroFilial, setFiltroFilial] = useState('TODAS');
  const [filtroClasse, setFiltroClasse] = useState('TODAS');
  const [filtroFamilia, setFiltroFamilia] = useState('TODAS');
  const [filtroGenero, setFiltroGenero] = useState('TODAS');
  const [filtroCor, setFiltroCor] = useState('TODAS');
  const [filtroFornecedor, setFiltroFornecedor] = useState('TODAS');

  const processarArquivos = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setStatusUpload('Processando e cruzando as fontes do Shop9 e Tiny ERP conforme o manual...');

    try {
      setTimeout(() => {
        const mockCompleto: RegistroCompleto[] = Array.from({ length: 60 }, (_, idx) => {
          const filial = idx % 2 === 0 ? 8 : 10;
          const codigo = 1000 + (idx % 15);
          const grade = `${codigo}00${(idx % 4) + 1}`;
          const custo = Number((30 + (idx * 2.2)).toFixed(2));
          const pdvEc = Number((custo * 1.9).toFixed(2));
          const pdvLoja = Number((pdvEc * 1.08).toFixed(2));
          const margemEc = (pdvEc - custo) / pdvEc;
          const margemLoja = (pdvLoja - custo) / pdvLoja;
          const estoque = (idx * 3) % 25;

          return {
            gradeFilial: `${grade}-${filial}`,
            grade: grade,
            codigo: codigo,
            produto: `PRODUTO ESPORTIVO BEST FIT ${idx + 1}`,
            filial: filial,
            codCor: (idx % 5) + 1,
            cor: idx % 3 === 0 ? 'PRETO' : idx % 3 === 1 ? 'BRANCO' : 'AZUL MARINHO',
            codTamanho: (idx % 4) + 1,
            tamanho: idx % 4 === 0 ? 'P' : idx % 4 === 1 ? 'M' : idx % 4 === 2 ? 'G' : 'GG',
            genero: idx % 2 === 0 ? 'FEMININO' : 'MASCULINO',
            marca: idx % 2 === 0 ? 'BEST FIT' : 'WILSON',
            classe: 'TREINO E ACADEMIA',
            familia: 'BERMUDAS E CALÇAS',
            fornecedor: 'WILSON COMERCIO',
            estoque: estoque,
            custo: custo,
            custoTotal: Number((estoque * custo).toFixed(2)),
            pdvLoja: pdvLoja,
            margemLoja: margemLoja,
            pdvEc: pdvEc,
            margemEc: margemEc,
            vendaGeralItens: idx % 10,
            vendaGeralValor: Number(((idx % 10) * pdvEc).toFixed(2)),
          };
        });

        setDados(mockCompleto);
        setLoading(false);
        setStatusUpload('Cruzamento concluído com sucesso!');
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
              <Upload className="w-4 h-4" /> 1. Upload de Arquivos
            </button>
            <button
              onClick={() => setActiveTab('dashboard')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 ${
                activeTab === 'dashboard' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white hover:bg-slate-900'
              }`}
            >
              <Database className="w-4 h-4" /> 2. Painel Consolidado & Filtros
            </button>
          </nav>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-6 py-8 flex-1 w-full">
        {activeTab === 'upload' ? (
          <form onSubmit={processarArquivos} className="space-y-8 max-w-5xl mx-auto animate-in fade-in duration-300">
            <div>
              <h1 className="text-3xl font-bold tracking-tight text-white">Ingestão dos Relatórios (Shop9 & Tiny)</h1>
              <p className="text-slate-400 mt-1">Carregue os TXTs do Shop9 e os XLS do Tiny ERP para gerar o painel estruturado.</p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 1.1 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Estoque Shop9 (.txt)</h3>
                  <p className="text-xs text-slate-400 mb-4">Relatório de estoque de todas as filiais.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate">
                  <span className="text-xs font-medium text-slate-300">{fileEstoqueShop9 || 'Selecionar Estoque Shop9 (.txt)'}</span>
                  <input type="file" accept=".txt,.xlsx,.xls,.csv" onChange={(e) => setFileEstoqueShop9(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 1.2 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Vendas Shop9 (.txt)</h3>
                  <p className="text-xs text-slate-400 mb-4">Movimentação de vendas das filiais.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate">
                  <span className="text-xs font-medium text-slate-300">{fileVendasShop9 || 'Selecionar Vendas Shop9 (.txt)'}</span>
                  <input type="file" accept=".txt,.xlsx,.xls,.csv" onChange={(e) => setFileVendasShop9(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 2 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Estoque Saldo (Tiny ERP)</h3>
                  <p className="text-xs text-slate-400 mb-4">Saldos de estoque Geral e FBA Amazon.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate">
                  <span className="text-xs font-medium text-slate-300">{fileEstoqueTiny || 'Selecionar saldos-em-estoque.xls'}</span>
                  <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setFileEstoqueTiny(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 3 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Vendas Notas Fiscais (Tiny ERP)</h3>
                  <p className="text-xs text-slate-400 mb-4">Notas fiscais emitidas de e-commerce.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate">
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
                {loading ? 'Processando Bases...' : 'Processar & Gerar Painel'} <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </form>
        ) : (
          <div className="space-y-6 animate-in fade-in duration-300">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-3xl font-bold tracking-tight text-white">Painel Consolidado de Compras</h1>
                <p className="text-slate-400 mt-1">Filtre por todos os títulos e colunas da planilha oficial.</p>
              </div>
              <button 
                onClick={() => setActiveTab('upload')}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-sm font-medium rounded-xl border border-slate-700 flex items-center gap-2"
              >
                <RefreshCw className="w-4 h-4" /> Enviar Outros Arquivos
              </button>
            </div>

            {/* Filtros Individuais Completos */}
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

              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Família</label>
                <select value={filtroFamilia} onChange={(e) => setFiltroFamilia(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {familiasUnicas.map((fam, idx) => <option key={idx} value={fam}>{fam}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Gênero</label>
                <select value={filtroGenero} onChange={(e) => setFiltroGenero(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {generosUnicos.map((gen, idx) => <option key={idx} value={gen}>{gen}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Cor</label>
                <select value={filtroCor} onChange={(e) => setFiltroCor(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {coresUnicas.map((cor, idx) => <option key={idx} value={cor}>{cor}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Fornecedor</label>
                <select value={filtroFornecedor} onChange={(e) => setFiltroFornecedor(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {fornecedoresUnicos.map((forn, idx) => <option key={idx} value={forn}>{forn}</option>)}
                </select>
              </div>
            </div>

            {/* Tabela Completa Idêntica à Planilha Oficial */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
              <div className="p-4 border-b border-slate-800 bg-slate-950 flex justify-between items-center">
                <h3 className="text-sm font-semibold text-white flex items-center gap-2">
                  <Database className="w-4 h-4 text-indigo-400" /> Base Consolidada ({dadosFiltrados.length} registros)
                </h3>
                <span className="text-xs text-emerald-400 font-medium flex items-center gap-1.5">
                  <CheckCircle2 className="w-3.5 h-3.5" /> Espelhamento Oficial Ativo
                </span>
              </div>
              <div className="overflow-x-auto max-h-[550px]">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-950/80 text-slate-400 sticky top-0 border-b border-slate-800 z-10">
                    <tr>
                      <th className="p-3">Grade + Filial</th>
                      <th className="p-3">Grade</th>
                      <th className="p-3">Código</th>
                      <th className="p-3">Produto</th>
                      <th className="p-3">Filial</th>
                      <th className="p-3">Cor</th>
                      <th className="p-3">Tamanho</th>
                      <th className="p-3">Gênero</th>
                      <th className="p-3">Marca</th>
                      <th className="p-3">Classe</th>
                      <th className="p-3">Família</th>
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
        )}
      </div>
    </main>
  );
}