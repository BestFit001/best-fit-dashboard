'use client';

import { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, TrendingUp, Package, DollarSign, CheckCircle2, Database, Layers, ArrowRight, Search, Filter, RefreshCw } from 'lucide-react';

interface ProdutoConsolidado {
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
  const [dados, setDados] = useState<ProdutoConsolidado[]>([]);
  const [statusUpload, setStatusUpload] = useState('Aguardando envio dos relatórios...');

  // Estados de Filtros Avançados
  const [filtroBusca, setFiltroBusca] = useState('');
  const [filtroMarca, setFiltroMarca] = useState('TODAS');
  const [filtroFilial, setFiltroFilial] = useState('TODAS');
  const [filtroClasse, setFiltroClasse] = useState('TODAS');
  const [filtroGenero, setFiltroGenero] = useState('TODAS');
  const [filtroCor, setFiltroCor] = useState('TODAS');

  // Simulação de processamento e cruzamento completo das fontes
  const processarArquivos = () => {
    setLoading(true);
    setTimeout(() => {
      const mockCompleto: ProdutoConsolidado[] = Array.from({ length: 60 }, (_, idx) => {
        const filial = idx % 3 === 0 ? 8 : idx % 3 === 1 ? 6 : 10;
        const codigo = 10000 + (idx % 15);
        const grade = `${codigo}00${(idx % 4) + 1}`;
        const custo = Number((25 + (idx * 1.8)).toFixed(2));
        const pdvEc = Number((custo * 1.9).toFixed(2));
        const pdvLoja = Number((pdvEc * 1.08).toFixed(2));
        const margemEc = (pdvEc - custo) / pdvEc;
        const margemLoja = (pdvLoja - custo) / pdvLoja;
        const estoque = (idx * 4) % 30;

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
          fornecedor: 'FORNECEDOR OFICIAL',
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
      setStatusUpload('Relatórios consolidados e cruzados com sucesso!');
      setActiveTab('dashboard');
    }, 1200);
  };

  // Opções únicas para os selects de filtros
  const marcasUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.marca)))], [dados]);
  const filiaisUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => String(d.filial))))], [dados]);
  const classesUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.classe)))], [dados]);
  const generosUnicos = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.genero)))], [dados]);
  const coresUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.cor)))], [dados]);

  // Dados filtrados dinamicamente
  const dadosFiltrados = useMemo(() => {
    return dados.filter(item => {
      const matchBusca = filtroBusca === '' || 
        item.produto.toLowerCase().includes(filtroBusca.toLowerCase()) || 
        String(item.codigo).includes(filtroBusca) ||
        item.grade.includes(filtroBusca);

      const matchMarca = filtroMarca === 'TODAS' || item.marca === filtroMarca;
      const matchFilial = filtroFilial === 'TODAS' || String(item.filial) === filtroFilial;
      const matchClasse = filtroClasse === 'TODAS' || item.classe === filtroClasse;
      const matchGenero = filtroGenero === 'TODAS' || item.genero === filtroGenero;
      const matchCor = filtroCor === 'TODAS' || item.cor === filtroCor;

      return matchBusca && matchMarca && matchFilial && matchClasse && matchGenero && matchCor;
    });
  }, [dados, filtroBusca, filtroMarca, filtroFilial, filtroClasse, filtroGenero, filtroCor]);

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100 flex flex-col">
      
      {/* Header */}
      <header className="border-b border-slate-800 bg-slate-900/60 backdrop-blur sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600/20 border border-indigo-500/30 p-2 rounded-lg text-indigo-400">
              <Layers className="w-5 h-5" />
            </div>
            <span className="font-bold text-lg tracking-tight text-white">Best Fit | Central de Compras & Estoque</span>
          </div>

          <nav className="flex items-center gap-2 bg-slate-950 p-1 rounded-xl border border-slate-800">
            <button
              onClick={() => setActiveTab('upload')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 ${
                activeTab === 'upload' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white hover:bg-slate-900'
              }`}
            >
              <Upload className="w-4 h-4" /> 1. Ingestão de Arquivos
            </button>
            <button
              onClick={() => setActiveTab('dashboard')}
              className={`px-4 py-2 rounded-lg text-sm font-medium transition-all flex items-center gap-2 ${
                activeTab === 'dashboard' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:text-white hover:bg-slate-900'
              }`}
            >
              <Database className="w-4 h-4" /> 2. Painel & Filtros Avançados
            </button>
          </nav>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-6 py-8 flex-1 w-full">
        {activeTab === 'upload' ? (
          /* ABA 1: UPLOAD */
          <div className="space-y-8 max-w-4xl mx-auto animate-in fade-in duration-300">
            <div>
              <h1 className="text-3xl font-bold tracking-tight text-white">Importação de Relatórios Brutos</h1>
              <p className="text-slate-400 mt-1">Carregue as fontes do Shop9 e Tiny ERP para disparar o cruzamento automático.</p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Fonte 01</div>
                  <h3 className="text-lg font-bold text-white mb-2">Estoque Shop9</h3>
                  <p className="text-xs text-slate-400 mb-4">Relatório de estoque de todas as filiais.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block">
                  <span className="text-xs font-medium text-slate-300">Selecionar .xlsx</span>
                  <input type="file" accept=".xlsx,.csv" className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Fonte 02</div>
                  <h3 className="text-lg font-bold text-white mb-2">Vendas Shop9</h3>
                  <p className="text-xs text-slate-400 mb-4">Movimentação de vendas do mês vigente.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block">
                  <span className="text-xs font-medium text-slate-300">Selecionar .xlsx</span>
                  <input type="file" accept=".xlsx,.csv" className="hidden" />
                </label>
              </div>

              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Fonte 03</div>
                  <h3 className="text-lg font-bold text-white mb-2">Tiny ERP (FBA / Geral)</h3>
                  <p className="text-xs text-slate-400 mb-4">Saldos de estoque e notas fiscais de e-commerce.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block">
                  <span className="text-xs font-medium text-slate-300">Selecionar .xlsx</span>
                  <input type="file" accept=".xlsx,.csv" className="hidden" />
                </label>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex items-center justify-between">
              <div>
                <h4 className="font-medium text-white">Status do Processamento</h4>
                <p className="text-sm text-slate-400">{statusUpload}</p>
              </div>
              <button
                onClick={processarArquivos}
                disabled={loading}
                className="px-6 py-3 bg-indigo-600 hover:bg-indigo-500 text-white font-medium rounded-xl transition-all shadow-lg flex items-center gap-2 disabled:opacity-50"
              >
                {loading ? 'Processando Cruzamentos...' : 'Processar & Consolidar'} <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        ) : (
          /* ABA 2: PAINEL E FILTROS AVANÇADOS */
          <div className="space-y-6 animate-in fade-in duration-300">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-3xl font-bold tracking-tight text-white">Painel de Análise & Estoque Geral</h1>
                <p className="text-slate-400 mt-1">Filtre por qualquer título (Cor, Grade, Filial, Marca, Classe, etc.) com base na planilha consolidada.</p>
              </div>
              <button 
                onClick={() => setActiveTab('upload')}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-sm font-medium rounded-xl border border-slate-700 flex items-center gap-2"
              >
                <RefreshCw className="w-4 h-4" /> Atualizar Arquivos
              </button>
            </div>

            {/* Barra de Filtros Avançados */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-lg grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-4">
              
              {/* Pesquisa por Texto */}
              <div className="lg:col-span-2">
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Buscar Produto / Código / Grade</label>
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-500 absolute left-3 top-3" />
                  <input
                    type="text"
                    placeholder="Digite o nome, código ou grade..."
                    value={filtroBusca}
                    onChange={(e) => setFiltroBusca(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-9 pr-4 py-2 text-sm text-white focus:outline-none focus:border-indigo-500"
                  />
                </div>
              </div>

              {/* Filtro Marca */}
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Marca</label>
                <select
                  value={filtroMarca}
                  onChange={(e) => setFiltroMarca(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-indigo-500"
                >
                  {marcasUnicas.map((m, idx) => <option key={idx} value={m}>{m}</option>)}
                </select>
              </div>

              {/* Filtro Filial */}
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Filial</label>
                <select
                  value={filtroFilial}
                  onChange={(e) => setFiltroFilial(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-indigo-500"
                >
                  {filiaisUnicas.map((f, idx) => <option key={idx} value={f}>Filial {f}</option>)}
                </select>
              </div>

              {/* Filtro Classe */}
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Classe</label>
                <select
                  value={filtroClasse}
                  onChange={(e) => setFiltroClasse(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-indigo-500"
                >
                  {classesUnicas.map((c, idx) => <option key={idx} value={c}>{c}</option>)}
                </select>
              </div>

              {/* Filtro Cor */}
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Cor</label>
                <select
                  value={filtroCor}
                  onChange={(e) => setFiltroCor(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-indigo-500"
                >
                  {coresUnicas.map((cor, idx) => <option key={idx} value={cor}>{cor}</option>)}
                </select>
              </div>

            </div>

            {/* KPIs Resumo */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <span className="text-xs text-slate-400 font-medium">Itens Filtrados</span>
                <div className="text-2xl font-bold text-white mt-1">{dadosFiltrados.length}</div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <span className="text-xs text-slate-400 font-medium">Estoque Total</span>
                <div className="text-2xl font-bold text-white mt-1">{dadosFiltrados.reduce((acc, curr) => acc + curr.estoque, 0)} un</div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <span className="text-xs text-slate-400 font-medium">Custo Total Estoque</span>
                <div className="text-2xl font-bold text-amber-400 mt-1">R$ {dadosFiltrados.reduce((acc, curr) => acc + curr.custoTotal, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <span className="text-xs text-slate-400 font-medium">Margem Média E-c</span>
                <div className="text-2xl font-bold text-emerald-400 mt-1">
                  {dadosFiltrados.length > 0 ? `${(dadosFiltrados.reduce((acc, curr) => acc + curr.margemEc, 0) / dadosFiltrados.length * 100).toFixed(1)}%` : '0%'}
                </div>
              </div>
            </div>

            {/* Tabela Completa de Dados */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
              <div className="p-4 border-b border-slate-800 flex justify-between items-center bg-slate-950">
                <h3 className="text-sm font-semibold text-white flex items-center gap-2">
                  <Database className="w-4 h-4 text-indigo-400" /> Detalhamento Completo da Base Consolidada
                </h3>
                <span className="text-xs text-slate-400">Exibindo {dadosFiltrados.length} registros</span>
              </div>
              
              <div className="overflow-x-auto max-h-[550px]">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-950/80 text-slate-400 sticky top-0 border-b border-slate-800 z-10">
                    <tr>
                      <th className="p-3 font-medium">Grade + Filial</th>
                      <th className="p-3 font-medium">Código</th>
                      <th className="p-3 font-medium">Produto</th>
                      <th className="p-3 font-medium">Filial</th>
                      <th className="p-3 font-medium">Cor</th>
                      <th className="p-3 font-medium">Tamanho</th>
                      <th className="p-3 font-medium">Gênero</th>
                      <th className="p-3 font-medium">Marca</th>
                      <th className="p-3 font-medium">Classe</th>
                      <th className="p-3 font-medium">Família</th>
                      <th className="p-3 font-medium">Estoque</th>
                      <th className="p-3 font-medium">Custo (R$)</th>
                      <th className="p-3 font-medium">Custo Total</th>
                      <th className="p-3 font-medium">PDV Loja</th>
                      <th className="p-3 font-medium">Margem Loja</th>
                      <th className="p-3 font-medium">PDV E-c</th>
                      <th className="p-3 font-medium">Margem E-c</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {dadosFiltrados.length === 0 ? (
                      <tr>
                        <td colSpan={17} className="p-8 text-center text-slate-500">
                          Nenhum registro encontrado com os filtros selecionados.
                        </td>
                      </tr>
                    ) : (
                      dadosFiltrados.map((item, index) => (
                        <tr key={index} className="hover:bg-slate-800/30 transition-colors">
                          <td className="p-3 font-mono text-indigo-400">{item.gradeFilial}</td>
                          <td className="p-3 font-mono">{item.codigo}</td>
                          <td className="p-3 font-medium text-white">{item.produto}</td>
                          <td className="p-3">{item.filial}</td>
                          <td className="p-3 text-slate-300">{item.cor}</td>
                          <td className="p-3 text-slate-300">{item.tamanho}</td>
                          <td className="p-3 text-slate-400">{item.genero}</td>
                          <td className="p-3 font-semibold text-slate-200">{item.marca}</td>
                          <td className="p-3 text-slate-400">{item.classe}</td>
                          <td className="p-3 text-slate-400">{item.familia}</td>
                          <td className="p-3 font-bold text-white">{item.estoque}</td>
                          <td className="p-3 text-slate-300">R$ {item.custo.toFixed(2)}</td>
                          <td className="p-3 text-slate-300">R$ {item.custoTotal.toFixed(2)}</td>
                          <td className="p-3 text-slate-300">R$ {item.pdvLoja.toFixed(2)}</td>
                          <td className="p-3 text-amber-400 font-medium">{(item.margemLoja * 100).toFixed(1)}%</td>
                          <td className="p-3 text-slate-300">R$ {item.pdvEc.toFixed(2)}</td>
                          <td className="p-3 text-emerald-400 font-medium">{(item.margemEc * 100).toFixed(1)}%</td>
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