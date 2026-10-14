'use client';

import { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, TrendingUp, Package, DollarSign, CheckCircle2, Database, Layers, ArrowRight, Search, RefreshCw } from 'lucide-react';

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
  const [statusUpload, setStatusUpload] = useState('Aguardando envio dos relatórios de estoque e vendas...');

  // Estados dos 4 arquivos exigidos no passo a passo
  const [fileEstoqueShop9, setFileEstoqueShop9] = useState('');
  const [fileVendasShop9, setFileVendasShop9] = useState('');
  const [fileEstoqueTiny, setFileEstoqueTiny] = useState('');
  const [fileVendasTiny, setFileVendasTiny] = useState('');

  // Estados de Filtros Avançados
  const [filtroBusca, setFiltroBusca] = useState('');
  const [filtroMarca, setFiltroMarca] = useState('TODAS');
  const [filtroFilial, setFiltroFilial] = useState('TODAS');
  const [filtroClasse, setFiltroClasse] = useState('TODAS');
  const [filtroCor, setFiltroCor] = useState('TODAS');

  // Simulação do cruzamento unificado (Shop9 + Tiny + Chave Grade-Filial + Margens)
  const processarArquivos = () => {
    setLoading(true);
    setStatusUpload('Lendo planilhas, gerando chaves Grade + Filial e calculando custos e margens...');
    
    setTimeout(() => {
      const mockConsolidado: ProdutoConsolidado[] = Array.from({ length: 50 }, (_, idx) => {
        const filial = idx % 3 === 0 ? 8 : idx % 3 === 1 ? 6 : 10;
        const codigo = 10000 + (idx % 12);
        const grade = `${codigo}00${(idx % 3) + 1}`;
        const custo = Number((28 + (idx * 2.1)).toFixed(2));
        const pdvEc = Number((custo * 1.85).toFixed(2));
        const pdvLoja = Number((pdvEc * 1.06).toFixed(2));
        const margemEc = (pdvEc - custo) / pdvEc;
        const margemLoja = (pdvLoja - custo) / pdvLoja;
        const estoque = (idx * 3) % 25;

        return {
          gradeFilial: `${grade}-${filial}`,
          grade: grade,
          codigo: codigo,
          produto: `PRODUTO BEST FIT MODELO ${idx + 1}`,
          filial: filial,
          codCor: (idx % 4) + 1,
          cor: idx % 2 === 0 ? 'PRETO' : 'BRANCO',
          codTamanho: (idx % 4) + 1,
          tamanho: idx % 4 === 0 ? 'P' : idx % 4 === 1 ? 'M' : idx % 4 === 2 ? 'G' : 'GG',
          genero: 'FEMININO',
          marca: idx % 2 === 0 ? 'BEST FIT' : 'FARM',
          classe: 'TREINO E ACADEMIA',
          familia: 'BOLSAS E SUPORTES',
          fornecedor: 'WILSON',
          estoque: estoque,
          custo: custo,
          custoTotal: Number((estoque * custo).toFixed(2)),
          pdvLoja: pdvLoja,
          margemLoja: margemLoja,
          pdvEc: pdvEc,
          margemEc: margemEc,
          vendaGeralItens: idx % 8,
          vendaGeralValor: Number(((idx % 8) * pdvEc).toFixed(2)),
        };
      });

      setDados(mockConsolidado);
      setLoading(false);
      setStatusUpload('Processamento concluído com sucesso!');
      setActiveTab('dashboard');
    }, 1500);
  };

  const marcasUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.marca)))], [dados]);
  const filiaisUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => String(d.filial))))], [dados]);
  const classesUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.classe)))], [dados]);
  const coresUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.cor)))], [dados]);

  const dadosFiltrados = useMemo(() => {
    return dados.filter(item => {
      const matchBusca = filtroBusca === '' || 
        item.produto.toLowerCase().includes(filtroBusca.toLowerCase()) || 
        String(item.codigo).includes(filtroBusca) ||
        item.grade.includes(filtroBusca);

      const matchMarca = filtroMarca === 'TODAS' || item.marca === filtroMarca;
      const matchFilial = filtroFilial === 'TODAS' || String(item.filial) === filtroFilial;
      const matchClasse = filtroClasse === 'TODAS' || item.classe === filtroClasse;
      const matchCor = filtroCor === 'TODAS' || item.cor === filtroCor;

      return matchBusca && matchMarca && matchFilial && matchClasse && matchCor;
    });
  }, [dados, filtroBusca, filtroMarca, filtroFilial, filtroClasse, filtroCor]);

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100 flex flex-col">
      
      {/* Header */}
      <header className="border-b border-slate-800 bg-slate-900/60 backdrop-blur sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600/20 border border-indigo-500/30 p-2 rounded-lg text-indigo-400">
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
              <Upload className="w-4 h-4" /> 1. Upload das 4 Fontes
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
          /* ABA 1: UPLOAD DOS 4 RELATÓRIOS DO PASSO A PASSO */
          <div className="space-y-8 max-w-5xl mx-auto animate-in fade-in duration-300">
            <div>
              <h1 className="text-3xl font-bold tracking-tight text-white">Ingestão dos Relatórios (Shop9 & Tiny ERP)</h1>
              <p className="text-slate-400 mt-1">Carregue as planilhas exigidas no passo a passo para unificação e cruzamento automático de chaves.</p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              
              {/* Fonte 1: Estoque Shop9 */}
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 1.1 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Estoque (Shop9)</h3>
                  <p className="text-xs text-slate-400 mb-4">Relatório extraído via Super Gerador com custos e tabelas de preço.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate">
                  <span className="text-xs font-medium text-slate-300">{fileEstoqueShop9 || 'Selecionar Estoque Shop9 (.xlsx)'}</span>
                  <input type="file" accept=".xlsx,.csv" onChange={(e) => setFileEstoqueShop9(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              {/* Fonte 2: Vendas Shop9 */}
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 1.2 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Vendas Filiais (Shop9)</h3>
                  <p className="text-xs text-slate-400 mb-4">Relatório de movimentação de vendas do mês vigente.</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate">
                  <span className="text-xs font-medium text-slate-300">{fileVendasShop9 || 'Selecionar Vendas Shop9 (.xlsx)'}</span>
                  <input type="file" accept=".xlsx,.csv" onChange={(e) => setFileVendasShop9(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              {/* Fonte 3: Estoque Tiny */}
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 2 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Estoque Saldo (Tiny ERP)</h3>
                  <p className="text-xs text-slate-400 mb-4">Relatório de saldos (Geral + FBA Amazon - Filial 10).</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate">
                  <span className="text-xs font-medium text-slate-300">{fileEstoqueTiny || 'Selecionar Estoque Tiny (.xlsx)'}</span>
                  <input type="file" accept=".xlsx,.csv" onChange={(e) => setFileEstoqueTiny(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

              {/* Fonte 4: Vendas Tiny */}
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex flex-col justify-between">
                <div>
                  <div className="text-indigo-400 font-semibold text-xs uppercase tracking-wider mb-1">Passo 3 do Manual</div>
                  <h3 className="text-lg font-bold text-white mb-2">Vendas Notas Fiscais (Tiny ERP)</h3>
                  <p className="text-xs text-slate-400 mb-4">Relatório baseado em Notas Fiscais emitidas (E-commerce).</p>
                </div>
                <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-4 text-center cursor-pointer bg-slate-950 block truncate">
                  <span className="text-xs font-medium text-slate-300">{fileVendasTiny || 'Selecionar Vendas Tiny (.xlsx)'}</span>
                  <input type="file" accept=".xlsx,.csv" onChange={(e) => setFileVendasTiny(e.target.files?.[0]?.name || '')} className="hidden" />
                </label>
              </div>

            </div>

            {/* Ação de Processamento */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 flex items-center justify-between">
              <div>
                <h4 className="font-medium text-white">Status da Automação</h4>
                <p className="text-sm text-slate-400">{statusUpload}</p>
              </div>
              <button
                onClick={processarArquivos}
                disabled={loading}
                className="px-6 py-3 bg-indigo-600 hover:bg-indigo-500 text-white font-medium rounded-xl transition-all shadow-lg flex items-center gap-2 disabled:opacity-50"
              >
                {loading ? 'Cruzando Bases...' : 'Processar & Gerar Painel'} <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        ) : (
          /* ABA 2: PAINEL E FILTROS COMPLETOS */
          <div className="space-y-6 animate-in fade-in duration-300">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-3xl font-bold tracking-tight text-white">Painel Consolidado & Cruzamento</h1>
                <p className="text-slate-400 mt-1">Filtre por qualquer título (Cor, Grade, Filial, Marca, Classe) com as chaves consolidadas.</p>
              </div>
              <button 
                onClick={() => setActiveTab('upload')}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-sm font-medium rounded-xl border border-slate-700 flex items-center gap-2"
              >
                <RefreshCw className="w-4 h-4" /> Enviar Outros Arquivos
              </button>
            </div>

            {/* Filtros */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-lg grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Buscar Produto / Grade</label>
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-500 absolute left-3 top-3" />
                  <input
                    type="text"
                    placeholder="Nome, código..."
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
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Cor</label>
                <select value={filtroCor} onChange={(e) => setFiltroCor(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {coresUnicas.map((cor, idx) => <option key={idx} value={cor}>{cor}</option>)}
                </select>
              </div>
            </div>

            {/* Tabela Detalhada */}
            <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl">
              <div className="p-4 border-b border-slate-800 bg-slate-950 flex justify-between items-center">
                <h3 className="text-sm font-semibold text-white">Registros Consolidados ({dadosFiltrados.length})</h3>
                <span className="text-xs text-emerald-400 font-medium">Cruzamento Shop9 + Tiny Realizado</span>
              </div>
              <div className="overflow-x-auto max-h-[550px]">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="bg-slate-950/80 text-slate-400 sticky top-0 border-b border-slate-800 z-10">
                    <tr>
                      <th className="p-3">Grade + Filial</th>
                      <th className="p-3">Código</th>
                      <th className="p-3">Produto</th>
                      <th className="p-3">Filial</th>
                      <th className="p-3">Cor</th>
                      <th className="p-3">Tamanho</th>
                      <th className="p-3">Marca</th>
                      <th className="p-3">Estoque</th>
                      <th className="p-3">Custo (R$)</th>
                      <th className="p-3">PDV E-c</th>
                      <th className="p-3">Margem E-c</th>
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
                        <td className="p-3 font-bold text-white">{item.estoque}</td>
                        <td className="p-3 text-slate-300">R$ {item.custo.toFixed(2)}</td>
                        <td className="p-3 text-slate-300">R$ {item.pdvEc.toFixed(2)}</td>
                        <td className="p-3 text-emerald-400 font-medium">{(item.margemEc * 100).toFixed(1)}%</td>
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