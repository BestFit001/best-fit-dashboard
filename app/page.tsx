'use client';

import { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '@/lib/supabase';
import { Upload, Database, Layers, ArrowRight, Search, RefreshCw, CheckCircle2, LayoutDashboard, Loader2, FileSpreadsheet, Trash2, Calendar } from 'lucide-react';

interface RegistroBase {
  id: number;
  dataReferencia: string;
  gradeFilial: string;
  grade: string;
  codigo: number;
  produto: string;
  filial: number;
  cor: string;
  tamanho: string;
  genero: string;
  marca: string;
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
  vendaGeralItens: number;
  vendaGeralValor: number;
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<'upload' | 'dashboard' | 'dinamica'>('upload');
  const [loading, setLoading] = useState(false);
  const [dados, setDados] = useState<RegistroBase[]>([]);
  const [statusUpload, setStatusUpload] = useState('Selecione a data e os relatórios para processamento.');

  const [dataEnvio, setDataEnvio] = useState<string>('');
  const [fileEstoqueShop, setFileEstoqueShop] = useState<File | null>(null);
  const [fileVendasShop, setFileVendasShop] = useState<File | null>(null);
  const [fileEstoqueTiny, setFileEstoqueTiny] = useState<File | null>(null);
  const [fileVendasTiny, setFileVendasTiny] = useState<File | null>(null);

  const [filtroBusca, setFiltroBusca] = useState('');
  const [filtroData, setFiltroData] = useState('TODAS');
  const [filtroMarca, setFiltroMarca] = useState('TODAS');
  const [filtroFilial, setFiltroFilial] = useState('TODAS');
  const [filtroClasse, setFiltroClasse] = useState('TODAS');

  const [linhaDinamica, setLinhaDinamica] = useState<keyof RegistroBase>('marca');
  const [colunaDinamica, setColunaDinamica] = useState<keyof RegistroBase>('filial');
  const [metricaDinamica, setMetricaDinamica] = useState<'estoque' | 'custoTotal' | 'pdvEc' | 'vendaGeralValor'>('estoque');

  useEffect(() => {
    setDataEnvio(new Date().toISOString().split('T')[0]);
    carregarDadosDoSupabase();
  }, []);

  const carregarDadosDoSupabase = async () => {
    try {
      const { data, error } = await supabase
        .from('produtos_compras')
        .select('*')
        .order('data_referencia', { ascending: false })
        .limit(6000);

      if (error) {
        console.error('Erro ao buscar dados do Supabase:', error.message);
        return;
      }

      if (data && data.length > 0) {
        const formatados: RegistroBase[] = data.map((item: any) => ({
          id: item.id,
          dataReferencia: item.data_referencia,
          gradeFilial: item.grade_filial,
          grade: item.grade,
          codigo: item.codigo,
          produto: item.produto,
          filial: item.filial,
          cor: item.cor || 'SEM COR',
          tamanho: item.tamanho || 'UNICO',
          genero: item.genero || 'Não Cadastrado',
          marca: item.marca || 'Não Cadastrado',
          classe: item.classe || 'NÃO CADASTRADO',
          familia: item.familia || 'Não Cadastrado',
          fornecedor: item.fornecedor || 'NÃO CADASTRADO',
          estoque: item.estoque || 0,
          custoTotal: Number(item.custo_total || 0),
          custo: Number(item.custo || 0),
          margemLoja: Number(item.margem_loja || 0),
          pdvLoja: Number(item.pdv_loja || 0),
          margemEc: Number(item.margem_ec || 0),
          pdvEc: Number(item.pdv_ec || 0),
          vendaGeralItens: Number(item.venda_geral_itens || 0),
          vendaGeralValor: Number(item.venda_geral_valor || 0),
        }));
        setDados(formatados);
      } else {
        setDados([]);
      }
    } catch (err) {
      console.error('Erro geral ao buscar dados:', err);
    }
  };

  const processarEEnviar = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!fileEstoqueShop) {
      alert('O arquivo principal "Estoque Filiais (Shop9)" é obrigatório!');
      return;
    }
    if (!dataEnvio) {
      alert('Selecione a data de referência!');
      return;
    }

    setLoading(true);

    try {
      // 1. Processamento de Estoque E-commerce / Tiny (Coluna A = Grade/SKU, Coluna C = Estoque)
      const mapaEstoqueEc: { [grade: string]: number } = {};
      if (fileEstoqueTiny) {
        setStatusUpload(`Lendo saldos de estoque E-commerce (${fileEstoqueTiny.name})...`);
        const bufTiny = await fileEstoqueTiny.arrayBuffer();
        const wbTiny = XLSX.read(bufTiny, { type: 'array' });
        const shTiny = wbTiny.Sheets[wbTiny.SheetNames[0]];
        const rowsTiny: any[] = XLSX.utils.sheet_to_json(shTiny, { header: 1 }); // Array of arrays para pegar por índice de coluna (A=0, C=2)

        rowsTiny.forEach((r: any[]) => {
          if (r && r.length >= 3) {
            const gradeSku = String(r[0] || '').trim(); // Coluna A (SKU/Grade)
            const qtdEstoqueEc = Number(r[2] || 0);     // Coluna C (Unidades de Estoque)
            if (gradeSku) {
              mapaEstoqueEc[gradeSku] = (mapaEstoqueEc[gradeSku] || 0) + qtdEstoqueEc;
            }
          }
        });
      }

      // 2. Processamento de Vendas (Mapeando Colunas I, J, K / Qtde_Vendida / Preco_Total / Preco_Médio e Média E-commerce)
      const mapaVendas: { [key: string]: { qtd: number, valor: number, somaPdv: number, countPdv: number } } = {};
      const arquivosVendas = [fileVendasShop, fileVendasTiny].filter(Boolean) as File[];

      for (const arquivoVenda of arquivosVendas) {
        setStatusUpload(`Lendo relatório de vendas (${arquivoVenda.name})...`);
        const bufVendas = await arquivoVenda.arrayBuffer();
        const wbVendas = XLSX.read(bufVendas, { type: 'array' });
        const shVendas = wbVendas.Sheets[wbVendas.SheetNames[0]];
        const rowsVendas: any[] = XLSX.utils.sheet_to_json(shVendas);

        rowsVendas.forEach((rv, idx) => {
          const f = Number(rv['Filial'] || rv['Depósito'] || 3);
          const g = String(rv['Codigo_da_Grade'] || rv['Grade'] || rv['Código'] || `1000${idx}`);
          const k = `${g}-${f}`;
          
          const qtd = Number(rv['Qtde_Vendida'] || rv['Quantidade'] || 0);
          const val = Number(rv['Preco_Total'] || rv['Valor_Total'] || rv['Valor'] || 0);
          const pdvUnitario = Number(rv['Preco_Médio'] || rv['Preço_Unitário'] || (qtd > 0 ? val / qtd : 0));

          if (!mapaVendas[k]) mapaVendas[k] = { qtd: 0, valor: 0, somaPdv: 0, countPdv: 0 };
          mapaVendas[k].qtd += qtd;
          mapaVendas[k].valor += val;
          if (pdvUnitario > 0) {
            mapaVendas[k].somaPdv += pdvUnitario;
            mapaVendas[k].countPdv += 1;
          }
        });
      }

      // 3. Processamento de Estoque Principal Filiais (Shop9) + Soma do Estoque E-commerce (Tiny)
      setStatusUpload(`Processando base principal de estoque (${fileEstoqueShop.name})...`);
      const bufferEstoque = await fileEstoqueShop.arrayBuffer();
      const wbEstoque = XLSX.read(bufferEstoque, { type: 'array' });
      const sheetEstoque = wbEstoque.Sheets[wbEstoque.SheetNames[0]];
      const rowsEstoque: any[] = XLSX.utils.sheet_to_json(sheetEstoque);

      const registrosProcessados = rowsEstoque.map((row, idx) => {
        const filial = Number(row['Filial'] || row['Depósito'] || 3);
        const grade = String(row['Grade'] || row['Código_da_Grade'] || `1000${idx}`);
        const gradeFilial = `${grade}-${filial}`;
        const codigo = Number(row['Codigo'] || row['Código_Produto'] || 1000 + idx);
        const produto = String(row['Produto'] || row['Nome_Produto'] || 'PRODUTO NÃO CADASTRADO');
        const cor = String(row['Cor'] || row['Nome_Cor'] || 'SEM COR');
        const tamanho = String(row['Tamanho'] || row['Nome_Tamanho'] || 'UNICO');
        
        const genero = String(row['Genero'] || row['Gênero'] || 'Não Cadastrado');
        const marca = String(row['Marca'] || 'Não Cadastrado');
        const classe = String(row['Classe'] || 'NÃO CADASTRADO');
        const familia = String(row['Familia'] || row['Família'] || 'Não Cadastrado');
        const fornecedor = String(row['Fornecedor'] || 'NÃO CADASTRADO');
        
        // Estoque Loja + Estoque E-commerce (Tiny coluna C com base na Grade SKU)
        const estoqueLoja = Number(row['Estoque'] || row['Qtde_Estoque_Atual'] || row['Saldo'] || 0);
        const estoqueEc = mapaEstoqueEc[grade] || 0;
        const estoqueTotal = estoqueLoja + estoqueEc;

        const custo = Number(row['Preco'] || row['Custo'] || 0);
        
        const vendaInfo = mapaVendas[gradeFilial] || { qtd: 0, valor: 0, somaPdv: 0, countPdv: 0 };
        const mediaPdvEc = vendaInfo.countPdv > 0 ? vendaInfo.somaPdv / vendaInfo.countPdv : 0;

        const pdvEc = Number(row['Preco1'] || row['Preço'] || (mediaPdvEc > 0 ? mediaPdvEc : custo * 1.9));
        const pdvLoja = Number(row['Preco2'] || (pdvEc * 1.05));
        const margemEc = pdvEc > 0 ? (pdvEc - custo) / pdvEc : 0;
        const margemLoja = pdvLoja > 0 ? (pdvLoja - custo) / pdvLoja : 0;
        const custoTotal = Number((estoqueTotal * custo).toFixed(2));

        return {
          data_referencia: dataEnvio,
          grade_filial: gradeFilial,
          grade,
          codigo,
          produto,
          filial,
          cor,
          tamanho,
          genero,
          marca,
          classe,
          familia,
          fornecedor,
          estoque: estoqueTotal,
          custo,
          custo_total: custoTotal,
          margem_loja: margemLoja,
          pdv_loja: pdvLoja,
          margem_ec: margemEc,
          pdv_ec: pdvEc,
          venda_geral_itens: vendaInfo.qtd,
          venda_geral_valor: Number(vendaInfo.valor.toFixed(2)),
        };
      });

      // 4. Envio seguro em lotes para o Supabase
      const tamanhoLote = 300;
      const totalLotes = Math.ceil(registrosProcessados.length / tamanhoLote);

      for (let i = 0; i < registrosProcessados.length; i += tamanhoLote) {
        const lote = registrosProcessados.slice(i, i + tamanhoLote);
        const numeroLote = Math.floor(i / tamanhoLote) + 1;
        setStatusUpload(`Enviando lote ${numeroLote} de ${totalLotes}...`);

        let tentativas = 0;
        let sucesso = false;

        while (tentativas < 3 && !sucesso) {
          try {
            const { error } = await supabase
              .from('produtos_compras')
              .upsert(lote, { onConflict: 'data_referencia,grade_filial' });

            if (error) throw error;
            sucesso = true;
          } catch (err) {
            tentativas++;
            if (tentativas >= 3) throw err;
            await new Promise(res => setTimeout(res, 1000));
          }
        }

        await new Promise(res => setTimeout(res, 200));
      }

      setStatusUpload('Base consolidada e cruzamento de estoque/vendas concluídos com sucesso!');
      setLoading(false);
      setFileEstoqueShop(null);
      setFileVendasShop(null);
      setFileEstoqueTiny(null);
      setFileVendasTiny(null);
      await carregarDadosDoSupabase();
      setActiveTab('dashboard');
    } catch (err: any) {
      console.error('Erro geral no upload:', err);
      setLoading(false);
      setStatusUpload('Erro ao processar os arquivos.');
      alert('Erro inesperado: ' + err.message);
    }
  };

  const apagarBaseInteira = async () => {
    if (!confirm('Tem certeza que deseja apagar TODOS os registros da base do Supabase?')) return;
    setLoading(true);
    const { error } = await supabase.from('produtos_compras').delete().neq('id', 0);
    if (error) alert('Erro ao apagar: ' + error.message);
    else {
      alert('Base apagada com sucesso!');
      setDados([]);
    }
    setLoading(false);
  };

  const apagarPorData = async () => {
    if (filtroData === 'TODAS') {
      alert('Selecione uma data específica no filtro para apagar.');
      return;
    }
    if (!confirm(`Deseja apagar todos os registros da data ${filtroData}?`)) return;
    setLoading(true);
    const { error } = await supabase.from('produtos_compras').delete().eq('data_referencia', filtroData);
    if (error) alert('Erro ao apagar data: ' + error.message);
    else {
      alert(`Registros de ${filtroData} apagados com sucesso!`);
      await carregarDadosDoSupabase();
    }
    setLoading(false);
  };

  const datasUnicas = useMemo(() => ['TODAS', ...Array.from(new Set(dados.map(d => d.dataReferencia)))], [dados]);
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
        (filtroData === 'TODAS' || item.dataReferencia === filtroData) &&
        (filtroMarca === 'TODAS' || item.marca === filtroMarca) &&
        (filtroFilial === 'TODAS' || String(item.filial) === filtroFilial) &&
        (filtroClasse === 'TODAS' || item.classe === filtroClasse)
      );
    });
  }, [dados, filtroBusca, filtroData, filtroMarca, filtroFilial, filtroClasse]);

  // Totais superiores do mês (acumulados na linha superior do dashboard)
  const totaisGeraisMes = useMemo(() => {
    const totalEstoque = dadosFiltrados.reduce((acc, item) => acc + item.estoque, 0);
    const totalCusto = dadosFiltrados.reduce((acc, item) => acc + item.custoTotal, 0);
    const totalVendasItens = dadosFiltrados.reduce((acc, item) => acc + item.vendaGeralItens, 0);
    const totalVendasValor = dadosFiltrados.reduce((acc, item) => acc + item.vendaGeralValor, 0);
    return { totalEstoque, totalCusto, totalVendasItens, totalVendasValor };
  }, [dadosFiltrados]);

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
      else if (metricaDinamica === 'vendaGeralValor') matriz[linhaVal][colunaVal] += item.vendaGeralValor;
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
              <Upload className="w-4 h-4" /> 1. Upload & Data
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
          <form onSubmit={processarEEnviar} className="space-y-8 max-w-4xl mx-auto animate-in fade-in duration-300">
            <div>
              <h1 className="text-3xl font-bold tracking-tight text-white">Upload de Relatórios & Histórico</h1>
              <p className="text-slate-400 mt-1">Selecione a data, envie os arquivos de loja e e-commerce para cruzar o estoque.</p>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl p-6 space-y-6">
              <div>
                <label className="block text-xs font-medium text-slate-400 mb-2 uppercase tracking-wider flex items-center gap-1.5">
                  <Calendar className="w-4 h-4 text-indigo-400" /> Data de Referência dos Dados
                </label>
                <input
                  type="date"
                  value={dataEnvio}
                  onChange={(e) => setDataEnvio(e.target.value)}
                  className="w-full sm:w-1/3 bg-slate-950 border border-slate-800 rounded-lg p-3 text-sm text-white focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
                <div className="bg-slate-950 border border-slate-800 p-4 rounded-xl">
                  <div className="text-xs text-indigo-400 font-semibold uppercase mb-1">Fonte 01 (Shop9)</div>
                  <h4 className="text-sm font-bold text-white mb-3">Estoque Filiais</h4>
                  <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-3 text-center cursor-pointer bg-slate-900 block truncate">
                    <span className="text-xs text-slate-300">{fileEstoqueShop ? fileEstoqueShop.name : 'Selecionar arquivo (.xls/.txt)'}</span>
                    <input type="file" accept=".xls,.xlsx,.txt,.csv" onChange={(e) => setFileEstoqueShop(e.target.files?.[0] || null)} className="hidden" />
                  </label>
                </div>

                <div className="bg-slate-950 border border-slate-800 p-4 rounded-xl">
                  <div className="text-xs text-indigo-400 font-semibold uppercase mb-1">Fonte 02 (Shop9)</div>
                  <h4 className="text-sm font-bold text-white mb-3">Vendas Filiais</h4>
                  <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-3 text-center cursor-pointer bg-slate-900 block truncate">
                    <span className="text-xs text-slate-300">{fileVendasShop ? fileVendasShop.name : 'Selecionar arquivo (.xls/.txt)'}</span>
                    <input type="file" accept=".xls,.xlsx,.txt,.csv" onChange={(e) => setFileVendasShop(e.target.files?.[0] || null)} className="hidden" />
                  </label>
                </div>

                <div className="bg-slate-950 border border-slate-800 p-4 rounded-xl">
                  <div className="text-xs text-indigo-400 font-semibold uppercase mb-1">Fonte 03 (Tiny ERP)</div>
                  <h4 className="text-sm font-bold text-white mb-3">Saldos Estoque E-c (Coluna A & C)</h4>
                  <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-3 text-center cursor-pointer bg-slate-900 block truncate">
                    <span className="text-xs text-slate-300">{fileEstoqueTiny ? fileEstoqueTiny.name : 'Selecionar arquivo (.xls)'}</span>
                    <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setFileEstoqueTiny(e.target.files?.[0] || null)} className="hidden" />
                  </label>
                </div>

                <div className="bg-slate-950 border border-slate-800 p-4 rounded-xl">
                  <div className="text-xs text-indigo-400 font-semibold uppercase mb-1">Fonte 04 (Tiny ERP)</div>
                  <h4 className="text-sm font-bold text-white mb-3">Relatório Vendas</h4>
                  <label className="border border-dashed border-slate-700 hover:border-indigo-500 rounded-lg p-3 text-center cursor-pointer bg-slate-900 block truncate">
                    <span className="text-xs text-slate-300">{fileVendasTiny ? fileVendasTiny.name : 'Selecionar arquivo (.xls)'}</span>
                    <input type="file" accept=".xls,.xlsx,.csv" onChange={(e) => setFileVendasTiny(e.target.files?.[0] || null)} className="hidden" />
                  </label>
                </div>
              </div>

              <div className="bg-slate-950 border border-slate-800 rounded-xl p-4 flex flex-col sm:flex-row items-center justify-between gap-4">
                <div>
                  <h4 className="font-medium text-white text-xs uppercase tracking-wider text-indigo-400">Status</h4>
                  <p className="text-sm text-slate-300">{statusUpload}</p>
                </div>
                <button
                  type="submit"
                  disabled={loading}
                  className="w-full sm:w-auto px-6 py-3 bg-indigo-600 hover:bg-indigo-500 text-white font-medium rounded-xl transition-all shadow-lg flex items-center justify-center gap-2 disabled:opacity-50 text-sm"
                >
                  {loading && <Loader2 className="w-4 h-4 animate-spin" />}
                  {loading ? 'Processando...' : 'Processar & Enviar'} <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          </form>
        ) : activeTab === 'dashboard' ? (
          <div className="space-y-6 animate-in fade-in duration-300">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div>
                <h1 className="text-3xl font-bold tracking-tight text-white">Base Consolidada (Supabase)</h1>
                <p className="text-slate-400 mt-1">Gerencie, filtre por data ou limpe registros indesejados.</p>
              </div>
              <div className="flex flex-wrap items-center gap-3">
                <button 
                  onClick={apagarPorData}
                  className="px-4 py-2 bg-amber-600/20 border border-amber-500/30 hover:bg-amber-600/30 text-amber-300 text-sm font-medium rounded-xl flex items-center gap-2"
                >
                  <Trash2 className="w-4 h-4" /> Apagar Data Selecionada
                </button>
                <button 
                  onClick={apagarBaseInteira}
                  className="px-4 py-2 bg-rose-600/20 border border-rose-500/30 hover:bg-rose-600/30 text-rose-300 text-sm font-medium rounded-xl flex items-center gap-2"
                >
                  <Trash2 className="w-4 h-4" /> Apagar Base Inteira
                </button>
              </div>
            </div>

            {/* Linha Superior: Totais de Vendas e Estoque no Mês */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-lg">
                <span className="text-xs text-slate-400 font-medium uppercase">Total Estoque (Loja + E-c)</span>
                <div className="text-2xl font-bold text-white mt-1">{totaisGeraisMes.totalEstoque.toLocaleString('pt-BR')} un</div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-lg">
                <span className="text-xs text-slate-400 font-medium uppercase">Total Custo Estoque</span>
                <div className="text-2xl font-bold text-indigo-400 mt-1">R$ {totaisGeraisMes.totalCusto.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-lg">
                <span className="text-xs text-slate-400 font-medium uppercase">Total Vendas (Itens no Mês)</span>
                <div className="text-2xl font-bold text-sky-400 mt-1">{totaisGeraisMes.totalVendasItens.toLocaleString('pt-BR')} un</div>
              </div>
              <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-lg">
                <span className="text-xs text-slate-400 font-medium uppercase">Total Vendas (Valor no Mês)</span>
                <div className="text-2xl font-bold text-emerald-400 mt-1">R$ {totaisGeraisMes.totalVendasValor.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
              </div>
            </div>

            <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-lg grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
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
                <label className="block text-xs font-medium text-slate-400 mb-1.5">Filtrar por Data</label>
                <select value={filtroData} onChange={(e) => setFiltroData(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-sm text-white focus:outline-none">
                  {datasUnicas.map((d, idx) => <option key={idx} value={d}>{d}</option>)}
                </select>
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
                <h3 className="text-sm font-semibold text-white">Registros Exibidos ({dadosFiltrados.length})</h3>
                <span className="text-xs text-emerald-400 font-medium flex items-center gap-1.5">
                  <CheckCircle2 className="w-3.5 h-3.5" /> Sincronizado
                </span>
              </div>
              <div className="overflow-x-auto max-h-[550px]">
                <table className="w-full text-left border-collapse text-xs whitespace-nowrap">
                  <thead className="bg-slate-950/90 text-slate-400 sticky top-0 border-b border-slate-800 z-10">
                    <tr>
                      <th className="p-3">Data Ref.</th>
                      <th className="p-3">Grade + filial</th>
                      <th className="p-3">Código</th>
                      <th className="p-3">Produto</th>
                      <th className="p-3">Filial</th>
                      <th className="p-3">Cor</th>
                      <th className="p-3">Tamanho</th>
                      <th className="p-3">Gênero</th>
                      <th className="p-3">Marca</th>
                      <th className="p-3">Classe</th>
                      <th className="p-3">Estoque (Loja+E-c)</th>
                      <th className="p-3">Custo Total</th>
                      <th className="p-3">PDV E-c</th>
                      <th className="p-3">Vendas (Itens)</th>
                      <th className="p-3">Vendas (Valor)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/50">
                    {dadosFiltrados.map((item, index) => (
                      <tr key={index} className="hover:bg-slate-800/30 transition-colors">
                        <td className="p-3 font-mono text-indigo-400">{item.dataReferencia}</td>
                        <td className="p-3 font-mono">{item.gradeFilial}</td>
                        <td className="p-3 font-mono">{item.codigo}</td>
                        <td className="p-3 font-medium text-white">{item.produto}</td>
                        <td className="p-3">{item.filial}</td>
                        <td className="p-3 text-slate-300">{item.cor}</td>
                        <td className="p-3 text-slate-300">{item.tamanho}</td>
                        <td className="p-3 text-slate-400">{item.genero}</td>
                        <td className="p-3 font-semibold text-slate-200">{item.marca}</td>
                        <td className="p-3 text-slate-400">{item.classe}</td>
                        <td className="p-3 font-bold text-white">{item.estoque} un</td>
                        <td className="p-3 text-slate-300">R$ {item.custoTotal.toFixed(2)}</td>
                        <td className="p-3 text-slate-300">R$ {item.pdvEc.toFixed(2)}</td>
                        <td className="p-3 font-bold text-indigo-300">{item.vendaGeralItens} un</td>
                        <td className="p-3 font-bold text-emerald-400">R$ {item.vendaGeralValor.toFixed(2)}</td>
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
                <p className="text-slate-400 mt-1">Cruzamento matricial por eixos (Linhas x Colunas) em tempo real.</p>
              </div>
              <div className="flex flex-wrap items-center gap-3 bg-slate-900 border border-slate-800 p-3 rounded-xl shadow-lg">
                <div>
                  <label className="block text-[10px] uppercase font-semibold text-slate-400 mb-1">Rótulos de Linha:</label>
                  <select value={linhaDinamica} onChange={(e) => setLinhaDinamica(e.target.value as any)} className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none">
                    <option value="marca">Marca</option>
                    <option value="classe">Classe</option>
                    <option value="familia">Família</option>
                    <option value="filial">Filial</option>
                    <option value="genero">Gênero</option>
                    <option value="cor">Cor</option>
                    <option value="tamanho">Tamanho</option>
                    <option value="fornecedor">Fornecedor</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] uppercase font-semibold text-slate-400 mb-1">Rótulos de Coluna:</label>
                  <select value={colunaDinamica} onChange={(e) => setColunaDinamica(e.target.value as any)} className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none">
                    <option value="filial">Filial</option>
                    <option value="marca">Marca</option>
                    <option value="classe">Classe</option>
                    <option value="genero">Gênero</option>
                    <option value="tamanho">Tamanho</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] uppercase font-semibold text-slate-400 mb-1">Valores (Métrica):</label>
                  <select value={metricaDinamica} onChange={(e) => setMetricaDinamica(e.target.value as any)} className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-white focus:outline-none">
                    <option value="estoque">Soma de Estoque (Loja + E-c)</option>
                    <option value="custoTotal">Soma de Custo Total</option>
                    <option value="pdvEc">Soma de PDV E-commerce</option>
                    <option value="vendaGeralValor">Soma de Valor de Vendas</option>
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
                          Nenhum dado disponível. Faça o upload na primeira aba.
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
                                  {metricaDinamica.includes('custo') || metricaDinamica.includes('pdv') || metricaDinamica.includes('Valor')
                                    ? `R$ ${val.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`
                                    : `${val.toLocaleString('pt-BR')} un`}
                                </td>
                              );
                            })}
                            <td className="p-4 text-center font-bold text-emerald-400">
                              {metricaDinamica.includes('custo') || metricaDinamica.includes('pdv') || metricaDinamica.includes('Valor')
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