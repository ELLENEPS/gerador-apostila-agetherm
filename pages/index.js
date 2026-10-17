import { useState, useRef, useEffect } from 'react';
import Head from 'next/head';
import * as XLSX from 'xlsx';
import bwipjs from 'bwip-js';

export default function GeradorApostila() {
    const [tipoApostila, setTipoApostila] = useState('BARRAS'); // 'BARRAS' ou 'OUTLET'
    const [paginasA4, setPaginasA4] = useState([]);
    const [paginasIndice, setPaginasIndice] = useState([]);
    const [totalItens, setTotalItens] = useState(0);
    const [carregando, setCarregando] = useState(false);
    const [nomeArquivo, setNomeArquivo] = useState('');
    const [tituloApostila, setTituloApostila] = useState('CATÁLOGO DE CÓDIGOS DE BARRAS');

    const handleTrocaTipo = (novoTipo) => {
        setTipoApostila(novoTipo);
        setPaginasA4([]);
        setPaginasIndice([]);
        setTotalItens(0);
        setNomeArquivo('');
        if (novoTipo === 'OUTLET') {
            setTituloApostila('OUTLET');
        } else {
            setTituloApostila('CATÁLOGO DE CÓDIGOS DE BARRAS');
        }
    };

    const handleDownloadModelo = () => {
        let dadosModelo = [];
        let filename = '';

        if (tipoApostila === 'OUTLET') {
            dadosModelo = [
                { 'CATEGORIA': 'CAPACITORES', 'FOTO PRODUTO': 'AGT-CPD', 'CÓDIGO': 'AGT-CPD10', 'DESCRIÇÃO': 'CAPACITOR DE PARTIDA (DUPLO) CBB65 380V - 30 + 4 µF', 'VALOR UN COM 5% DESC': 14.31, 'PREÇO COM 45% DESCONTO': 7.87 },
                { 'CATEGORIA': 'CAPACITORES', 'FOTO PRODUTO': '', 'CÓDIGO': 'AGT-CPD15', 'DESCRIÇÃO': 'CAPACITOR DE PARTIDA (DUPLO) CBB65 380V - 40 + 5 µF', 'VALOR UN COM 5% DESC': 18.03, 'PREÇO COM 45% DESCONTO': 9.92 },
                { 'CATEGORIA': 'FERRAMENTAS', 'FOTO PRODUTO': 'AGT-CT102', 'CÓDIGO': 'AGT-CT102-10', 'DESCRIÇÃO': 'MOLA CURVADORA DE TUBOS 5/8"', 'VALOR UN COM 5% DESC': 25.50, 'PREÇO COM 45% DESCONTO': 14.02 }
            ];
            filename = 'modelo_apostila_outlet.xlsx';
        } else {
            dadosModelo = [
                { 'CATEGORIA': 'CONEXÕES', 'REFERÊNCIA': 'REF-001', 'MÚLTIPLO VENDA': '1 PCT', 'DESCRIÇÃO': 'JOELHO 90 GRAUS 22MM PARA TUBULAÇÃO DE COBRE E SOLDA', 'CÓDIGO DE BARRAS': '7891234567890' },
                { 'CATEGORIA': 'CONEXÕES', 'REFERÊNCIA': 'REF-002', 'MÚLTIPLO VENDA': '5 UN', 'DESCRIÇÃO': 'LUVA REDUÇÃO 28X22MM', 'CÓDIGO DE BARRAS': '7891234567891' }
            ];
            filename = 'modelo_apostila_barras.xlsx';
        }

        const worksheet = XLSX.utils.json_to_sheet(dadosModelo);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Modelo');
        XLSX.writeFile(workbook, filename);
    };

    const handleFileUpload = (e) => {
        const file = e.target.files[0];
        if (!file) return;

        setNomeArquivo(file.name);
        setCarregando(true);

        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const data = new Uint8Array(evt.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

                const gruposMap = new Map();
                let contadorTotal = 0;
                let ultimaFotoValida = '';

                jsonData.forEach((row) => {
                    if (tipoApostila === 'OUTLET') {
                        let fotoRaw = String(row['FOTO PRODUTO'] || row['Foto Produto'] || row['FOTO'] || row['Foto'] || '').trim();
                        if (fotoRaw) {
                            ultimaFotoValida = fotoRaw;
                        }
                        const fotoFinal = ultimaFotoValida;

                        const categoria = String(row['CATEGORIA'] || row['Categoria'] || 'GERAL').trim().toUpperCase();
                        const codigo = String(row['CÓDIGO'] || row['CODIGO'] || row['Código'] || row['Codigo'] || row['REFERÊNCIA'] || '').trim();
                        const descricao = String(row['DESCRIÇÃO'] || row['DESCRICAO'] || row['Descrição'] || row['Descricao'] || '').trim();
                        
                        const precoOriginal = parseFloat(row['VALOR UN COM 5% DESC'] || row['VALOR UN COM 5% DESCONTO'] || row['VALOR UNITÁRIO 5% DESC'] || row['VALOR UNITARIO'] || 0);
                        const precoDesconto = parseFloat(row['PREÇO COM 45% DESCONTO'] || row['PRECO COM 45% DESCONTO'] || 0);

                        if (!codigo && !descricao) return;

                        if (!gruposMap.has(categoria)) {
                            gruposMap.set(categoria, []);
                        }

                        gruposMap.get(categoria).push({
                            categoria,
                            foto: fotoFinal,
                            codigo,
                            descricao,
                            precoOriginal,
                            precoDesconto
                        });

                        contadorTotal++;
                    } else {
                        const categoria = String(row['CATEGORIA'] || row['Categoria'] || 'GERAL').trim().toUpperCase();
                        const codigo = String(row['REFERÊNCIA'] || row['REFERENCIA'] || row['Referência'] || row['Referencia'] || row['Código'] || 'SEM REFERÊNCIA').trim();
                        const multiplo = String(row['MÚLTIPLO VENDA'] || row['MULTIPLO VENDA'] || row['Múltiplo Venda'] || row['MÚLTIPLO'] || '').trim();
                        const descricao = String(row['DESCRIÇÃO'] || row['DESCRICAO'] || row['Descrição'] || '').trim();
                        const codBarras = String(row['CÓDIGO DE BARRAS'] || row['CODIGO DE BARRAS'] || row['Código de Barras'] || row['EAN'] || codigo).trim();

                        if (!gruposMap.has(categoria)) {
                            gruposMap.set(categoria, []);
                        }

                        gruposMap.get(categoria).push({
                            categoria,
                            codigo,
                            multiplo,
                            descricao,
                            codBarras
                        });

                        contadorTotal++;
                    }
                });

                if (tipoApostila === 'OUTLET') {
                    const listagemLinear = [];
                    
                    gruposMap.forEach((itensCategoria, nomeCategoria) => {
                        itensCategoria.sort((a, b) => a.codigo.localeCompare(b.codigo, undefined, { numeric: true, sensitivity: 'base' }));

                        itensCategoria.forEach((item, idx) => {
                            listagemLinear.push({
                                ...item,
                                ehInicioCategoria: idx === 0
                            });
                        });
                    });

                    // 1. Calcula quantidade de categorias e monta as páginas do Índice (até 18 itens por página A4)
                    const categoriasLista = Array.from(gruposMap.keys());
                    const CATEGORIAS_POR_PAGINA_INDICE = 18;
                    const qtdPaginasIndice = Math.ceil(categoriasLista.length / CATEGORIAS_POR_PAGINA_INDICE) || 1;

                    // Página 1 = Capa | Índice = Páginas seguintes | Produtos começam após o índice
                    const numeroPaginaInicialProdutos = 1 + qtdPaginasIndice + 1;

                    const paginasProdutosTemp = [];
                    let paginaAtual = [];
                    let pesoAtual = 0;
                    let numPagina = numeroPaginaInicialProdutos;

                    // LIMITE EXPANDIDO PARA PREENCHER A FOLHA A4 QUASE 100%
                    const LIMITE_PESO_PAGINA = 12.5;
                    const categoriasVistasEmPagina = new Map();
                    const resumoIndice = [];

                    listagemLinear.forEach((item) => {
                        const pesoItem = item.ehInicioCategoria ? 1.3 : 1.0;

                        if (pesoAtual + pesoItem > LIMITE_PESO_PAGINA && paginaAtual.length > 0) {
                            paginasProdutosTemp.push({
                                itens: paginaAtual,
                                numPagina: numPagina
                            });
                            paginaAtual = [];
                            pesoAtual = 0;
                            numPagina++;
                        }

                        if (item.ehInicioCategoria && !categoriasVistasEmPagina.has(item.categoria)) {
                            categoriasVistasEmPagina.set(item.categoria, numPagina);
                            resumoIndice.push({
                                categoria: item.categoria,
                                paginaInicial: numPagina
                            });
                        }

                        paginaAtual.push(item);
                        pesoAtual += pesoItem;
                    });

                    if (paginaAtual.length > 0) {
                        paginasProdutosTemp.push({
                            itens: paginaAtual,
                            numPagina: numPagina
                        });
                    }

                    // Monta as páginas do Índice
                    const paginasIndiceMontadas = [];
                    for (let i = 0; i < resumoIndice.length; i += CATEGORIAS_POR_PAGINA_INDICE) {
                        paginasIndiceMontadas.push({
                            itens: resumoIndice.slice(i, i + CATEGORIAS_POR_PAGINA_INDICE),
                            numPagina: 2 + paginasIndiceMontadas.length
                        });
                    }

                    if (paginasIndiceMontadas.length === 0) {
                        paginasIndiceMontadas.push({ itens: [], numPagina: 2 });
                    }

                    setPaginasA4(paginasProdutosTemp);
                    setPaginasIndice(paginasIndiceMontadas);
                } else {
                    const paginasMontadas = [];
                    let numPagina = 1;

                    gruposMap.forEach((itensCategoria, nomeCategoria) => {
                        let indiceItemLocal = 1;
                        for (let i = 0; i < itensCategoria.length; i += 12) {
                            const fatiaItens = itensCategoria.slice(i, i + 12).map((item) => ({
                                ...item,
                                numeroSequencial: indiceItemLocal++
                            }));

                            paginasMontadas.push({
                                categoria: nomeCategoria,
                                itens: fatiaItens,
                                numPagina: numPagina++
                            });
                        }
                    });

                    setPaginasA4(paginasMontadas);
                }

                setTotalItens(contadorTotal);
            } catch (err) {
                alert('Erro ao ler a planilha. Certifique-se de que é um arquivo Excel ou CSV válido.');
            } finally {
                setCarregando(false);
            }
        };

        reader.readAsArrayBuffer(file);
    };

    const BarcodeImage = ({ text }) => {
        const canvasRef = useRef(null);

        useEffect(() => {
            if (canvasRef.current && text) {
                try {
                    bwipjs.toCanvas(canvasRef.current, {
                        bcid: 'code128',
                        text: text,
                        scale: 2,
                        height: 7.5,
                        includetext: true,
                        textxalign: 'center',
                        textsize: 8.5,
                        textcolor: '000000',
                    });
                } catch (e) {
                    console.error('Erro ao gerar código de barras:', e);
                }
            }
        }, [text]);

        return <canvas ref={canvasRef} className="max-w-full h-auto" />;
    };

    const handleImprimir = () => {
        window.print();
    };

    const formatarMoeda = (valor) => {
        if (isNaN(valor)) return 'R$ 0,00';
        return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    };

    const dataAtualExtenso = new Date().toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' }).toUpperCase();
    const totalPaginasPDF = tipoApostila === 'OUTLET' ? 1 + paginasIndice.length + paginasA4.length : paginasA4.length;

    return (
        <div className="min-h-screen bg-slate-100 py-8 px-4 font-sans text-slate-800 print:bg-white print:py-0 print:px-0">
            <Head>
                <title>Gerador de Apostila - AGETHERM</title>
                <link rel="icon" href="/favicon.ico" />
                <link rel="shortcut icon" href="/logo.png" />
            </Head>

            {/* PAINEL DE CONTROLE DE CONFIGURAÇÕES */}
            <div className="max-w-4xl mx-auto bg-white rounded-xl shadow-md border-t-4 border-[#FF8C00] overflow-hidden mb-10 print:hidden">
                
                <div className="p-6 md:p-8 border-b border-slate-100 flex flex-col md:flex-row items-center justify-between gap-4">
                    <div>
                        <h1 className="text-xl md:text-2xl font-black text-[#002B5B] uppercase tracking-tight">
                            Gerador de Apostilas
                        </h1>
                        <p className="text-xs text-slate-500 font-medium">Selecione o modelo e importe sua planilha</p>
                    </div>
                    <div className="h-12 flex items-center justify-center">
                        <img src="/logo.png" alt="Agetherm Logo" className="max-h-full w-auto object-contain" />
                    </div>
                </div>

                <div className="p-6 md:p-8 space-y-6">

                    {/* SELETOR DE MODELO */}
                    <div>
                        <label className="block text-xs font-bold text-[#002B5B] uppercase mb-2 tracking-wider">
                            Modelo de Apostila
                        </label>
                        <div className="grid grid-cols-2 gap-3">
                            <button
                                type="button"
                                onClick={() => handleTrocaTipo('BARRAS')}
                                className={`py-3 px-4 rounded-xl font-bold text-xs uppercase tracking-wider border-2 transition-all flex items-center justify-center gap-2 ${
                                    tipoApostila === 'BARRAS'
                                        ? 'bg-[#002B5B] text-white border-[#002B5B] shadow-md'
                                        : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                                }`}
                            >
                                Código de Barras
                            </button>

                            <button
                                type="button"
                                onClick={() => handleTrocaTipo('OUTLET')}
                                className={`py-3 px-4 rounded-xl font-bold text-xs uppercase tracking-wider border-2 transition-all flex items-center justify-center gap-2 ${
                                    tipoApostila === 'OUTLET'
                                        ? 'bg-[#FF8C00] text-white border-[#FF8C00] shadow-md'
                                        : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                                }`}
                            >
                                Outlet (Promoções 45% OFF)
                            </button>
                        </div>
                    </div>

                    <div>
                        <label className="block text-xs font-bold text-[#002B5B] uppercase mb-1.5 tracking-wider">
                            Título do Documento
                        </label>
                        <input
                            type="text"
                            value={tituloApostila}
                            onChange={(e) => setTituloApostila(e.target.value)}
                            className="w-full text-slate-800 px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:ring-2 focus:ring-[#FF8C00] focus:border-transparent outline-none text-sm transition-all font-medium"
                        />
                    </div>

                    {/* BLOCO DE IMPORTAÇÃO */}
                    <div className="bg-slate-50 rounded-xl p-6 border border-slate-200/80 text-left space-y-4">
                        <div>
                            <h2 className="text-sm font-bold text-[#002B5B] uppercase tracking-wide">
                                Importar Planilha do {tipoApostila === 'OUTLET' ? 'OUTLET' : 'Catálogo'}
                            </h2>
                            <p className="text-xs text-slate-500 mt-1">
                                {tipoApostila === 'OUTLET'
                                    ? 'Envie a planilha contendo CATEGORIA, FOTO PRODUTO, CÓDIGO, DESCRIÇÃO, VALOR UN COM 5% DESC e PREÇO COM 45% DESCONTO.'
                                    : 'Envie a planilha contendo CATEGORIA, REFERÊNCIA, MÚLTIPLO VENDA, DESCRIÇÃO e CÓDIGO DE BARRAS.'}
                            </p>
                        </div>

                        <div className="flex flex-col sm:flex-row items-center justify-start gap-3 pt-2">
                            <label className="w-full sm:w-auto cursor-pointer bg-[#002B5B] text-white px-5 py-2.5 rounded-lg font-bold text-xs uppercase tracking-wider hover:bg-[#001d3d] active:scale-95 transition-all shadow-sm flex items-center justify-center gap-2">
                                <span>{nomeArquivo ? 'Trocar Planilha' : 'Escolher Planilha'}</span>
                                <input
                                    type="file"
                                    accept=".xlsx, .xls, .csv"
                                    onChange={handleFileUpload}
                                    className="hidden"
                                />
                            </label>

                            <button
                                type="button"
                                onClick={handleDownloadModelo}
                                className="w-full sm:w-auto bg-white text-slate-700 border border-slate-300 px-5 py-2.5 rounded-lg font-bold text-xs uppercase tracking-wider hover:bg-slate-100 active:scale-95 transition-all shadow-sm flex items-center justify-center gap-2"
                            >
                                <span>Baixar Modelo (.XLSX)</span>
                            </button>
                        </div>

                        {nomeArquivo && (
                            <div className="inline-flex items-center gap-2 text-xs font-semibold text-emerald-700 bg-emerald-50 px-3 py-1 rounded-md border border-emerald-200 mt-2">
                                <span>{nomeArquivo}</span>
                            </div>
                        )}
                    </div>

                    <div className="pt-4 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4">
                        <div className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                            Total de Itens: <span className="text-base font-black text-[#002B5B] ml-1">{totalItens}</span> | Páginas A4: <span className="text-base font-black text-[#FF8C00] ml-1">{totalPaginasPDF}</span>
                        </div>

                        <button
                            onClick={handleImprimir}
                            disabled={totalItens === 0}
                            className="w-full sm:w-auto bg-[#FF8C00] text-white px-7 py-3 rounded-lg font-bold text-xs uppercase tracking-wider hover:bg-[#e67e00] active:scale-95 transition-all disabled:bg-slate-200 disabled:text-slate-400 disabled:cursor-not-allowed shadow-md flex items-center justify-center gap-2"
                        >
                            <span>Gerar Apostila (PDF)</span>
                        </button>
                    </div>

                </div>
            </div>

            {/* VISUALIZAÇÃO A4 DAS PÁGINAS FÍSICAS */}
            {paginasA4.length > 0 ? (
                <div className="space-y-10 print:space-y-0">
                    
                    {/* CAPA E ÍNDICE(S) (SOMENTE MODO OUTLET) */}
                    {tipoApostila === 'OUTLET' && (
                        <>
                            {/* PÁGINA 1 - CAPA ELEGANTE */}
                            <div
                                style={{ backgroundColor: '#002B5B', color: '#ffffff' }}
                                className="pagina-a4 max-w-[210mm] mx-auto p-12 shadow-xl rounded-xl print:shadow-none print:p-[12mm] print:rounded-none print:w-full flex flex-col justify-between h-[297mm] relative overflow-hidden"
                            >
                                <div className="flex justify-between items-start z-10">
                                    <div className="h-16">
                                        <img src="/logo2.png" alt="Agetherm" className="h-full object-contain" />
                                    </div>
                                    <span style={{ backgroundColor: '#FF8C00', color: '#ffffff' }} className="text-xs font-black tracking-widest uppercase px-3 py-1 rounded-full shadow-sm">
                                        OFERTAS ESPECIAIS
                                    </span>
                                </div>

                                <div className="my-auto z-10 text-center space-y-4">
                                    <h1 style={{ color: '#ffffff' }} className="text-6xl font-black tracking-tight uppercase">
                                        {tituloApostila}
                                    </h1>
                                    <p style={{ color: '#FF8C00' }} className="text-lg font-bold uppercase tracking-widest">
                                        CATÁLOGO DE PRODUTOS EM PROMOÇÃO
                                    </p>
                                    <div style={{ backgroundColor: '#FF8C00' }} className="w-24 h-1.5 mx-auto rounded-full"></div>
                                </div>

                                <div style={{ borderColor: 'rgba(255,255,255,0.3)' }} className="border-t pt-6 flex justify-between items-end z-10">
                                    <div>
                                        <div style={{ color: '#cbd5e1' }} className="text-xs font-semibold uppercase">Data de Emissão</div>
                                        <div style={{ color: '#ffffff' }} className="text-sm font-black">{dataAtualExtenso}</div>
                                    </div>
                                    <div className="text-right">
                                        <span style={{ color: '#cbd5e1' }} className="text-xs font-bold">OUTLET AGETHERM</span>
                                    </div>
                                </div>
                            </div>

                            {/* PÁGINAS DO ÍNDICE DINÂMICO MULTIPÁGINA */}
                            {paginasIndice.map((pagInd, idxInd) => (
                                <div key={idxInd} className="pagina-a4 max-w-[210mm] mx-auto bg-white p-0 shadow-xl rounded-xl print:shadow-none print:rounded-none print:w-full flex flex-col justify-between h-[297mm] overflow-hidden">
                                    <div style={{ backgroundColor: '#002B5B' }} className="px-10 py-5 flex justify-between items-center shrink-0">
                                        <h1 className="text-2xl font-black text-white uppercase tracking-tight">
                                            ÍNDICE DE CATEGORIAS {paginasIndice.length > 1 ? `(${idxInd + 1}/${paginasIndice.length})` : ''}
                                        </h1>
                                        <div className="h-8">
                                            <img src="/logo2.png" alt="Agetherm" className="h-full object-contain" />
                                        </div>
                                    </div>

                                    <div className="p-10 flex-1 space-y-3.5">
                                        {pagInd.itens.map((item, idx) => (
                                            <div key={idx} className="flex items-center justify-between text-sm py-1.5 border-b border-slate-100">
                                                <span className="font-black text-[#002B5B] uppercase tracking-wide">
                                                    {item.categoria}
                                                </span>
                                                <span className="flex-1 mx-4 border-b-2 border-dotted border-slate-300 h-3"></span>
                                                <span className="font-black text-[#FF8C00]">
                                                    PÁG. {item.paginaInicial}
                                                </span>
                                            </div>
                                        ))}
                                    </div>

                                    <div className="px-10 pb-6 pt-3 border-t border-slate-200 text-xs font-bold text-slate-400 flex justify-between items-center shrink-0">
                                        <span>OUTLET AGETHERM</span>
                                        <span>PÁG. {pagInd.numPagina} DE {totalPaginasPDF}</span>
                                    </div>
                                </div>
                            ))}
                        </>
                    )}

                    {/* PÁGINAS DE PRODUTOS COMPLETA MÁXIMA */}
                    {paginasA4.map((pagina, pIndex) => (
                        <div
                            key={pIndex}
                            className="pagina-a4 max-w-[210mm] mx-auto bg-white p-0 shadow-xl rounded-xl print:shadow-none print:rounded-none print:w-full flex flex-col justify-between h-[297mm] overflow-hidden"
                        >
                            {/* CABEÇALHO FAIXA AZUL SÓLIDA */}
                            <div style={{ backgroundColor: '#002B5B' }} className="px-8 py-4 flex justify-between items-center shrink-0">
                                <h2 className="text-xl font-black text-white uppercase tracking-wider">
                                    {tituloApostila}
                                </h2>
                                <div className="h-8">
                                    <img src="/logo2.png" alt="Agetherm" className="h-full object-contain" />
                                </div>
                            </div>

                            {/* LISTA DE PRODUTOS COMPLETA E SEM SOBRA EXAGERADA */}
                            {tipoApostila === 'OUTLET' ? (
                                <div className="px-8 pt-1 pb-1 flex-1 flex flex-col justify-start gap-0.5 overflow-hidden">
                                    {pagina.itens.map((item, itemIdx) => (
                                        <div key={itemIdx} className="flex flex-col gap-0.5">
                                            {/* BARRA DA CATEGORIA */}
                                            {item.ehInicioCategoria && (
                                                <div className="bg-[#002B5B] text-white px-3 py-1 rounded font-black text-xs uppercase tracking-wider border-l-4 border-[#FF8C00] shadow-sm my-0.5">
                                                    {item.categoria}
                                                </div>
                                            )}

                                            {/* CARD DO PRODUTO */}
                                            <div className="py-1.5 px-3 flex items-center justify-between gap-4 print:break-inside-avoid border-b border-slate-200 min-h-[74px]">
                                                {/* FOTO E DADOS */}
                                                <div className="flex items-center gap-4 overflow-hidden min-w-0 flex-1">
                                                    <div className="w-16 h-16 rounded-xl border border-slate-200 overflow-hidden bg-white shrink-0 flex items-center justify-center p-1 shadow-sm">
                                                        <img
                                                            src={`/produtos/${item.foto}.jpg`}
                                                            alt={item.codigo}
                                                            className="max-h-full max-w-full object-contain"
                                                            onError={(e) => {
                                                                if (e.target.src.endsWith('.jpg')) {
                                                                    e.target.src = `/produtos/${item.foto}.png`;
                                                                } else if (e.target.src.endsWith('.png')) {
                                                                    e.target.src = `/produtos/${item.foto}.webp`;
                                                                }
                                                            }}
                                                        />
                                                    </div>

                                                    <div className="flex flex-col min-w-0 space-y-0.5">
                                                        <span className="text-sm font-black text-[#002B5B] tracking-tight leading-tight">
                                                            {item.codigo}
                                                        </span>
                                                        {item.descricao && (
                                                            <span className="text-xs font-bold text-slate-600 uppercase tracking-tight leading-snug line-clamp-2">
                                                                {item.descricao}
                                                            </span>
                                                        )}
                                                    </div>
                                                </div>

                                                {/* PREÇO DE/POR DESTACADO */}
                                                <div className="flex items-center gap-3 shrink-0 bg-slate-50 px-3.5 py-1.5 rounded-xl border border-slate-200 shadow-inner">
                                                    <div className="bg-[#FF8C00] text-white text-xs font-black px-2 py-1 rounded shadow-sm">
                                                        -45% OFF
                                                    </div>

                                                    <div className="text-right min-w-[115px]">
                                                        <div className="text-xs font-bold text-rose-500 line-through leading-none mb-0.5">
                                                            DE: {formatarMoeda(item.precoOriginal)}
                                                        </div>
                                                        <div className="text-base font-black text-emerald-600 leading-tight">
                                                            POR: {formatarMoeda(item.precoDesconto)}
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            ) : (
                                /* MODO CÓDIGO DE BARRAS */
                                <div className="px-8 py-4 divide-y divide-slate-100 flex-1 flex flex-col justify-start gap-2">
                                    {pagina.itens.map((item, itemIdx) => (
                                        <div
                                            key={itemIdx}
                                            className="py-1 flex items-center justify-between gap-4 print:break-inside-avoid min-h-[48px]"
                                        >
                                            <div className="flex items-start gap-3 overflow-hidden min-w-0 flex-1">
                                                <span className="text-xs font-black text-slate-400 w-6 text-right shrink-0 pt-0.5">
                                                    #{item.numeroSequencial}
                                                </span>
                                                <div className="flex flex-col min-w-0">
                                                    <div className="flex items-center gap-2 flex-wrap">
                                                        <span className="text-sm font-black text-[#002B5B] tracking-tight leading-tight">
                                                            {item.codigo}
                                                        </span>
                                                        {item.multiplo && (
                                                            <span className="text-xs font-bold text-slate-500 uppercase tracking-tight">
                                                                - {item.multiplo}
                                                            </span>
                                                        )}
                                                    </div>
                                                    {item.descricao && (
                                                        <span className="text-[11px] font-semibold text-slate-600 uppercase tracking-tight leading-snug line-clamp-2">
                                                            {item.descricao}
                                                        </span>
                                                    )}
                                                </div>
                                            </div>

                                            <div className="flex flex-col items-center justify-center min-w-[160px] shrink-0">
                                                <BarcodeImage text={item.codBarras} />
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            )}

                            {/* RODAPÉ LIMPO */}
                            <div className="px-8 pb-5 pt-3 border-t border-slate-200 text-xs font-bold text-slate-400 flex justify-between items-center shrink-0">
                                <span>OUTLET AGETHERM</span>
                                <span>PÁG. {pagina.numPagina} DE {totalPaginasPDF}</span>
                            </div>
                        </div>
                    ))}
                </div>
            ) : (
                <div className="max-w-4xl mx-auto bg-white/60 backdrop-blur p-12 text-center text-slate-400 rounded-xl border-2 border-dashed border-slate-300 font-medium text-sm print:hidden">
                    Aguardando importação da planilha para gerar a visualização...
                </div>
            )}

            {/* REGRAS ESTRITAS DE IMPRESSÃO A4 */}
            <style jsx global>{`
                @media print {
                    @page {
                        size: A4 portrait;
                        margin: 0mm;
                    }
                    html, body {
                        background-color: #ffffff !important;
                        background: #ffffff !important;
                        padding: 0 !important;
                        margin: 0 !important;
                        -webkit-print-color-adjust: exact !important;
                        print-color-adjust: exact !important;
                    }
                    .pagina-a4 {
                        box-shadow: none !important;
                        width: 210mm !important;
                        height: 297mm !important;
                        max-height: 297mm !important;
                        display: flex !important;
                        flex-direction: column !important;
                        justify-content: space-between !important;
                        page-break-before: always !important;
                        break-before: page !important;
                        page-break-after: always !important;
                        break-after: page !important;
                        box-sizing: border-box !important;
                        overflow: hidden !important;
                        -webkit-print-color-adjust: exact !important;
                        print-color-adjust: exact !important;
                    }
                    .pagina-a4:first-child {
                        page-break-before: avoid !important;
                        break-before: avoid !important;
                    }
                }
            `}</style>
        </div>
    );
}