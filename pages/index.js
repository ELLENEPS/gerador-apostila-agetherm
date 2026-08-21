import { useState, useRef, useEffect } from 'react';
import Head from 'next/head';
import * as XLSX from 'xlsx';
import bwipjs from 'bwip-js';

export default function GeradorApostila() {
    const [paginasA4, setPaginasA4] = useState([]);
    const [totalItens, setTotalItens] = useState(0);
    const [carregando, setCarregando] = useState(false);
    const [nomeArquivo, setNomeArquivo] = useState('');
    const [tituloApostila, setTituloApostila] = useState('CATÁLOGO DE CÓDIGOS DE BARRAS');

    const ITENS_POR_PAGINA = 12;

    const handleDownloadModelo = () => {
        const dadosModelo = [
            { 'CATEGORIA': 'CONEXÕES', 'REFERÊNCIA': 'REF-001', 'DESCRIÇÃO': 'JOELHO 90 GRAUS 22MM PARA TUBULAÇÃO DE COBRE E SOLDA', 'CÓDIGO DE BARRAS': '7891234567890' },
            { 'CATEGORIA': 'CONEXÕES', 'REFERÊNCIA': 'REF-002', 'DESCRIÇÃO': 'LUVA REDUÇÃO 28X22MM', 'CÓDIGO DE BARRAS': '7891234567891' },
            { 'CATEGORIA': 'VÁLVULAS', 'REFERÊNCIA': 'REF-003', 'DESCRIÇÃO': 'VÁLVULA ESFERA 3/4 PASSO PLENO ROSTÁVEL', 'CÓDIGO DE BARRAS': '7891234567892' },
            { 'CATEGORIA': 'VÁLVULAS', 'REFERÊNCIA': 'REF-004', 'DESCRIÇÃO': 'VÁLVULA RETENÇÃO 1 POL', 'CÓDIGO DE BARRAS': '7891234567893' }
        ];

        const worksheet = XLSX.utils.json_to_sheet(dadosModelo);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Modelo');
        XLSX.writeFile(workbook, 'modelo_apostila_agetherm.xlsx');
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
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                const gruposMap = new Map();
                let contadorTotal = 0;

                jsonData.forEach((row) => {
                    const categoria = String(row['CATEGORIA'] || row['Categoria'] || 'GERAL').trim().toUpperCase();
                    const codigo = String(row['REFERÊNCIA'] || row['REFERENCIA'] || row['Referência'] || row['Referencia'] || row['Código'] || row['Codigo'] || 'SEM REFERÊNCIA').trim();
                    const descricao = String(row['DESCRIÇÃO'] || row['DESCRICAO'] || row['Descrição'] || row['Descricao'] || row['PRODUTO'] || row['Produto'] || row['NOME'] || row['Nome'] || '').trim();
                    const codBarras = String(row['CÓDIGO DE BARRAS'] || row['CODIGO DE BARRAS'] || row['Código de Barras'] || row['Codigo de Barras'] || row['BARRAS'] || row['EAN'] || row['BARCODE'] || codigo).trim();

                    if (!gruposMap.has(categoria)) {
                        gruposMap.set(categoria, []);
                    }

                    gruposMap.get(categoria).push({
                        codigo,
                        descricao,
                        codBarras
                    });

                    contadorTotal++;
                });

                const paginasMontadas = [];

                gruposMap.forEach((itensCategoria, nomeCategoria) => {
                    let indiceItemLocal = 1;
                    
                    for (let i = 0; i < itensCategoria.length; i += ITENS_POR_PAGINA) {
                        const fatiaItens = itensCategoria.slice(i, i + ITENS_POR_PAGINA).map((item) => ({
                            ...item,
                            numeroSequencial: indiceItemLocal++
                        }));

                        paginasMontadas.push({
                            categoria: nomeCategoria,
                            itens: fatiaItens
                        });
                    }
                });

                setPaginasA4(paginasMontadas);
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

    return (
        <div className="min-h-screen bg-slate-100 py-8 px-4 font-sans text-slate-800 print:bg-white print:py-0 print:px-0">
            <Head>
                <title>Gerador de Apostila - AGETHERM</title>
            </Head>

            {/* PAINEL DE CONTROLE DE CONFIGURAÇÕES */}
            <div className="max-w-4xl mx-auto bg-white rounded-xl shadow-md border-t-4 border-[#FF8C00] overflow-hidden mb-10 print:hidden">
                
                <div className="p-6 md:p-8 border-b border-slate-100 flex flex-col md:flex-row items-center justify-between gap-4">
                    <div>
                        <h1 className="text-xl md:text-2xl font-black text-[#002B5B] uppercase tracking-tight">
                            Gerador de Apostila de Códigos
                        </h1>
                    </div>
                    <div className="h-12 flex items-center justify-center">
                        <img src="/logo.png" alt="Agetherm Logo" className="max-h-full w-auto object-contain" />
                    </div>
                </div>

                <div className="p-6 md:p-8 space-y-6">
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

                    {/* BLOCO DE IMPORTAÇÃO ALINHADO À ESQUERDA */}
                    <div className="bg-slate-50 rounded-xl p-6 border border-slate-200/80 text-left space-y-4">
                        <div>
                            <h2 className="text-sm font-bold text-[#002B5B] uppercase tracking-wide">
                                Importar Dados da Planilha
                            </h2>
                            <p className="text-xs text-slate-500 mt-1">
                                Envie um arquivo contendo as colunas <span className="font-bold text-slate-700">CATEGORIA</span>, <span className="font-bold text-slate-700">REFERÊNCIA</span>, <span className="font-bold text-slate-700">DESCRIÇÃO</span> e <span className="font-bold text-slate-700">CÓDIGO DE BARRAS</span>.
                            </p>
                        </div>

                        <div className="flex flex-col sm:flex-row items-center justify-start gap-3 pt-2">
                            <label className="w-full sm:w-auto cursor-pointer bg-[#002B5B] text-white px-5 py-2.5 rounded-lg font-bold text-xs uppercase tracking-wider hover:bg-[#001d3d] active:scale-95 transition-all shadow-sm flex items-center justify-center gap-2">
                                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                                </svg>
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
                                <svg className="w-4 h-4 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                                </svg>
                                <span>Baixar Modelo (.XLSX)</span>
                            </button>
                        </div>

                        {nomeArquivo && (
                            <div className="inline-flex items-center gap-2 text-xs font-semibold text-emerald-700 bg-emerald-50 px-3 py-1 rounded-md border border-emerald-200 mt-2">
                                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" />
                                </svg>
                                <span>{nomeArquivo}</span>
                            </div>
                        )}
                    </div>

                    <div className="pt-4 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4">
                        <div className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                            Total de Itens: <span className="text-base font-black text-[#002B5B] ml-1">{totalItens}</span> | Páginas A4: <span className="text-base font-black text-[#FF8C00] ml-1">{paginasA4.length}</span>
                        </div>

                        <button
                            onClick={handleImprimir}
                            disabled={totalItens === 0}
                            className="w-full sm:w-auto bg-[#FF8C00] text-white px-7 py-3 rounded-lg font-bold text-xs uppercase tracking-wider hover:bg-[#e67e00] active:scale-95 transition-all disabled:bg-slate-200 disabled:text-slate-400 disabled:cursor-not-allowed disabled:transform-none shadow-md flex items-center justify-center gap-2"
                        >
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z" />
                            </svg>
                            <span>Gerar Apostila (PDF)</span>
                        </button>
                    </div>

                </div>
            </div>

            {/* VISUALIZAÇÃO A4 DAS PÁGINAS FÍSICAS */}
            {paginasA4.length > 0 ? (
                <div className="space-y-10 print:space-y-0">
                    {paginasA4.map((pagina, pIndex) => (
                        <div
                            key={pIndex}
                            className="pagina-a4 max-w-[210mm] mx-auto bg-white pt-[6mm] pb-[6mm] pr-[16mm] pl-[22mm] shadow-xl rounded-xl print:shadow-none print:p-0 print:rounded-none print:w-full flex flex-col justify-between"
                        >
                            {/* BLOCO SUPERIOR (CABEÇALHO + CATEGORIA ORIGINAL) */}
                            <div>
                                <div className="flex justify-between items-center border-b-2 border-[#002B5B] pb-1.5 mb-2">
                                    <div>
                                        <h1 className="text-xl font-black text-[#002B5B] uppercase tracking-tight">
                                            {tituloApostila}
                                        </h1>
                                    </div>
                                    <div className="h-8">
                                        <img src="/logo.png" alt="Agetherm" className="h-full object-contain" />
                                    </div>
                                </div>

                                <div className="box-categoria bg-[#002B5B] text-white px-4 py-2.5 rounded-md font-black text-xs uppercase tracking-wider mb-2 border-l-8 border-[#FF8C00] shadow-sm flex justify-between items-center">
                                    <span className="text-white font-black">{pagina.categoria}</span>
                                    <span className="text-white font-black text-[11px]">
                                        PÁG. {pIndex + 1} DE {paginasA4.length}
                                    </span>
                                </div>
                            </div>

                            {/* LISTA DE ITENS */}
                            <div className="divide-y divide-slate-100 my-1">
                                {pagina.itens.map((item, itemIdx) => (
                                    <div
                                        key={itemIdx}
                                        className="py-1 mb-[2.5mm] flex items-center justify-between gap-4 print:break-inside-avoid min-h-[48px]"
                                    >
                                        <div className="flex items-start gap-3 overflow-hidden min-w-0 flex-1">
                                            <span className="text-xs font-black text-slate-400 w-6 text-right shrink-0 pt-0.5">
                                                #{item.numeroSequencial}
                                            </span>
                                            <div className="flex flex-col min-w-0">
                                                <span className="text-sm font-black text-[#002B5B] tracking-tight shrink-0 leading-tight">
                                                    {item.codigo}
                                                </span>
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
                        </div>
                    ))}
                </div>
            ) : (
                <div className="max-w-4xl mx-auto bg-white/60 backdrop-blur p-12 text-center text-slate-400 rounded-xl border-2 border-dashed border-slate-300 font-medium text-sm print:hidden">
                    Aguardando importação da planilha para gerar a visualização...
                </div>
            )}

            {/* REGRAS DE IMPRESSÃO A4 */}
            <style jsx global>{`
                @media print {
                    @page {
                        size: A4 portrait;
                        margin: 5mm 16mm 5mm 22mm;
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
                        background-color: #ffffff !important;
                        padding: 0 !important;
                        width: 100% !important;
                        page-break-after: auto !important;
                        break-after: auto !important;
                    }
                    .pagina-a4:not(:last-child) {
                        page-break-after: always !important;
                        break-after: page !important;
                    }
                    .box-categoria {
                        background-color: #002B5B !important;
                        color: #ffffff !important;
                        -webkit-print-color-adjust: exact !important;
                        print-color-adjust: exact !important;
                    }
                    .box-categoria span {
                        color: #ffffff !important;
                    }
                }
            `}</style>
        </div>
    );
}