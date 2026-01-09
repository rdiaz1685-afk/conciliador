import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Transaction, performConciliation, cleanAmount } from './conciliadorService';

const App = () => {
    const [innovatData, setInnovatData] = useState<Transaction[]>([]);
    const [bancoData, setBancoData] = useState<Transaction[]>([]);
    const [results, setResults] = useState<any>(null);
    const [isProcessing, setIsProcessing] = useState(false);

    const exportToExcel = (data: any) => {
        if (!data) return;

        const worksheetData = [
            ["ID Innovat", "Nombre Innovat", "Factura", "Método Pago", "Referencia (I)", "Monto Innovat", "Fecha Banco", "Confianza", "Status"],
            ...data.matches.map((m: any) => [
                m.a.id,
                m.a.name,
                m.a.factura || "-",
                m.a.metodoPago || "-",
                m.a.referencia || "-",
                m.a.amount,
                m.b.date,
                m.confidence + "%",
                m.confidence === 100 ? "Conciliado" : "Sugerido"
            ]),
            [],
            ["Solo en Innovat", "Nombre", "Factura", "Pago", "Ref", "Importe"],
            ...data.onlyInnovat.map((t: any) => [t.id, t.name, t.factura || "-", t.metodoPago || "-", t.referencia || "-", t.amount]),
            [],
            ["Solo en Banco", "", "", "", "", "Importe"],
            ...data.onlyBanco.map((t: any) => [t.id, t.name, "-", "-", "-", t.amount])
        ];

        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Conciliación");
        XLSX.writeFile(workbook, `Conciliacion_Diciembre_${Date.now()}.xlsx`);
    };

    // Función para partir CSV respetando comillas y detectando separador (; o ,)
    const splitCSV = (line: string) => {
        const commaCount = (line.match(/,/g) || []).length;
        const semiCount = (line.match(/;/g) || []).length;
        const sep = semiCount > commaCount ? ';' : ',';

        const result = [];
        let cur = '';
        let inQuote = false;
        for (let char of line) {
            if (char === '"') inQuote = !inQuote;
            else if (char === sep && !inQuote) {
                result.push(cur.trim());
                cur = '';
            } else cur += char;
        }
        result.push(cur.trim());
        return result;
    }

    // Simulador de carga de archivos (Pronto lo conectaremos a tu CSV real)
    const handleFileUpload = (e: any, type: 'innovat' | 'banco') => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            const text = event.target?.result as string;
            const lines = text.split('\n');
            console.log(`Analizando ${lines.length} líneas del archivo...`);

            const processed: Transaction[] = lines.map((line) => {
                if (!line.trim()) return null;
                const parts = splitCSV(line);

                if (type === 'innovat') {
                    // 1. Identificar campos básicos primero
                    const dateIdx = parts.findIndex(p => /\d{1,2}\/\d{1,2}\/\d{4}/.test(p));
                    if (dateIdx === -1) return null;
                    const date = parts[dateIdx];

                    const amountVal = cleanAmount(parts.find(p => p.includes('.') && cleanAmount(p) > 0) || "0");
                    if (amountVal <= 0) return null;

                    // El ID suele ser una parte numérica de 4-7 dígitos que NO es el monto
                    const idPart = parts.find(p => {
                        const clean = p.replace(/\D/g, "");
                        return clean.length >= 4 && clean.length <= 7 && cleanAmount(p) !== amountVal;
                    }) || "S/R";

                    const namePart = parts.find(p => p.length > 12 && !p.includes('/') && !p.includes('$')) || "S/N";

                    // 2. BUSCADOR INTELIGENTE DE DATOS EXTRA (Por eliminación)
                    const candidates = parts.filter((p, idx) => {
                        const clean = p.replace(/["'$]/g, "").trim();
                        if (!clean || clean === "-" || clean.length < 2) return false;
                        if (idx === dateIdx) return false;
                        if (cleanAmount(p) === amountVal) return false;
                        if (clean === idPart || p.includes(idPart)) return false;
                        if (p === namePart || namePart.includes(p)) return false;
                        if (p.includes('/')) return false;
                        return true;
                    });

                    const pms = ['TARJETA', 'EFECTIVO', 'STP', 'TRANSFERENCIA', 'CHEQUE', 'DEPOSITO', 'TERMINAL', 'TRA', 'EFEC', 'SANTANDER', 'BANCOMER', 'SPEI'];
                    const metodo = candidates.find(p => pms.some(m => p.toUpperCase().includes(m))) || "-";

                    const others = candidates.filter(p => p !== metodo);
                    const factura = others[0] || (parts[10] && parts[10].length > 2 ? parts[10] : "-");
                    const referencia = others[1] || (parts[8] && parts[8].length > 2 ? parts[8] : "-");

                    return {
                        date,
                        name: namePart,
                        id: idPart.replace(/["']/g, ""),
                        amount: amountVal,
                        source: type,
                        status: 'pending',
                        originalLine: line,
                        factura: factura.replace(/["']/g, "").trim(),
                        metodoPago: metodo.replace(/["']/g, "").trim(),
                        referencia: referencia.replace(/["']/g, "").trim()
                    };
                } else {
                    // BANCO: Detección Automática Estándar
                    const dateIdx = parts.findIndex(p => /\d{1,2}\/\d{1,2}\/\d{4}/.test(p));
                    if (dateIdx === -1) return null;

                    const amount = cleanAmount(parts.find(p => p.includes('.') && cleanAmount(p) > 0) || "0");
                    if (amount <= 0) return null;

                    const namePart = parts.find(p => p.length > 10 && !p.includes('/') && !p.includes('$')) || "S/N";
                    const idPart = parts.find(p => p.length >= 4 && p.length <= 8 && !isNaN(Number(p.replace(/\D/g, "")))) || "S/R";

                    return {
                        date: parts[dateIdx],
                        name: namePart,
                        id: idPart,
                        amount: amount,
                        source: type,
                        status: 'pending',
                        originalLine: line,
                        factura: "-",
                        metodoPago: "-",
                        referencia: "-"
                    };
                }
            }).filter(x => x !== null) as Transaction[];

            console.log(`✅ Procesados ${processed.length} registros válidos.`);
            if (type === 'innovat') setInnovatData(processed);
            else setBancoData(processed);
        };
        reader.readAsText(file);
    };

    const runAnalysis = () => {
        setIsProcessing(true);
        setTimeout(() => {
            const res = performConciliation(innovatData, bancoData);
            setResults(res);
            setIsProcessing(false);
        }, 1500);
    };

    return (
        <div className="min-h-screen bg-[#0f172a] text-slate-200 p-8 font-sans">
            <header className="mb-12 flex justify-between items-end border-b border-emerald-500/20 pb-8">
                <div>
                    <p className="text-emerald-500 font-black uppercase tracking-[0.4em] text-[10px] mb-2">Módulo de Auditoría</p>
                    <h1 className="text-5xl font-black text-white uppercase tracking-tighter italic">Conciliador <span className="text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 to-teal-500">Pro</span></h1>
                </div>
                <div className="text-right flex items-end gap-6">
                    <div>
                        <p className="text-[9px] font-bold text-slate-500 uppercase tracking-widest">Colegio Dominio • Diciembre 2025</p>
                    </div>
                    <button
                        onClick={() => {
                            setInnovatData([]);
                            setBancoData([]);
                            setResults(null);
                        }}
                        className="px-4 py-2 border border-red-500/30 text-red-500 text-[9px] font-black uppercase rounded-lg hover:bg-red-500 hover:text-white transition-all"
                    >
                        ✕ Salir / Limpiar
                    </button>
                </div>
            </header>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                {/* Panel de Carga */}
                <div className="lg:col-span-1 space-y-6">
                    <div className="bg-slate-900/50 p-8 rounded-[40px] border border-white/5 backdrop-blur-xl">
                        <h2 className="text-xs font-black uppercase tracking-widest mb-6 flex items-center gap-2">
                            <span className="w-2 h-2 rounded-full bg-emerald-500"></span> Fuente de Datos
                        </h2>

                        <div className="space-y-4">
                            <div className="p-4 bg-black/40 rounded-2xl border border-white/5 hover:border-emerald-500/30 transition-all cursor-pointer relative group">
                                <p className="text-[10px] font-black uppercase text-slate-500 mb-2">Reporte Innovat (Ingresos)</p>
                                <input type="file" onChange={(e) => handleFileUpload(e, 'innovat')} className="opacity-0 absolute inset-0 cursor-pointer" />
                                <p className="text-sm font-bold text-white">{innovatData.length > 0 ? `✅ ${innovatData.length} Registros` : '✚ Seleccionar Archivo'}</p>
                            </div>

                            <div className="p-4 bg-black/40 rounded-2xl border border-white/5 hover:border-emerald-500/30 transition-all cursor-pointer relative group">
                                <p className="text-[10px] font-black uppercase text-slate-500 mb-2">Estado de Cuenta (Banco)</p>
                                <input type="file" onChange={(e) => handleFileUpload(e, 'banco')} className="opacity-0 absolute inset-0 cursor-pointer" />
                                <p className="text-sm font-bold text-white">{bancoData.length > 0 ? `✅ ${bancoData.length} Registros` : '✚ Seleccionar Archivo'}</p>
                            </div>
                        </div>

                        <button
                            onClick={runAnalysis}
                            disabled={innovatData.length === 0 || bancoData.length === 0 || isProcessing}
                            className="w-full mt-8 py-5 bg-emerald-500 text-slate-950 rounded-2xl font-black uppercase text-[10px] tracking-[0.3em] hover:bg-emerald-400 transition-all active:scale-95 disabled:opacity-30 flex items-center justify-center gap-3"
                        >
                            {isProcessing ? 'Procesando Inteligencia...' : 'Ejecutar Conciliación ➔'}
                        </button>
                    </div>

                    {results && (
                        <div className="bg-slate-900/50 p-8 rounded-[40px] border border-white/5 space-y-6">
                            <div className="flex justify-between items-center py-4 border-b border-white/5">
                                <span className="text-[10px] uppercase font-black text-slate-500">Total Innovat</span>
                                <span className="text-lg font-black text-white">${results.totalAmountInnovat.toLocaleString()}</span>
                            </div>
                            <div className="flex justify-between items-center py-4 border-b border-white/5 text-emerald-400">
                                <span className="text-[10px] uppercase font-black">Total Banco</span>
                                <span className="text-lg font-black">${results.totalAmountBanco.toLocaleString()}</span>
                            </div>
                            <div className="flex justify-between items-center pt-4">
                                <span className="text-[10px] uppercase font-black text-slate-500">Diferencia</span>
                                <span className={`text-xl font-black ${results.totalAmountBanco - results.totalAmountInnovat === 0 ? 'text-emerald-500' : 'text-red-500'}`}>
                                    ${(results.totalAmountBanco - results.totalAmountInnovat).toLocaleString()}
                                </span>
                            </div>
                        </div>
                    )}
                </div>

                {/* Panel de Resultados */}
                <div className="lg:col-span-2 space-y-6">
                    {!results ? (
                        <div className="h-full min-h-[500px] border-4 border-dashed border-white/5 rounded-[60px] flex flex-col items-center justify-center text-slate-700">
                            <span className="text-6xl mb-6">📊</span>
                            <p className="text-[10px] font-black uppercase tracking-[0.5em]">Esperando Datos para Análisis</p>
                        </div>
                    ) : (
                        <div className="space-y-6">
                            <div className="grid grid-cols-3 gap-4">
                                <div className="bg-emerald-500/10 border border-emerald-500/20 p-6 rounded-3xl text-center">
                                    <p className="text-[10px] font-black uppercase text-emerald-500 mb-1">Cruce Exitoso</p>
                                    <p className="text-3xl font-black text-white">{results.matches.length}</p>
                                </div>
                                <div className="bg-red-500/10 border border-red-500/20 p-6 rounded-3xl text-center">
                                    <p className="text-[10px] font-black uppercase text-red-500 mb-1">Solo en Banco</p>
                                    <p className="text-3xl font-black text-white">{results.onlyBanco.length}</p>
                                </div>
                                <div className="bg-amber-500/10 border border-amber-500/20 p-6 rounded-3xl text-center">
                                    <p className="text-[10px] font-black uppercase text-amber-500 mb-1">Pendiente Innovat</p>
                                    <p className="text-3xl font-black text-white">{results.onlyInnovat.length}</p>
                                </div>
                            </div>

                            <div className="bg-slate-900/50 rounded-[40px] border border-white/5 p-8">
                                <div className="flex justify-between items-center mb-8">
                                    <h3 className="text-xs font-black uppercase tracking-widest">Detalle de Conciliación</h3>
                                    <button
                                        onClick={() => exportToExcel(results)}
                                        className="px-6 py-3 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded-2xl text-[9px] font-black uppercase tracking-widest hover:bg-emerald-500 hover:text-slate-950 transition-all"
                                    >
                                        ⬇ Descargar Excel
                                    </button>
                                </div>
                                <div className="space-y-4 max-h-[600px] overflow-y-auto pr-4 custom-scrollbar">
                                    {results.matches.map((m: any, idx: number) => (
                                        <div key={idx} className={`flex items-center gap-4 p-4 bg-black/40 rounded-2xl border transition-all ${m.confidence === 100 ? 'border-emerald-500/20 hover:border-emerald-500/40' : 'border-amber-500/20 hover:border-amber-500/40'}`}>
                                            <div className="flex-1">
                                                <p className="text-[9px] font-black text-slate-500 uppercase tracking-tighter">Innovat: {m.a.id}</p>
                                                <p className="text-sm font-bold text-white truncate">{m.a.name}</p>
                                                {m.a.referencia && m.a.referencia !== '-' && (
                                                    <p className="text-[10px] text-emerald-500/80 mt-1 font-bold italic">Ref: {m.a.referencia}</p>
                                                )}
                                            </div>
                                            <div className="w-24 text-center">
                                                <span className={`text-[9px] font-black px-4 py-1.5 rounded-full border ${m.confidence === 100
                                                    ? 'text-emerald-500 bg-emerald-500/10 border-emerald-500/20'
                                                    : 'text-amber-500 bg-amber-500/10 border-amber-500/20'
                                                    }`}>
                                                    {m.confidence === 100 ? 'EXACTO 100%' : ' REVISAR ' + m.confidence + '%'}
                                                </span>
                                            </div>
                                            <div className="flex-1 text-right">
                                                <p className="text-[9px] font-black text-slate-500 uppercase tracking-tighter">Banco: {m.b.date}</p>
                                                <p className={`text-sm font-black ${m.confidence === 100 ? 'text-emerald-400' : 'text-amber-400'}`}>
                                                    ${m.a.amount.toLocaleString()}
                                                </p>
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
};

export default App;
