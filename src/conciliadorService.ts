
/**
 * Motor de Conciliación "Conciliador Pro"
 * Lógica para procesar reportes de Innovat vs Estados de Cuenta del Banco
 */

export interface Transaction {
    date: string;
    name: string;
    id: string; // Matrícula o Clave
    amount: number;
    source: 'innovat' | 'banco';
    status: 'pending' | 'matched' | 'suggested';
    matchId?: string;
    originalLine?: any;
}

export interface ConciliationResult {
    matches: { a: Transaction, b: Transaction, confidence: number }[];
    onlyInnovat: Transaction[];
    onlyBanco: Transaction[];
    totalAmountInnovat: number;
    totalAmountBanco: number;
}

/**
 * Limpia montos como "$10,500.00" y los convierte a números puros
 */
export const cleanAmount = (amountStr: string): number => {
    if (!amountStr) return 0;
    const cleaned = amountStr.replace(/[$,\s]/g, "");
    return parseFloat(cleaned) || 0;
};

/**
 * Normaliza nombres para comparación (quita acentos, mayúsculas, etc)
 */
export const normalizeName = (name: string): string => {
    return name.toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim();
};

/**
 * LÓGICA DE CONCILIACIÓN
 */
export const performConciliation = (innovat: Transaction[], banco: Transaction[]): ConciliationResult => {
    const matches: { a: Transaction, b: Transaction, confidence: number }[] = [];
    const onlyInnovat = [...innovat];
    const onlyBanco = [...banco];

    // PASO 1: Match Exacto (Monto + ID/Matrícula)
    for (let i = onlyInnovat.length - 1; i >= 0; i--) {
        const inv = onlyInnovat[i];
        if (!inv.id) continue;

        const bankIdx = onlyBanco.findIndex(b =>
            b.amount === inv.amount &&
            (b.id === inv.id || b.originalLine.toString().includes(inv.id))
        );

        if (bankIdx !== -1) {
            matches.push({ a: inv, b: onlyBanco[bankIdx], confidence: 100 });
            onlyInnovat.splice(i, 1);
            onlyBanco.splice(bankIdx, 1);
        }
    }

    // PASO 2: Match Lógico (Monto + Nombre similar)
    for (let i = onlyInnovat.length - 1; i >= 0; i--) {
        const inv = onlyInnovat[i];
        const bankIdx = onlyBanco.findIndex(b => {
            const sameAmount = Math.abs(b.amount - inv.amount) < 0.01;
            const nameMatch = normalizeName(b.name).includes(normalizeName(inv.name)) ||
                normalizeName(inv.name).includes(normalizeName(b.name));
            return sameAmount && nameMatch;
        });

        if (bankIdx !== -1) {
            matches.push({ a: inv, b: onlyBanco[bankIdx], confidence: 85 });
            onlyInnovat.splice(i, 1);
            onlyBanco.splice(bankIdx, 1);
        }
    }

    // PASO 3: Match por Fecha y Monto (Terminales / STP sin datos)
    for (let i = onlyInnovat.length - 1; i >= 0; i--) {
        const inv = onlyInnovat[i];
        const bankIdx = onlyBanco.findIndex(b => {
            const sameAmount = Math.abs(b.amount - inv.amount) < 0.01;
            const dateDiff = Math.abs(new Date(b.date).getTime() - new Date(inv.date).getTime());
            const maxDiff = 48 * 60 * 60 * 1000; // 48 horas de margen
            return sameAmount && dateDiff <= maxDiff;
        });

        if (bankIdx !== -1) {
            matches.push({ a: inv, b: onlyBanco[bankIdx], confidence: 60 });
            onlyInnovat.splice(i, 1);
            onlyBanco.splice(bankIdx, 1);
        }
    }

    return {
        matches,
        onlyInnovat,
        onlyBanco,
        totalAmountInnovat: innovat.reduce((sum, t) => sum + t.amount, 0),
        totalAmountBanco: banco.reduce((sum, t) => sum + t.amount, 0)
    };
};
