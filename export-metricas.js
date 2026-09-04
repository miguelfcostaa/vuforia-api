const admin = require('firebase-admin');
const { getFirestore } = require('firebase-admin/firestore');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// --- Firebase Init (same logic as server.js) ---
let serviceAccount;
const renderSecretPath = '/etc/secrets/firebase-key.json';
const localPath = './firebase-key.json';

try {
    if (fs.existsSync(renderSecretPath)) {
        serviceAccount = require(renderSecretPath);
    } else {
        serviceAccount = require(localPath);
    }
    if (serviceAccount && serviceAccount.private_key) {
        serviceAccount.private_key = serviceAccount.private_key.replace(/\\n/g, '\n');
    }
} catch (error) {
    console.error("ERRO FATAL: Não foi possível ler o ficheiro firebase-key.json.", error);
    process.exit(1);
}

const firebaseApp = admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
    storageBucket: 'prevention-vuforia-api'
});

const db = getFirestore(firebaseApp, 'prevention-game');

// Pivot configuration for special array fields
// - 'by_field': each item's field value becomes the column name (e.g. AGUA, MAÇA)
// - 'by_index': array index becomes the column name (e.g. PERGUNTA 1, PERGUNTA 2)
const PIVOT_FIELDS = {
    'detalhes_cortes_errados': {
        groupName: 'CORTE ERRADO',
        mode: 'by_field',
        columnField: 'item_name',
        valueField: 'quantidade'
    },
    'detalhes_respostas': {
        groupName: 'QUIZ',
        mode: 'by_index',
        columnPrefix: 'PERGUNTA',
        valueField: 'estava_correta'
    }
};

// --- Helper: extract the display value from a pivot item ---
function extractPivotValue(item, valueField) {
    // Try the configured field first
    if (item[valueField] !== undefined) {
        const val = item[valueField];
        // Convert booleans to ACERTOU/ERROU
        if (typeof val === 'boolean') {
            return val ? 'ACERTOU' : 'ERROU';
        }
        return val;
    }
    return '';
}

// --- Helper: flatten nested objects and arrays, with pivot support ---
function flattenObject(obj, prefix = '') {
    const result = {};
    for (const [key, value] of Object.entries(obj)) {
        const newKey = prefix ? `${prefix}.${key}` : key;

        // Check if this is a pivot field (at any nesting level)
        if (PIVOT_FIELDS[key] && Array.isArray(value)) {
            const config = PIVOT_FIELDS[key];

            if (config.mode === 'by_field') {
                // Column name comes from a field value (e.g. item_name)
                for (const item of value) {
                    if (item && item[config.columnField] !== undefined) {
                        const colName = `${config.groupName} > ${item[config.columnField]}`;
                        result[colName] = item[config.valueField] !== undefined ? item[config.valueField] : '';
                    }
                }
            } else if (config.mode === 'by_index') {
                // Column name comes from array position (PERGUNTA 1, PERGUNTA 2, ...)
                value.forEach((item, index) => {
                    const colName = `${config.groupName} > ${config.columnPrefix} ${index + 1}`;
                    result[colName] = extractPivotValue(item, config.valueField);
                });
            }
        } else if (Array.isArray(value)) {
            // Generic array expansion for other fields
            value.forEach((item, index) => {
                const arrayKey = `${newKey}[${index}]`;
                if (item !== null && typeof item === 'object' && !(item instanceof Date) && !(item.toDate instanceof Function)) {
                    Object.assign(result, flattenObject(item, arrayKey));
                } else {
                    result[arrayKey] = item;
                }
            });
        } else if (value !== null && typeof value === 'object' && !(value instanceof Date) && !(value.toDate instanceof Function)) {
            Object.assign(result, flattenObject(value, newKey));
        } else {
            result[newKey] = value;
        }
    }
    return result;
}

// --- Main export function ---
async function exportMetricas() {
    console.log('A ler métricas do Firestore...');

    const snapshot = await db.collection('metricas').get();

    if (snapshot.empty) {
        console.log('A coleção "metricas" está vazia. Nada para exportar.');
        process.exit(0);
    }

    console.log(`Encontrados ${snapshot.size} documentos.`);

    // 1. Flatten all documents and collect all unique column names
    const rows = [];
    const allColumns = new Set();

    snapshot.forEach(doc => {
        const data = doc.data();
        const flat = flattenObject(data);

        // Convert Firestore Timestamps to JS Dates
        for (const [key, value] of Object.entries(flat)) {
            if (value && typeof value.toDate === 'function') {
                flat[key] = value.toDate();
            }
        }

        flat['_doc_id'] = doc.id;

        for (const key of Object.keys(flat)) {
            allColumns.add(key);
        }

        rows.push(flat);
    });

    // 2. Define column order: _doc_id first, then sessao_id, username, timestamp, then the rest alphabetically
    const priorityColumns = ['_doc_id', 'sessao_id', 'username', 'timestamp'];
    const remainingColumns = [...allColumns]
        .filter(c => !priorityColumns.includes(c))
        .sort();
    const orderedColumns = [...priorityColumns.filter(c => allColumns.has(c)), ...remainingColumns];

    // 3. Build the Excel workbook
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'vuforia-api export';
    workbook.created = new Date();

    const sheet = workbook.addWorksheet('Métricas');

    // Header row
    sheet.columns = orderedColumns.map(col => ({
        header: col,
        key: col,
        width: Math.max(col.length + 4, 16)
    }));

    // Style the header row
    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4472C4' }
    };
    headerRow.alignment = { horizontal: 'center' };

    // Data rows
    for (const row of rows) {
        const rowData = {};
        for (const col of orderedColumns) {
            const value = row[col];
            rowData[col] = value !== undefined ? value : '';
        }
        sheet.addRow(rowData);
    }

    // Auto-filter on all columns
    sheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: orderedColumns.length }
    };

    // 4. Save to file
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const outputFile = path.join(__dirname, `metricas_${timestamp}.xlsx`);

    await workbook.xlsx.writeFile(outputFile);

    console.log(`\n✅ Excel exportado com sucesso!`);
    console.log(`   Ficheiro: ${outputFile}`);
    console.log(`   Linhas: ${rows.length}`);
    console.log(`   Colunas: ${orderedColumns.length}`);

    process.exit(0);
}

exportMetricas().catch(err => {
    console.error('Erro ao exportar métricas:', err);
    process.exit(1);
});
