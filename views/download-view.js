import { html, nothing } from 'lit-html';
import * as XLSX from 'https://esm.sh/xlsx@0.18.5';
import { getAllData, importBulkData, resetData, getNextCorrelatives, setSalesCorrelatives, getCompanyInfo, setCompanyInfo, importFullBackup } from '../data-store.js';
import { forceRender, navigate } from '../router.js';

function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

/* ----------------------------
   Template download utilities
   ---------------------------- */

// Convert template definition to CSV (semicolon separated for consistency with exports)
function convertTemplateDefToCsv(def) {
    const delim = ';';
    const headers = def.keys;
    const desc = def.descriptions || headers;
    const example = def.example || [];
    const rows = [];

    // Header (keys)
    rows.push(headers.join(delim));
    // Descriptions row
    rows.push(desc.map(d => `"${String(d).replace(/"/g, '""')}"`).join(delim));
    // Example row (if provided)
    if (example.length) {
        rows.push(example.map(v => {
            let val = v === undefined || v === null ? '' : String(v);
            val = val.replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/"/g, '""');
            if (val.includes(delim) || val.includes('"') || /^\s|\s$/.test(val)) val = `"${val}"`;
            return val;
        }).join(delim));
    }

    return rows.join('\n');
}

// Download a CSV template for a given key (clients/suppliers)
function downloadTemplateCsvFor(key) {
    const def = TEMPLATE_DEFINITIONS[key];
    if (!def) return;
    const csv = convertTemplateDefToCsv(def);
    const blobContent = "\uFEFF" + csv; // BOM for Excel compatibility
    downloadFile(blobContent, `${key}_template_${new Date().toISOString().split('T')[0]}.csv`, 'text/csv;charset=utf-8;');
}

// Generate a minimal XLSX template for a given key using SheetJS
function downloadTemplateXlsxFor(key) {
    const def = TEMPLATE_DEFINITIONS[key];
    if (!def) return;
    const wb = XLSX.utils.book_new();
    const headers = def.keys;
    const descriptions = def.descriptions || headers;
    const example = def.example || [];

    // Build AOa with headers, descriptions and example (if any)
    const aoa = [];
    aoa.push(headers);
    aoa.push(descriptions);
    if (example.length) aoa.push(example);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, key.substring(0, 31));
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filename = `${key}_template_${new Date().toISOString().split('T')[0]}.xlsx`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function downloadJson() {
    // Build a complete export payload explicitly to guarantee all main datasets are present
    const raw = getAllData();

    const exportData = {
        // metadata for easier identification when restoring
        meta: {
            exportedAt: new Date().toISOString(),
            app: 'Sistema VTA-IVA',
            version: '1.0'
        },
        // Ensure arrays exist even if empty to avoid missing sections on restore
        clients: Array.isArray(raw.clients) ? raw.clients.map(c => ({ ...c, email: c?.email || '' })) : [],
        suppliers: Array.isArray(raw.suppliers) ? raw.suppliers.map(s => ({ ...s, email: s?.email || '' })) : [],
        salesRecords: Array.isArray(raw.salesRecords) ? raw.salesRecords.slice().map(s => ({ ...s })) : [],
        purchaseRecords: Array.isArray(raw.purchaseRecords) ? raw.purchaseRecords.slice().map(p => ({ ...p })) : [],
        // Persist config/state objects with defaults if absent
        nextCorrelatives: raw.nextCorrelatives || { salesCF: 1, salesCCF: 1, purchases: 1 },
        companyInfo: raw.companyInfo || {
            name: '', nit: '', nrc: '', dui: '', activity: '', address: '', phone: ''
        }
    };

    // Order sales by correlativo (numeric asc) to make reading predictable
    exportData.salesRecords.sort((a, b) => (Number(a.correlative) || 0) - (Number(b.correlative) || 0));
    // Order purchases by date asc (oldest first)
    exportData.purchaseRecords.sort((a, b) => new Date(a.date) - new Date(b.date));

    const json = JSON.stringify(exportData, null, 2);
    downloadFile(json, `control_tributario_sv_${new Date().toISOString().split('T')[0]}.json`, 'application/json');
}

function convertToCsv(data, entityName) {
    // Use semicolon (;) as delimiter for CSV as requested (text/plain semicolon-separated)
    const delim = ';';
    if (!data || data.length === 0) {
        return `${entityName.toUpperCase()}\nNo data`;
    }
    const headers = Object.keys(data.reduce((acc,row)=>({ ...acc, ...row }), {})); // union of keys
    const csvRows = [];
    // Header row
    csvRows.push(headers.join(delim));
    for (const row of data) {
        const values = headers.map(header => {
            let value = row[header];
            if (value === null || value === undefined) value = '';
            if (typeof value === 'object') value = JSON.stringify(value);
            // Normalize newlines and quotes
            value = String(value).replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/"/g, '""');
            // If value contains delimiter or quotes or leading/trailing spaces, wrap in quotes
            if (value.includes(delim) || value.includes('"') || /^\s|\s$/.test(value)) {
                value = `"${value}"`;
            }
            return value;
        });
        csvRows.push(values.join(delim));
    }
    return csvRows.join('\n');
}

function downloadCsv() {
    const data = getAllData();
    const allCsv = [];
    for (const key in data) {
        if (Array.isArray(data[key])) {
            const csvContent = convertToCsv(data[key], key);
            allCsv.push(`\n\n#== START_DATASET: ${key.toUpperCase()} ==#`);
            allCsv.push(csvContent);
        }
    }
    // Include non-array datasets (companyInfo, nextCorrelatives)
    if (data.companyInfo) {
        allCsv.push(`\n\n#== START_DATASET: COMPANY_INFO ==#`);
        const infoRows = Object.entries(data.companyInfo).map(([k,v]) => `${k},${String(v).replace(/"/g,'""')}`);
        allCsv.push(['key,value', ...infoRows].join('\n'));
    }
    if (data.nextCorrelatives) {
        allCsv.push(`\n\n#== START_DATASET: NEXT_CORRELATIVES ==#`);
        const corrRows = Object.entries(data.nextCorrelatives).map(([k,v]) => `${k},${v}`);
        allCsv.push(['key,value', ...corrRows].join('\n'));
    }
    const finalCsv = allCsv.join('\n');
    const csvWithEncoding = "\uFEFF" + finalCsv;
    downloadFile(csvWithEncoding, `control_tributario_sv_data_${new Date().toISOString().split('T')[0]}.csv`, 'text/csv;charset=utf-8;');
}

const TEMPLATE_DEFINITIONS = {
    clients: {
        keys: ['name', 'nrc', 'nit', 'address', 'activity', 'contact', 'phone', 'email'],
        descriptions: [
            'Nombre/Razon Social',
            'NRC',
            'NIT',
            'Direccion',
            'Giro/Actividad Economica',
            'Contacto (Nombre o cargo)',
            'Telefono (7-10 dígitos)',
            'Correo Electrónico'
        ],
        example: [
            'Ejemplo Cliente S.A.',
            '1234567-8',
            '0101-123456-101-1',
            'San Salvador, Av. Principal',
            'Venta de Servicios',
            'Juan Perez - Compras',
            '12345678',
            'cliente@ejemplo.com'
        ]
    },
    // Suppliers now include contact and phone/email to match exports/imports and the clients template
    suppliers: {
        keys: ['name', 'nrc', 'nit', 'address', 'activity', 'contact', 'phone', 'email'],
        descriptions: [
            'Nombre/Razon Social',
            'NRC',
            'NIT',
            'Direccion',
            'Giro/Actividad Economica',
            'Contacto (Nombre o cargo)',
            'Telefono (7-10 dígitos)',
            'Correo Electrónico'
        ],
        example: [
            'Ejemplo Proveedor LTDA',
            '8765432-1',
            '0202-654321-202-2',
            'Santa Ana, Calle Central',
            'Compra de Mercaderia',
            'María López - Ventas',
            '98765432',
            'proveedor@ejemplo.com'
        ]
    },
    salesRecords: {
        // Include dteCode so exports/imports cover the DTE field used in CCF records
        keys: ['date', 'documentType', 'clientNrc', 'dteCode', 'taxableAmount', 'exemptAmount', 'total'],
        descriptions: [
            'Fecha (YYYY-MM-DD)', 
            'Tipo Doc (CCF/CF)', 
            'NRC Cliente (Obligatorio si CCF)',
            'Código DTE (opcional, si aplica al comprobante)',
            'Venta Gravada (Obligatorio si CCF, 0 si CF)', 
            'Venta Exenta (0 si CF)', 
            'Total Bruto (Obligatorio si CF, 0 si CCF)'
        ],
        // Example 1: CCF (Total calculated internally) including an example DTE code
        example: ['2023-11-20', 'CCF', '1234567-8', 'DTE-20231120-001', '100.00', '0.00', '0.00'],
        // Example 2: CF (Gravado/Exento calculated internally)
        example2: ['2023-11-21', 'CF', '', '', '0.00', '0.00', '113.00']
    },
    // Purchase records now include supplier email so Excel/CSV/JSON templates cover contact info
    purchaseRecords: {
        keys: ['date', 'documentType', 'supplierNrc', 'supplierEmail', 'documentNumber', 'taxableAmount', 'exemptAmount', 'ivaCredit', 'ivaWithheld'],
        descriptions: [
            'Fecha (YYYY-MM-DD)', 
            'Tipo Doc (CCF/Importacion/Otros)', 
            'NRC Proveedor', 
            'Correo Electrónico del Proveedor (opcional)', 
            'Número Documento', 
            'Gravado', 
            'Exento', 
            'IVA Crédito (Debe ser 13% Gravado)', 
            'Retención IVA (2%)'
        ],
        example: ['2023-11-20', 'CCF', '8765432-1', 'proveedor@ejemplo.com', 'F12345', '200.00', '0.00', '26.00', '0.00']
    }
};

function generateExcelTemplate() {
    // Create a proper XLSX workbook with one sheet per dataset using SheetJS (XLSX)
    const data = getAllData();
    const wb = XLSX.utils.book_new();

    // Helper to convert array of objects to worksheet
    const addSheetFromArray = (key, rows) => {
        // Ensure consistent headers (union of keys)
        const headers = rows.length ? Object.keys(rows.reduce((acc,r)=>({ ...acc, ...r }), {})) : [];
        // Map rows into array of arrays for worksheet
        const aoa = [headers];
        rows.forEach(r => {
            const row = headers.map(h => {
                let v = r[h];
                if (v === null || v === undefined) return '';
                if (typeof v === 'object') return JSON.stringify(v);
                return v;
            });
            aoa.push(row);
        });
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, key.substring(0, 31)); // Sheet name limited to 31 chars
    };

    // Add array datasets as separate sheets
    for (const key in data) {
        if (Array.isArray(data[key])) {
            addSheetFromArray(key, data[key]);
        }
    }

    // Add companyInfo as a simple key/value sheet
    const company = data.companyInfo || {};
    const companyAoa = [['key','value']];
    Object.entries(company).forEach(([k,v]) => companyAoa.push([k, v === undefined || v === null ? '' : v]));
    const wsCompany = XLSX.utils.aoa_to_sheet(companyAoa);
    XLSX.utils.book_append_sheet(wb, wsCompany, 'company_info');

    // Add nextCorrelatives as key/value sheet
    const corr = data.nextCorrelatives || {};
    const corrAoa = [['key','value']];
    Object.entries(corr).forEach(([k,v]) => corrAoa.push([k, v]));
    const wsCorr = XLSX.utils.aoa_to_sheet(corrAoa);
    XLSX.utils.book_append_sheet(wb, wsCorr, 'next_correlatives');

    // Generate binary and trigger download as true .xlsx
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filename = `control_tributario_sv_${new Date().toISOString().split('T')[0]}.xlsx`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

let uploadNotification = null;

function setUploadNotification(message, isSuccess = true, duration = 7000) {
    uploadNotification = { message, isSuccess };
    // Force re-render of the current view immediately to show status update
    forceRender();
    setTimeout(() => {
        uploadNotification = null;
        forceRender(); // Clear notification
    }, duration);
}

function parseCsvSection(sectionContent, entityType) {
    const def = TEMPLATE_DEFINITIONS[entityType];
    if (!def) return [];

    // Expects sectionContent starting from the key header row
    const lines = sectionContent.trim().split(/\r?\n/).filter(line => line.trim() !== '');
    if (lines.length < 2) return []; // Needs at least keys (0) and descriptions (1)

    // Detect delimiter: prefer semicolon if present, otherwise comma
    const headerLine = lines[0];
    const delim = headerLine.includes(';') ? ';' : ',';
    // Build a split regex that respects quoted fields
    const splitRegex = new RegExp(`${delim}(?=(?:(?:[^"]*"){2})*[^"]*$)`, 'g');

    // Line 0 is the keys (e.g., 'name;nrc;nit...' or 'name,nrc,nit...')
    const keyHeaders = headerLine.split(splitRegex).map(h => h.replace(/\\"/g, '').replace(/^"|"$/g, '').trim());
    const expectedKeys = def.keys;

    if (keyHeaders.length !== expectedKeys.length || keyHeaders.some((h, i) => h !== expectedKeys[i])) {
         console.warn(`Header keys mismatch for ${entityType}. Expected: ${expectedKeys}, Found: ${keyHeaders}.`);
         return [];
    }
    
    const data = [];
    
    // Determine starting line for actual data (skip Keys (0) and Descriptions (1))
    let startLineIndex = 2; 

    // Start parsing from the determined index
    for (let i = startLineIndex; i < lines.length; i++) {
        const line = lines[i];
        if (!line.trim()) continue; // Skip empty lines

        // Split using detected delimiter while preserving quoted values
        const values = line.split(splitRegex).map(v => v.replace(/^"|"$/g, '').trim());
        
        if (values.length === keyHeaders.length) {
            const record = {};
            keyHeaders.forEach((key, index) => {
                let value = values[index];
                
                // Attempt basic type conversion for numeric fields
                if (key.toLowerCase().includes('amount') || key.toLowerCase().includes('iva') || key.toLowerCase().includes('total')) {
                    // Clean symbols/separators before parsing float
                    value = (value || '').replace(/[^0-9.-]+/g,""); 
                    record[key] = parseFloat(value) || 0;
                } else {
                    record[key] = value;
                }
            });
            data.push(record);
        } else if (line.trim().length > 0) {
             console.warn(`Skipping malformed line in ${entityType}: Found ${values.length} fields, expected ${keyHeaders.length}. Line: ${line}`);
        }
    }
    return data;
}

async function bulkImportDataFromContent(content) {
    const results = {
        clients: [],
        suppliers: [],
        salesRecords: [],
        purchaseRecords: []
    };
    
    // Split the content by the dataset markers
    const sections = content.split(/#== TEMPLATE_START: ([A-Z_]+) ==#/g).slice(1); 
    
    if (sections.length === 0) {
        return { success: false, error: "No se encontraron marcadores de sección en el archivo. Asegúrate de usar la plantilla correcta." };
    }
    
    for (let i = 0; i < sections.length; i += 2) {
        const typeKey = sections[i].toLowerCase().trim();
        const contentBlock = sections[i + 1];
        
        let entityType;
        if (typeKey === 'clients') entityType = 'clients';
        else if (typeKey === 'suppliers') entityType = 'suppliers';
        else if (typeKey === 'salesrecords') entityType = 'salesRecords';
        else if (typeKey === 'purchaserecords') entityType = 'purchaseRecords';
        else continue; 
        
        const parsedData = parseCsvSection(contentBlock, entityType);
        results[entityType] = parsedData;
    }
    
    // Call the actual bulk import function exported from data-store.js
    const importStats = importBulkData(results); 
    
    // Save data after all imports (Handled internally by importBulkData now)
    return importStats;
}

async function handleFileUploadFor(entityHint, e) {
    // entityHint: 'clients' | 'suppliers' | null (best-effort)
    const file = e.target.files[0];
    if (!file) return;

    // Accept CSV, .xls and .xlsx files
    if (!file.name.toLowerCase().endsWith('.csv') && !file.name.toLowerCase().endsWith('.xls') && !file.name.toLowerCase().endsWith('.xlsx')) {
        setUploadNotification("Por favor, selecciona un archivo CSV, XLS o XLSX compatible.", false);
        return;
    }
    
    setUploadNotification("Cargando y procesando datos...", true);

    try {
        if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });

            const sections = [];
            workbook.SheetNames.forEach(sheetName => {
                const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
                if (!csv || !csv.trim()) return;
                const markerName = sheetName.toUpperCase().replace(/\s+/g, '_');
                sections.push(`#== TEMPLATE_START: ${markerName} ==#\n${csv}`);
            });

            const compositeContent = sections.join('\n\n');
            const importResult = await bulkImportDataFromContent(compositeContent);

            if (importResult && typeof importResult.clients === 'number') {
                // show success briefly so user sees immediate confirmation
                setUploadNotification(`Datos cargados exitosamente. Importados: Clientes (${importResult.clients}), Proveedores (${importResult.suppliers}), Ventas (${importResult.salesRecords}), Compras (${importResult.purchaseRecords}).`, true, 2000);
                e.target.value = '';
                setTimeout(() => { navigate('#/'); }, 1000);
            } else {
                setUploadNotification(`Error al importar datos desde el archivo Excel: formato inválido o datos inconsistente.`, false);
            }

        } else {
            // CSV path
            const text = await file.text();
            const hasMarkers = /#== TEMPLATE_START:/i.test(text);
            let contentToImport = text;

            if (!hasMarkers) {
                // If caller provided an entityHint, force-wrap the CSV into that section.
                const firstLine = text.split(/\r?\n/).find(l => l.trim()) || '';
                const header = firstLine.toLowerCase();

                if (entityHint === 'clients') {
                    contentToImport = `#== TEMPLATE_START: CLIENTS ==#\n${text}`;
                } else if (entityHint === 'suppliers') {
                    contentToImport = `#== TEMPLATE_START: SUPPLIERS ==#\n${text}`;
                } else {
                    // previous heuristic fallback: try to detect
                    if (header.includes('name') && header.includes('nrc') && header.includes('nit') && header.includes('email')) {
                        contentToImport = `#== TEMPLATE_START: CLIENTS ==#\n${text}`;
                    } else if (header.includes('name') && header.includes('nrc') && header.includes('nit')) {
                        contentToImport = `#== TEMPLATE_START: SUPPLIERS ==#\n${text}`;
                    } else {
                        contentToImport = text; // pass-through; parser may reject
                    }
                }
            }

            const importResult = await bulkImportDataFromContent(contentToImport);

            if (importResult && typeof importResult.clients === 'number') {
                // show brief success notification for uploaded templates
                setUploadNotification(`Datos cargados exitosamente. Importados: Clientes (${importResult.clients}), Proveedores (${importResult.suppliers}), Ventas (${importResult.salesRecords}), Compras (${importResult.purchaseRecords}).`, true, 2000);
                e.target.value = '';
                setTimeout(() => { navigate('#/'); }, 1000);
            } else {
                setUploadNotification(`Error al importar datos: Formato de archivo inválido o datos inconsistentes.`, false);
            }
        }
    } catch (error) {
        console.error("Upload error:", error);
        setUploadNotification(`Ocurrió un error inesperado durante la carga.`, false);
    }
}

// --- NEW JSON Upload Handler ---
async function handleJsonUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.endsWith('.json')) {
        setUploadNotification("Por favor, selecciona un archivo JSON compatible.", false);
        return;
    }

    setUploadNotification("Cargando y restaurando datos desde JSON...", true);

    try {
        const reader = new FileReader();
        reader.onload = async (event) => {
            const content = event.target.result;
            let data;
            
            try {
                data = JSON.parse(content);
            } catch (parseError) {
                setUploadNotification("Error al parsear el archivo JSON. Asegúrate de que el formato sea válido (debe ser un backup completo del sistema).", false);
                e.target.value = '';
                return;
            }

            try {
                // Use the new full backup import function
                importFullBackup(data);
                
                setUploadNotification(`Datos restaurados exitosamente desde el archivo JSON.`, true);
                
                e.target.value = '';

                setTimeout(() => {
                    // Navigate home to refresh dashboard/state
                    navigate('#/'); 
                }, 1000);
                
            } catch (importError) {
                console.error("JSON import error:", importError);
                setUploadNotification(`Error al restaurar el backup: ${importError.message}`, false);
            }
        };
        reader.readAsText(file); 
    } catch (error) {
        console.error("JSON Upload error:", error);
        setUploadNotification(`Ocurrió un error inesperado durante la carga del JSON.`, false);
    }
}
// --- END NEW JSON Upload Handler ---

function handleClearData() {
    // Replace window.confirm with custom modal
    showConfirmation(
        "ADVERTENCIA: ¿Está seguro que desea ELIMINAR TODOS los datos registrados (Clientes, Proveedores, Ventas, Compras)? Esta acción es irreversible.",
        () => {
            resetData();
            // Set notification, which triggers a render
            setUploadNotification("Todos los datos han sido eliminados exitosamente. El sistema ha sido reseteado.", true);
            
            // Navigate home after successful reset
            setTimeout(() => {
                navigate('#/'); 
            }, 1000);
        }
    );
}

let confirmAction = null; // Stores the action to be executed if confirmed
let confirmMessage = null; // Stores the message to display

function showConfirmation(message, action) {
    confirmMessage = message;
    confirmAction = action;
    forceRender();
}

function closeConfirmation() {
    confirmMessage = null;
    confirmAction = null;
    forceRender();
}

function handleConfirmExecute() {
    if (confirmAction) {
        confirmAction();
    }
    closeConfirmation();
}

function ConfirmationModal() {
    if (!confirmMessage) return nothing;

    return html`
        <div class="modal-overlay" @click=${closeConfirmation}>
            <div class="modal-content" @click=${(e) => e.stopPropagation()}>
                <h3>Confirmación Requerida</h3>
                <p>${confirmMessage}</p>
                <div class="modal-actions">
                    <button class="cancel-btn" @click=${closeConfirmation} style="background-color: #6c757d;">Cancelar</button>
                    <button @click=${handleConfirmExecute} style="background-color: var(--error-color);">Confirmar Eliminación</button>
                </div>
            </div>
        </div>
    `;
}

// --- Company Info Configuration Logic ---
function handleCompanyInfoSubmit(e) {
    e.preventDefault();
    const form = e.target;
    // Using simple FormData extraction as values are text fields
    const formData = new FormData(form);
    
    const info = {
        name: formData.get('companyName').trim(),
        nit: formData.get('companyNit').trim(),
        nrc: formData.get('companyNrc').trim(),
        dui: formData.get('companyDui').trim(),
        activity: formData.get('companyActivity').trim(),
        address: formData.get('companyAddress').trim(),
        phone: formData.get('companyPhone').trim(),
    };
    
    let success = false;
    let message = "Datos de empresa no actualizados.";

    try {
        if (!info.name) {
            throw new Error("El Nombre de la empresa o Razón Social es obligatorio.");
        }
        if (!info.nit) {
            throw new Error("El NIT es obligatorio.");
        }
        
        setCompanyInfo(info);
        success = true;
        message = "Datos de la Empresa actualizados exitosamente.";
    } catch (e) {
        console.error("Error setting company info:", e);
        message = `Error al guardar los datos de la empresa: ${e.message}`;
    }

    setUploadNotification(message, success);
}

function CompanyInfoConfig() {
    const info = getCompanyInfo();

    return html`
        <div class="card" style="margin-bottom: 20px;">
            <h3>Datos de la Empresa</h3>
            <p>Información utilizada para reportes y referencia interna.</p>
            <form @submit=${handleCompanyInfoSubmit} style="margin-top: 15px;">
                <div class="form-group">
                    <label for="companyName">Nombre Empresa o Razón Social</label>
                    <input type="text" id="companyName" name="companyName" value="${info.name}" required>
                </div>
                
                <div class="dashboard-grid" style="grid-template-columns: 1fr 1fr 1fr; gap: 15px;">
                    <div class="form-group">
                        <label for="companyNit">NIT</label>
                        <input type="text" id="companyNit" name="companyNit" value="${info.nit}" required>
                    </div>
                    <div class="form-group">
                        <label for="companyNrc">NRC</label>
                        <input type="text" id="companyNrc" name="companyNrc" value="${info.nrc}">
                    </div>
                    <div class="form-group">
                        <label for="companyDui">DUI</label>
                        <input type="text" id="companyDui" name="companyDui" value="${info.dui}">
                    </div>
                </div>
                
                <div class="form-group">
                    <label for="companyActivity">Giro de la empresa</label>
                    <input type="text" id="companyActivity" name="companyActivity" value="${info.activity}">
                </div>

                <div class="form-group">
                    <label for="companyAddress">Dirección</label>
                    <textarea id="companyAddress" name="companyAddress" rows="2">${info.address}</textarea>
                </div>
                
                <div class="form-group">
                    <label for="companyPhone">Teléfono</label>
                    <input type="text" id="companyPhone" name="companyPhone" value="${info.phone}">
                </div>

                <button type="submit" style="background-color: var(--primary-color); margin-top: 10px;">Guardar Datos de Empresa</button>
            </form>
        </div>
    `;
}

let correlativeConfigActive = false; // State for new configuration modal

function openCorrelativeConfig() {
    // Fetch latest values on open
    correlativeConfigActive = true;
    forceRender();
}

function closeCorrelativeConfig() {
    correlativeConfigActive = false;
    forceRender();
}

function handleCorrelativeSubmit(e) {
    e.preventDefault();
    const form = e.target;
    
    // Values are expected to be numeric input strings
    const ccfInput = form.elements.ccfCorrelative.value;
    const cfInput = form.elements.cfCorrelative.value;
    
    let success = false;
    let message = "Configuración no actualizada.";

    if (parseInt(ccfInput) >= 1 && parseInt(cfInput) >= 1) {
        try {
            // setSalesCorrelatives handles parsing and saving
            setSalesCorrelatives(ccfInput, cfInput);
            success = true;
            message = `Correlativos actualizados: CCF a ${parseInt(ccfInput)}, CF a ${parseInt(cfInput)}.`;
        } catch (e) {
            console.error(e);
            message = "Error al guardar los correlativos.";
        }
    } else {
        message = "Los números iniciales deben ser enteros mayores o iguales a 1.";
    }

    setUploadNotification(message, success);
    closeCorrelativeConfig();
}

function CorrelativeConfigModal() {
    if (!correlativeConfigActive) return nothing;

    // Fetch defaults here when rendering the modal
    const defaults = getNextCorrelatives();

    return html`
        <div class="modal-overlay" @click=${closeCorrelativeConfig}>
            <div class="modal-content" @click=${(e) => e.stopPropagation()} style="max-width: 500px;">
                <h3>Configuración Inicial de Correlativos de Venta</h3>
                <form @submit=${handleCorrelativeSubmit}>
                    <p>Establece el número con el que se iniciará el siguiente comprobante de cada tipo (Crédito Fiscal y Consumidor Final). El valor actual es el siguiente número disponible.</p>
                    
                    <div class="form-group" style="margin-top: 15px;">
                        <label for="ccfCorrelative">Siguiente Comprobante CCF (Crédito Fiscal)</label>
                        <input 
                            type="number" 
                            id="ccfCorrelative" 
                            name="ccfCorrelative" 
                            value="${defaults.salesCCF}" 
                            min="1" 
                            required
                        >
                    </div>
                    
                    <div class="form-group">
                        <label for="cfCorrelative">Siguiente Comprobante CF (Consumidor Final)</label>
                        <input 
                            type="number" 
                            id="cfCorrelative" 
                            name="cfCorrelative" 
                            value="${defaults.salesCF}" 
                            min="1" 
                            required
                        >
                    </div>

                    <div class="modal-actions">
                        <button type="button" class="cancel-btn" @click=${closeCorrelativeConfig} style="background-color: #6c757d;">Cancelar</button>
                        <button type="submit" style="background-color: var(--primary-color);">Guardar Configuración</button>
                    </div>
                </form>
            </div>
        </div>
    `;
}

export function DownloadView() {
    return html`
        ${ConfirmationModal()}
        ${CorrelativeConfigModal()}
        <h2>Herramientas de Gestión de Datos</h2>
        
        ${uploadNotification ? html`
            <div class="message ${uploadNotification.isSuccess ? 'success' : 'error'}">
                ${uploadNotification.message}
            </div>
        ` : nothing}
        
        ${CompanyInfoConfig()}

        <div class="card" style="margin-bottom: 20px;">
            <h3>Configuración Inicial de Correlativos</h3>
            <p>Ajusta el punto de inicio para la numeración automática de tus documentos de venta (CCF y CF).</p>
            <button @click=${openCorrelativeConfig} style="background-color: var(--accent-color); margin-top: 10px;">Establecer Correlativos de Venta</button>
        </div>

        <div class="card">
            <h3>Exportar Respaldo</h3>
            <p>Exporta toda la información registrada en el sistema (Clientes, Proveedores, Ventas y Compras), configuración de empresa y correlativos.</p>

            <div class="dashboard-grid" style="grid-template-columns: repeat(3, 1fr); margin-top: 20px;">
                <button @click=${downloadJson}>Descargar Datos en JSON</button>
                <button @click=${downloadCsv}>Descargar Datos en CSV</button>
                <button @click=${generateExcelTemplate} style="background-color: var(--success-color);">Descargar Respaldo Excel (.xls)</button>
            </div>
        </div>
        
        <div class="card" style="margin-top: 20px;">
            <h3>Carga Inicial/Masiva de Datos (Desde Plantilla Excel/CSV/JSON)</h3>
            <p>Utilice la plantilla descargada (Excel/CSV) o un archivo JSON compatible para preparar la información y subirla. Esto agregará nuevos registros al sistema.</p>

            <div class="form-group" style="margin-top: 20px;">
                    <label>Seleccionar Archivo de Plantilla (.csv o .xlsx)</label>
                    <div style="display:flex; gap:10px; flex-wrap:wrap; align-items:center;">
                        <div style="display:flex; flex-direction:column;">
                            <label for="fileUploadClients" style="font-weight:600; font-size:0.95em;">Cargar Plantilla - Clientes</label>
                            <input type="file" id="fileUploadClients" accept=".csv, .xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" @change=${(e)=>handleFileUploadFor('clients', e)}>
                        </div>
                        <div style="display:flex; flex-direction:column;">
                            <label for="fileUploadSuppliers" style="font-weight:600; font-size:0.95em;">Cargar Plantilla - Proveedores</label>
                            <input type="file" id="fileUploadSuppliers" accept=".csv, .xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" @change=${(e)=>handleFileUploadFor('suppliers', e)}>
                        </div>

                        <div style="display:flex; flex-direction:column;">
                            <label for="fileUploadSales" style="font-weight:600; font-size:0.95em;">Cargar Plantilla - Registro de Ventas</label>
                            <input type="file" id="fileUploadSales" accept=".csv, .xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" @change=${(e)=>handleFileUploadFor('salesRecords', e)}>
                        </div>
                        <div style="display:flex; flex-direction:column;">
                            <label for="fileUploadPurchases" style="font-weight:600; font-size:0.95em;">Cargar Plantilla - Registro de Compras</label>
                            <input type="file" id="fileUploadPurchases" accept=".csv, .xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" @change=${(e)=>handleFileUploadFor('purchaseRecords', e)}>
                        </div>

                        <div style="display:flex; flex-direction:column; align-items:flex-start;">
                            <small style="color:#666;">Usa el control correspondiente para subir plantillas específicas de Clientes, Proveedores, Ventas o Compras.</small>
                        </div>
                    </div>
                </div>

            <div style="margin-top:12px; display:flex; gap:10px; flex-wrap:wrap; align-items:center;">
                <div>
                    <strong>Descargar plantillas individuales (XLSX recomendado):</strong>
                </div>
                <div style="display:flex; gap:8px; flex-wrap:wrap;">
                    <button type="button" style="background-color:#2d6a4f; color:#fff;" @click=${() => downloadTemplateXlsxFor('clients')}>Clientes .xlsx</button>

                    <button type="button" style="background-color:#5b2e2e; color:#fff;" @click=${() => downloadTemplateXlsxFor('suppliers')}>Proveedores .xlsx</button>

                    <!-- New templates for sales and purchases (XLSX only, abbreviated labels and distinct colors) -->
                    <button type="button" style="background-color:#1b3b6f; color:#fff;" @click=${() => downloadTemplateXlsxFor('salesRecords')}>Ventas .xlsx</button>

                    <button type="button" style="background-color:#6f4f1b; color:#fff;" @click=${() => downloadTemplateXlsxFor('purchaseRecords')}>Compras .xlsx</button>
                </div>
            </div>
            
            <!-- NEW JSON Upload Section -->
            <h4 style="margin-top: 20px; border-bottom: 1px solid var(--border-color); padding-bottom: 5px;">Restaurar Backup JSON</h4>
            <p>Carga un archivo JSON previamente exportado para restaurar completamente el estado del sistema, sobrescribiendo los datos existentes.</p>
            <div class="form-group" style="margin-top: 10px;">
                <label for="jsonFileUpload">Seleccionar Archivo JSON de Respaldo</label>
                <input type="file" id="jsonFileUpload" accept="application/json" @change=${handleJsonUpload}>
            </div>
        </div>
        
        <div class="card" style="margin-top: 20px; border: 1px solid var(--error-color);">
            <h3>Eliminación Masiva de Datos</h3>
            <p>Esta acción ELIMINARÁ PERMANENTEMENTE TODOS los registros almacenados en el sistema (Clientes, Proveedores, Ventas, Compras).</p>
            <button @click=${handleClearData} style="background-color: var(--error-color); margin-top: 10px;">Eliminar Todos los Datos</button>
        </div>
    `;
}