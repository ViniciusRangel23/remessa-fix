(function() {
    'use strict';
    
    // Conversor de planilhas (XLS/XLSX/CSV) para uso em remessa:
    // formata K/L/M/S/T/U com 4 casas decimais (vírgula) e força texto.
    const CONFIG = {
        allowedExtensions: ['.xls', '.xlsx', '.csv'],
        targetColumns: [10, 11, 12, 18, 19, 20], // K, L, M, S, T, U
        maxColumnWidth: 30,
        minColumnWidth: 8,
        decimalPlaces: 4
    };

    // Referências de DOM
    const DOM = {
        uploader: null,
        fileInput: null,
        browseBtn: null,
        gallery: null
    };

    init();

    function init() {
        if (!cacheDOM()) {
            console.error('Elementos DOM obrigatórios não encontrados');
            return;
        }
        setupEventListeners();
    }

    function cacheDOM() {
        DOM.uploader = document.getElementById('uploader');
        DOM.fileInput = document.getElementById('fileInput');
        DOM.browseBtn = document.getElementById('browseBtn');
        DOM.gallery = document.getElementById('gallery');
        
        return !![DOM.uploader, DOM.fileInput, DOM.browseBtn, DOM.gallery].every(el => el);
    }

    // Eventos de arrastar e soltar + seleção de arquivo
    function setupEventListeners() {
        const dragDropEvents = ['dragenter', 'dragover', 'dragleave', 'drop'];
        dragDropEvents.forEach(event => {
            DOM.uploader.addEventListener(event, preventDefault);
        });

        DOM.uploader.addEventListener('dragover', () => DOM.uploader.classList.add('dragging'));
        DOM.uploader.addEventListener('dragleave', () => DOM.uploader.classList.remove('dragging'));
        DOM.uploader.addEventListener('drop', handleDrop);

        // Botão "procurar"
        DOM.browseBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            DOM.fileInput.click();
        });

        
        DOM.uploader.addEventListener('click', () => DOM.fileInput.click());

        
        DOM.fileInput.addEventListener('change', handleFileInputChange);
    }

   
    // Dispara processamento ao escolher arquivos pelo input
    function handleFileInputChange(e) {
        if (e.target.files?.length) {
            handleFiles(e.target.files);
        }
        e.target.value = ''; 
    }

    function preventDefault(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function handleDrop(e) {
        DOM.uploader.classList.remove('dragging');
        const dt = e.dataTransfer;
        if (dt?.files?.length) {
            handleFiles(dt.files);
        }
    }


    // Filtra extensões permitidas e cria cards
    function handleFiles(files) {
        Array.from(files).forEach(file => {
            if (isValidFile(file)) {
                createFileCard(file);
            }
        });
    }


    function isValidFile(file) {
        const name = file.name.toLowerCase();
        return CONFIG.allowedExtensions.some(ext => name.endsWith(ext));
    }


    // Monta o card do arquivo
    function createFileCard(file) {
        const card = document.createElement('div');
        card.className = 'file-card';

        card.appendChild(createFileIcon(file));
        card.appendChild(createFileMeta(file));
        card.appendChild(createFileActions(file));
        
        DOM.gallery.appendChild(card);
    }

    // Ícone com a extensão do arquivo
    function createFileIcon(file) {
        const icon = document.createElement('div');
        icon.className = 'file-icon';
        icon.textContent = file.name.split('.').pop().toUpperCase();
        return icon;
    }

    // Metadados: nome e tamanho
    function createFileMeta(file) {
        const meta = document.createElement('div');
        meta.className = 'meta';

        const nameEl = document.createElement('div');
        nameEl.className = 'name';
        nameEl.textContent = file.name;

        const sizeEl = document.createElement('div');
        sizeEl.className = 'size';
        sizeEl.textContent = formatBytes(file.size);

        meta.appendChild(nameEl);
        meta.appendChild(sizeEl);
        return meta;
    }

    // Ações: baixar e remover
    function createFileActions(file) {
        const actions = document.createElement('div');
        actions.className = 'actions';

        const processBtn = document.createElement('button');
        processBtn.className = 'download';
        processBtn.type = 'button';
        processBtn.textContent = 'Baixar';
        processBtn.addEventListener('click', () => processAndDownload(file, processBtn));

        const removeBtn = document.createElement('button');
        removeBtn.className = 'remove';
        removeBtn.type = 'button';
        removeBtn.textContent = 'Remover';
        removeBtn.addEventListener('click', () => {
            const card = removeBtn.closest('.file-card');
            if (card) card.remove();
        });

        actions.appendChild(processBtn);
        actions.appendChild(removeBtn);
        return actions;
    }

    // Processa e baixa o arquivo
    async function processAndDownload(file, button) {
        const originalText = button.textContent;
        button.disabled = true;
        button.textContent = 'Processando...';

        try {
            await processFile(file);
            button.textContent = 'Baixando...';
            await new Promise(resolve => setTimeout(resolve, 1000));
            button.textContent = 'Baixar novamente';
        } catch (error) {
            console.error('Erro ao processar arquivo:', error);
            showError(`Erro: ${error.message}`);
            button.textContent = originalText;
        } finally {
            button.disabled = false;
        }
    }

    // Exibe mensagem de erro
    function showError(message) {
        alert(message);
    }

    // Lê e processa um arquivo
    function processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    if (!e.target?.result) {
                        throw new Error('Falha ao ler arquivo');
                    }
                    
                    const workbook = XLSX.read(e.target.result, { type: 'array' });
                    processWorkbook(workbook);
                    downloadWorkbook(workbook, file);
                    resolve();
                } catch (error) {
                    reject(new Error(`Falha ao processar ${file.name}: ${error.message}`));
                }
            };

            reader.onerror = () => reject(new Error('Erro ao ler o arquivo'));
            reader.readAsArrayBuffer(file);
        });
    }

    // Percorre as abas da planilha
    function processWorkbook(workbook) {
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error('Planilha vazia ou inválida');
        }

        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            if (worksheet) {
                processWorksheet(worksheet);
            }
        });
    }

    // Converte células para texto e formata K/L/M/S/T/U
    function processWorksheet(worksheet) {
        if (!worksheet['!ref']) return; // Planilha vazia
        
        const maxLengths = {};

        // Converte todas as células; K/L/M/S/T/U recebem 4 casas decimais
        Object.keys(worksheet).forEach(key => {
            if (key.startsWith('!')) return; // Ignorar metadados

            const cell = worksheet[key];
            if (!cell?.v) return;

            const col = XLSX.utils.decode_cell(key).c;
            cell.v = formatCellValue(cell.v, col);
            cell.t = 's'; // Tipo texto
            cell.z = '@';

            // Rastreia comprimentos para ajuste de coluna
            const displayLength = String(cell.v).length;
            maxLengths[col] = Math.max(maxLengths[col] || 0, displayLength);

            // Remove fórmulas
            if (cell.f) delete cell.f;
        });

        adjustColumnWidths(worksheet, maxLengths);
    }

    // Formata valor da célula: K/L/M/S/T/U com 4 casas (vírgula); demais como texto
    function formatCellValue(value, columnIndex) {
        const stringValue = String(value).trim();
        
        if (CONFIG.targetColumns.includes(columnIndex)) {
            const parsed = parseNumberLike(stringValue);
            if (parsed != null) {
                return parsed.toFixed(CONFIG.decimalPlaces).replace('.', ',');
            }
            return formatDecimalTo4(stringValue);
        }
        
        return stringValue;
    }

    // Ajusta larguras de coluna pelo conteúdo
    function adjustColumnWidths(worksheet, maxLengths) {
        worksheet['!cols'] = worksheet['!cols'] || [];
        
        // Ajusta colunas de A até X (0 até 23)
        for (let col = 0; col <= 23; col++) {
            const existingWidth = worksheet['!cols'][col]?.wch || CONFIG.minColumnWidth;
            const contentWidth = (maxLengths[col] || 0) + 2;
            
            let finalWidth = Math.min(CONFIG.maxColumnWidth, Math.max(existingWidth, contentWidth));
            
            // Colunas A-H: reduzir tamanho
            if (col <= 7) {
                finalWidth *= 0.6;
            }
            
            worksheet['!cols'][col] = { wch: finalWidth };
        }
    }

    // Gera o arquivo final e baixa
    function downloadWorkbook(workbook, originalFile) {
        const extension = originalFile.name.split('.').pop().toLowerCase();
        const baseName = originalFile.name.replace(/\.[^/.]+$/, '');
        
        let blob, filename;

        if (extension === 'csv') {
            const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]);
            blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
            filename = `${baseName}_fixed.csv`;
        } else {
            const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
            blob = new Blob([wbout], { type: 'application/octet-stream' });
            filename = `${baseName}_fixed.xlsx`;
        }

        triggerDownload(blob, filename);
    }

    // Inicia o download de um Blob
    function triggerDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // Cleanup
        URL.revokeObjectURL(url);
    }

    // Tenta interpretar número (EN-US e PT-BR). Retorna número ou null.
    function parseNumberLike(s) {
        if (s == null) return null;
        
        const str = String(s).trim();
        if (!str) return null;

        try {
            const hasDot = str.includes('.');
            const hasComma = str.includes(',');
            let cleaned;

            if (hasDot && hasComma) {
                // Ambos separadores presentes: usar o último como decimal
                cleaned = str.lastIndexOf(',') > str.lastIndexOf('.')
                    ? str.replace(/\./g, '').replace(',', '.')  // PT-BR
                    : str.replace(/,/g, '');                     // EN-US
            } else if (hasComma) {
                // Apenas vírgula: tratar como separador decimal (PT-BR)
                cleaned = str.replace(/\./g, '').replace(',', '.');
            } else {
                // Apenas ponto ou nenhum: usar como está
                cleaned = str.replace(/,/g, '');
            }

            const num = Number(cleaned.replace(/\s/g, ''));
            return Number.isFinite(num) ? num : null;
        } catch {
            return null;
        }
    }

    // Formata com 4 casas; mantém original se inválido
    function formatDecimalTo4(value) {
        const str = String(value).trim();
        if (!str) return str;

        const normalized = str.replace(',', '.');
        // Validar se é número válido
        if (!/^[+-]?\d+(\.\d+)?$/.test(normalized)) return str;

        const [intPart, fracPart = ''] = normalized.split('.');
        const frac = fracPart.length < 4 
            ? fracPart.padEnd(4, '0') 
            : fracPart.slice(0, 4);
        
        return `${intPart},${frac}`;
    }

    // Formata bytes em B, KB, MB, GB
    function formatBytes(bytes) {
        if (bytes === 0) return '0 B';
        
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return `${(bytes / Math.pow(k, i)).toFixed(2)} ${sizes[i]}`;
    }

    
})();