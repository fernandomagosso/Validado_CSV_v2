
import { GoogleGenAI, Type } from "@google/genai";

// --- STATE MANAGEMENT ---
let apiKey = '';
let ai;
let csvData = []; // Array of objects, where each object is a row
let csvHeaders = [];
let htmlTemplate = ''; // Stores the user-provided HTML template string
let htmlFieldMapping = {}; // Stores mapping from {placeholder} to csvHeader
let activeRowIndex = null;
let layoutMode = null; // 'ai' or 'html'
let previewMode = null; // 'single' or 'bulk'
let layoutConfigPromise = {}; // Used to resolve layout choice from modal
let aiLayoutInstructionsText = ''; // Stores user instructions for AI layout
let currentTourStep = 0;
// Cache for AI-generated single previews to reduce API calls.
let aiPreviewCache = new Map(); 

// --- DOM ELEMENTS ---
const apiKeyInput = document.getElementById('apiKey');
const saveApiKeyBtn = document.getElementById('saveApiKey');
const csvFileInput = document.getElementById('csvFile');
const runAiBtn = document.getElementById('runAi');
const downloadCsvBtn = document.getElementById('downloadCsv');
const downloadDocBtn = document.getElementById('downloadDoc');
const clearDataBtn = document.getElementById('clearDataBtn');
const clearLayoutBtn = document.getElementById('clearLayoutBtn');
const dataGridContainer = document.getElementById('data-grid-container');
const docPreviewContent = document.getElementById('doc-preview-content');
const aiOutputContainer = document.getElementById('ai-output-container');
const aiOutputContent = document.getElementById('ai-output-content');
const welcomeMessage = document.querySelector('.welcome-message');
const modeChoiceContainer = document.getElementById('mode-choice-container');
const bulkPreviewBtn = document.getElementById('bulk-preview-btn');
const changeModeBtn = document.getElementById('changeModeBtn');
const mappingModalOverlay = document.getElementById('mapping-modal-overlay');
const mappingFormContainer = document.getElementById('mapping-form-container');
const confirmMappingBtn = document.getElementById('confirm-mapping-btn');
const cancelMappingBtn = document.getElementById('cancel-mapping-btn');
const aiPromptModalOverlay = document.getElementById('ai-prompt-modal-overlay');
const aiPromptTextarea = document.getElementById('aiPromptTextarea');
const confirmAiPromptBtn = document.getElementById('confirm-ai-prompt-btn');
const cancelAiPromptBtn = document.getElementById('cancel-ai-prompt-btn');
const layoutConfigModalOverlay = document.getElementById('layout-config-modal-overlay');
const useAiLayoutBtn = document.getElementById('useAiLayoutBtn');
const useHtmlTemplateBtn = document.getElementById('useHtmlTemplateBtn');
const htmlTemplateTextarea = document.getElementById('htmlTemplateTextarea');
const aiLayoutInstructions = document.getElementById('aiLayoutInstructions');
const aiTransformRulesSection = document.getElementById('ai-transform-rules-section');
const aiTransformRulesTextarea = document.getElementById('aiTransformRulesTextarea');
const applyTransformRulesBtn = document.getElementById('applyTransformRulesBtn');
const generateAnalysisBtn = document.getElementById('generateAnalysisBtn');
const analysisModalOverlay = document.getElementById('analysis-modal-overlay');
const analysisSummaryContent = document.getElementById('analysis-summary-content');
const analysisChartContent = document.getElementById('analysis-chart-content');
const closeAnalysisModalBtn = document.getElementById('close-analysis-modal-btn');
const tourOverlay = document.getElementById('tour-overlay');
const tourPopover = document.getElementById('tour-popover');
const tourTitle = document.getElementById('tour-title');
const tourText = document.getElementById('tour-text');
const tourStepCounter = document.getElementById('tour-step-counter');
const tourPrevBtn = document.getElementById('tour-prev');
const tourNextBtn = document.getElementById('tour-next');
const mainViewContainer = document.getElementById('main-view-container');


// --- TOUR DEFINITION ---
const tourSteps = [
    { elementId: 'tour-step-1', title: 'Boas-vindas!', text: 'Este é o cabeçalho. O primeiro passo é sempre carregar seu arquivo de dados clicando em "Carregar CSV".' },
    { elementId: 'tour-step-2', title: 'Chave da API', text: 'Insira sua chave da API Gemini aqui. Ela é essencial para ativar todas as funcionalidades de inteligência artificial, como validação e análise.' },
    { elementId: 'tour-step-3', title: 'Grade de Dados', text: 'Seus dados do CSV aparecerão aqui. Você pode clicar em qualquer célula para editar as informações diretamente.' },
    { elementId: 'tour-step-4', title: 'Pré-visualização do Documento', text: 'A pré-visualização do seu documento, seja com um modelo ou gerada por IA, será exibida nesta área.' },
    { elementId: 'tour-step-5', title: 'Controles da IA', text: 'Depois de carregar os dados e configurar um layout, use os botões nesta seção para executar validações e análises com a IA.' }
];

// --- UTILITY FUNCTIONS ---
function debounce(func, delay) {
    let timeout;
    return function(...args) {
        const context = this;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(context, args), delay);
    };
}

function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = message;
    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('hide');
        toast.addEventListener('transitionend', () => toast.remove());
    }, 5000);
}

function sanitizeHtml(str) {
    const temp = document.createElement('div');
    temp.textContent = str;
    return temp.innerHTML;
}

// --- INITIALIZATION ---
function initializeApp() {
  const savedApiKey = sessionStorage.getItem('geminiApiKey');
  if (savedApiKey) {
    apiKey = savedApiKey;
    apiKeyInput.value = apiKey;
    try {
        ai = new GoogleGenAI({ apiKey });
        console.log("API Key loaded from session.");
    } catch (e) {
        console.error("Failed to initialize GoogleGenAI:", e);
        showToast("Falha ao inicializar a API. Verifique sua chave.", 'error');
    }
  }
  
  addEventListeners();
  updateButtonStates();
  
  if (!sessionStorage.getItem('tourShown')) {
      startTour();
  }
}

// Debounced version of the render function to prevent API spam
const debouncedRenderSinglePreview = debounce(renderSinglePreview, 500);

// --- EVENT LISTENERS ---
function addEventListeners() {
    saveApiKeyBtn.addEventListener('click', handleSaveApiKey);
    csvFileInput.addEventListener('change', handleCsvUpload);
    bulkPreviewBtn.addEventListener('click', handleBulkPreviewClick);
    useAiLayoutBtn.addEventListener('click', () => handleLayoutChoice('ai'));
    useHtmlTemplateBtn.addEventListener('click', () => handleLayoutChoice('html'));
    changeModeBtn.addEventListener('click', handleChangeMode);
    runAiBtn.addEventListener('click', handleAiValidation);
    downloadCsvBtn.addEventListener('click', handleDownloadCsv);
    downloadDocBtn.addEventListener('click', handleDownloadZip);
    clearDataBtn.addEventListener('click', handleClearData);
    clearLayoutBtn.addEventListener('click', handleClearLayout);
    applyTransformRulesBtn.addEventListener('click', applyBulkTransformations);
    generateAnalysisBtn.addEventListener('click', handleGenerateAnalysis);

    // Modal close buttons
    document.querySelectorAll('.modal-close-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            btn.closest('.modal-overlay').style.display = 'none';
        });
    });

    confirmMappingBtn.addEventListener('click', handleConfirmMapping);
    cancelMappingBtn.addEventListener('click', () => mappingModalOverlay.style.display = 'none');
    
    // Tour listeners
    tourOverlay.addEventListener('click', (e) => {
        if (e.target === tourOverlay || e.target.classList.contains('tour-close-btn')) endTour();
    });
    tourPrevBtn.addEventListener('click', () => showTourStep(currentTourStep - 1));
    tourNextBtn.addEventListener('click', () => showTourStep(currentTourStep + 1));
}

// --- CORE LOGIC ---

function handleSaveApiKey() {
    const key = apiKeyInput.value.trim();
    if (key) {
        apiKey = key;
        try {
            ai = new GoogleGenAI({ apiKey });
            sessionStorage.setItem('geminiApiKey', apiKey);
            showToast('Chave da API salva com sucesso!');
            updateButtonStates();
        } catch (e) {
            console.error("Failed to initialize GoogleGenAI:", e);
            showToast("Falha ao inicializar com a nova chave de API.", 'error');
        }
    } else {
        showToast('Por favor, insira uma chave de API válida.', 'error');
    }
}

async function handleCsvUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const text = e.target.result;
            const rows = text.split(/\r?\n/).filter(row => row.trim() !== '');
            if (rows.length < 2) throw new Error("CSV precisa de cabeçalho e pelo menos uma linha de dados.");

            csvHeaders = rows[0].split(';');
            csvData = rows.slice(1).map(row => {
                const values = row.split(';');
                return csvHeaders.reduce((obj, header, index) => {
                    obj[header] = values[index] || '';
                    return obj;
                }, {});
            });

            renderDataGrid();
            handleClearLayout(); // Reset layout when new data is loaded
            showToast(`CSV com ${csvData.length} registros carregado.`);
        } catch (error) {
            console.error("Error parsing CSV:", error);
            showToast(`Erro ao ler o CSV: ${error.message}`, 'error');
            handleClearData();
        }
    };
    reader.readAsText(file);
    csvFileInput.value = '';
}

function renderDataGrid() {
    const table = document.querySelector('.data-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    thead.innerHTML = `<tr>${csvHeaders.map(h => `<th>${sanitizeHtml(h)}</th>`).join('')}</tr>`;
    tbody.innerHTML = '';

    csvData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        tr.dataset.rowIndex = rowIndex;
        tr.innerHTML = csvHeaders.map(header => `<td contenteditable="true">${sanitizeHtml(row[header])}</td>`).join('');
        tbody.appendChild(tr);
    });

    tbody.addEventListener('click', handleRowClick);
    tbody.addEventListener('input', handleCellEdit);
}

function handleRowClick(event) {
    const tr = event.target.closest('tr');
    if (!tr || !tr.dataset.rowIndex) return;

    const rowIndex = parseInt(tr.dataset.rowIndex, 10);
    
    if (activeRowIndex !== null) {
        const previousRow = document.querySelector(`tr[data-row-index='${activeRowIndex}']`);
        if (previousRow) previousRow.classList.remove('active');
    }
    
    tr.classList.add('active');
    activeRowIndex = rowIndex;
    updateButtonStates();

    if (layoutMode) {
        previewMode = 'single';
        changeModeBtn.textContent = 'Ver Todos';
        debouncedRenderSinglePreview(rowIndex);
    }
}

async function renderSinglePreview(rowIndex) {
    if (rowIndex === null || !csvData[rowIndex]) return;

    docPreviewContent.innerHTML = `<div class="welcome-message"><p>Gerando pré-visualização...</p></div>`;
    const rowData = csvData[rowIndex];
    let previewHtml = '';

    try {
        if (layoutMode === 'ai') {
            if (aiPreviewCache.has(rowIndex)) {
                previewHtml = aiPreviewCache.get(rowIndex);
                showToast("Layout carregado do cache.");
            } else {
                previewHtml = await generatePreviewWithAI(rowData, rowIndex);
                if (previewHtml) aiPreviewCache.set(rowIndex, previewHtml);
            }
        } else if (layoutMode === 'html') {
            previewHtml = generatePreviewWithHtmlTemplate(rowData);
        }
    } catch (error) {
        console.error("Error rendering single preview:", error);
        previewHtml = `<p style="color:red;">Falha ao gerar preview: ${error.message}</p>`;
    }

    docPreviewContent.innerHTML = `<div class="preview-document">${previewHtml}</div>`;
}

function handleCellEdit(event) {
    const td = event.target;
    const tr = td.closest('tr');
    if (!tr) return;

    const rowIndex = parseInt(tr.dataset.rowIndex, 10);
    const cellIndex = td.cellIndex;
    const header = csvHeaders[cellIndex];
    
    csvData[rowIndex][header] = td.textContent;
    aiPreviewCache.delete(rowIndex); // Invalidate cache for this row
}

function updateButtonStates() {
    const hasData = csvData.length > 0;
    const hasApiKey = !!apiKey;
    const hasLayout = !!layoutMode;
    const rowSelected = activeRowIndex !== null;

    clearDataBtn.disabled = !hasData;
    clearLayoutBtn.disabled = !hasLayout;
    downloadCsvBtn.disabled = !hasData;
    downloadDocBtn.disabled = !hasData || !hasLayout;
    generateAnalysisBtn.disabled = !hasData || !hasApiKey;
    runAiBtn.disabled = !hasData || !hasApiKey || !rowSelected;
    
    if (hasData) {
        welcomeMessage.style.display = 'none';
        modeChoiceContainer.style.display = 'flex';
        changeModeBtn.style.display = 'block';
    } else {
        welcomeMessage.style.display = 'flex';
        modeChoiceContainer.style.display = 'none';
        changeModeBtn.style.display = 'none';
    }
}

function handleApiError(error, button) {
    console.error("Gemini API Error:", error);
    let userMessage = "Ocorreu um erro na chamada da API.";
    if (error.message && error.message.includes('429')) {
        userMessage = "Limite de requisições da API atingido. Tente novamente em um minuto.";
    } else if (error.message) {
        userMessage = `Erro da API: ${error.message.substring(0, 100)}`;
    }
    showToast(userMessage, 'error');
    if (button) {
        button.classList.remove('btn-loading');
        button.disabled = false;
    }
}

// --- PREVIEW & LAYOUT ---

async function handleBulkPreviewClick() {
    if (!layoutMode) {
        await promptForLayoutChoice();
        if (!layoutMode) return;
    }
    
    previewMode = 'bulk';
    changeModeBtn.textContent = 'Ver Individual';
    bulkPreviewBtn.classList.add('btn-loading');
    bulkPreviewBtn.disabled = true;

    const wrapper = document.createElement('div');
    wrapper.className = 'bulk-preview-wrapper';
    docPreviewContent.innerHTML = '';
    docPreviewContent.appendChild(wrapper);

    // Process sequentially to avoid rate limiting
    for (let i = 0; i < csvData.length; i++) {
        wrapper.innerHTML = `<p>Processando registro ${i + 1} de ${csvData.length}...</p>${wrapper.innerHTML}`;
        const rowData = csvData[i];
        let previewHtml = '';

        try {
            if (layoutMode === 'ai') {
                previewHtml = await generatePreviewWithAI(rowData, i);
            } else {
                previewHtml = generatePreviewWithHtmlTemplate(rowData);
            }
            
            const docElement = document.createElement('div');
            docElement.className = 'preview-document';
            docElement.innerHTML = previewHtml;
            wrapper.prepend(docElement);
        
            if (layoutMode === 'ai' && i < csvData.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 1100)); // Stay under 60 RPM
            }
        } catch (error) {
            handleApiError(error, bulkPreviewBtn);
            const errorElement = document.createElement('div');
            errorElement.className = 'preview-document';
            errorElement.innerHTML = `<p style="color:red">Erro ao processar linha ${i + 1}.</p>`;
            wrapper.prepend(errorElement);
            // Stop processing on API error
            break; 
        } finally {
            wrapper.querySelector('p')?.remove();
        }
    }
    
    bulkPreviewBtn.classList.remove('btn-loading');
    bulkPreviewBtn.disabled = false;
}

function handleChangeMode() {
    if (previewMode === 'bulk') {
        previewMode = 'single';
        changeModeBtn.textContent = 'Ver Todos';
        if (activeRowIndex !== null) {
            renderSinglePreview(activeRowIndex);
        } else {
            docPreviewContent.innerHTML = '<p>Selecione uma linha para ver a pré-visualização individual.</p>';
        }
    } else {
        handleBulkPreviewClick();
    }
}

async function generatePreviewWithAI(rowData, rowIndex) {
    if (!ai) {
        showToast('API não inicializada. Salve sua chave primeiro.', 'error');
        return '';
    }
    
    const prompt = `
        Com base nestes dados de uma linha de um CSV:
        ${JSON.stringify(rowData, null, 2)}

        E estas instruções de layout: "${aiLayoutInstructionsText || 'Crie um layout limpo e profissional.'}"
        
        Gere APENAS o código HTML para um documento que represente esses dados. 
        - O HTML não deve conter <html>, <head>, ou <body>. Apenas o conteúdo interno.
        - Use estilos inline para garantir que a aparência seja consistente.
        - Não adicione comentários ou explicações. Apenas o código HTML.
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
        });
        return response.text;
    } catch (error) {
        handleApiError(error);
        return `<p style="color:red;">Falha ao gerar preview com IA para a linha ${rowIndex + 1}.</p>`;
    }
}

function generatePreviewWithHtmlTemplate(rowData) {
    let populatedTemplate = htmlTemplate;
    for (const placeholder in htmlFieldMapping) {
        const csvHeader = htmlFieldMapping[placeholder];
        const value = rowData[csvHeader] || '';
        populatedTemplate = populatedTemplate.replace(new RegExp(`{${placeholder}}`, 'g'), sanitizeHtml(value));
    }
    return populatedTemplate;
}

// --- AI FEATURES ---

async function handleAiValidation() {
    if (activeRowIndex === null || !ai) return;

    runAiBtn.classList.add('btn-loading');
    runAiBtn.disabled = true;

    const rowData = csvData[activeRowIndex];
    const prompt = `
        Você é um assistente de validação de dados. Analise esta linha de dados de um CSV:
        ${JSON.stringify(rowData, null, 2)}
        
        Com base nesses dados, identifique possíveis erros, inconsistências ou campos que precisam de atenção.
        Seja conciso e direto. Retorne APENAS uma lista de problemas em formato JSON.
        O JSON deve ser um array de objetos, onde cada objeto tem duas chaves: "field" (o nome da coluna com problema) e "issue" (a descrição do problema).
        Se nenhum problema for encontrado, retorne um array vazio [].
        
        Exemplo de saída com erros:
        [
          {"field": "CPF", "issue": "O CPF parece ser inválido ou está mal formatado."},
          {"field": "data_nascimento", "issue": "O formato da data é inconsistente."}
        ]

        Exemplo de saída sem erros:
        []
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: { responseMimeType: "application/json" }
        });
        const validationResult = JSON.parse(response.text);
        displayValidationResults(validationResult);
    } catch (error) {
        handleApiError(error, runAiBtn);
    } finally {
        runAiBtn.classList.remove('btn-loading');
        runAiBtn.disabled = false;
    }
}

async function handleGenerateAnalysis() {
    if (!ai || csvData.length === 0) return;
    
    generateAnalysisBtn.classList.add('btn-loading');
    generateAnalysisBtn.disabled = true;

    const dataSample = csvData.slice(0, 20); // Use a sample for large files
    const prompt = `
        Analise o seguinte conjunto de dados (amostra de um arquivo CSV):
        ${JSON.stringify(dataSample, null, 2)}

        Forneça um resumo executivo e uma análise de infográfico.
        A resposta DEVE ser um único objeto JSON com duas chaves: "summary" e "chartData".
        - "summary": Uma string contendo o resumo em HTML (use <p>, <strong>, <ul>, <li>).
        - "chartData": Um array de objetos, cada um com "label" (string) e "value" (número), representando uma métrica importante (ex: contagem por categoria, média de valores). Escolha a métrica mais relevante dos dados. Limite a 5 itens.

        Exemplo de Resposta:
        {
          "summary": "<p>A análise revela <strong>3 clientes</strong> distintos. A maioria das compras está na categoria 'Eletrônicos'.</p><ul><li>Ticket médio: R$ 150,00</li></ul>",
          "chartData": [
            {"label": "Eletrônicos", "value": 2},
            {"label": "Livros", "value": 1}
          ]
        }
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: { responseMimeType: "application/json" }
        });
        const analysis = JSON.parse(response.text);
        displayAnalysisModal(analysis);
    } catch (error) {
        handleApiError(error, generateAnalysisBtn);
    } finally {
        generateAnalysisBtn.classList.remove('btn-loading');
        generateAnalysisBtn.disabled = false;
    }
}

function displayValidationResults(results) {
    // Clear previous highlights
    document.querySelectorAll('.invalid-field').forEach(el => {
        el.classList.remove('invalid-field');
        delete el.dataset.tooltip;
    });

    if (results.length === 0) {
        showToast('Nenhum problema encontrado!', 'success');
        return;
    }

    showToast(`${results.length} problema(s) de validação encontrado(s).`, 'error');

    const activeRowElement = document.querySelector(`tr[data-row-index='${activeRowIndex}']`);
    if (!activeRowElement) return;

    results.forEach(item => {
        const headerIndex = csvHeaders.indexOf(item.field);
        if (headerIndex !== -1) {
            const cell = activeRowElement.cells[headerIndex];
            cell.classList.add('invalid-field');
            cell.dataset.tooltip = item.issue;
        }
    });
}


// --- MODALS & CONFIG ---

function promptForLayoutChoice() {
    layoutConfigModalOverlay.style.display = 'flex';
    return new Promise((resolve) => {
        layoutConfigPromise.resolve = resolve;
    });
}

function handleLayoutChoice(choice) {
    if (choice === 'ai') {
        layoutMode = 'ai';
        aiLayoutInstructionsText = aiLayoutInstructions.value;
        showToast('Modo de layout IA ativado.');
    } else if (choice === 'html') {
        const template = htmlTemplateTextarea.value.trim();
        if (!template) {
            showToast('Por favor, insira um template HTML.', 'error');
            return;
        }
        htmlTemplate = template;
        const placeholders = htmlTemplate.match(/\{(\w+)\}/g)?.map(p => p.slice(1, -1)) || [];
        if (placeholders.length === 0) {
            showToast('Nenhum placeholder (ex: {NOME}) encontrado no template.', 'error');
            return;
        }
        layoutMode = 'html';
        showMappingModal(placeholders);
    }
    layoutConfigModalOverlay.style.display = 'none';
    if (layoutConfigPromise.resolve) {
        layoutConfigPromise.resolve();
        layoutConfigPromise = {};
    }
    updateButtonStates();
}

function showMappingModal(placeholders) {
    mappingFormContainer.innerHTML = '';
    const uniquePlaceholders = [...new Set(placeholders)];

    uniquePlaceholders.forEach(ph => {
        const row = document.createElement('div');
        row.className = 'mapping-row';
        
        const options = csvHeaders.map(h => `<option value="${h}">${h}</option>`).join('');
        
        row.innerHTML = `
            <label for="map-${ph}">{${ph}}</label>
            <select id="map-${ph}" data-placeholder="${ph}">
                <option value="">Selecione uma coluna</option>
                ${options}
            </select>
        `;
        mappingFormContainer.appendChild(row);
    });
    mappingModalOverlay.style.display = 'flex';
}

function handleConfirmMapping() {
    htmlFieldMapping = {};
    let allMapped = true;
    mappingFormContainer.querySelectorAll('select').forEach(select => {
        const placeholder = select.dataset.placeholder;
        if (select.value) {
            htmlFieldMapping[placeholder] = select.value;
        } else {
            allMapped = false;
        }
    });

    if (!allMapped) {
        showToast('Por favor, mapeie todos os placeholders.', 'error');
        return;
    }
    
    mappingModalOverlay.style.display = 'none';
    showToast('Mapeamento salvo com sucesso!');
}

function displayAnalysisModal({ summary, chartData }) {
    analysisSummaryContent.innerHTML = summary;
    analysisChartContent.innerHTML = ''; // Clear previous

    if (chartData && chartData.length > 0) {
        const maxValue = Math.max(...chartData.map(d => d.value));
        
        chartData.forEach(item => {
            const percentage = (item.value / maxValue) * 100;
            const barRow = `
                <div class="chart-bar-row">
                    <div class="chart-label" title="${item.label}">${item.label}</div>
                    <div class="chart-bar-container">
                        <div class="chart-bar" style="width: ${percentage}%;">${item.value}</div>
                    </div>
                </div>
            `;
            analysisChartContent.innerHTML += barRow;
        });
    } else {
        analysisChartContent.innerHTML = "<p>Nenhuma métrica visualizável foi extraída.</p>";
    }
    
    analysisModalOverlay.style.display = 'flex';
}


// --- DATA MANAGEMENT ---

function handleClearData() {
    csvData = [];
    csvHeaders = [];
    activeRowIndex = null;
    document.querySelector('.data-table thead').innerHTML = '';
    document.querySelector('.data-table tbody').innerHTML = `<tr><td class="placeholder-cell">Carregue um arquivo CSV para visualizar os dados aqui.</td></tr>`;
    handleClearLayout(); // Also clear layout
    showToast('Dados limpos.');
}

function handleClearLayout() {
    layoutMode = null;
    previewMode = null;
    htmlTemplate = '';
    htmlFieldMapping = {};
    aiLayoutInstructionsText = '';
    aiPreviewCache.clear();
    docPreviewContent.innerHTML = '';
    docPreviewContent.appendChild(welcomeMessage);
    docPreviewContent.appendChild(modeChoiceContainer);
    updateButtonStates();
    showToast('Configuração de layout limpa.');
}

function handleDownloadCsv() {
    const csvContent = [
        csvHeaders.join(';'),
        ...csvData.map(row => csvHeaders.map(header => row[header]).join(';'))
    ].join('\n');
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "dados_editados.csv");
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

async function handleDownloadZip() {
    if (!layoutMode) {
        showToast('Configure um layout primeiro.', 'error');
        return;
    }
    
    const zip = new PizZip();
    downloadDocBtn.classList.add('btn-loading');
    downloadDocBtn.disabled = true;

    for (let i = 0; i < csvData.length; i++) {
        const rowData = csvData[i];
        let content = '';
        if (layoutMode === 'ai') {
            content = await generatePreviewWithAI(rowData, i);
            if (i < csvData.length - 1) { // Rate limit
                await new Promise(resolve => setTimeout(resolve, 1100));
            }
        } else {
            content = generatePreviewWithHtmlTemplate(rowData);
        }
        const fullHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Documento ${i+1}</title></head><body>${content}</body></html>`;
        zip.file(`documento_${i + 1}.html`, fullHtml);
    }

    const zipContent = zip.generate({ type: "blob" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(zipContent);
    link.setAttribute("href", url);
    link.setAttribute("download", "documentos.zip");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    downloadDocBtn.classList.remove('btn-loading');
    downloadDocBtn.disabled = false;
}

// --- TOUR LOGIC ---
function startTour() {
    sessionStorage.setItem('tourShown', 'true');
    currentTourStep = 0;
    tourOverlay.style.display = 'block';
    showTourStep(0);
}

function endTour() {
    tourOverlay.style.display = 'none';
    const highlighted = document.querySelector('.tour-highlight');
    if (highlighted) highlighted.classList.remove('tour-highlight');
}

function showTourStep(stepIndex) {
    if (stepIndex < 0 || stepIndex >= tourSteps.length) {
        endTour();
        return;
    }
    
    const highlighted = document.querySelector('.tour-highlight');
    if (highlighted) highlighted.classList.remove('tour-highlight');
    
    currentTourStep = stepIndex;
    const step = tourSteps[stepIndex];
    const element = document.querySelector(`[data-tour-id="${step.elementId}"]`);

    if (!element) {
        endTour();
        return;
    }
    
    element.classList.add('tour-highlight');
    const rect = element.getBoundingClientRect();
    
    tourPopover.style.top = `${rect.bottom + 15}px`;
    tourPopover.style.left = `${rect.left}px`;
    
    // Adjust popover position if it goes off-screen
    if (rect.left + tourPopover.offsetWidth > window.innerWidth) {
        tourPopover.style.left = `${rect.right - tourPopover.offsetWidth}px`;
    }
    if (rect.bottom + tourPopover.offsetHeight > window.innerHeight) {
        tourPopover.style.top = `${rect.top - tourPopover.offsetHeight - 15}px`;
    }

    tourTitle.textContent = step.title;
    tourText.textContent = step.text;
    tourStepCounter.textContent = `Passo ${stepIndex + 1} de ${tourSteps.length}`;
    tourPrevBtn.style.visibility = stepIndex === 0 ? 'hidden' : 'visible';
    tourNextBtn.textContent = stepIndex === tourSteps.length - 1 ? 'Finalizar' : 'Próximo';
}

// --- BULK TRANSFORM (Placeholder) ---
async function applyBulkTransformations() {
    showToast('Funcionalidade de transformação em massa ainda não implementada.', 'error');
}


// --- START ---
initializeApp();
