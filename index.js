import { GoogleGenAI, Type } from "@google/genai";

// --- STATE MANAGEMENT ---
let apiKey = '';
let ai;
let csvData = []; // Array of objects, where each object is a row
let csvHeaders = [];
let templateFileContent = null; // Will store the ArrayBuffer of the template
let activeRowIndex = null;
let layoutMode = null; // 'ai' or 'template'
let previewMode = null; // 'single' or 'bulk'
let fieldMapping = {}; // Stores the confirmed mapping from template placeholder to CSV header
let layoutConfigPromise = {}; // Used to resolve layout choice from modal
let aiLayoutInstructionsText = ''; // Stores user instructions for AI layout
let currentTourStep = 0;

// --- DOM ELEMENTS ---
const apiKeyInput = document.getElementById('apiKey');
const saveApiKeyBtn = document.getElementById('saveApiKey');
const csvFileInput = document.getElementById('csvFile');
const templateFileInputHidden = document.getElementById('templateFileInputHidden');
const aiControls = document.getElementById('ai-controls');
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


// --- TOUR DEFINITION ---
const tourSteps = [
    {
        elementId: 'tour-step-1',
        title: 'Boas-vindas!',
        text: 'Este é o cabeçalho. O primeiro passo é sempre carregar seu arquivo de dados clicando em "Carregar CSV".'
    },
    {
        elementId: 'tour-step-2',
        title: 'Chave da API',
        text: 'Insira sua chave da API Gemini aqui. Ela é essencial para ativar todas as funcionalidades de inteligência artificial, como validação e análise.'
    },
    {
        elementId: 'tour-step-3',
        title: 'Grade de Dados',
        text: 'Seus dados do CSV aparecerão aqui. Você pode clicar em qualquer célula para editar as informações diretamente.'
    },
    {
        elementId: 'tour-step-4',
        title: 'Pré-visualização do Documento',
        text: 'A pré-visualização do seu documento, seja com um modelo ou gerada por IA, será exibida nesta área.'
    },
    {
        elementId: 'tour-step-5',
        title: 'Controles da IA',
        text: 'Depois de carregar os dados e configurar um layout, use os botões na barra lateral para executar validações e análises com a IA.'
    }
];


// --- INITIALIZATION ---
function initializeApp() {
  const savedApiKey = sessionStorage.getItem('geminiApiKey');
  if (savedApiKey) {
    apiKey = savedApiKey;
    apiKeyInput.value = apiKey;
    ai = new GoogleGenAI({ apiKey });
    console.log("API Key loaded from session.");
  }
  
  addEventListeners();
  updateButtonStates();
  
  if (!sessionStorage.getItem('tourShown')) {
      startTour();
  }
}

// --- EVENT LISTENERS ---
function addEventListeners() {
  saveApiKeyBtn.addEventListener('click', handleSaveApiKey);
  csvFileInput.addEventListener('change', handleCsvUpload);
  bulkPreviewBtn.addEventListener('click', handleBulkPreviewClick);
  useAiLayoutBtn.addEventListener('click', () => handleLayoutChoice('ai'));
  templateFileInputHidden.addEventListener('change', (e) => handleLayoutChoice('template', e));
  changeModeBtn.addEventListener('click', handleChangeMode);
  runAiBtn.addEventListener('click', showAiPromptModal);
  downloadCsvBtn.addEventListener('click', handleDownloadCsv);
  downloadDocBtn.addEventListener('click', handleDownloadZip);
  clearDataBtn.addEventListener('click', handleClearData);
  clearLayoutBtn.addEventListener('click', handleClearLayout);
  confirmMappingBtn.addEventListener('click', handleMappingConfirmation);
  cancelMappingBtn.addEventListener('click', () => mappingModalOverlay.style.display = 'none');
  confirmAiPromptBtn.addEventListener('click', handleAiPrompt);
  cancelAiPromptBtn.addEventListener('click', () => aiPromptModalOverlay.style.display = 'none');
  applyTransformRulesBtn.addEventListener('click', handleApplyTransformRules);
  generateAnalysisBtn.addEventListener('click', handleGenerateAnalysis);
  closeAnalysisModalBtn.addEventListener('click', () => analysisModalOverlay.style.display = 'none');
  tourNextBtn.addEventListener('click', nextTourStep);
  tourPrevBtn.addEventListener('click', prevTourStep);
}

// --- HANDLER FUNCTIONS ---

function handleSaveApiKey() {
  const key = apiKeyInput.value.trim();
  if (key) {
    apiKey = key;
    sessionStorage.setItem('geminiApiKey', apiKey);
    ai = new GoogleGenAI({ apiKey });
    showToast('Chave API salva com sucesso.');
    document.getElementById('api-section').classList.remove('needs-attention');
    console.log("API Key saved.");
  } else {
    showToast('Por favor, insira uma chave API válida.', 'error');
  }
}

async function handleCsvUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const text = e.target.result;
    parseCsv(text);
    renderDataTable(); 
    
    resetPreviewState();
    welcomeMessage.style.display = 'none';
    modeChoiceContainer.style.display = 'flex';
    docPreviewContent.innerHTML = '';
    docPreviewContent.appendChild(modeChoiceContainer);

    showToast('Arquivo CSV carregado com sucesso.');
    updateButtonStates();
  };
  reader.readAsText(file);
  event.target.value = ''; // Reset input so the same file can be reloaded
}

function handleBulkPreviewClick() {
    previewMode = 'bulk';
    showLayoutConfigModal().then(success => {
        if(success) renderBulkPreview();
    });
}

function handleRowClick(rowIndex, renderPreview = true) {
    previewMode = 'single';
    
    const currentActiveRow = dataGridContainer.querySelector('.data-table tbody tr.active');
    currentActiveRow?.classList.remove('active');
    
    const newActiveRow = dataGridContainer.querySelector(`tr[data-row-index="${rowIndex}"]`);
    newActiveRow?.classList.add('active');
    
    activeRowIndex = rowIndex;
    
    if (renderPreview) {
      if (!layoutMode) {
           showLayoutConfigModal().then(success => {
              if(success) renderSinglePreview(rowIndex);
          });
      } else {
          renderSinglePreview(rowIndex);
      }
    }
}


function handleLayoutChoice(choice, event = null) {
    if (choice === 'ai') {
        if (!ai) {
            showToast('Por favor, salve sua chave da API Gemini primeiro.', 'error');
            layoutConfigPromise.reject?.();
            return;
        }
        layoutMode = 'ai';
        aiLayoutInstructionsText = aiLayoutInstructions.value.trim();
        layoutConfigModalOverlay.style.display = 'none';
        layoutConfigPromise.resolve?.(true);
    } else if (choice === 'template') {
        const file = event.target.files[0];
        if (!file) {
            layoutConfigPromise.reject?.();
            return;
        };

        const reader = new FileReader();
        reader.onload = async (e) => {
            templateFileContent = e.target.result;
            layoutMode = 'template';
            console.log("Template file loaded.");
            showToast('Arquivo de modelo carregado.');
            layoutConfigModalOverlay.style.display = 'none';
            await startFieldMappingProcess(); // This resolves the promise on its own
        };
        reader.readAsArrayBuffer(file);
        event.target.value = ''; // Reset input
    }
    updateButtonStates();
}

function handleChangeMode() {
    resetPreviewState();
    welcomeMessage.style.display = 'none';
    modeChoiceContainer.style.display = 'flex';
    docPreviewContent.innerHTML = '';
    docPreviewContent.appendChild(modeChoiceContainer);

    const activeRow = dataGridContainer.querySelector('.data-table tbody tr.active');
    activeRow?.classList.remove('active');
}

function updateDataFromTable(e) {
    const cell = e.target;
    if (cell.tagName === 'TD' && cell.hasAttribute('contenteditable')) {
        const rowIndex = parseInt(cell.parentElement.dataset.rowIndex, 10);
        const header = cell.dataset.header;
        
        if (!isNaN(rowIndex) && header && csvData[rowIndex]) {
            csvData[rowIndex][header] = cell.textContent;
            
            if (previewMode === 'single' && rowIndex === activeRowIndex) {
                renderSinglePreview(rowIndex);
            }
        }
    }
}

function showAiPromptModal() {
    if (!ai) {
        showToast('Por favor, salve sua chave da API Gemini primeiro.', 'error');
        return;
    }
    if (previewMode !== 'single' && previewMode !== 'bulk') {
        showToast('Gere uma pré-visualização (individual ou em massa) antes de validar.', 'error');
        return;
    }
    if (previewMode === 'single' && activeRowIndex === null) {
        showToast('Por favor, selecione uma linha para validar.', 'error');
        return;
    }
    aiPromptTextarea.value = '';
    aiPromptModalOverlay.style.display = 'flex';
}

async function handleAiPrompt() {
    const prompt = aiPromptTextarea.value.trim();
    if (!prompt) {
        showToast('Por favor, insira as instruções de validação.', 'error');
        return;
    }

    confirmAiPromptBtn.disabled = true;
    confirmAiPromptBtn.textContent = 'Validando...';
    aiOutputContainer.style.display = 'flex';
    aiOutputContent.textContent = 'Aguarde, a IA está validando os dados...';
    
    clearAllValidationHighlights();

    const schema = {
        type: Type.ARRAY,
        items: {
            type: Type.OBJECT,
            properties: {
                fieldName: { type: Type.STRING },
                suggestion: { type: Type.STRING },
            },
            required: ["fieldName", "suggestion"],
        },
    };

    const runValidationOnRow = async (rowData) => {
        const fullPrompt = `
            Given this JSON data for a single record: ${JSON.stringify(rowData, null, 2)}.
            And this validation instruction from the user: "${prompt}".
            Please analyze the data based on the instruction.
            Respond ONLY with a JSON array of objects. Each object in the array should represent a validation error and have TWO keys:
            1. "fieldName": The exact key from the JSON with the error.
            2. "suggestion": A brief explanation of the error and how to fix it.
            If there are no errors, return an empty array [].
        `;
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: fullPrompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: schema,
                temperature: 0.1,
                systemInstruction: "Você é um especialista em UX com foco em frontend, retornando apenas o JSON solicitado.",
            },
        });
        return JSON.parse(response.text);
    };

    try {
        let allResults = [];
        if (previewMode === 'single') {
            const validationErrors = await runValidationOnRow(csvData[activeRowIndex]);
            if (validationErrors.length > 0) {
                 allResults = validationErrors.map(error => ({ ...error, rowIndex: activeRowIndex }));
            }
        } else { // previewMode === 'bulk'
            const validationPromises = csvData.map((row, index) => 
                runValidationOnRow(row).then(errors => 
                    errors.map(error => ({ ...error, rowIndex: index }))
                )
            );
            const nestedResults = await Promise.all(validationPromises);
            allResults = nestedResults.flat();
        }
        
        applyValidationToPreview(allResults);

    } catch (error) {
        handleApiError(error);
        aiOutputContent.textContent = `Ocorreu um erro ao validar os dados.`;
    } finally {
        confirmAiPromptBtn.disabled = false;
        confirmAiPromptBtn.textContent = 'Validar';
        aiPromptModalOverlay.style.display = 'none';
    }
}

async function handleApplyTransformRules() {
    const rules = aiTransformRulesTextarea.value.trim();
    if (!rules) {
        showToast('Por favor, insira as regras de transformação.', 'error');
        return;
    }
    if (!ai) {
        showToast('Por favor, salve sua chave da API Gemini primeiro.', 'error');
        return;
    }
    if (csvData.length === 0) {
        showToast('Carregue um arquivo CSV antes de aplicar regras.', 'error');
        return;
    }

    applyTransformRulesBtn.disabled = true;
    applyTransformRulesBtn.textContent = 'Processando...';
    showToast('A IA está processando as transformações em todos os dados...');

    try {
        const properties = csvHeaders.reduce((acc, header) => {
            acc[header] = { type: Type.STRING };
            return acc;
        }, {});
        const schema = {
            type: Type.ARRAY,
            items: {
                type: Type.OBJECT,
                properties: properties,
                required: csvHeaders,
            },
        };

        const prompt = `
            Aja como um especialista em limpeza e formatação de dados.
            Aqui está um array de objetos JSON com dados de um CSV: ${JSON.stringify(csvData)}.
            Aqui estão as regras de transformação que você deve aplicar a CADA objeto no array: "${rules}".
            Sua tarefa é processar todo o array e retornar um NOVO array de objetos JSON com as regras aplicadas.
            - MANTENHA a mesma estrutura de dados (mesmas chaves e mesma quantidade de objetos).
            - Se uma regra não se aplica a um campo ou linha, mantenha o valor original.
            - Responda APENAS com o array JSON final. Não inclua texto explicativo, markdown ou qualquer outra coisa.
        `;

        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: schema,
                temperature: 0.1,
                systemInstruction: "Você é um especialista em UX e limpeza de dados com foco em frontend, retornando apenas o JSON solicitado.",
            },
        });

        const transformedData = JSON.parse(response.text);
        if (transformedData.length !== csvData.length) {
            throw new Error("A IA retornou um número diferente de registros. A transformação falhou.");
        }

        csvData = transformedData;
        renderDataTable();
        showToast('Regras de transformação aplicadas com sucesso!', 'success');
        
    } catch (error) {
        handleApiError(error);
    } finally {
        applyTransformRulesBtn.disabled = false;
        applyTransformRulesBtn.textContent = 'Aplicar Regras';
    }
}


function applyValidationToPreview(results) {
    if (results.length === 0) {
        aiOutputContent.textContent = "Nenhuma inconsistência encontrada.";
        showToast('Validação concluída: Nenhuma inconsistência encontrada.');
        return;
    }

    const uniqueErrorRows = new Set(results.map(r => r.rowIndex + 1)).size;
    aiOutputContent.textContent = `Foram encontradas ${results.length} inconsistências em ${uniqueErrorRows} registro(s). Passe o mouse sobre os campos destacados para ver as sugestões e clique para ir ao campo e editar.`;
    showToast(`${results.length} inconsistências encontradas.`, 'error');

    results.forEach(error => {
        const { fieldName: csvHeaderWithError, suggestion, rowIndex } = error;
        
        const documentPreview = docPreviewContent.querySelector(`.preview-document[data-row-index="${rowIndex}"]`);
        if (!documentPreview) return;

        let fieldElement;
        
        if (layoutMode === 'template') {
            const placeholder = Object.keys(fieldMapping).find(key => fieldMapping[key] === csvHeaderWithError);
            if (placeholder) {
                fieldElement = documentPreview.querySelector(`[data-field-name="${placeholder}"]`);
            }
        } else { // layoutMode === 'ai'
            fieldElement = documentPreview.querySelector(`[data-field-name="${csvHeaderWithError}"]`);
        }

        if (fieldElement) {
            fieldElement.classList.add('invalid-field');
            fieldElement.dataset.tooltip = `${suggestion}\nClique para editar este campo na grade.`;
            
            const focusHandler = () => {
                handleRowClick(rowIndex, false); // select row without re-rendering
                focusCellInGrid(rowIndex, csvHeaderWithError);
            };
            
            fieldElement.addEventListener('click', focusHandler, { once: true });
        }
    });
}

function handleDownloadCsv() {
    if (csvData.length === 0 && csvHeaders.length === 0) {
        showToast("Não há dados para baixar.", 'error');
        return;
    }
    const csvContent = [
        csvHeaders.join(';'),
        ...csvData.map(row => csvHeaders.map(header => `"${(row[header] || '').toString().replace(/"/g, '""')}"`).join(';'))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "modified_data.csv");
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

async function handleDownloadZip() {
    if (!templateFileContent || Object.keys(fieldMapping).length === 0) {
        showToast("Carregue um modelo e confirme o mapeamento de campos.", 'error');
        return;
    }
    if (csvData.length === 0) {
        showToast("Não há dados CSV para gerar os documentos.", 'error');
        return;
    }
    
    downloadDocBtn.disabled = true;
    downloadDocBtn.textContent = 'Gerando...';
    showToast('Iniciando a geração dos documentos...');

    try {
        const zip = new JSZip();
        for (let i = 0; i < csvData.length; i++) {
            const row = csvData[i];
            const mappedData = getMappedData(row);

            const doc = new docxtemplater(new PizZip(templateFileContent), {
                paragraphLoop: true,
                linebreaks: true,
            });
            doc.setData(mappedData);
            doc.render();
            const blob = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            
            const fileName = row[csvHeaders[0]] || `document_${i + 1}`;
            zip.file(`${fileName}.docx`, blob);
        }

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const link = document.createElement("a");
        const url = URL.createObjectURL(zipBlob);
        link.setAttribute("href", url);
        link.setAttribute("download", "documentos_gerados.zip");
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        showToast('Documentos gerados com sucesso!');

    } catch (error) {
        console.error('Error generating documents:', error);
        showToast('Ocorreu um erro ao gerar os documentos.', 'error');
    } finally {
        downloadDocBtn.disabled = false;
        downloadDocBtn.textContent = 'Baixar Documentos';
    }
}

function handleClearData() {
    csvData = [];
    csvHeaders = [];
    renderDataTable(); 
    resetPreviewState(); 

    docPreviewContent.innerHTML = '';
    docPreviewContent.appendChild(welcomeMessage);
    welcomeMessage.style.display = 'flex';
    modeChoiceContainer.style.display = 'none';
    
    aiOutputContainer.style.display = 'none';
    aiOutputContent.textContent = '';
    
    showToast('Base de dados limpa.');
    updateButtonStates();
}

function handleClearLayout() {
    const wasInPreview = !!layoutMode;
    resetPreviewState();

    if (wasInPreview && csvData.length > 0) {
        welcomeMessage.style.display = 'none';
        modeChoiceContainer.style.display = 'flex';
        docPreviewContent.innerHTML = '';
        docPreviewContent.appendChild(modeChoiceContainer);

        const currentActiveRow = dataGridContainer.querySelector('.data-table tbody tr.active');
        currentActiveRow?.classList.remove('active');
    }
    
    showToast('Configuração de layout limpa.');
    updateButtonStates();
}

async function handleGenerateAnalysis() {
    if (!ai) {
        showToast('Por favor, salve sua chave da API Gemini primeiro.', 'error');
        return;
    }
    if (csvData.length === 0) {
        showToast('Carregue um arquivo CSV para gerar uma análise.', 'error');
        return;
    }
    
    analysisModalOverlay.style.display = 'flex';
    analysisSummaryContent.innerHTML = '<p>Analisando dados...</p>';
    analysisChartContent.innerHTML = '';
    
    const schema = {
        type: Type.OBJECT,
        properties: {
            summary: { type: Type.STRING },
            chartData: {
                type: Type.OBJECT,
                properties: {
                    title: { type: Type.STRING },
                    labelHeader: { type: Type.STRING },
                    valueHeader: { type: Type.STRING },
                    data: {
                        type: Type.ARRAY,
                        items: {
                            type: Type.OBJECT,
                            properties: {
                                label: { type: Type.STRING },
                                value: { type: Type.NUMBER },
                            },
                            required: ["label", "value"],
                        },
                    },
                },
                 required: ["title", "labelHeader", "valueHeader", "data"],
            },
        },
        required: ["summary", "chartData"],
    };

    const prompt = `
        Aja como um analista de dados e especialista em visualização.
        Aqui está um conjunto de dados em formato JSON: ${JSON.stringify(csvData.slice(0, 50))}
        (Nota: apenas as primeiras 50 linhas foram fornecidas para a análise, mas assuma que elas são representativas do conjunto completo).

        Sua tarefa é realizar duas coisas:

        1.  **Criar um Resumo Executivo:** Analise os dados e escreva um resumo conciso em HTML (usando <p>, <ul>, <li>, <strong>) destacando:
            - A qualidade geral dos dados (ex: campos vazios, inconsistências).
            - Principais insights ou tendências observadas.
            - Estatísticas interessantes (ex: valor médio de uma coluna numérica, o item mais frequente em uma coluna categórica).

        2.  **Sugerir um Gráfico:** Identifique a melhor visualização para um gráfico de barras a partir destes dados. Para isso:
            - Escolha UMA coluna categórica (para os rótulos do gráfico).
            - Escolha UMA coluna numérica (para os valores do gráfico). Se não houver uma coluna numérica óbvia, você deve fazer uma contagem de frequência da coluna categórica que escolheu.
            - Crie um título claro para o gráfico.
            - Agrupe os dados se necessário (ex: some os valores para categorias repetidas).

        Responda APENAS com um único objeto JSON que corresponda ao esquema fornecido. Não inclua texto explicativo fora do JSON.
    `;

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: schema,
                temperature: 0.1,
                systemInstruction: "Você é um especialista em UX e análise de dados, retornando apenas o JSON estruturado solicitado.",
            },
        });

        const result = JSON.parse(response.text);
        analysisSummaryContent.innerHTML = result.summary;
        renderBarChart(result.chartData);

    } catch (error) {
        handleApiError(error);
        analysisSummaryContent.innerHTML = `<p style="color: var(--error-color);">Ocorreu um erro ao gerar a análise.</p>`;
    }
}


// --- PREVIEW LOGIC ---

async function renderSinglePreview(rowIndex) {
    if (layoutMode === null) return;

    switchToPreviewMode();
    docPreviewContent.innerHTML = `<p style="text-align: center; padding: 2rem;">Gerando pré-visualização...</p>`;
    
    const row = csvData[rowIndex];
    const html = await getPreviewHtml(row);
    
    docPreviewContent.innerHTML = `<div class="preview-document" data-row-index="${rowIndex}">${html}</div>`;
}

async function renderBulkPreview() {
    if (layoutMode === null) return;

    switchToPreviewMode();
    docPreviewContent.innerHTML = `<p style="text-align: center; padding: 2rem;">Gerando pré-visualizações em massa...</p>`;

    const allHtmlPromises = csvData.map(row => getPreviewHtml(row));
    const allHtmls = await Promise.all(allHtmlPromises);

    docPreviewContent.innerHTML = allHtmls.map((html, index) => `<div class="preview-document" data-row-index="${index}">${html}</div>`).join('');
    docPreviewContent.parentElement.classList.add('bulk-preview-wrapper');
}

// --- MAPPING MODAL LOGIC ---

async function startFieldMappingProcess() {
    if (!ai) {
        showToast('A chave da API Gemini é necessária para o mapeamento inteligente.', 'error');
        layoutConfigPromise.reject?.();
        return;
    }
    mappingModalOverlay.style.display = 'flex';
    mappingFormContainer.innerHTML = '<p>Analisando modelo e CSV...</p>';

    try {
        const placeholders = extractPlaceholdersFromDoc(templateFileContent);
        if (placeholders.length === 0) {
            showToast("Nenhum placeholder (ex: {campo}) foi encontrado no modelo.", 'error');
            mappingModalOverlay.style.display = 'none';
            layoutConfigPromise.resolve?.(true); // Proceed even without placeholders
            return;
        }

        mappingFormContainer.innerHTML = '<p>A IA está sugerindo o mapeamento...</p>';
        const aiSuggestions = await getAiFieldMapping(placeholders, csvHeaders);
        showMappingModal(placeholders, csvHeaders, aiSuggestions);

    } catch (error) {
        handleApiError(error);
        mappingFormContainer.innerHTML = `<p style="color: var(--error-color);">Ocorreu um erro ao se comunicar com a IA.</p>`;
        layoutConfigPromise.reject?.();
    }
}

function extractPlaceholdersFromDoc(fileBuffer) {
    try {
        const zip = new PizZip(fileBuffer);
        const doc = new docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });
        
        const xmlContent = zip.files["word/document.xml"].asText();
        const regex = /\{(.*?)\}/g;
        const matches = new Set();
        let match;
        while ((match = regex.exec(xmlContent)) !== null) {
            const placeholder = match[1].replace(/<[^>]*>/g, "").trim();
            if (placeholder && !placeholder.startsWith('#') && !placeholder.startsWith('/')) {
                 matches.add(placeholder);
            }
        }
        return Array.from(matches);
    } catch(error) {
        console.error("Error extracting placeholders:", error);
        throw new Error("Não foi possível analisar o arquivo .docx. Verifique se o arquivo não está corrompido.");
    }
}

async function getAiFieldMapping(placeholders, headers) {
    const prompt = `
      Aja como um assistente de mapeamento de dados.
      Aqui estão os placeholders de um modelo de documento: ${JSON.stringify(placeholders)}.
      Aqui estão os cabeçalhos de um arquivo CSV: ${JSON.stringify(headers)}.
      Sua tarefa é encontrar a melhor correspondência semântica de um cabeçalho CSV para cada placeholder.
      Responda APENAS com um objeto JSON onde cada chave é um placeholder do modelo e seu valor é o cabeçalho CSV correspondente que você encontrou.
      Se você não encontrar uma correspondência razoável para um placeholder, defina seu valor como null.
      Exemplo de resposta: {"nome_cliente": "Nome Completo", "data_pedido": "Data da Compra", "produto_comprado": null}
    `;

    const schema = {
        type: Type.OBJECT,
        properties: placeholders.reduce((acc, placeholder) => {
            acc[placeholder] = { type: Type.STRING, nullable: true };
            return acc;
        }, {}),
    };
    
    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                responseSchema: schema,
                temperature: 0.1,
                systemInstruction: "Você é um especialista em UX com foco em frontend, especialista em mapeamento de dados. Retorne apenas o JSON solicitado.",
            },
        });
        return JSON.parse(response.text);
    } catch(error) {
        console.error("AI Mapping Error:", error);
        throw error; // Re-throw original error to be handled by the caller
    }
}

function showMappingModal(placeholders, headers, suggestions) {
    mappingFormContainer.innerHTML = '';
    const optionsHtml = [
        `<option value="">-- Não Mapear --</option>`,
        ...headers.map(h => `<option value="${h}">${h}</option>`)
    ].join('');

    placeholders.forEach(ph => {
        const row = document.createElement('div');
        row.className = 'mapping-row';
        row.innerHTML = `
            <label for="map-${ph}">{${ph}}</label>
            <select id="map-${ph}" data-placeholder="${ph}">
                ${optionsHtml}
            </select>
        `;
        mappingFormContainer.appendChild(row);
        
        const suggestedHeader = suggestions[ph];
        if (suggestedHeader && headers.includes(suggestedHeader)) {
            row.querySelector('select').value = suggestedHeader;
        }
    });
}

function handleMappingConfirmation() {
    fieldMapping = {};
    const selects = mappingFormContainer.querySelectorAll('select');
    selects.forEach(select => {
        const placeholder = select.dataset.placeholder;
        const selectedHeader = select.value;
        if (placeholder && selectedHeader) {
            fieldMapping[placeholder] = selectedHeader;
        }
    });
    
    mappingModalOverlay.style.display = 'none';
    console.log("Mapping confirmed:", fieldMapping);
    showToast('Mapeamento de campos confirmado.');
    
    aiTransformRulesSection.style.display = 'block'; // Show transform rules section
    updateButtonStates();
    layoutConfigPromise.resolve?.(true);
}

// --- UI & STATE HELPERS ---

function showLayoutConfigModal() {
    aiLayoutInstructions.value = ''; // Reset instructions
    return new Promise((resolve, reject) => {
        layoutConfigPromise = { resolve, reject };
        layoutConfigModalOverlay.style.display = 'flex';
    });
}

function switchToPreviewMode() {
    modeChoiceContainer.style.display = 'none';
    changeModeBtn.style.display = 'block';
}

function resetPreviewState() {
    layoutMode = null;
    previewMode = null;
    templateFileContent = null;
    activeRowIndex = null;
    fieldMapping = {};
    aiLayoutInstructionsText = '';
    changeModeBtn.style.display = 'none';
    docPreviewContent.parentElement.classList.remove('bulk-preview-wrapper');
    aiTransformRulesSection.style.display = 'none';
    updateButtonStates();
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
  if (lines.length === 0) {
    csvData = [];
    csvHeaders = [];
    return;
  }
  csvHeaders = lines[0].split(';').map(h => h.trim().replace(/^"|"$/g, ''));
  csvData = lines.slice(1).map(line => {
    const values = line.split(/;(?=(?:(?:[^"]*"){2})*[^"]*$)/);
    const rowObject = {};
    csvHeaders.forEach((header, index) => {
      rowObject[header] = (values[index] || '').trim().replace(/^"|"$/g, '');
    });
    return rowObject;
  });
}

function renderDataTable() {
    const table = dataGridContainer.querySelector('.data-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    if (csvData.length === 0) {
        thead.innerHTML = '';
        tbody.innerHTML = `<tr><td class="placeholder-cell">Carregue um arquivo CSV para visualizar os dados aqui.</td></tr>`;
        return;
    }

    thead.innerHTML = `<tr><th style="min-width: 60px;">#</th>${csvHeaders.map(h => `<th>${h}</th>`).join('')}</tr>`;

    tbody.innerHTML = csvData.map((row, rowIndex) => `
        <tr data-row-index="${rowIndex}">
            <td>${rowIndex + 1}</td>
            ${csvHeaders.map(header => `<td contenteditable="true" data-header="${header}">${row[header] || ''}</td>`).join('')}
        </tr>
    `).join('');
    
    tbody.querySelectorAll('tr').forEach(tr => {
        tr.addEventListener('click', (e) => {
            // Prevent row click from firing when editing text
            if (e.target.isContentEditable) return;
            const newIndex = parseInt(tr.dataset.rowIndex, 10);
            handleRowClick(newIndex);
        });
    });

    table.addEventListener('input', updateDataFromTable);
}

function getMappedData(row) {
    const mapped = {};
    for (const placeholder in fieldMapping) {
        const csvHeader = fieldMapping[placeholder];
        mapped[placeholder] = row[csvHeader] || '';
    }
    return mapped;
}

async function getPreviewHtml(row) {
    if (layoutMode === 'template' && templateFileContent) {
        try {
            const mappedData = getMappedData(row);
            const doc = new docxtemplater(new PizZip(templateFileContent), {
                paragraphLoop: true,
                linebreaks: true,
            });
            doc.setData(mappedData);
            doc.render();
            const blob = doc.getZip().generate({ type: 'blob' });
            const arrayBuffer = await blob.arrayBuffer();
            const result = await mammoth.convertToHtml({ arrayBuffer });
            
            let processedHtml = result.value;
            for (const key in mappedData) {
                const value = mappedData[key];
                if (value) {
                    const regex = new RegExp(value.toString().replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + '(?![^<>]*>)', 'g');
                    processedHtml = processedHtml.replace(regex, `<span data-field-name="${key}">${value}</span>`);
                }
            }
            return processedHtml;

        } catch (error) {
            console.error('Error generating .docx preview:', error);
            return `<p style="color: red;">Erro ao gerar pré-visualização. Verifique se os campos foram mapeados corretamente e se as chaves no seu modelo (ex: {Nome}) estão corretas.</p><pre>${error.message}</pre>`;
        }
    } else if (layoutMode === 'ai') {
        if (!ai) return `<p style="color: orange;">Chave da API Gemini não configurada.</p>`;
        try {
            const dataJson = JSON.stringify(row, null, 2);
            const instructions = aiLayoutInstructionsText ? `Com as seguintes instruções adicionais: "${aiLayoutInstructionsText}"` : "";
            const prompt = `
                Aja como um designer de documentos. Crie um HTML bem formatado para os dados JSON abaixo.
                ${instructions}
                Use títulos (h3, h4), parágrafos e listas para uma apresentação clara e profissional. 
                Quando você incluir um valor dos dados JSON, **obrigatoriamente** envolva-o em um span com um atributo 'data-field-name', onde o valor do atributo é a chave do JSON.
                Exemplo: Para a chave "Nome" com valor "João", o HTML deve conter <span data-field-name="Nome">João</span>.
                Não inclua \`<html>\`, \`<body>\`, ou \`\`\`html\`\`\`. O resultado deve ser apenas o código HTML para ser inserido em uma div.

                Dados JSON:
                ${dataJson}
            `;
            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: prompt,
                config: {
                    temperature: 0.1,
                    systemInstruction: "Você é um especialista em UX com foco em frontend e design de documentos. Retorne apenas o código HTML solicitado.",
                }
            });
            return response.text.trim().replace(/^```html\s*|```\s*$/g, '');
        } catch (error) {
            handleApiError(error);
            return `<p style="color: red;">Erro na API ao gerar pré-visualização. Verifique sua chave.</p>`;
        }
    }
    return '';
}

function updateButtonStates() {
    const hasCsvData = csvData.length > 0;
    const hasLayout = !!layoutMode;
    const canDownloadZip = hasCsvData && templateFileContent && layoutMode === 'template' && Object.keys(fieldMapping).length > 0;

    downloadCsvBtn.disabled = !hasCsvData;
    runAiBtn.disabled = !hasCsvData || !hasLayout;
    downloadDocBtn.disabled = !canDownloadZip;
    clearDataBtn.disabled = !hasCsvData;
    clearLayoutBtn.disabled = !hasLayout;
    generateAnalysisBtn.disabled = !hasCsvData;
    
    const canApplyTransform = hasCsvData && layoutMode === 'template' && Object.keys(fieldMapping).length > 0;
    applyTransformRulesBtn.disabled = !canApplyTransform;
}


function showToast(message, type = 'success', duration = 3000) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;

    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('hide');
        toast.addEventListener('transitionend', () => toast.remove());
    }, duration);
}

function focusCellInGrid(rowIndex, header) {
    const tableWrapper = dataGridContainer.querySelector('.table-wrapper');
    const cell = tableWrapper.querySelector(`tr[data-row-index="${rowIndex}"] td[data-header="${header}"]`);

    if (cell) {
        cell.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        cell.focus();
        showToast(`Campo "${header}" pronto para edição.`);
    } else {
        console.warn(`Não foi possível encontrar a célula para a linha ${rowIndex} e cabeçalho ${header}`);
    }
}

function clearAllValidationHighlights() {
    const highlightedFields = docPreviewContent.querySelectorAll('.invalid-field');
    highlightedFields.forEach(field => {
        const cleanField = field.cloneNode(true);
        cleanField.classList.remove('invalid-field');
        cleanField.removeAttribute('data-tooltip');
        field.parentNode.replaceChild(cleanField, field);
    });
}

function handleApiError(error) {
    console.error("Gemini API Error:", error);
    const message = error.toString().toLowerCase();
    
    // Check for common API key / quota error messages
    if (message.includes('quota') || message.includes('api key not valid') || message.includes('api_key')) {
        showToast('Chave da API inválida ou cota excedida. Por favor, insira uma nova chave.', 'error');
        apiKey = '';
        ai = null;
        sessionStorage.removeItem('geminiApiKey');
        apiKeyInput.value = '';
        
        const apiSection = document.getElementById('api-section');
        apiSection.classList.add('needs-attention');
        apiKeyInput.focus();
    } else {
        showToast(`Ocorreu um erro na API: ${error.message}`, 'error');
    }
}

function renderBarChart(chartData) {
    analysisChartContent.innerHTML = ''; // Clear previous chart
    if (!chartData || !chartData.data || chartData.data.length === 0) {
        analysisChartContent.innerHTML = '<p>Não foi possível gerar um gráfico para estes dados.</p>';
        return;
    }

    const titleEl = document.createElement('h4');
    titleEl.className = 'chart-title';
    titleEl.textContent = chartData.title;
    analysisChartContent.appendChild(titleEl);

    const maxValue = Math.max(...chartData.data.map(item => item.value));
    
    chartData.data.forEach(item => {
        const rowEl = document.createElement('div');
        rowEl.className = 'chart-bar-row';
        
        const labelEl = document.createElement('div');
        labelEl.className = 'chart-label';
        labelEl.textContent = item.label;
        labelEl.title = item.label;

        const barContainerEl = document.createElement('div');
        barContainerEl.className = 'chart-bar-container';

        const barEl = document.createElement('div');
        barEl.className = 'chart-bar';
        const barWidth = maxValue > 0 ? (item.value / maxValue) * 100 : 0;
        barEl.style.width = `${barWidth}%`;
        barEl.textContent = item.value.toLocaleString('pt-BR');

        barContainerEl.appendChild(barEl);
        rowEl.appendChild(labelEl);
        rowEl.appendChild(barContainerEl);
        analysisChartContent.appendChild(rowEl);
    });
}


// --- TOUR LOGIC ---
function startTour() {
    currentTourStep = 0;
    tourOverlay.style.display = 'block';
    showTourStep(currentTourStep);
}

function endTour() {
    const currentElement = document.getElementById(tourSteps[currentTourStep].elementId);
    currentElement?.classList.remove('tour-highlight');
    tourOverlay.style.display = 'none';
    sessionStorage.setItem('tourShown', 'true');
}

function nextTourStep() {
    if (currentTourStep < tourSteps.length - 1) {
        currentTourStep++;
        showTourStep(currentTourStep);
    } else {
        endTour();
    }
}

function prevTourStep() {
    if (currentTourStep > 0) {
        currentTourStep--;
        showTourStep(currentTourStep);
    }
}

function showTourStep(stepIndex) {
    // Remove highlight from previous step
    if (stepIndex > 0) {
        document.getElementById(tourSteps[stepIndex - 1].elementId)?.classList.remove('tour-highlight');
    }
    if (stepIndex < tourSteps.length -1 ) {
         document.getElementById(tourSteps[stepIndex + 1].elementId)?.classList.remove('tour-highlight');
    }


    const step = tourSteps[stepIndex];
    const targetElement = document.getElementById(step.elementId);

    if (!targetElement) {
        console.error(`Tour element not found: ${step.elementId}`);
        endTour();
        return;
    }

    targetElement.classList.add('tour-highlight');
    targetElement.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });

    tourTitle.textContent = step.title;
    tourText.textContent = step.text;
    tourStepCounter.textContent = `${stepIndex + 1} de ${tourSteps.length}`;
    
    // Position popover
    const targetRect = targetElement.getBoundingClientRect();
    const popoverRect = tourPopover.getBoundingClientRect();
    
    let top = targetRect.bottom + 15;
    let left = targetRect.left + (targetRect.width / 2) - (popoverRect.width / 2);
    
    // Adjust if it overflows
    if (top + popoverRect.height > window.innerHeight) {
        top = targetRect.top - popoverRect.height - 15;
    }
    if (left < 10) {
        left = 10;
    }
    if (left + popoverRect.width > window.innerWidth) {
        left = window.innerWidth - popoverRect.width - 10;
    }

    tourPopover.style.top = `${top}px`;
    tourPopover.style.left = `${left}px`;
    
    // Update buttons
    tourPrevBtn.style.display = stepIndex === 0 ? 'none' : 'inline-block';
    if (stepIndex === tourSteps.length - 1) {
        tourNextBtn.textContent = 'Concluir';
    } else {
        tourNextBtn.textContent = 'Próximo';
    }
}


// --- START ---
initializeApp();