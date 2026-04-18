// ═══════════════════════════════════════════════════════════
// APPS SCRIPT — Cole este código inteiro no Editor de Scripts
// da sua planilha (Extensões > Apps Script).
// Depois publique como Web App (Deploy > New deployment > Web app)
// Execute as: Me  |  Who has access: Anyone
// ═══════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1FmJy-gHRMpF4thwWotX-tFQUoUGJ4yCbd4eIsjnBlVg';

// ── Helpers ──
function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CORS-friendly: GET e POST ──
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    // Prioriza parâmetros GET (e.parameter) — evita bloqueios de CORS do browser
    let params = e.parameter || {};

    // Fallback para POST body se não tiver action no GET
    if (!params.action && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    }

    const action = params.action;

    switch (action) {
      case 'checkCode':
        return checkCode(params.codigo);
      case 'getClientData':
        return getClientData(params.codigo);
      case 'createClientTab':
        return createClientTab(params.codigo);
      case 'saveAnswer':
        return saveAnswer(params.codigo, Number(params.rowIndex), params.resposta);
      case 'getFormQuestions':
        return getFormQuestions();
      default:
        return jsonResponse({ error: 'Ação desconhecida: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════
// 1. VERIFICAR CÓDIGO — Busca na aba "Painel"
// ═══════════════════════════════════════════════════
function checkCode(codigo) {
  const sheet = getSheet('Painel');
  const data = sheet.getDataRange().getValues();
  
  // Cabeçalhos na linha 1: Nome (A), Código (B), Status (C)
  for (let i = 1; i < data.length; i++) {
    const nome   = String(data[i][0]).trim();
    const cod    = String(data[i][1]).trim();
    const status = String(data[i][2]).trim();

    if (cod === String(codigo).trim()) {
      return jsonResponse({
        found: true,
        nome: nome,
        codigo: cod,
        status: status,
        row: i + 1 // linha real na planilha (1-indexed)
      });
    }
  }

  return jsonResponse({ found: false });
}

// ═══════════════════════════════════════════════════
// 2. BUSCAR AS PERGUNTAS DO FORMULÁRIO
// ═══════════════════════════════════════════════════
function getFormQuestions() {
  const sheet = getSheet('Formulário');
  const data = sheet.getDataRange().getValues();
  const questions = [];

  // Coluna A = Formulário  |  Coluna B = Informações
  // Linha 1 é cabeçalho, pula
  for (let i = 1; i < data.length; i++) {
    const q = String(data[i][0]).trim();
    if (q) {
      questions.push({ rowIndex: i + 1, question: q });
    }
  }

  return jsonResponse({ questions: questions });
}

// ═══════════════════════════════════════════════════
// 3. CRIAR ABA DO CLIENTE
// ═══════════════════════════════════════════════════
function createClientTab(codigo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const painel = ss.getSheetByName('Painel');
  const formulario = ss.getSheetByName('Formulário');

  // Encontrar a linha do cliente no Painel
  const painelData = painel.getDataRange().getValues();
  let clientRow = -1;
  let clientName = '';
  
  for (let i = 1; i < painelData.length; i++) {
    if (String(painelData[i][1]).trim() === String(codigo).trim()) {
      clientRow = i + 1;
      clientName = String(painelData[i][0]).trim();
      break;
    }
  }

  if (clientRow === -1) {
    return jsonResponse({ error: 'Código não encontrado no Painel.' });
  }

  // Nome da aba = nome do cliente
  const tabName = clientName;
  let clientSheet = ss.getSheetByName(tabName);

  if (!clientSheet) {
    // Copiar estrutura do Formulário
    const formData = formulario.getDataRange().getValues();
    clientSheet = ss.insertSheet(tabName);

    // Escrever cabeçalhos e perguntas
    for (let i = 0; i < formData.length; i++) {
      clientSheet.getRange(i + 1, 1).setValue(formData[i][0]); // Coluna A: Perguntas
      clientSheet.getRange(i + 1, 2).setValue(formData[i][1]); // Coluna B: (vazio ou cabeçalho)
    }

    // Formatar a aba do cliente
    clientSheet.setColumnWidth(1, 300);
    clientSheet.setColumnWidth(2, 400);
    
    // Cabeçalho em negrito
    clientSheet.getRange(1, 1, 1, 2).setFontWeight('bold')
      .setBackground('#4f8eff').setFontColor('#ffffff');

    // Criar link clicável no Painel (coluna A = Nome vira hyperlink)
    const sheetUrl = ss.getUrl() + '#gid=' + clientSheet.getSheetId();
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(clientName)
      .setLinkUrl(sheetUrl)
      .build();
    painel.getRange(clientRow, 1).setRichTextValue(richText);

    // Atualizar status para "Em andamento"
    painel.getRange(clientRow, 3).setValue('Em andamento');
  }

  return jsonResponse({
    success: true,
    tabName: tabName,
    tabId: clientSheet.getSheetId()
  });
}

// ═══════════════════════════════════════════════════
// 4. BUSCAR DADOS DO CLIENTE (aba individual)
// ═══════════════════════════════════════════════════
function getClientData(codigo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const painel = ss.getSheetByName('Painel');
  const painelData = painel.getDataRange().getValues();

  let clientName = '';
  for (let i = 1; i < painelData.length; i++) {
    if (String(painelData[i][1]).trim() === String(codigo).trim()) {
      clientName = String(painelData[i][0]).trim();
      break;
    }
  }

  if (!clientName) return jsonResponse({ error: 'Código não encontrado.' });

  const clientSheet = ss.getSheetByName(clientName);
  if (!clientSheet) {
    return jsonResponse({ tabExists: false, nome: clientName });
  }

  // Carregar aba Formulário para pegar as observações também (para funcionar inclusive com abas de clientes já criadas)
  const formSheet = ss.getSheetByName('Formulário');
  let formObsMap = {};
  if (formSheet) {
    const formData = formSheet.getDataRange().getValues();
    const formHeader = formData[0] || [];
    let obsColIndex = -1;
    // Tenta encontrar a coluna de observações pelo cabeçalho
    for(let c=0; c<formHeader.length; c++) {
      const headerText = String(formHeader[c]).toLowerCase();
      if(headerText.includes('observaç') || headerText.includes('observac')) {
        obsColIndex = c;
        break;
      }
    }
    // Fallback para a coluna C (índice 2) se não achar pelo nome, mas houver mais de 2 colunas
    if(obsColIndex === -1 && formData[0].length > 2) {
      obsColIndex = 2;
    }

    if(obsColIndex !== -1) {
      for(let i=1; i<formData.length; i++) {
        const q = String(formData[i][0]).trim();
        const obs = String(formData[i][obsColIndex]).trim();
        if(q) formObsMap[q] = obs;
      }
    }
  }

  const data = clientSheet.getDataRange().getValues();
  const fields = [];

  // Linha 1 é cabeçalho, começa em 2
  for (let i = 1; i < data.length; i++) {
    const question = String(data[i][0]).trim();
    const answer   = String(data[i][1]).trim();
    if (question) {
      fields.push({
        rowIndex: i + 1,
        question: question,
        answer: answer || '',
        observacao: formObsMap[question] || ''
      });
    }
  }

  return jsonResponse({
    tabExists: true,
    nome: clientName,
    fields: fields
  });
}

// ═══════════════════════════════════════════════════
// 5. SALVAR RESPOSTA NA ABA DO CLIENTE
// ═══════════════════════════════════════════════════
function saveAnswer(codigo, rowIndex, resposta) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const painel = ss.getSheetByName('Painel');
  const painelData = painel.getDataRange().getValues();

  let clientName = '';
  let painelRow = -1;
  for (let i = 1; i < painelData.length; i++) {
    if (String(painelData[i][1]).trim() === String(codigo).trim()) {
      clientName = String(painelData[i][0]).trim();
      painelRow = i + 1;
      break;
    }
  }

  if (!clientName) return jsonResponse({ error: 'Código não encontrado.' });

  const clientSheet = ss.getSheetByName(clientName);
  if (!clientSheet) return jsonResponse({ error: 'Aba do cliente não encontrada.' });

  // Salvar resposta na coluna B da linha indicada
  clientSheet.getRange(rowIndex, 2).setValue(resposta);

  // Verificar se todas as perguntas foram respondidas
  const data = clientSheet.getDataRange().getValues();
  let allDone = true;
  for (let i = 1; i < data.length; i++) {
    const q = String(data[i][0]).trim();
    const a = String(data[i][1]).trim();
    if (q && !a) {
      allDone = false;
      break;
    }
  }

  // Atualizar status no Painel
  if (allDone) {
    painel.getRange(painelRow, 3).setValue('Completo');
  } else {
    painel.getRange(painelRow, 3).setValue('Em andamento');
  }

  return jsonResponse({ success: true, allDone: allDone });
}
