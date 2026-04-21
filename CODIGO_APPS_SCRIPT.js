// ═══════════════════════════════════════════════════════════
// APPS SCRIPT — Cole este código inteiro no Editor de Scripts
// da sua planilha (Extensões > Apps Script).
// ═══════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1FmJy-gHRMpF4thwWotX-tFQUoUGJ4yCbd4eIsjnBlVg';

// ── Helpers ──
function getSs() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function getSheet(name) { return getSs().getSheetByName(name); }

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Conversão segura → impede que "undefined" apareça como texto na planilha
function safeStr(val) {
  if (val === undefined || val === null) return '';
  return String(val);
}

// Localiza cabeçalhos dinamicamente usando aliases (para aba Painel)
function getTableMap(sheet, colsConfig) {
  const data = sheet.getDataRange().getValues();
  let headerRow = -1;
  let mapping = {};
  const limit = Math.min(5, data.length);

  for (let r = 0; r < limit; r++) {
    let matches = 0;
    let tempMap = {};
    for (let c = 0; c < data[r].length; c++) {
      let val = safeStr(data[r][c]).toLowerCase().trim();
      val = val.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

      Object.keys(colsConfig).forEach(key => {
        const aliases = colsConfig[key];
        if (tempMap[key] === undefined) {
          if (aliases.some(alias => val.includes(alias))) {
            tempMap[key] = c;
            matches++;
          }
        }
      });
    }

    const needed = Object.keys(colsConfig).length;
    if (matches >= Math.min(2, needed) || (needed === 1 && matches === 1)) {
      headerRow = r;
      mapping = tempMap;
      break;
    }
  }
  return { headerRow, mapping, data };
}

// ── CORS-friendly ──
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    let params = e.parameter || {};
    if (!params.action && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    }
    const action = params.action;
    switch (action) {
      case 'checkCode':     return checkCode(params.codigo);
      case 'getClientData': return getClientData(params.codigo);
      case 'createClientTab': return createClientTab(params.codigo);
      case 'saveAnswer':    return saveAnswer(params.codigo, Number(params.rowIndex), params.resposta);
      case 'getConfig':     return getConfig();
      default: return jsonResponse({ error: 'Ação desconhecida: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.toString(), stack: err.stack });
  }
}

// ═══════════════════════════════════════════════════════════
// 1. VERIFICAR CÓDIGO (Aba Painel)
// ═══════════════════════════════════════════════════════════
function checkCode(codigo) {
  const sheet = getSheet('Painel');
  if (!sheet) return jsonResponse({ error: 'Aba "Painel" não encontrada.' });

  const { headerRow, mapping, data } = getTableMap(sheet, {
    'Nome': ['nome', 'cliente'],
    'Código': ['codigo', 'acesso'],
    'Status': ['status', 'estado']
  });

  if (headerRow === -1 || mapping['Código'] === undefined) {
    return jsonResponse({ error: 'Cabeçalhos do Painel não encontrados.' });
  }

  const cNome = mapping['Nome'] !== undefined ? mapping['Nome'] : 0;
  const cCod  = mapping['Código'];
  const cStat = mapping['Status'] !== undefined ? mapping['Status'] : 2;

  for (let i = headerRow + 1; i < data.length; i++) {
    if (safeStr(data[i][cCod]).trim() === String(codigo).trim()) {
      return jsonResponse({
        found: true,
        nome:   safeStr(data[i][cNome]).trim(),
        codigo: safeStr(data[i][cCod]).trim(),
        status: safeStr(data[i][cStat]).trim()
      });
    }
  }
  return jsonResponse({ found: false });
}

// ═══════════════════════════════════════════════════════════
// FUNÇÃO CENTRAL: Detecta colunas da aba Formulário
// Retorna { headerRow, qCol, obsCol } com segurança total
// ═══════════════════════════════════════════════════════════
function detectFormularioColumns(formularioSheet) {
  const fData = formularioSheet.getDataRange().getValues();

  let headerRow = -1;
  let qCol = -1;    // Coluna com perguntas ("Formulário" ou "Pergunta")
  let obsCol = -1;  // Coluna com observações ("Observações")

  for (let r = 0; r < Math.min(5, fData.length); r++) {
    for (let c = 0; c < fData[r].length; c++) {
      const val = safeStr(fData[r][c]).toLowerCase().trim()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "");

      // Procura a coluna de perguntas
      if (qCol === -1) {
        if (val === 'formulario' || val === 'pergunta' || val === 'perguntas' ||
            val === 'questao'    || val === 'questoes') {
          qCol = c;
          headerRow = r;
        }
      }

      // Procura a coluna de observações
      if (obsCol === -1) {
        if (val === 'observacoes' || val === 'observacao' || val === 'obs' ||
            val === 'notas'       || val === 'nota') {
          obsCol = c;
          if (headerRow === -1) headerRow = r;
        }
      }
    }
    // Se já achou a de perguntas, podemos parar a busca de linhas
    if (qCol >= 0) break;
  }

  // Fallback seguro: se nenhum header achado, assume col 0 = perguntas
  if (headerRow === -1) { headerRow = 0; qCol = 0; }
  if (qCol === -1) { qCol = 0; }

  return { fData, headerRow, qCol, obsCol };
}

// ═══════════════════════════════════════════════════════════
// SINCRONIZAR PERGUNTAS: Formulário → Aba do Cliente
// ★ Conserta abas corrompidas e preserva respostas existentes
// ═══════════════════════════════════════════════════════════
function syncQuestionsFromFormulario(clientSheet, formularioSheet) {
  const { fData, headerRow, qCol, obsCol } = detectFormularioColumns(formularioSheet);

  // Extrair perguntas e observações válidas do Formulário
  const questions = [];
  const obsFromForm = [];
  for (let i = headerRow + 1; i < fData.length; i++) {
    const raw = (qCol < fData[i].length) ? fData[i][qCol] : null;
    const q = safeStr(raw).trim();
    if (q && q !== 'undefined') {
      questions.push(q);
      const obs = (obsCol >= 0 && obsCol < fData[i].length) ? safeStr(fData[i][obsCol]).trim() : '';
      obsFromForm.push(obs);
    }
  }

  // Ler respostas já existentes na aba do cliente (para preservá-las)
  const existingAnswers = {};
  if (clientSheet.getLastRow() > 1) {
    const existingData = clientSheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      const eq = safeStr(existingData[i][0]).trim();
      const ea = safeStr(existingData[i][1]).trim();
      if (eq && ea && eq !== 'undefined') {
        existingAnswers[eq] = ea;
      }
    }
  }

  // Escrever cabeçalhos da aba do cliente (3 colunas)
  clientSheet.getRange(1, 1).setValue('Pergunta')
    .setFontWeight('bold').setBackground('#4f8eff').setFontColor('#ffffff');
  clientSheet.getRange(1, 2).setValue('Respostas')
    .setFontWeight('bold').setBackground('#4f8eff').setFontColor('#ffffff');
  clientSheet.getRange(1, 3).setValue('Observações')
    .setFontWeight('bold').setBackground('#f59e0b').setFontColor('#ffffff');

  // Escrever perguntas + preservar respostas + copiar observações do formulário
  // ★ A coluna C é SEMPRE criada (fica vazia se o Formulário base não tiver obs)
  for (let i = 0; i < questions.length; i++) {
    const row = i + 2;
    clientSheet.getRange(row, 1).setValue(questions[i]);
    clientSheet.getRange(row, 2).setValue(existingAnswers[questions[i]] || '');
    const obsValue = (obsFromForm[i] !== undefined) ? obsFromForm[i] : '';
    clientSheet.getRange(row, 3).setValue(obsValue);
  }

  // Limpar linhas extras (dados antigos/corrompidos de tentativas anteriores)
  const lastRow = clientSheet.getLastRow();
  const expectedLastRow = questions.length + 1;
  if (lastRow > expectedLastRow) {
    clientSheet.getRange(expectedLastRow + 1, 1, lastRow - expectedLastRow, 3).clearContent();
  }

  return questions.length;
}

// ═══════════════════════════════════════════════════════════
// 2. CRIAR ABA DO CLIENTE
// ★ Sincroniza APENAS na criação. Abas existentes são preservadas.
// ═══════════════════════════════════════════════════════════
function createClientTab(codigo) {
  const ss = getSs();
  const painel = ss.getSheetByName('Painel');
  const formulario = ss.getSheetByName('Formulário');
  if (!painel || !formulario) return jsonResponse({ error: 'Abas principais não encontradas.' });

  // Localizar cliente no Painel
  const pMap = getTableMap(painel, {
    'Nome': ['nome'], 'Código': ['codigo'], 'Status': ['status']
  });

  let clientName = '', pRow = -1;
  const cCodPainel = pMap.mapping['Código'];
  if (cCodPainel === undefined) return jsonResponse({ error: 'Coluna Código não achada no Painel.' });

  for (let i = pMap.headerRow + 1; i < pMap.data.length; i++) {
    if (safeStr(pMap.data[i][cCodPainel]).trim() === String(codigo).trim()) {
      clientName = safeStr(pMap.data[i][pMap.mapping['Nome']]).trim();
      pRow = i + 1;
      break;
    }
  }
  if (!clientName) return jsonResponse({ error: 'Cliente não localizado no Painel.' });

  // Criar aba ou reusar existente
  let clientSheet = ss.getSheetByName(clientName);
  let isNew = false;

  if (!clientSheet) {
    clientSheet = ss.insertSheet(clientName);
    isNew = true;
  }

  // ★ Sincroniza APENAS se a aba for nova.
  // Se a aba já existe, respeita o conteúdo manual do usuário.
  let totalQ = 0;
  if (isNew) {
    totalQ = syncQuestionsFromFormulario(clientSheet, formulario);
    clientSheet.setColumnWidth(1, 450);
    clientSheet.setColumnWidth(2, 450);
    clientSheet.setColumnWidth(3, 300);
    clientSheet.setFrozenRows(1);
  } else {
    // Conta as perguntas existentes sem modificar nada
    const rows = clientSheet.getLastRow();
    totalQ = rows > 1 ? rows - 1 : 0;
  }

  // Link e status no Painel (só para abas novas)
  if (isNew && pMap.mapping['Nome'] !== undefined) {
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(clientName)
      .setLinkUrl(ss.getUrl() + '#gid=' + clientSheet.getSheetId())
      .build();
    painel.getRange(pRow, pMap.mapping['Nome'] + 1).setRichTextValue(richText);
  }
  if (isNew && pMap.mapping['Status'] !== undefined) {
    painel.getRange(pRow, pMap.mapping['Status'] + 1).setValue('Em andamento');
  }

  // ★ Reordenar abas: Painel, Formulário, Configuração sempre primeiro
  reorderSheets(ss);

  return jsonResponse({ success: true, questionsCount: totalQ });
}

// ═══════════════════════════════════════════════════════════
// 3. BUSCAR DADOS DO CLIENTE
// ═══════════════════════════════════════════════════════════
function getClientData(codigo) {
  const ss = getSs();
  const painel = ss.getSheetByName('Painel');
  const pMap = getTableMap(painel, { 'Nome': ['nome'], 'Código': ['codigo'] });

  let clientName = '';
  const cCodPainel = pMap.mapping['Código'];

  if (cCodPainel !== undefined) {
    for (let i = pMap.headerRow + 1; i < pMap.data.length; i++) {
      if (safeStr(pMap.data[i][cCodPainel]).trim() === String(codigo).trim()) {
        clientName = safeStr(pMap.data[i][pMap.mapping['Nome']]).trim();
        break;
      }
    }
  }
  if (!clientName) return jsonResponse({ error: 'Código inválido.' });

  const clientSheet = ss.getSheetByName(clientName);
  const formSheet   = ss.getSheetByName('Formulário');

  if (!clientSheet) return jsonResponse({ tabExists: false });

  // ── Carregar perguntas, respostas e observações diretamente da aba do cliente ──
  // Estrutura: col A = Pergunta, col B = Informações, col C = Observações (editável manualmente)
  const cData = clientSheet.getDataRange().getValues();
  const fields = [];

  for (let i = 1; i < cData.length; i++) {
    const q   = safeStr(cData[i][0]).trim();  // Coluna A = pergunta
    const a   = safeStr(cData[i][1]).trim();  // Coluna B = resposta
    const obs = safeStr(cData[i][2]).trim();  // Coluna C = observação (editável manualmente)
    if (q && q !== 'undefined') {
      fields.push({
        rowIndex:   i + 1,   // Linha real na planilha (1-indexed)
        question:   q,
        answer:     a,
        observacao: obs
      });
    }
  }

  return jsonResponse({ tabExists: true, nome: clientName, fields: fields });
}

// ═══════════════════════════════════════════════════════════
// 4. SALVAR RESPOSTA
// ═══════════════════════════════════════════════════════════
function saveAnswer(codigo, rowIndex, resposta) {
  const ss = getSs();
  const painel = ss.getSheetByName('Painel');
  const pMap = getTableMap(painel, { 'Nome': ['nome'], 'Código': ['codigo'], 'Status': ['status'] });

  let clientName = '', pRow = -1;
  const cCodPainel = pMap.mapping['Código'];
  if (cCodPainel !== undefined) {
    for (let i = pMap.headerRow + 1; i < pMap.data.length; i++) {
      if (safeStr(pMap.data[i][cCodPainel]).trim() === String(codigo).trim()) {
        clientName = safeStr(pMap.data[i][pMap.mapping['Nome']]).trim();
        pRow = i + 1;
        break;
      }
    }
  }

  const clientSheet = ss.getSheetByName(clientName);
  if (!clientSheet) return jsonResponse({ error: 'Aba do cliente não encontrada.' });

  // Coluna B (col 2 no getRange, 1-indexed) = Informações — sempre fixa
  clientSheet.getRange(rowIndex, 2).setValue(resposta);

  // Re-checar se TODAS as perguntas foram respondidas
  const updatedData = clientSheet.getDataRange().getValues();
  let allDone = true;
  for (let i = 1; i < updatedData.length; i++) {
    const q = safeStr(updatedData[i][0]).trim();
    const a = safeStr(updatedData[i][1]).trim();
    if (q && q !== 'undefined' && !a) { allDone = false; break; }
  }

  if (allDone && pMap.mapping['Status'] !== undefined) {
    painel.getRange(pRow, pMap.mapping['Status'] + 1).setValue('Completo');
  }

  return jsonResponse({ success: true, allDone: allDone });
}
// ═══════════════════════════════════════════════════════════
// HELPER: Reordenar abas — Painel, Formulário, Configuração sempre primeiro
// ═══════════════════════════════════════════════════════════
function reorderSheets(ss) {
  const priority = ['Painel', 'Formulário', 'Configuração'];
  let pos = 0;
  priority.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(pos + 1);
      pos++;
    }
  });
}

// ═══════════════════════════════════════════════════════════
// 5. BUSCAR CONFIGURAÇÕES (Prompt do Sistema)
// ═══════════════════════════════════════════════════════════
function getConfig() {
  const ss = getSs();
  let configSheet = ss.getSheetByName('Configuração');
  
  if (!configSheet) {
    // Criar aba de configuração se não existir
    configSheet = ss.insertSheet('Configuração');
    configSheet.getRange(1, 1).setValue('Chave').setFontWeight('bold');
    configSheet.getRange(1, 2).setValue('Valor').setFontWeight('bold');
    
    configSheet.getRange(2, 1).setValue('SystemPrompt');
    configSheet.getRange(2, 2).setValue('Você é uma assistente virtual carismática e profissional da C.E. Afonso Soluções Digitais. Você ajuda donos de pequenos negócios a preencher o cadastro da empresa.\n\nREGRAS:\n1. Uma pergunta por vez.\n2. Seja breve e cordial.\n3. Use o marcador [SALVAR|ID|resposta] no final.');
    
    configSheet.setColumnWidth(1, 150);
    configSheet.setColumnWidth(2, 600);
  }

  const data = configSheet.getDataRange().getValues();
  let config = {};
  for (let i = 1; i < data.length; i++) {
    const key = safeStr(data[i][0]).trim();
    const val = safeStr(data[i][1]).trim();
    if (key) config[key] = val;
  }

  return jsonResponse({ success: true, config: config });
}
