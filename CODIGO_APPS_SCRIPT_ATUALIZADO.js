// ═══════════════════════════════════════════════════════════
// APPS SCRIPT ATUALIZADO (Integração com Supabase PDL FLOW)
// ═══════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1FmJy-gHRMpF4thwWotX-tFQUoUGJ4yCbd4eIsjnBlVg';

// --- CONFIGURAÇÕES SUPABASE ---
const SUPABASE_URL = "https://vntfebxbsipumjswimmo.supabase.co"; // URL do seu Supabase
const SUPABASE_KEY = "sb_publishable_-Aef7s60PytmsxHfMGmKlA_olOmo3YG"; // Chave Anon
const ADMIN_UUID   = "COLE_AQUI_SEU_USER_UID"; // Vá em Authentication > Users e copie seu User UID

// ── Helpers ──
function getSs() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function getSheet(name) { return getSs().getSheetByName(name); }

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeStr(val) {
  if (val === undefined || val === null) return '';
  return String(val);
}

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
      case 'syncAllClients': return syncAllClients();
      default: return jsonResponse({ error: 'Ação desconhecida: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.toString(), stack: err.stack });
  }
}

// ═══════════════════════════════════════════════════════════
// INTEGRAÇÃO COM SUPABASE (Sincronização do Briefing)
// ═══════════════════════════════════════════════════════════
function syncToSupabase(codigo) {
  try {
    const ss = getSs();
    const painel = ss.getSheetByName('Painel');
    const pMap = getTableMap(painel, { 'Nome': ['nome'], 'Código': ['codigo'], 'Status': ['status'] });
    
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
    if (!clientName) return;

    // Pega todas as respostas da aba do cliente
    let briefingData = {};
    const clientSheet = ss.getSheetByName(clientName);
    if (clientSheet) {
      const cData = clientSheet.getDataRange().getValues();
      for (let i = 1; i < cData.length; i++) {
        const q = safeStr(cData[i][0]).trim();
        const a = safeStr(cData[i][1]).trim();
        if (q && q !== 'undefined' && a) {
          // Mapeia perguntas da planilha para chaves snake_case do PDL Flow
          let key = q; // fallback: usa a pergunta como chave (KEY_MAP cobre via alias)
          const qLower = q.toLowerCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // remove acentos para comparação

          // Identificação
          if (qLower.includes("nome") && !qLower.includes("empresa") && !qLower.includes("negocio")) key = "responsible_name";
          if (qLower.includes("empresa") || qLower.includes("negocio") || qLower.includes("estabelecimento")) key = "company_name";
          if ((qLower.includes("segmento") || qLower.includes("nicho")) && !qLower.includes("empresa")) key = "segment";

          // Contato
          if (qLower.includes("telefone") || (qLower.includes("whatsapp") && qLower.includes("ddd"))) key = "phone";
          if (qLower.includes("email") || qLower.includes("e-mail") || qLower.includes("correio eletronico")) key = "email";

          // Localização
          if (qLower.includes("cidade") || qLower.includes("estado") || qLower.includes("municipio")) key = "city_state";
          if ((qLower.includes("bairro") || qLower.includes("regiao") || qLower.includes("cidades atendidas")) && !qLower.includes("cidade e estado")) key = "areas";

          // Presença digital
          if (qLower.includes("site") && (qLower.includes("tem") || qLower.includes("qual") || qLower.includes("endereco"))) key = "website";
          if (qLower.includes("instagram") || qLower.includes("redes sociais") || qLower.includes("social")) key = "socials";

          // Negócio
          if (qLower.includes("principal produto") || qLower.includes("principal servico") || qLower.includes("mais vende")) key = "main_service";
          if ((qLower.includes("outros") && (qLower.includes("produto") || qLower.includes("servico")))) key = "other_services";
          if (qLower.includes("problema") && qLower.includes("resolve")) key = "problem_solved";
          if (qLower.includes("quem costuma comprar") || (qLower.includes("publico") && !qLower.includes("nome"))) key = "audience";
          if (qLower.includes("clientes chegam") || qLower.includes("adquire clientes") || qLower.includes("captacao")) key = "acquisition";
          if (qLower.includes("escolher voce") || qLower.includes("diferencial") || qLower.includes("diferencia")) key = "differentiator";
          if (qLower.includes("elogiam") || qLower.includes("elogio")) key = "praises";
          if (qLower.includes("concorrente")) key = "competitors";

          // Operacional
          if (qLower.includes("horario") && (qLower.includes("funcionamento") || qLower.includes("atendimento"))) key = "hours";
          if (qLower.includes("atende") && (qLower.includes("local") || qLower.includes("online") || qLower.includes("delivery") || qLower.includes("forma"))) key = "service_modes";
          if (qLower.includes("pagamento")) key = "payment_methods";
          if (qLower.includes("agendamento") && qLower.includes("ordem")) key = "scheduling";
          if (qLower.includes("sem agendamento") || (qLower.includes("atende") && qLower.includes("sem"))) key = "walkin";
          if (qLower.includes("atendimentos") && qLower.includes("dia")) key = "daily_capacity";
          if (qLower.includes("quanto tempo") && qLower.includes("atendimento")) key = "avg_duration";
          if (qLower.includes("trabalha sozinho") || (qLower.includes("equipe") && qLower.includes("quantas"))) key = "team";

          // Dúvidas / Restrições
          if (qLower.includes("duvidas") || qLower.includes("perguntas frequentes") || qLower.includes("faq")) key = "faq";
          if (qLower.includes("nao faz") || qLower.includes("nao atende") || qLower.includes("restricao") || qLower.includes("limitacao")) key = "restrictions";

          // Estrutura física / local
          if (qLower.includes("ambiente") && (qLower.includes("interno") || qLower.includes("externo"))) key = "ambient";
          if (qLower.includes("estacionamento")) key = "parking";
          if (qLower.includes("acessibilidade") || qLower.includes("locomocao") || qLower.includes("cadeirante")) key = "accessibility";
          if (qLower.includes("crianca")) key = "kid_friendly";
          if (qLower.includes("wi-fi") || qLower.includes("wifi") || qLower.includes("internet para cliente")) key = "wifi";
          if (qLower.includes("coberto") || qLower.includes("ar livre")) key = "covered";
          if (qLower.includes("espera") || qLower.includes("fila")) key = "wait_time";
          if (qLower.includes("banheiro")) key = "restroom";
          if (qLower.includes("facil acesso")) key = "easy_access";

          // Identidade
          if (qLower.includes("historia") || qLower.includes("2 frases sobre voce") || qLower.includes("quem e voce")) key = "bio";
          if (qLower.includes("frase curta") || qLower.includes("slogan") || qLower.includes("tagline") || qLower.includes("resuma o que")) key = "slogan";
          if (qLower.includes("abriu") || qLower.includes("abertura") || qLower.includes("fundacao") || qLower.includes("quando comecou")) key = "opening_date";

          // Promoções / Fotos
          if (qLower.includes("promocao") || qLower.includes("oferta")) key = "promotions";
          if (qLower.includes("fotos") || qLower.includes("imagens") || qLower.includes("fotografias")) key = "has_photos";
          
          briefingData[key] = a;
        }
      }
    }

    // Identificar nome da empresa vs nome da pessoa
    const companyName = briefingData["company_name"] || clientName;

    const payload = {
      briefing_token: codigo,
      name: clientName,
      company_name: companyName,
      user_id: ADMIN_UUID,
      briefing_data: briefingData,
      status: "active",
      current_phase_id: 1 // Fase de Onboarding
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "apikey": SUPABASE_KEY,
        "Authorization": "Bearer " + SUPABASE_KEY,
        "Prefer": "resolution=merge-duplicates" // Se já existir o token, ele atualiza
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    UrlFetchApp.fetch(SUPABASE_URL + "/rest/v1/clients", options);
  } catch(e) {
    console.error("Erro no sync Supabase", e);
  }
}

// ═══════════════════════════════════════════════════════════
// O RESTO DO CÓDIGO PERMANECE IGUAL, APENAS CHAMA syncToSupabase()
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

function detectFormularioColumns(formularioSheet) {
  const fData = formularioSheet.getDataRange().getValues();
  let headerRow = -1;
  let qCol = -1;
  let obsCol = -1;

  for (let r = 0; r < Math.min(5, fData.length); r++) {
    for (let c = 0; c < fData[r].length; c++) {
      const val = safeStr(fData[r][c]).toLowerCase().trim()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "");

      if (qCol === -1) {
        if (val === 'formulario' || val === 'pergunta' || val === 'perguntas' ||
            val === 'questao'    || val === 'questoes') {
          qCol = c;
          headerRow = r;
        }
      }

      if (obsCol === -1) {
        if (val === 'observacoes' || val === 'observacao' || val === 'obs' ||
            val === 'notas'       || val === 'nota') {
          obsCol = c;
          if (headerRow === -1) headerRow = r;
        }
      }
    }
    if (qCol >= 0) break;
  }

  if (headerRow === -1) { headerRow = 0; qCol = 0; }
  if (qCol === -1) { qCol = 0; }

  return { fData, headerRow, qCol, obsCol };
}

function syncQuestionsFromFormulario(clientSheet, formularioSheet) {
  const { fData, headerRow, qCol, obsCol } = detectFormularioColumns(formularioSheet);

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

  clientSheet.getRange(1, 1).setValue('Pergunta').setFontWeight('bold').setBackground('#4f8eff').setFontColor('#ffffff');
  clientSheet.getRange(1, 2).setValue('Respostas').setFontWeight('bold').setBackground('#4f8eff').setFontColor('#ffffff');
  clientSheet.getRange(1, 3).setValue('Observações').setFontWeight('bold').setBackground('#f59e0b').setFontColor('#ffffff');

  for (let i = 0; i < questions.length; i++) {
    const row = i + 2;
    clientSheet.getRange(row, 1).setValue(questions[i]);
    clientSheet.getRange(row, 2).setValue(existingAnswers[questions[i]] || '');
    const obsValue = (obsFromForm[i] !== undefined) ? obsFromForm[i] : '';
    clientSheet.getRange(row, 3).setValue(obsValue);
  }

  const lastRow = clientSheet.getLastRow();
  const expectedLastRow = questions.length + 1;
  if (lastRow > expectedLastRow) {
    clientSheet.getRange(expectedLastRow + 1, 1, lastRow - expectedLastRow, 3).clearContent();
  }

  return questions.length;
}

function createClientTab(codigo) {
  const ss = getSs();
  const painel = ss.getSheetByName('Painel');
  const formulario = ss.getSheetByName('Formulário');
  if (!painel || !formulario) return jsonResponse({ error: 'Abas principais não encontradas.' });

  const pMap = getTableMap(painel, { 'Nome': ['nome'], 'Código': ['codigo'], 'Status': ['status'] });

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

  let clientSheet = ss.getSheetByName(clientName);
  let isNew = false;

  if (!clientSheet) {
    clientSheet = ss.insertSheet(clientName);
    isNew = true;
  }

  let totalQ = 0;
  if (isNew) {
    totalQ = syncQuestionsFromFormulario(clientSheet, formulario);
    clientSheet.setColumnWidth(1, 450);
    clientSheet.setColumnWidth(2, 450);
    clientSheet.setColumnWidth(3, 300);
    clientSheet.setFrozenRows(1);
  } else {
    const rows = clientSheet.getLastRow();
    totalQ = rows > 1 ? rows - 1 : 0;
  }

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

  reorderSheets(ss);
  
  // ---> SINCRONIZA COM O SUPABASE ASSIM QUE INICIA
  syncToSupabase(codigo);

  return jsonResponse({ success: true, questionsCount: totalQ });
}

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

  const cData = clientSheet.getDataRange().getValues();
  const fields = [];

  for (let i = 1; i < cData.length; i++) {
    const q   = safeStr(cData[i][0]).trim(); 
    const a   = safeStr(cData[i][1]).trim(); 
    const obs = safeStr(cData[i][2]).trim(); 
    if (q && q !== 'undefined') {
      fields.push({
        rowIndex:   i + 1,   
        question:   q,
        answer:     a,
        observacao: obs
      });
    }
  }

  return jsonResponse({ tabExists: true, nome: clientName, fields: fields });
}

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

  clientSheet.getRange(rowIndex, 2).setValue(resposta);

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

  // ---> SINCRONIZA COM O SUPABASE A CADA NOVA RESPOSTA
  syncToSupabase(codigo);

  return jsonResponse({ success: true, allDone: allDone });
}

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

function getConfig() {
  const ss = getSs();
  let configSheet = ss.getSheetByName('Configuração');
  
  if (!configSheet) {
    configSheet = ss.insertSheet('Configuração');
    configSheet.getRange(1, 1).setValue('Chave').setFontWeight('bold');
    configSheet.getRange(1, 2).setValue('Valor').setFontWeight('bold');
    configSheet.getRange(2, 1).setValue('SystemPrompt');
    configSheet.getRange(2, 2).setValue('Você é uma assistente virtual...');
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

// ═══════════════════════════════════════════════════════════
// 6. SINCRONIZAÇÃO EM MASSA (Planilha → Supabase)
// Chamada pelo botão "Sincronizar" nas Configurações do PDL Flow
// ═══════════════════════════════════════════════════════════
function syncAllClients() {
  try {
    const ss = getSs();
    const painel = ss.getSheetByName('Painel');
    if (!painel) return jsonResponse({ error: 'Aba Painel não encontrada.' });

    const pMap = getTableMap(painel, {
      'Nome': ['nome', 'cliente'],
      'Código': ['codigo', 'acesso'],
      'Status': ['status']
    });

    if (pMap.headerRow === -1 || pMap.mapping['Código'] === undefined) {
      return jsonResponse({ error: 'Cabeçalhos do Painel não encontrados.' });
    }

    const cNome = pMap.mapping['Nome'] !== undefined ? pMap.mapping['Nome'] : 0;
    const cCod  = pMap.mapping['Código'];
    let synced = 0;

    for (let i = pMap.headerRow + 1; i < pMap.data.length; i++) {
      const nome   = safeStr(pMap.data[i][cNome]).trim();
      const codigo = safeStr(pMap.data[i][cCod]).trim();
      if (!nome || !codigo) continue;

      // Chama a lógica de sync individual para cada cliente
      syncToSupabase(codigo);
      synced++;
    }

    return jsonResponse({ success: true, synced: synced });
  } catch(err) {
    return jsonResponse({ error: err.toString() });
  }
}
