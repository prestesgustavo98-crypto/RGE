/**
 * ROSE NAILS — Sistema de Gestão
 * Backend: Google Apps Script (Web App / API REST) + Google Sheets como banco de dados
 *
 * ─────────────────────────────────────────────────────────────
 * CONFIGURAÇÃO OBRIGATÓRIA ANTES DE PUBLICAR
 * ─────────────────────────────────────────────────────────────
 * 1) Abra Extensões > Apps Script no seu Google Sheets e cole este arquivo.
 * 2) Vá em "Configurações do projeto" (ícone de engrenagem) > "Propriedades do script"
 *    e crie as seguintes Script Properties:
 *      ADMIN_USER   -> seu usuário de login (ex: "gustavo")
 *      ADMIN_PASS   -> sua senha de login (ex: "SenhaForte123!")
 *      SECRET_KEY   -> uma string aleatória longa, usada para assinar o token
 *                      (ex: gere em https://randomkeygen.com e cole aqui)
 * 3) Substitua SPREADSHEET_ID abaixo pelo ID da sua planilha
 *    (fica na URL: docs.google.com/spreadsheets/d/ESTE_TRECHO/edit)
 * 4) Rode a função `setupSheets()` uma única vez (menu Executar > setupSheets)
 *    para criar automaticamente todas as abas e colunas necessárias.
 *
 * ─────────────────────────────────────────────────────────────
 * NOTA IMPORTANTE SOBRE DATAS E HORAS (leia antes de editar)
 * ─────────────────────────────────────────────────────────────
 * O Google Sheets AUTO-CONVERTE textos como "2026-07-06" ou "14:00" em
 * objetos Date internos, mesmo quando gravados via script. Isso quebrava
 * a agenda do site (eventos com data inválida eram descartados pelo
 * FullCalendar) e o filtro por mês do financeiro (deslocamento de 1 dia
 * por causa de fuso horário).
 *
 * Este arquivo neutraliza o problema em DUAS camadas:
 *   (a) NA ESCRITA: appendObject()/updateRowByField() forçam formato de
 *       texto ('@') nas colunas de data/hora ANTES de gravar o valor,
 *       impedindo a auto-conversão do Sheets.
 *   (b) NA LEITURA: normalizarData()/normalizarHora() reconvertem para
 *       string sempre que o valor vier como Date (cobre também linhas
 *       antigas já gravadas incorretamente, antes desta correção).
 * Não remova essas camadas ao editar o código.
 */

// ====================== CONFIGURAÇÃO ======================

const SPREADSHEET_ID = '1LjL9v8Eubd2kEQGZs_FbHqLfsN_GKRdEIvu15JrT9GQ';

const SHEETS = {
  CLIENTES: 'Clientes',
  AGENDAMENTOS: 'Agendamentos',
  FINANCEIRO: 'Financeiro',
  SERVICOS: 'Servicos'
};

const TOKEN_VALID_HOURS = 168; // 7 dias — reduz a frequência de re-login (antes: 12h)

// Horário de funcionamento (usado no cálculo de taxa de ocupação)
const HORARIO_ABERTURA = 8;   // 08:00
const HORARIO_FECHAMENTO = 19; // 19:00

// ====================== SETUP INICIAL ======================

/**
 * Execute esta função UMA VEZ pelo editor do Apps Script para criar
 * a estrutura de abas e colunas na planilha automaticamente.
 */
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const schemas = {
    [SHEETS.CLIENTES]: ['ID', 'Nome', 'Telefone', 'DataCadastro', 'HistoricoServicos', 'Observacoes'],
    [SHEETS.AGENDAMENTOS]: ['ID', 'ClienteID', 'ClienteNome', 'ClienteTelefone', 'Servico', 'Data', 'HoraInicio', 'HoraFim', 'Status', 'Valor', 'Observacoes'],
    [SHEETS.FINANCEIRO]: ['ID', 'Data', 'Tipo', 'Descricao', 'Categoria', 'Valor', 'FormaPagamento', 'StatusPagamento', 'AgendamentoID'],
    [SHEETS.SERVICOS]: ['ID', 'Nome', 'DuracaoMinutos', 'Preco']
  };

  Object.keys(schemas).forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const headers = schemas[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2E2E2E').setFontColor('#FFFFFF');
  });

  // Força as colunas de data/hora a começarem como texto puro, para
  // impedir que o Sheets auto-converta valores futuros em Date.
  const agSheet = ss.getSheetByName(SHEETS.AGENDAMENTOS);
  agSheet.getRange('F2:H1000').setNumberFormat('@'); // Data, HoraInicio, HoraFim
  const cliSheet = ss.getSheetByName(SHEETS.CLIENTES);
  cliSheet.getRange('D2:D1000').setNumberFormat('@'); // DataCadastro
  const finSheet = ss.getSheetByName(SHEETS.FINANCEIRO);
  finSheet.getRange('B2:B1000').setNumberFormat('@'); // Data

  // Serviços de exemplo (edite/apague livremente depois)
  const servicosSheet = ss.getSheetByName(SHEETS.SERVICOS);
  if (servicosSheet.getLastRow() === 1) {
    servicosSheet.appendRow(['SRV_1', 'Manicure Tradicional', 45, 35]);
    servicosSheet.appendRow(['SRV_2', 'Esmaltação em Gel', 60, 55]);
    servicosSheet.appendRow(['SRV_3', 'Alongamento em Fibra', 120, 130]);
    servicosSheet.appendRow(['SRV_4', 'Pedicure', 50, 40]);
  }

  Logger.log('Estrutura criada com sucesso. Remova/ajuste os serviços de exemplo se necessário.');
}

// ====================== ROTEAMENTO HTTP ======================

function doGet(e) {
  return handleRequest(e, 'GET');
}

function doPost(e) {
  return handleRequest(e, 'POST');
}

function handleRequest(e, method) {
  try {
    let params = {};

    if (method === 'GET') {
      params = e.parameter || {};
    } else {
      // Frontend envia Content-Type: text/plain para evitar preflight CORS.
      params = JSON.parse(e.postData.contents || '{}');
    }

    const action = params.action;
    if (!action) return respond({ error: true, message: 'Parâmetro "action" ausente.' });

    // ---- Ações públicas (não exigem token) ----
    if (action === 'login') {
      return respond(login(params.usuario, params.senha));
    }

    // ---- Todas as demais ações exigem token válido ----
    const auth = verifyToken(params.token);
    if (!auth.valid) {
      // authError permite ao frontend detectar expiração de sessão de forma
      // confiável (sem depender de comparação de texto) e forçar novo login.
      return respond({ error: true, authError: true, message: 'Sessão inválida ou expirada. Faça login novamente.' });
    }

    switch (action) {
      // Clientes
      case 'getClientes': return respond({ error: false, data: getClientes() });
      case 'addCliente': return respond(addCliente(params.data));
      case 'updateCliente': return respond(updateCliente(params.data));
      case 'deleteCliente': return respond(deleteCliente(params.id));

      // Agenda
      case 'getAgendamentos': return respond({ error: false, data: getAgendamentos(params.dataInicio, params.dataFim) });
      case 'checkDisponibilidade': return respond(checkDisponibilidade(params.data, params.horaInicio, params.horaFim, params.excludeId));
      case 'addAgendamento': return respond(addAgendamento(params.data));
      case 'updateAgendamento': return respond(updateAgendamento(params.data));
      case 'deleteAgendamento': return respond(deleteAgendamento(params.id));

      // Financeiro
      case 'getFinanceiro': return respond({ error: false, data: getFinanceiro(params.mes) });
      case 'addFinanceiro': return respond(addFinanceiro(params.data));
      case 'updateFinanceiro': return respond(updateFinanceiro(params.data));
      case 'deleteFinanceiro': return respond(deleteFinanceiro(params.id));

      // Serviços
      case 'getServicos': return respond({ error: false, data: getServicos() });
      case 'addServico': return respond(addServico(params.data));
      case 'updateServico': return respond(updateServico(params.data));
      case 'deleteServico': return respond(deleteServico(params.id));

      // Dashboard
      case 'getDashboard': return respond({ error: false, data: getDashboardMetrics(params.mes) });

      default:
        return respond({ error: true, message: 'Ação desconhecida: ' + action });
    }
  } catch (err) {
    return respond({ error: true, message: 'Erro no servidor: ' + err.message });
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====================== AUTENTICAÇÃO ======================

function login(usuario, senha) {
  const props = PropertiesService.getScriptProperties();
  const adminUser = props.getProperty('ADMIN_USER');
  const adminPass = props.getProperty('ADMIN_PASS');

  if (!adminUser || !adminPass) {
    return { error: true, message: 'Credenciais de administrador não configuradas no servidor.' };
  }
  if (usuario !== adminUser || senha !== adminPass) {
    return { error: true, message: 'Usuário ou senha incorretos.' };
  }
  return { error: false, token: generateToken(usuario) };
}

function generateToken(usuario) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET_KEY');
  const expiry = new Date().getTime() + TOKEN_VALID_HOURS * 60 * 60 * 1000;
  const payload = usuario + '|' + expiry;
  const signatureBytes = Utilities.computeHmacSha256Signature(payload, secret);
  const signature = Utilities.base64EncodeWebSafe(signatureBytes);
  const payloadEncoded = Utilities.base64EncodeWebSafe(payload);
  return payloadEncoded + '.' + signature;
}

function verifyToken(token) {
  try {
    if (!token || token.indexOf('.') === -1) return { valid: false };
    const secret = PropertiesService.getScriptProperties().getProperty('SECRET_KEY');
    const parts = token.split('.');
    const payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    const expectedSignature = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret));

    if (expectedSignature !== parts[1]) return { valid: false };

    const expiry = Number(payload.split('|')[1]);
    if (new Date().getTime() > expiry) return { valid: false };

    return { valid: true, usuario: payload.split('|')[0] };
  } catch (err) {
    return { valid: false };
  }
}

// ====================== HELPERS DE PLANILHA ======================

/**
 * Cache da planilha e das abas por execução. Cada requisição do Apps Script
 * roda em um contexto novo, então estas variáveis vivem apenas durante uma
 * única chamada — tempo suficiente para eliminar as reaberturas repetidas.
 * Antes, getSheet() chamava SpreadsheetApp.openById() a CADA uso (4 a 5 vezes
 * por gravação de agendamento, ~200-400ms cada). Agora abre no máximo uma vez.
 */
let _spreadsheetCache = null;
function getSpreadsheet_() {
  if (!_spreadsheetCache) _spreadsheetCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _spreadsheetCache;
}

const _sheetCache = {};
function getSheet(name) {
  if (!_sheetCache[name]) _sheetCache[name] = getSpreadsheet_().getSheetByName(name);
  return _sheetCache[name];
}

/** Normaliza texto para comparação: sem acentos, minúsculo, espaços colapsados. */
function normalizarTexto_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

/** Canoniza o Tipo de lançamento financeiro para "Entrada" | "Saida". */
function normalizarTipo_(value) {
  const s = normalizarTexto_(value);
  if (s === 'saida') return 'Saida';
  if (s === 'entrada') return 'Entrada';
  return String(value == null ? '' : value).trim();
}

/** Canoniza o status de pagamento para "Pago" | "Pendente". */
function normalizarStatusPagamento_(value) {
  const s = normalizarTexto_(value);
  if (s === 'pago') return 'Pago';
  if (s === 'pendente') return 'Pendente';
  return String(value == null ? '' : value).trim();
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(function (row) { return row.some(function (cell) { return cell !== ''; }); })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

/**
 * Grava uma nova linha. `textColumns` lista os nomes de coluna que DEVEM
 * permanecer como texto puro (datas/horas) — o formato '@' é aplicado
 * ANTES do valor ser escrito, o que impede o Sheets de auto-converter
 * o texto em Date/hora internamente.
 */
function appendObject(sheet, obj, textColumns) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  const newRowIndex = sheet.getLastRow() + 1;
  const range = sheet.getRange(newRowIndex, 1, 1, headers.length);

  if (textColumns && textColumns.length) {
    const formats = headers.map(function (h) { return textColumns.indexOf(h) > -1 ? '@' : 'General'; });
    range.setNumberFormats([formats]);
  }
  range.setValues([row]);
  return obj;
}

/**
 * Atualiza campos de uma linha pelo valor de `matchField`.
 * `textFields` (opcional) força formato de texto nas colunas de data/hora
 * antes de escrever, pela mesma razão de appendObject().
 */
function updateRowByField(sheet, matchField, matchValue, updates, textFields) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = headers.indexOf(matchField);
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][colIndex]) === String(matchValue)) {
      Object.keys(updates).forEach(function (key) {
        const ci = headers.indexOf(key);
        if (ci > -1) {
          const cell = sheet.getRange(r + 1, ci + 1);
          if (textFields && textFields.indexOf(key) > -1) cell.setNumberFormat('@');
          cell.setValue(updates[key]);
        }
      });
      // Retorna o objeto completo da linha já com as atualizações aplicadas.
      // Isso permite aos chamadores obter o registro final SEM reler a planilha
      // inteira (updateAgendamento fazia um getAgendamentos() completo só para
      // pegar o registro que acabara de escrever).
      const merged = {};
      headers.forEach(function (h, i) {
        merged[h] = updates[h] !== undefined ? updates[h] : data[r][i];
      });
      return merged;
    }
  }
  return null;
}

function deleteRowByField(sheet, matchField, matchValue) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = headers.indexOf(matchField);
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][colIndex]) === String(matchValue)) {
      sheet.deleteRow(r + 1);
      return true;
    }
  }
  return false;
}

function generateId(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

/**
 * Converte QUALQUER representação de data (string "YYYY-MM-DD", string
 * com timestamp completo, ou objeto Date vindo de auto-conversão do
 * Sheets) para uma string "YYYY-MM-DD" estável, SEM reprocessar strings
 * que já estão no formato correto — isso é o que evita o bug de
 * deslocamento de 1 dia por fuso horário (new Date("2026-07-06") é
 * interpretado como UTC e "volta" um dia em fusos negativos como
 * America/Porto_Velho).
 */
function normalizarData(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(value);
  const match = str.match(/^(\d{4}-\d{2}-\d{2})/);
  if (match) return match[1];
  const d = new Date(str);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Mesma lógica de normalizarData(), mas para horários ("HH:mm").
 */
function normalizarHora(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  const str = String(value);
  const match = str.match(/^(\d{2}:\d{2})/);
  return match ? match[1] : str;
}

/** Retorna uma cópia do agendamento com Data/HoraInicio/HoraFim normalizados. */
function normalizarAgendamento(a) {
  return Object.assign({}, a, {
    Data: normalizarData(a.Data),
    HoraInicio: normalizarHora(a.HoraInicio),
    HoraFim: normalizarHora(a.HoraFim)
  });
}

// ====================== CLIENTES ======================

function getClientes() {
  return sheetToObjects(getSheet(SHEETS.CLIENTES)).map(function (c) {
    return Object.assign({}, c, { DataCadastro: normalizarData(c.DataCadastro) });
  });
}

function addCliente(data) {
  if (!data || !data.Nome || !data.Telefone) {
    return { error: true, message: 'Nome e telefone são obrigatórios.' };
  }
  const obj = {
    ID: generateId('CLI'),
    Nome: data.Nome,
    Telefone: data.Telefone,
    DataCadastro: normalizarData(new Date()),
    HistoricoServicos: data.HistoricoServicos || '',
    Observacoes: data.Observacoes || ''
  };
  appendObject(getSheet(SHEETS.CLIENTES), obj, ['DataCadastro']);
  return { error: false, data: obj };
}

function updateCliente(data) {
  if (!data || !data.ID) return { error: true, message: 'ID do cliente é obrigatório.' };
  const ok = updateRowByField(getSheet(SHEETS.CLIENTES), 'ID', data.ID, data, ['DataCadastro']);
  return ok ? { error: false } : { error: true, message: 'Cliente não encontrado.' };
}

function deleteCliente(id) {
  const ok = deleteRowByField(getSheet(SHEETS.CLIENTES), 'ID', id);
  return ok ? { error: false } : { error: true, message: 'Cliente não encontrado.' };
}

/**
 * Procura um cliente pelo telefone (ignorando formatação: espaços, +, -, etc.).
 * Se não existir, cadastra automaticamente. É isto que fazia falta ser
 * chamado a partir de addAgendamento() — antes, nada disparava esta busca.
 */
function encontrarOuCriarCliente(nome, telefone) {
  const sheet = getSheet(SHEETS.CLIENTES);
  const clientes = sheetToObjects(sheet);
  const telefoneDigits = String(telefone || '').replace(/\D/g, '');
  const nomeNorm = normalizarTexto_(nome);

  // 1) Critério primário: telefone (ignorando formatação).
  // 2) Critério alternativo: nome normalizado (sem acento/caixa/espaços extras).
  //    Antes, só o telefone era usado — se o telefone viesse vazio ou diferente,
  //    um cliente já cadastrado era duplicado.
  const existente = clientes.find(function (c) {
    const telMatch = telefoneDigits !== '' &&
      String(c.Telefone || '').replace(/\D/g, '') === telefoneDigits;
    const nomeMatch = nomeNorm !== '' && normalizarTexto_(c.Nome) === nomeNorm;
    return telMatch || nomeMatch;
  });
  if (existente) return existente;

  const novo = {
    ID: generateId('CLI'),
    Nome: nome,
    Telefone: telefone,
    DataCadastro: normalizarData(new Date()),
    HistoricoServicos: '',
    Observacoes: 'Cadastrado automaticamente ao agendar.'
  };
  appendObject(sheet, novo, ['DataCadastro']);
  return novo;
}

// ====================== AGENDA ======================

function getAgendamentos(dataInicio, dataFim) {
  const all = sheetToObjects(getSheet(SHEETS.AGENDAMENTOS)).map(normalizarAgendamento);
  if (!dataInicio && !dataFim) return all;
  return all.filter(function (a) {
    if (dataInicio && a.Data < dataInicio) return false;
    if (dataFim && a.Data > dataFim) return false;
    return true;
  });
}

/**
 * Verifica conflito de horário no mesmo dia.
 * Retorna { error:false, disponivel: true|false }
 */
function checkDisponibilidade(data, horaInicio, horaFim, excludeId) {
  const dataAlvo = normalizarData(data);
  const inicioAlvo = normalizarHora(horaInicio);
  const fimAlvo = normalizarHora(horaFim);

  const agendamentos = getAgendamentos(dataAlvo, dataAlvo).filter(function (a) {
    return a.Status !== 'Cancelado' && String(a.ID) !== String(excludeId || '');
  });

  const conflito = agendamentos.some(function (a) {
    // Sobreposição de intervalos: início A < fim B  E  início B < fim A
    return inicioAlvo < a.HoraFim && a.HoraInicio < fimAlvo;
  });

  return { error: false, disponivel: !conflito };
}

/**
 * Cria o agendamento e dispara a cascata completa esperada pelo negócio:
 *   1. valida disponibilidade
 *   2. grava a linha em Agendamentos (com formato de texto forçado em Data/Horas)
 *   3. localiza ou cadastra o cliente em Clientes
 *   4. cria o lançamento financeiro correspondente em Financeiro
 * Isso é exatamente o que faltava antes: a função original parava no passo 2.
 */
function addAgendamento(data) {
  if (!data || !data.ClienteNome || !data.Data || !data.HoraInicio || !data.HoraFim) {
    return { error: true, message: 'Dados incompletos para o agendamento.' };
  }

  const dataNormalizada = normalizarData(data.Data);
  const inicioNormalizado = normalizarHora(data.HoraInicio);
  const fimNormalizado = normalizarHora(data.HoraFim);

  const disponibilidade = checkDisponibilidade(dataNormalizada, inicioNormalizado, fimNormalizado, null);
  if (!disponibilidade.disponivel) {
    return { error: true, message: 'Horário indisponível: já existe um agendamento nesse intervalo.' };
  }

  // 1) Cliente: busca por telefone, cadastra se necessário
  const cliente = encontrarOuCriarCliente(data.ClienteNome, data.ClienteTelefone);

  // 2) Agendamento
  const obj = {
    ID: generateId('AG'),
    ClienteID: cliente.ID,
    ClienteNome: data.ClienteNome,
    ClienteTelefone: data.ClienteTelefone || '',
    Servico: data.Servico || '',
    Data: dataNormalizada,
    HoraInicio: inicioNormalizado,
    HoraFim: fimNormalizado,
    Status: data.Status || 'Confirmado',
    Valor: data.Valor || 0,
    Observacoes: data.Observacoes || ''
  };
  appendObject(getSheet(SHEETS.AGENDAMENTOS), obj, ['Data', 'HoraInicio', 'HoraFim']);

  // 3) Financeiro: cria a movimentação prevista vinculada a este agendamento
  sincronizarFinanceiroDoAgendamento(obj);

  return { error: false, data: obj, cliente: cliente };
}

function updateAgendamento(data) {
  if (!data || !data.ID) return { error: true, message: 'ID do agendamento é obrigatório.' };

  const camposParaAtualizar = Object.assign({}, data);

  // Se a data/horário está sendo alterado, valida disponibilidade novamente
  if (data.Data && data.HoraInicio && data.HoraFim) {
    const dataNormalizada = normalizarData(data.Data);
    const inicioNormalizado = normalizarHora(data.HoraInicio);
    const fimNormalizado = normalizarHora(data.HoraFim);

    const disponibilidade = checkDisponibilidade(dataNormalizada, inicioNormalizado, fimNormalizado, data.ID);
    if (!disponibilidade.disponivel) {
      return { error: true, message: 'Horário indisponível para o novo período informado.' };
    }
    camposParaAtualizar.Data = dataNormalizada;
    camposParaAtualizar.HoraInicio = inicioNormalizado;
    camposParaAtualizar.HoraFim = fimNormalizado;
  }

  const sheet = getSheet(SHEETS.AGENDAMENTOS);
  const merged = updateRowByField(sheet, 'ID', data.ID, camposParaAtualizar, ['Data', 'HoraInicio', 'HoraFim']);
  if (!merged) return { error: true, message: 'Agendamento não encontrado.' };

  // Usa o registro completo devolvido por updateRowByField (sem reler a planilha)
  // e normaliza datas/horas para manter o Financeiro sincronizado.
  const atualizado = normalizarAgendamento(merged);
  sincronizarFinanceiroDoAgendamento(atualizado);

  return { error: false, data: atualizado };
}

function deleteAgendamento(id) {
  const sheet = getSheet(SHEETS.AGENDAMENTOS);
  const ok = deleteRowByField(sheet, 'ID', id);
  if (!ok) return { error: true, message: 'Agendamento não encontrado.' };

  // Remove o lançamento financeiro vinculado, se existir, para não deixar
  // um registro "fantasma" de receita esperada de um agendamento que não existe mais.
  const financeiro = sheetToObjects(getSheet(SHEETS.FINANCEIRO));
  const vinculado = financeiro.find(function (f) { return String(f.AgendamentoID) === String(id); });
  if (vinculado) deleteRowByField(getSheet(SHEETS.FINANCEIRO), 'ID', vinculado.ID);

  return { error: false };
}

// ====================== FINANCEIRO ======================

function getFinanceiro(mes) {
  const all = sheetToObjects(getSheet(SHEETS.FINANCEIRO)).map(function (f) {
    return Object.assign({}, f, {
      Data: normalizarData(f.Data),
      // Normaliza Tipo e StatusPagamento na LEITURA para que variações digitadas
      // direto na planilha ("Saída", "saida", "SAIDA") não quebrem os filtros
      // estritos (===) do frontend e do Dashboard.
      Tipo: normalizarTipo_(f.Tipo),
      StatusPagamento: normalizarStatusPagamento_(f.StatusPagamento)
    });
  });
  if (!mes) return all; // mes no formato "YYYY-MM"
  return all.filter(function (f) { return f.Data.indexOf(mes) === 0; });
}

function addFinanceiro(data) {
  if (!data || !data.Tipo || !data.Valor) {
    return { error: true, message: 'Tipo e valor são obrigatórios.' };
  }
  const obj = {
    ID: generateId('FIN'),
    Data: normalizarData(data.Data || new Date()),
    Tipo: normalizarTipo_(data.Tipo), // "Entrada" | "Saida"
    Descricao: data.Descricao || '',
    Categoria: data.Categoria || '',
    Valor: data.Valor,
    FormaPagamento: data.FormaPagamento || 'Pix',
    StatusPagamento: normalizarStatusPagamento_(data.StatusPagamento || 'Pago'), // "Pago" | "Pendente"
    AgendamentoID: data.AgendamentoID || ''
  };
  appendObject(getSheet(SHEETS.FINANCEIRO), obj, ['Data']);
  return { error: false, data: obj };
}

function updateFinanceiro(data) {
  if (!data || !data.ID) return { error: true, message: 'ID do lançamento é obrigatório.' };
  if (data.Data) data.Data = normalizarData(data.Data);
  if (data.Tipo !== undefined) data.Tipo = normalizarTipo_(data.Tipo);
  if (data.StatusPagamento !== undefined) data.StatusPagamento = normalizarStatusPagamento_(data.StatusPagamento);
  const ok = updateRowByField(getSheet(SHEETS.FINANCEIRO), 'ID', data.ID, data, ['Data']);
  return ok ? { error: false } : { error: true, message: 'Lançamento não encontrado.' };
}

function deleteFinanceiro(id) {
  const ok = deleteRowByField(getSheet(SHEETS.FINANCEIRO), 'ID', id);
  return ok ? { error: false } : { error: true, message: 'Lançamento não encontrado.' };
}

/**
 * Mantém o Financeiro sincronizado com o estado atual de um agendamento:
 *  - sem valor definido            -> não lança nada
 *  - agendamento Cancelado         -> remove o lançamento vinculado (se existir)
 *  - agendamento ativo, sem vínculo-> cria um lançamento (Entrada, Pendente)
 *  - agendamento ativo, já vinculado -> atualiza descrição/valor/data do lançamento
 *
 * O lançamento nasce como "Pendente" propositalmente: ele deve aparecer na
 * aba/tela Financeiro imediatamente (o que o negócio pediu), mas só deve
 * contar como receita realizada no Dashboard depois que for marcado como
 * "Pago" (produção manual ou confirmação de recebimento via Pix).
 */
function sincronizarFinanceiroDoAgendamento(agendamento) {
  const sheet = getSheet(SHEETS.FINANCEIRO);
  const registros = sheetToObjects(sheet);
  const vinculado = registros.find(function (f) { return String(f.AgendamentoID) === String(agendamento.ID); });

  if (agendamento.Status === 'Cancelado') {
    if (vinculado) deleteRowByField(sheet, 'ID', vinculado.ID);
    return;
  }

  if (!agendamento.Valor || Number(agendamento.Valor) <= 0) return;

  const descricao = 'Agendamento: ' + agendamento.Servico + ' - ' + agendamento.ClienteNome;

  if (vinculado) {
    updateRowByField(sheet, 'ID', vinculado.ID, {
      Data: agendamento.Data,
      Descricao: descricao,
      Valor: agendamento.Valor
    }, ['Data']);
  } else {
    const obj = {
      ID: generateId('FIN'),
      Data: agendamento.Data,
      Tipo: 'Entrada',
      Descricao: descricao,
      Categoria: 'Serviço',
      Valor: agendamento.Valor,
      FormaPagamento: 'Pix',
      StatusPagamento: 'Pendente',
      AgendamentoID: agendamento.ID
    };
    appendObject(sheet, obj, ['Data']);
  }
}

// ====================== SERVIÇOS ======================

function getServicos() {
  return sheetToObjects(getSheet(SHEETS.SERVICOS)).map(function (s) {
    return Object.assign({}, s, {
      DuracaoMinutos: Number(s.DuracaoMinutos) || 0,
      Preco: Number(s.Preco) || 0
    });
  });
}

function addServico(data) {
  if (!data || !data.Nome) {
    return { error: true, message: 'Nome do serviço é obrigatório.' };
  }
  const obj = {
    ID: generateId('SRV'),
    Nome: data.Nome,
    DuracaoMinutos: Number(data.DuracaoMinutos) || 60,
    Preco: Number(data.Preco) || 0
  };
  appendObject(getSheet(SHEETS.SERVICOS), obj);
  return { error: false, data: obj };
}

function updateServico(data) {
  if (!data || !data.ID) return { error: true, message: 'ID do serviço é obrigatório.' };
  const updates = {};
  if (data.Nome !== undefined) updates.Nome = data.Nome;
  if (data.DuracaoMinutos !== undefined) updates.DuracaoMinutos = Number(data.DuracaoMinutos) || 0;
  if (data.Preco !== undefined) updates.Preco = Number(data.Preco) || 0;
  const ok = updateRowByField(getSheet(SHEETS.SERVICOS), 'ID', data.ID, updates);
  return ok ? { error: false } : { error: true, message: 'Serviço não encontrado.' };
}

function deleteServico(id) {
  const ok = deleteRowByField(getSheet(SHEETS.SERVICOS), 'ID', id);
  return ok ? { error: false } : { error: true, message: 'Serviço não encontrado.' };
}

// ====================== DASHBOARD ======================

/**
 * Calcula métricas do mês informado (formato "YYYY-MM"). Se omitido, usa o mês atual.
 * - receitaMensal: soma de Financeiro (Tipo=Entrada, StatusPagamento=Pago) no mês
 * - ticketMedio: receitaMensal / nº de agendamentos concluídos no mês
 * - taxaOcupacao: horas agendadas (status != Cancelado) / horas totais disponíveis no mês
 */
function getDashboardMetrics(mes) {
  const mesAlvo = mes || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

  const financeiro = getFinanceiro(mesAlvo);
  const receitaMensal = financeiro
    .filter(function (f) { return f.Tipo === 'Entrada' && f.StatusPagamento === 'Pago'; })
    .reduce(function (sum, f) { return sum + Number(f.Valor || 0); }, 0);

  const saidas = financeiro
    .filter(function (f) { return f.Tipo === 'Saida'; })
    .reduce(function (sum, f) { return sum + Number(f.Valor || 0); }, 0);

  const agendamentosMes = getAgendamentos(mesAlvo + '-01', mesAlvo + '-31')
    .filter(function (a) { return a.Status !== 'Cancelado'; });

  const concluidos = agendamentosMes.filter(function (a) { return a.Status === 'Concluído'; });
  const ticketMedio = concluidos.length > 0 ? receitaMensal / concluidos.length : 0;

  // Produtividade por período do mês: distribui os atendimentos (não cancelados)
  // em 4 semanas conforme o dia do mês (1-7, 8-14, 15-21, 22+).
  const produtividadePorSemana = [
    { semana: 1, label: 'Semana 1', qtd: 0 },
    { semana: 2, label: 'Semana 2', qtd: 0 },
    { semana: 3, label: 'Semana 3', qtd: 0 },
    { semana: 4, label: 'Semana 4', qtd: 0 }
  ];
  agendamentosMes.forEach(function (a) {
    const dia = Number(String(a.Data).slice(8, 10)) || 1;
    let idx = Math.floor((dia - 1) / 7);
    if (idx > 3) idx = 3; // dias 29-31 caem na Semana 4
    produtividadePorSemana[idx].qtd++;
  });

  // Horas ocupadas: soma da duração de cada agendamento não cancelado
  const minutosOcupados = agendamentosMes.reduce(function (sum, a) {
    return sum + minutosEntreHorarios(a.HoraInicio, a.HoraFim);
  }, 0);

  // Horas disponíveis no mês: dias úteis (seg-sáb) * horas de expediente
  const [ano, mesNum] = mesAlvo.split('-').map(Number);
  const diasNoMes = new Date(ano, mesNum, 0).getDate();
  let diasUteis = 0;
  for (let d = 1; d <= diasNoMes; d++) {
    const diaSemana = new Date(ano, mesNum - 1, d).getDay();
    if (diaSemana !== 0) diasUteis++; // considera domingo como fechado
  }
  const minutosDisponiveis = diasUteis * (HORARIO_FECHAMENTO - HORARIO_ABERTURA) * 60;

  const taxaOcupacao = minutosDisponiveis > 0 ? (minutosOcupados / minutosDisponiveis) * 100 : 0;

  return {
    mes: mesAlvo,
    receitaMensal: round2(receitaMensal),
    saidasMensal: round2(saidas),
    saldoMensal: round2(receitaMensal - saidas),
    ticketMedio: round2(ticketMedio),
    taxaOcupacao: round2(taxaOcupacao),
    atendimentosMes: agendamentosMes.length,
    totalAgendamentos: agendamentosMes.length,
    totalConcluidos: concluidos.length,
    produtividadePorSemana: produtividadePorSemana
  };
}

function minutosEntreHorarios(inicio, fim) {
  const [h1, m1] = String(inicio).split(':').map(Number);
  const [h2, m2] = String(fim).split(':').map(Number);
  return (h2 * 60 + m2) - (h1 * 60 + m1);
}

function round2(n) {
  return Math.round(n * 100) / 100;
}