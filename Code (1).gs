// ============================================================
// NANINHAS SAGRADAS - Apps Script (back-end)
// Cole este código no Apps Script da planilha e crie uma nova versão
// da implantação (Implantar > Gerenciar implantações > Editar > Nova versão).
//
// ANTES DE USAR: defina o código de acesso da equipe em
// Configurações do projeto > Propriedades do script > Adicionar propriedade
//   Propriedade: CODIGO
//   Valor: o código que a equipe vai digitar na primeira vez que abrir o app
// ============================================================

function doGet() {
  // Não devolve dado nenhum por GET. Serve só para testar se a implantação está de pé.
  return resposta({ ok: true, app: 'Naninhas Sagradas' });
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');

    const codigo = PropertiesService.getScriptProperties().getProperty('CODIGO');
    if (!codigo || normalizar(body.token) !== normalizar(codigo)) throw new Error('Não autorizado');

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      result = executar(body);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    result = { error: err.message };
  }
  return resposta(result);
}

// Ignora maiúsculas, acentos e espaços, porque o teclado do celular costuma pôr a primeira letra em maiúscula.
function normalizar(v) {
  return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '').toLowerCase();
}

function executar(body) {
  switch (body.action) {
    case 'getRegistros':    return getRegistros();
    case 'getNaninhas':     return getNaninhas();
    case 'addRegistro':     return addRegistro(body.data);
    case 'devolverNaninha': return devolverNaninha(body.id, body.dataEntrada, body.localEntrada);
    case 'addNaninha':      return addNaninha(body.data);
    case 'deleteNaninha':   return deleteNaninha(body.id);
    case 'updateRegistro':  return updateRegistro(body.id, body.data);
    case 'updateNaninha':   return updateNaninha(body.id, body.data);
    default: throw new Error('Ação não reconhecida: ' + body.action);
  }
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function aba(nome) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sheet) throw new Error('Aba não encontrada: ' + nome);
  return sheet;
}

// Datas voltam como texto yyyy-MM-dd no fuso do script, para o app não depender de fuso.
function valor(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (v === 'TRUE') return true;
  if (v === 'FALSE') return false;
  return v;
}

// A planilha antiga tem o cabeçalho "naniaId"; o app usa "naninaId".
function chave(h) {
  return h === 'naniaId' ? 'naninaId' : h;
}

function lerAba(nome) {
  const data = aba(nome).getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0].map(chave);
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = valor(row[i]));
    return obj;
  });
}

function getRegistros() { return lerAba('registros'); }
function getNaninhas()  { return lerAba('naninhas'); }

function addRegistro(data) {
  if (!data || !data.id || !data.crianca) throw new Error('Registro incompleto');
  const sheet = aba('registros');
  garantirColunasLocal(sheet);
  checarLocal(data.localSaida);
  sheet.appendRow([
    data.id, data.crianca, data.responsavel, data.telefone,
    data.naninaId, data.naniaNome, data.naniaEmoji, data.oracao,
    data.dataSaida, data.dadoOracao === true, false, '',
    data.localSaida || '', ''
  ]);
  return { ok: true };
}

function devolverNaninha(id, dataEntrada, localEntrada) {
  const sheet = aba('registros');
  garantirColunasLocal(sheet);
  checarLocal(localEntrada);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 11).setValue(true);        // coluna K: devolvida
      sheet.getRange(i + 1, 12).setValue(dataEntrada); // coluna L: dataEntrada
      sheet.getRange(i + 1, 14).setValue(localEntrada || ''); // coluna N: localEntrada
      return { ok: true };
    }
  }
  throw new Error('Registro não encontrado');
}

function addNaninha(data) {
  if (!data || !data.id || !data.nome) throw new Error('Naninha incompleta');
  const sheet = aba('naninhas');
  garantirColunaFoto(sheet);
  checarFoto(data.foto);
  sheet.appendRow([data.id, data.nome, data.emoji, data.oracao, data.foto || '']);
  return { ok: true };
}

function deleteNaninha(id) {
  const sheet = aba('naninhas');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error('Naninha não encontrada');
}

// Edita um empréstimo (colunas B a N). A coluna A (id) nunca muda.
function updateRegistro(id, data) {
  if (!data || !data.crianca) throw new Error('Registro incompleto');
  const sheet = aba('registros');
  garantirColunasLocal(sheet);
  checarLocal(data.localSaida);
  checarLocal(data.localEntrada);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      const devolvida = data.devolvida === true;
      sheet.getRange(i + 1, 2, 1, 13).setValues([[
        data.crianca, data.responsavel, data.telefone,
        data.naninaId, data.naniaNome, data.naniaEmoji, data.oracao,
        data.dataSaida, data.dadoOracao === true,
        devolvida, devolvida ? (data.dataEntrada || '') : '',
        data.localSaida || '', devolvida ? (data.localEntrada || '') : ''
      ]]);
      return { ok: true };
    }
  }
  throw new Error('Registro não encontrado');
}

// Edita nome, emoji e oração de uma naninha. Empréstimos antigos mantêm a cópia do que foi enviado.
function updateNaninha(id, data) {
  if (!data || !data.nome) throw new Error('Naninha incompleta');
  const sheet = aba('naninhas');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      garantirColunaFoto(sheet);
      checarFoto(data.foto);
      const foto = data.foto === undefined ? (rows[i][4] || '') : data.foto;
      sheet.getRange(i + 1, 2, 1, 4).setValues([[data.nome, data.emoji, data.oracao, foto]]);
      return { ok: true };
    }
  }
  throw new Error('Naninha não encontrada');
}

// A foto da naninha fica na coluna E da aba naninhas (miniatura em texto). O cabeçalho é criado sozinho.
function garantirColunaFoto(sheet) {
  if (sheet.getRange(1, 5).getValue() !== 'foto') sheet.getRange(1, 5).setValue('foto');
}
// Uma célula do Sheets aceita até 50.000 caracteres.
function checarFoto(f) {
  if (f && String(f).length > 45000) throw new Error('Foto muito grande');
}

// Onde aconteceu o empréstimo e a devolução: colunas M (localSaida) e N (localEntrada). Os cabeçalhos são criados sozinhos.
function garantirColunasLocal(sheet) {
  if (sheet.getRange(1, 13).getValue() !== 'localSaida') sheet.getRange(1, 13).setValue('localSaida');
  if (sheet.getRange(1, 14).getValue() !== 'localEntrada') sheet.getRange(1, 14).setValue('localEntrada');
}
function checarLocal(l) {
  if (l && ['Paróquia', 'Capela'].indexOf(l) < 0) throw new Error('Local inválido');
}
