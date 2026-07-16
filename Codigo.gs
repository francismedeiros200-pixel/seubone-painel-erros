/**
 * ============================================================================
 *  BACKEND — Painel de Erros SeuBoné (Google Apps Script / Web App)
 * ----------------------------------------------------------------------------
 *  Este script liga a planilha "Cadastro de erros" ao painel (index.html).
 *
 *  Ele faz 3 coisas:
 *    - doGet()               → lê a planilha e devolve os registros em JSON
 *    - doPost() action:criar → registra um novo erro (append de linha)
 *    - doPost() action:audit → atualiza os campos de auditoria de uma linha
 *
 *  IMPORTANTE — O mapeamento de colunas é feito por NOME DO CABEÇALHO
 *  (tolerante a acento/maiúscula), não por posição fixa. Assim, se alguém
 *  reordenar colunas na planilha, o painel continua funcionando.
 *
 *  Se o seu script atual já funciona para LEITURA, você pode:
 *    (a) me mandar o código atual para eu mesclar só a parte de escrita, ou
 *    (b) publicar este aqui como uma implantação NOVA e comparar a saída do
 *        /exec com a antiga ANTES de trocar a URL no painel (ver README).
 * ============================================================================
 */

/**
 * Aba onde ficam os registros. O gid é o identificador fixo da aba (vem da URL:
 * .../edit?gid=396842648) e NÃO muda mesmo que renomeiem a aba ou reordenem.
 * É a forma mais segura de mirar a aba certa. O nome fica só como reserva.
 */
var SHEET_GID = 396842648;
var SHEET_NAME = 'Respostas do Form';

/**
 * Mapa lógico: chave que o painel usa  →  lista de "pistas" de cabeçalho.
 * A primeira pista que casar com um cabeçalho da planilha define a coluna.
 * A ordem importa (a primeira que casar vence).
 */
var COLUNAS = {
  data:          ['carimbo de data/hora'],  // a data real fica na 1ª coluna; o fallback abaixo garante o índice 0
  auditoria:     ['auditoria'],
  idVenda:       ['id da venda'],
  nomeCard:      ['nome do card'],
  descricao:     ['descricao e solucao', 'descricao do erro', 'descricao'],
  quemCadastrou: ['quem cadastrou'],
  culpaDe:       ['culpa de', 'culpa'],
  setor:         ['setor do problema', 'setor'],
  responsavel:   ['responsavel'],
  empresa:       ['empresa'],
  tipoProblema:  ['tipo de problema'],
  subproblema:   ['subproblema'],
  qtd:           ['quantidade de produtos', 'quantidade'],
  custo:         ['custo do erro', 'custo'],
  tipoProduto:   ['tipo de produto'],
  queFim:        ['que fim'],
  // QUIRK confirmado por engenharia reversa: o "tipo de resolução" que o painel
  // usa vem da coluna "Solução". A coluna "Tipo de resolução" tende a ficar vazia.
  tipoResolucao: ['solucao', 'tipo de resolucao'],
  // Coluna NOVA (adicione um cabeçalho "Status" na planilha para ativar o workflow de status).
  // Sem essa coluna, o painel deriva o status de "Auditoria realizada?" automaticamente.
  status:        ['status'],
};

/* ============================ HELPERS ============================ */

function norm_(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // remove acentos
    .replace(/\s+/g, ' ')
    .trim();
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  // 1º: pela gid (identificador fixo da aba) — o mais confiável
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  // 2º: pelo nome; 3º: primeira aba (último recurso)
  var byName = ss.getSheetByName(SHEET_NAME);
  return byName || sheets[0];
}

/** Constrói { chaveLogica: indiceDaColuna(0-based) } a partir da linha 1. */
function buildColMap_(header) {
  var normHeaders = header.map(norm_);
  var map = {};
  Object.keys(COLUNAS).forEach(function (key) {
    var pistas = COLUNAS[key];
    var found = null;
    // 1ª passada: match EXATO — evita colisões de substring
    // (ex.: "Solução" vs "Descrição e solução do erro").
    for (var p = 0; p < pistas.length && found === null; p++) {
      var pe = norm_(pistas[p]);
      for (var c = 0; c < normHeaders.length; c++) {
        if (normHeaders[c] === pe) { found = c; break; }
      }
    }
    // 2ª passada: match por inclusão (para cabeçalhos com sufixos/pontuação).
    for (var p2 = 0; p2 < pistas.length && found === null; p2++) {
      var pi = norm_(pistas[p2]);
      for (var c2 = 0; c2 < normHeaders.length; c2++) {
        if (normHeaders[c2].indexOf(pi) !== -1) { found = c2; break; }
      }
    }
    if (found !== null) map[key] = found;
  });
  // Fallback: a data real está na 1ª coluna do layout atual.
  if (map.data == null) map.data = 0;
  return map;
}

/** "R$ 1.245,60" | "1.245,60" | "-" | 80  →  Number (ou '' se vazio/inválido). */
function parseNumber_(v) {
  if (v === '' || v == null) return '';
  if (typeof v === 'number') return v;
  var s = String(v).replace(/r\$/i, '').replace(/\s/g, '').trim();
  if (s === '' || s === '-') return '';
  // remove separador de milhar "." e troca vírgula decimal por "."
  s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? '' : n;
}

function parseBool_(v) {
  if (v === true) return true;
  var s = norm_(v);
  return s === 'true' || s === 'sim' || s === 'x' || s === 'verdadeiro' || s === '1';
}

/** Data (Date ou "dd/mm/aaaa") → "dd/mm/aaaa". */
function fmtDate_(v) {
  if (v instanceof Date) {
    var d = ('0' + v.getDate()).slice(-2);
    var m = ('0' + (v.getMonth() + 1)).slice(-2);
    return d + '/' + m + '/' + v.getFullYear();
  }
  return String(v == null ? '' : v).trim();
}

/** Extrai a 1ª URL de um texto (o painel usa isso como "link do pedido"). */
function extractUrl_(texto) {
  var m = String(texto || '').match(/https?:\/\/[^\s)]+/i);
  return m ? m[0] : '';
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================ HISTÓRICO ============================ */

var HIST_SHEET_NAME = 'Historico';

/** Aba de histórico (cria com cabeçalho se não existir). */
function getHistSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HIST_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(HIST_SHEET_NAME);
    sh.appendRow(['Data/Hora', 'ID do caso', 'ID venda', 'Usuário', 'Ação', 'Detalhe']);
  }
  return sh;
}

/** Registra um evento no histórico. Nunca deixa um erro aqui quebrar a operação principal. */
function logHist_(caseRow, idVenda, usuario, acao, detalhe) {
  try {
    getHistSheet_().appendRow([new Date(), caseRow, idVenda || '', usuario || '—', acao || '', detalhe || '']);
  } catch (e) { /* silencioso */ }
}

/** Retorna os eventos de histórico de um caso (mais recentes primeiro). */
function histFor_(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HIST_SHEET_NAME);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][1]) !== String(rowIndex)) continue;
    var dt = values[r][0];
    out.push({
      quando: (dt instanceof Date) ? Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(dt),
      usuario: String(values[r][3] || '—'),
      acao: String(values[r][4] || ''),
      detalhe: String(values[r][5] || ''),
    });
  }
  return out.reverse();
}

/* ============================ LEITURA (GET) ============================ */

function doGet(e) {
  try {
    // GET de histórico: ?action=historico&rowIndex=123
    if (e && e.parameter && e.parameter.action === 'historico') {
      return jsonOut_({ ok: true, eventos: histFor_(e.parameter.rowIndex) });
    }
    var sh = getSheet_();
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return jsonOut_({ ok: true, rows: [] });

    var header = values[0];
    var col = buildColMap_(header);
    var get = function (row, key) {
      var i = col[key];
      return (i == null) ? '' : row[i];
    };

    var rows = [];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var idVenda = get(row, 'idVenda');
      var nomeCard = get(row, 'nomeCard');
      // pula linhas totalmente vazias
      if (String(idVenda).trim() === '' && String(nomeCard).trim() === '') continue;

      var descricao = String(get(row, 'descricao') || '');
      rows.push({
        rowIndex:      r + 1, // linha real na planilha (1-based, +1 do cabeçalho)
        data:          fmtDate_(get(row, 'data')),
        auditoria:     parseBool_(get(row, 'auditoria')),
        idVenda:       String(idVenda || '').trim(),
        nomeCard:      String(nomeCard || '').trim(),
        descricao:     descricao,
        quemCadastrou: String(get(row, 'quemCadastrou') || '').trim(),
        culpaDe:       String(get(row, 'culpaDe') || '').trim(),
        setor:         String(get(row, 'setor') || '').trim(),
        responsavel:   String(get(row, 'responsavel') || '').trim(),
        empresa:       String(get(row, 'empresa') || '').trim(),
        tipoProblema:  String(get(row, 'tipoProblema') || '').trim(),
        subproblema:   String(get(row, 'subproblema') || '').trim(),
        qtd:           parseNumber_(get(row, 'qtd')),
        custo:         parseNumber_(get(row, 'custo')),
        tipoProduto:   String(get(row, 'tipoProduto') || '').trim(),
        queFim:        String(get(row, 'queFim') || '').trim(),
        tipoResolucao: String(get(row, 'tipoResolucao') || '').trim(),
        status:        String(get(row, 'status') || '').trim(),
        linkPedido:    extractUrl_(descricao),
      });
    }
    return jsonOut_({ ok: true, version: 'hist-2026-07', aba: sh.getName(), rows: rows });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ============================ ESCRITA (POST) ============================ */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'criar') return criarCaso_(body.fields || {}, body.usuario);
    if (action === 'audit') return auditarCaso_(body.rowIndex, body.fields || {}, body.usuario);
    if (action === 'setStatus') return setStatus_(body.rowIndex, body.status, body.usuario);

    return jsonOut_({ ok: false, error: 'Ação desconhecida: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Grava valor em uma coluna mapeada (ignora se a coluna não existir/veio vazia). */
function setCell_(sh, rowIndex1, col, key, value) {
  var i = col[key];
  if (i == null) return;
  if (value === undefined || value === null || value === '') return;
  sh.getRange(rowIndex1, i + 1).setValue(value);
}

/** action:criar — adiciona uma nova linha de erro no fim da planilha. */
function criarCaso_(f, usuario) {
  var sh = getSheet_();
  var header = sh.getDataRange().getValues()[0];
  var col = buildColMap_(header);

  var novaLinha = sh.getLastRow() + 1;

  // data de hoje na coluna de data
  var hoje = new Date();
  setCell_(sh, novaLinha, col, 'data', fmtDate_(hoje));

  setCell_(sh, novaLinha, col, 'auditoria',     f.auditoria ? 'TRUE' : 'FALSE');
  setCell_(sh, novaLinha, col, 'status',        f.status || (f.auditoria ? 'resolvido' : 'novo'));
  setCell_(sh, novaLinha, col, 'idVenda',       f.idVenda);
  setCell_(sh, novaLinha, col, 'nomeCard',      f.nomeCard);
  setCell_(sh, novaLinha, col, 'descricao',     montarDescricao_(f));
  setCell_(sh, novaLinha, col, 'quemCadastrou', f.quemCadastrou);
  setCell_(sh, novaLinha, col, 'culpaDe',       f.culpaDe);
  setCell_(sh, novaLinha, col, 'setor',         f.setor);
  setCell_(sh, novaLinha, col, 'responsavel',   f.responsavel);
  setCell_(sh, novaLinha, col, 'empresa',       f.empresa);
  setCell_(sh, novaLinha, col, 'tipoProblema',  f.tipoProblema);
  setCell_(sh, novaLinha, col, 'subproblema',   f.subproblema);
  setCell_(sh, novaLinha, col, 'qtd',           f.qtd);
  setCell_(sh, novaLinha, col, 'custo',         f.custo);
  setCell_(sh, novaLinha, col, 'tipoProduto',   f.tipoProduto);
  setCell_(sh, novaLinha, col, 'queFim',        f.queFim);
  setCell_(sh, novaLinha, col, 'tipoResolucao', f.tipoResolucao);

  logHist_(novaLinha, f.idVenda, usuario || f.quemCadastrou, 'Caso registrado',
    f.auditoria ? 'já auditado (' + (f.status || 'resolvido') + ')' : 'pendente de auditoria');

  return jsonOut_({ ok: true, rowIndex: novaLinha });
}

/** Se veio link separado, garante que ele apareça na descrição (o GET lê o link de lá). */
function montarDescricao_(f) {
  var desc = String(f.descricao || '').trim();
  var link = String(f.linkPedido || '').trim();
  if (link && desc.indexOf(link) === -1) desc = link + '\n' + desc;
  return desc;
}

/** action:audit — atualiza os campos de auditoria de uma linha existente. */
function auditarCaso_(rowIndex, f, usuario) {
  if (!rowIndex) return jsonOut_({ ok: false, error: 'rowIndex ausente' });
  var sh = getSheet_();
  var header = sh.getDataRange().getValues()[0];
  var col = buildColMap_(header);

  setCell_(sh, rowIndex, col, 'auditoria',     'TRUE');
  setCell_(sh, rowIndex, col, 'status',        f.status || 'resolvido');
  setCell_(sh, rowIndex, col, 'culpaDe',       f.culpaDe);
  setCell_(sh, rowIndex, col, 'setor',         f.setor);
  setCell_(sh, rowIndex, col, 'responsavel',   f.responsavel);
  setCell_(sh, rowIndex, col, 'empresa',       f.empresa);
  setCell_(sh, rowIndex, col, 'tipoProduto',   f.tipoProduto);
  setCell_(sh, rowIndex, col, 'tipoProblema',  f.tipoProblema);
  setCell_(sh, rowIndex, col, 'subproblema',   f.subproblema);
  setCell_(sh, rowIndex, col, 'qtd',           f.qtd);
  setCell_(sh, rowIndex, col, 'custo',         f.custo);
  setCell_(sh, rowIndex, col, 'queFim',        f.queFim);
  setCell_(sh, rowIndex, col, 'tipoResolucao', f.tipoResolucao);

  logHist_(rowIndex, f.idVenda, usuario, 'Auditoria salva',
    [f.setor, f.tipoResolucao, (f.custo ? 'R$ ' + f.custo : '')].filter(String).join(' · '));

  return jsonOut_({ ok: true, rowIndex: rowIndex });
}

/** action:setStatus — muda o status de workflow de uma linha (e sincroniza a auditoria). */
function setStatus_(rowIndex, status, usuario) {
  if (!rowIndex || !status) return jsonOut_({ ok: false, error: 'rowIndex/status ausente' });
  var sh = getSheet_();
  var col = buildColMap_(sh.getDataRange().getValues()[0]);
  if (col.status == null) return jsonOut_({ ok: false, error: 'Coluna "Status" não existe na planilha. Adicione um cabeçalho "Status".' });
  var idVenda = (col.idVenda != null) ? sh.getRange(rowIndex, col.idVenda + 1).getValue() : '';
  setCell_(sh, rowIndex, col, 'status', status);
  // "Resolvido" conta como auditado; os demais estados, não.
  setCell_(sh, rowIndex, col, 'auditoria', status === 'resolvido' ? 'TRUE' : 'FALSE');
  logHist_(rowIndex, idVenda, usuario, 'Status alterado', '→ ' + status);
  return jsonOut_({ ok: true, rowIndex: rowIndex, status: status });
}

/**
 * MIGRAÇÃO (rode UMA vez no editor após criar a coluna "Status"): preenche o status
 * das linhas antigas a partir de "Auditoria realizada?" — auditado→resolvido, senão→novo.
 * Não sobrescreve linhas que já tenham status.
 */
function migrarStatus() {
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var col = buildColMap_(values[0]);
  if (col.status == null) { Logger.log('Crie a coluna "Status" antes de migrar.'); return; }
  var n = 0;
  for (var r = 1; r < values.length; r++) {
    var atual = String(values[r][col.status] || '').trim();
    var idv = String(values[r][col.idVenda] || '').trim();
    if (atual !== '' || idv === '') continue; // já tem status ou linha vazia
    var status = parseBool_(values[r][col.auditoria]) ? 'resolvido' : 'novo';
    sh.getRange(r + 1, col.status + 1).setValue(status);
    n++;
  }
  Logger.log('Migração concluída: ' + n + ' linha(s) preenchida(s).');
}

/* ============================ DIAGNÓSTICO ============================ */
/**
 * Rode esta função uma vez no editor (menu ▶ Executar → verColunas) e veja em
 * "Registros de execução" como cada coluna da sua planilha foi mapeada.
 * Serve para conferir se o mapeamento está correto ANTES de confiar no painel.
 */
function verColunas() {
  var sh = getSheet_();
  Logger.log('Aba lida: "' + sh.getName() + '" (gid ' + sh.getSheetId() + ') — ' + sh.getLastColumn() + ' colunas, ' + sh.getLastRow() + ' linhas');
  var header = sh.getDataRange().getValues()[0];
  var col = buildColMap_(header);
  Logger.log('Cabeçalhos da planilha:');
  header.forEach(function (h, i) { Logger.log('  [' + i + '] ' + h); });
  Logger.log('---');
  Logger.log('Mapeamento (chave do painel → coluna):');
  Object.keys(COLUNAS).forEach(function (k) {
    var i = col[k];
    Logger.log('  ' + k + ' → ' + (i == null ? '(NÃO ENCONTRADA)' : '[' + i + '] ' + header[i]));
  });
}
