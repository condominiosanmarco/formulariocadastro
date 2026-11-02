/**
 * Cadastro de moradores · Condomínio San Marco
 *
 * API usada pela página https://condominiosanmarco.github.io/formulariocadastro/
 * Os dados ficam na aba "Cadastro" desta planilha, um morador por linha:
 *   Bloco | Unidade | Nome | Telefone | E-mail | Atualizado em
 * sempre em ordem de bloco e apartamento (101, 102, ... 201, 202 ...).
 *
 * Publicar: Implantar > Nova implantação > App da Web
 *   - Executar como: Eu
 *   - Quem pode acessar: Qualquer pessoa
 * Copie o link /exec e cole na variável API do index.html.
 * Mudou este arquivo? Implantar > Gerenciar implantações > editar > Nova versão (o link continua o mesmo).
 */

// Mesmo desenho do index.html: andares por bloco e apartamentos por andar (colunas)
var BLOCOS = [
  { id: '1', andares: 7, colunas: 6 },
  { id: '2', andares: 7, colunas: 7 }
];
var ABA = 'Cadastro';
var CABECALHO = ['Bloco', 'Unidade', 'Nome', 'Telefone', 'E-mail', 'Atualizado em'];
var MAX_MORADORES = 15;

function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'status') return _json({ ok: true, contagem: _contagem() });
  return ContentService.createTextOutput('Cadastro de moradores · Condomínio San Marco');
}

/** Recebe { acao: "abrir" | "salvar", bloco, unidade, telefone, moradores } da página. */
function doPost(e) {
  var d;
  try {
    d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return _json({ ok: false, erro: 'Pedido inválido.' });
  }
  if (d.acao === 'abrir') return _json(_abrir(d));
  if (d.acao === 'salvar') return _json(_salvar(d));
  return _json({ ok: false, erro: 'Ação desconhecida.' });
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _aba() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName(ABA);
  if (!aba) {
    aba = planilha.insertSheet(ABA, 0);
    aba.getRange('A:F').setNumberFormat('@');
    aba.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

/** Todas as linhas de moradores, como texto: [bloco, unidade, nome, telefone, email, atualizado]. */
function _linhas() {
  var aba = _aba();
  var n = aba.getLastRow();
  if (n < 2) return [];
  return aba.getRange(2, 1, n - 1, CABECALHO.length).getDisplayValues();
}

function _existe(bloco, unidade) {
  var b = BLOCOS.filter(function (x) { return x.id === bloco; })[0];
  if (!b || !/^\d{3,4}$/.test(unidade)) return false;
  var u = parseInt(unidade, 10);
  var andar = Math.floor(u / 100);
  var coluna = u % 100;
  return andar >= 1 && andar <= b.andares && coluna >= 1 && coluna <= b.colunas;
}

/** { "1-101": 2, ... }: quantos moradores cada apartamento tem. Não devolve nomes nem telefones. */
function _contagem() {
  var c = {};
  _linhas().forEach(function (l) {
    var k = l[0] + '-' + l[1];
    c[k] = (c[k] || 0) + 1;
  });
  return c;
}

function _digitos(t) {
  return String(t || '').replace(/\D/g, '');
}

/** O telefone informado bate (últimos 8 dígitos) com o de algum morador já cadastrado no apartamento? */
function _confere(linhasDoApto, telefone) {
  var t = _digitos(telefone);
  if (t.length < 8) return false;
  var fim = t.slice(-8);
  return linhasDoApto.some(function (l) {
    var d = _digitos(l[3]);
    return d.length >= 8 && d.slice(-8) === fim;
  });
}

function _limpar(texto, max) {
  return String(texto || '').replace(/[\r\n\t]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}

function _fone(texto) {
  var d = _digitos(texto);
  if (d.length === 13 && d.indexOf('55') === 0) d = d.slice(2);
  if (d.length === 11) return '(' + d.slice(0, 2) + ') ' + d.slice(2, 7) + '-' + d.slice(7);
  if (d.length === 10) return '(' + d.slice(0, 2) + ') ' + d.slice(2, 6) + '-' + d.slice(6);
  return _limpar(texto, 20);
}

function _abrir(d) {
  var bloco = String(d.bloco || '');
  var unidade = String(d.unidade || '');
  if (!_existe(bloco, unidade)) return { ok: false, erro: 'Apartamento não encontrado.' };
  var doApto = _linhas().filter(function (l) { return l[0] === bloco && l[1] === unidade; });
  if (!doApto.length) return { ok: true, moradores: [] };
  if (!_confere(doApto, d.telefone)) {
    return { ok: false, erro: 'Esse telefone não confere com nenhum morador cadastrado neste apartamento.' };
  }
  return {
    ok: true,
    moradores: doApto.map(function (l) { return { nome: l[2], telefone: l[3], email: l[4] }; })
  };
}

function _salvar(d) {
  var bloco = String(d.bloco || '');
  var unidade = String(d.unidade || '');
  if (!_existe(bloco, unidade)) return { ok: false, erro: 'Apartamento não encontrado.' };

  var moradores = (d.moradores || []).map(function (m) {
    return { nome: _limpar(m.nome, 80), telefone: _fone(m.telefone), email: _limpar(m.email, 80).toLowerCase() };
  }).filter(function (m) { return m.nome; });
  if (!moradores.length) return { ok: false, erro: 'Preencha o nome de pelo menos um morador.' };
  if (moradores.length > MAX_MORADORES) return { ok: false, erro: 'No máximo ' + MAX_MORADORES + ' moradores por apartamento.' };
  if (_digitos(moradores[0].telefone).length < 10) return { ok: false, erro: 'Informe o telefone com DDD do 1º morador.' };
  for (var i = 0; i < moradores.length; i++) {
    var m = moradores[i];
    if (m.telefone && _digitos(m.telefone).length < 10) return { ok: false, erro: 'O telefone de ' + m.nome + ' está incompleto (use o DDD).' };
    if (m.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(m.email)) return { ok: false, erro: 'O e-mail de ' + m.nome + ' parece errado.' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var todas = _linhas();
    var doApto = todas.filter(function (l) { return l[0] === bloco && l[1] === unidade; });
    if (doApto.length && !_confere(doApto, d.telefone)) {
      return { ok: false, erro: 'Para alterar este apartamento, confirme o telefone de um morador já cadastrado.' };
    }
    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
    var novas = moradores.map(function (m) { return [bloco, unidade, m.nome, m.telefone, m.email, agora]; });
    // Ordem da síndica: bloco, depois apartamento; dentro do apartamento fica a ordem digitada (sort estável)
    var final = todas.filter(function (l) { return !(l[0] === bloco && l[1] === unidade); })
      .concat(novas)
      .sort(function (a, b) { return (Number(a[0]) - Number(b[0])) || (Number(a[1]) - Number(b[1])); });

    var aba = _aba();
    var n = aba.getLastRow();
    if (n > 1) aba.getRange(2, 1, n - 1, CABECALHO.length).clearContent();
    aba.getRange(2, 1, final.length, CABECALHO.length).setNumberFormat('@').setValues(final);
    return { ok: true, contagem: _contagem() };
  } finally {
    lock.releaseLock();
  }
}
