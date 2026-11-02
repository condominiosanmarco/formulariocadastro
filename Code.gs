/**
 * Cadastro de moradores · Condomínio San Marco
 *
 * API usada pela página https://condominiosanmarco.github.io/formulariocadastro/
 * Os dados ficam na aba "Cadastro" desta planilha, um morador por linha:
 *   Bloco | Unidade | Nome | Telefone | E-mail | Atualizado em
 * sempre em ordem de bloco e apartamento (101, 102, ... 201, 202 ...).
 *
 * Quem vê o quê:
 *   - Morador: a lista pública mostra só se o apartamento está cadastrado. Para ver ou corrigir o próprio
 *     apartamento, confirma o telefone de um morador já cadastrado nele.
 *   - Síndica: entra com a senha e vê/corrige todos os apartamentos e baixa a lista (PDF, Excel, CSV).
 *     A senha NÃO fica neste arquivo: em Configurações do projeto > Propriedades do script, a propriedade
 *     SENHA_SINDICA_SHA256 guarda o SHA-256 (hex, minúsculo) da senha.
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
  if (e && e.parameter && e.parameter.api === 'status') return _json({ ok: true, cadastrados: _cadastrados() });
  return ContentService.createTextOutput('Cadastro de moradores · Condomínio San Marco');
}

/** Recebe { acao: "abrir" | "salvar" | "sindica", bloco, unidade, telefone, senha, moradores } da página. */
function doPost(e) {
  var d;
  try {
    d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return _json({ ok: false, erro: 'Pedido inválido.' });
  }
  if (d.acao === 'abrir') return _json(_abrir(d));
  if (d.acao === 'salvar') return _json(_salvar(d));
  if (d.acao === 'sindica') return _json(_sindica(d));
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
  // Ignora linhas vazias (sobra de quem apaga uma linha à mão): o próximo salvamento reorganiza a planilha
  return aba.getRange(2, 1, n - 1, CABECALHO.length).getDisplayValues()
    .filter(function (l) { return l[0] || l[1] || l[2]; });
}

function _existe(bloco, unidade) {
  var b = BLOCOS.filter(function (x) { return x.id === bloco; })[0];
  if (!b || !/^\d{3,4}$/.test(unidade)) return false;
  var u = parseInt(unidade, 10);
  var andar = Math.floor(u / 100);
  var coluna = u % 100;
  return andar >= 1 && andar <= b.andares && coluna >= 1 && coluna <= b.colunas;
}

/** { "1-101": true, ... }: só diz quais apartamentos já têm cadastro, sem nomes nem quantidades. */
function _cadastrados() {
  var c = {};
  _linhas().forEach(function (l) { c[l[0] + '-' + l[1]] = true; });
  return c;
}

function _totalApartamentos() {
  return BLOCOS.reduce(function (s, b) { return s + b.andares * b.colunas; }, 0);
}

function _digitos(t) {
  return String(t || '').replace(/\D/g, '');
}

function _sha256(texto) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(texto), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); })
    .join('');
}

/** A senha da síndica confere? Errar atrasa a resposta, para dificultar quem tenta adivinhar. */
function _ehSindica(senha) {
  if (!senha) return false;
  var hash = PropertiesService.getScriptProperties().getProperty('SENHA_SINDICA_SHA256');
  if (hash && _sha256(String(senha).trim()) === hash.trim().toLowerCase()) return true;
  Utilities.sleep(1500);
  return false;
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

/** Síndica: com a senha certa, devolve a lista completa (para a tabela e para baixar em PDF, Excel ou CSV). */
function _sindica(d) {
  if (!_ehSindica(d.senha)) return { ok: false, erro: 'Senha incorreta.' };
  return { ok: true, cabecalho: CABECALHO, linhas: _linhas(), totalApartamentos: _totalApartamentos() };
}

function _abrir(d) {
  var bloco = String(d.bloco || '');
  var unidade = String(d.unidade || '');
  if (!_existe(bloco, unidade)) return { ok: false, erro: 'Apartamento não encontrado.' };
  var doApto = _linhas().filter(function (l) { return l[0] === bloco && l[1] === unidade; });
  if (!doApto.length) return { ok: true, moradores: [] };
  if (!_confere(doApto, d.telefone) && !_ehSindica(d.senha)) {
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

  var sindica = d.senha ? _ehSindica(d.senha) : false;
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var todas = _linhas();
    var doApto = todas.filter(function (l) { return l[0] === bloco && l[1] === unidade; });
    if (doApto.length && !sindica && !_confere(doApto, d.telefone)) {
      return { ok: false, erro: 'Para alterar este apartamento, confirme o telefone de um morador já cadastrado.' };
    }
    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
    var novas = moradores.map(function (m) { return [bloco, unidade, m.nome, m.telefone, m.email, agora]; });
    // Ordem da síndica: bloco, depois apartamento; dentro do apartamento fica a ordem digitada (sort estável)
    var final = todas.filter(function (l) { return !(l[0] === bloco && l[1] === unidade); })
      .concat(novas)
      .sort(function (a, b) { return (Number(a[0]) - Number(b[0])) || (Number(a[1]) - Number(b[1])); });

    // Uma escrita só; limpa apenas as linhas que sobraram quando a lista diminui
    var aba = _aba();
    var n = aba.getLastRow();
    aba.getRange(2, 1, final.length, CABECALHO.length).setValues(final);
    if (n > final.length + 1) aba.getRange(final.length + 2, 1, n - final.length - 1, CABECALHO.length).clearContent();
    SpreadsheetApp.flush(); // grava de verdade antes de soltar a trava, para o próximo envio ler a planilha atualizada
    var cadastrados = {};
    final.forEach(function (l) { cadastrados[l[0] + '-' + l[1]] = true; });
    return { ok: true, cadastrados: cadastrados };
  } finally {
    lock.releaseLock();
  }
}
