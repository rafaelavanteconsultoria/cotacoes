/**
 * COTAÇÕES - servidor na Planilha Google (Apps Script)
 * Guarda clientes, itens, cotações, configurações e numeração.
 * Cada pessoa entra com e-mail e senha cadastrados na aba "Usuarios".
 */

var ABAS = {
  clients: 'Clientes', catalog: 'Itens', history: 'Cotacoes',
  cfg: 'Config', counters: 'Contadores', users: 'Usuarios',
  sessions: 'Sessoes', audit: 'Auditoria'
};
// Segurança (ajuste se quiser)
var SESSAO_OCIOSA_MS = 72 * 3600 * 1000;      // sessão expira após 72 h sem uso
var SESSAO_MAX_MS = 30 * 24 * 3600 * 1000;    // e no máximo 30 dias
var MAX_FALHAS = 5;                           // tentativas erradas de senha
var BLOQUEIO_S = 900;                         // bloqueio de 15 minutos
var MAX_BYTES = 6000000;                      // tamanho máximo de cada envio
var MAX_TEXTO = 4000;                         // tamanho máximo de cada campo de texto
var VERSAO_AVISO = 'aviso-2026-10';               // versão do aviso de privacidade
var COLS = {
  clients: ['nome', 'att', 'end', 'cep', 'cidade', 'uf', 'pais', 'fone', 'doc', 'email', 'last', 'extra'],
  catalog: ['key', 'pn', 'nome', 'preco_R$', 'preco_US$', 'extra'],
  history: ['id', 'numero', 'ano', 'data', 'cliente', 'moeda', 'total', 'json'],
  counters: ['ano', 'ultimo'],
  cfg: ['chave', 'valor'],
  users: ['email', 'nome', 'hash', 'perfil', 'ativo', 'aceite', 'trocar', 'permissoes', 'ultimo_acesso'],
  sessions: ['token_hash', 'email', 'criada', 'ultimo_uso'],
  audit: ['data_hora', 'usuario', 'acao', 'detalhe']
};
var KNOWN = {
  clients: ['nome', 'att', 'end', 'cep', 'cidade', 'uf', 'pais', 'fone', 'doc', 'email', 'last'],
  catalog: ['pn', 'nome', 'unit', 'unitUsd']
};

/* ---------- instalação ---------- */
function configurarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(COLS).forEach(function (k) {
    var sh = ss.getSheetByName(ABAS[k]) || ss.insertSheet(ABAS[k]);
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, COLS[k].length).setValues([COLS[k]]).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });
  // Células de texto puro: impede que um valor digitado como =FORMULA(...) seja executado
  ['clients', 'history'].forEach(function (k) { var sh = aba(k); sh.getRange(2, 1, sh.getMaxRows() - 1, COLS[k].length).setNumberFormat('@'); });
  var shI = aba('catalog');
  [1, 2, 3, 6].forEach(function (c) { shI.getRange(2, c, shI.getMaxRows() - 1, 1).setNumberFormat('@'); });
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid() + Utilities.getUuid());
  var padrao = ss.getSheetByName('Planilha1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1) ss.deleteSheet(padrao);
  SpreadsheetApp.getUi().alert('Planilha preparada. Agora use o menu "Cotações" > "Definir ou redefinir o usuário MASTER" para criar o seu acesso. Os demais usuários você cria dentro do aplicativo.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Cotações')
    .addItem('Preparar planilha (1ª vez)', 'configurarPlanilha')
    .addItem('Definir ou redefinir o usuário MASTER', 'menuMaster')
    .addItem('Encerrar todas as sessões (forçar novo login)', 'menuEncerrarSessoes')
    .addToUi();
}

// Só existe UM usuário MASTER (o dono da planilha). Os demais usuários e suas permissões
// são criados e controlados pelo MASTER dentro do aplicativo (Configurações > Usuários e permissões).
function menuMaster() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('E-mail do MASTER', 'Digite o e-mail do MASTER (use um e-mail que você acessa, pois é para onde vai o código de recuperação de senha):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var r2 = ui.prompt('Senha do MASTER', 'Mínimo 8 caracteres, com letras e números:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var r3 = ui.prompt('Nome', 'Nome do MASTER (aparece no app):', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  var msg = salvarMaster(r1.getResponseText(), r2.getResponseText(), r3.getResponseText());
  registrar('(planilha)', 'master_definido', norm(r1.getResponseText()));
  ui.alert(msg);
}

function menuEncerrarSessoes() {
  var sh = aba('sessions'), n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, 4).clearContent();
  registrar('(planilha)', 'sessoes_encerradas', 'todas');
  SpreadsheetApp.getUi().alert('Todas as sessões foram encerradas. Todos precisarão entrar de novo.');
}

function senhaForte(senha) {
  senha = String(senha || '');
  return senha.length >= 8 && /[A-Za-z]/.test(senha) && /[0-9]/.test(senha);
}

function salvarMaster(email, senha, nome) {
  email = norm(email);
  if (!email || email.indexOf('@') < 0) return 'E-mail inválido.';
  if (!senhaForte(senha)) return 'A senha precisa ter pelo menos 8 caracteres, com letras e números.';
  var sh = aba('users'), v = sh.getDataRange().getValues(), hash = novoHash(senha), agora = '';
  // rebaixa qualquer outro master (só pode haver um)
  for (var i = 1; i < v.length; i++) {
    if (upper(v[i][3]) === 'MASTER' && norm(v[i][0]) !== email) { sh.getRange(i + 1, 4).setValue('USUARIO'); revogarSessoes(v[i][0]); }
  }
  var linha = [email, nome || email, hash, 'MASTER', 'SIM', '', '', '', agora];
  for (var j = 1; j < v.length; j++) {
    if (norm(v[j][0]) === email) { linha[5] = v[j][5]; sh.getRange(j + 1, 1, 1, 9).setValues([linha]); revogarSessoes(email); return 'MASTER atualizado (sessões antigas encerradas).'; }
  }
  sh.appendRow(linha);
  return 'MASTER criado.';
}

/* ---------- utilidades ---------- */
function norm(s) { return String(s == null ? '' : s).trim().toLowerCase(); }
function upper(s) { return String(s == null ? '' : s).trim().toUpperCase(); }
function aba(k) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABAS[k]);
  if (!sh) throw new Error('Aba ausente: ' + ABAS[k] + '. Rode "Preparar planilha".');
  return sh;
}
function sha(txt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(txt), Utilities.Charset.UTF_8));
}
function derivar(salt, senha) {
  var pepper = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  var h = pepper + ':' + salt + ':' + senha;
  for (var i = 0; i < 3000; i++) h = sha(h + salt);
  return h;
}
function novoHash(senha) {
  var salt = Utilities.getUuid();
  return 'v2$' + salt + '$' + derivar(salt, senha);
}
function confereSenha(senha, guardado) {
  var p = String(guardado).split('$');
  if (p.length !== 3 || p[0] !== 'v2') return false;
  return derivar(p[1], senha) === p[2];
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function reais(c) { return (c === '' || c == null || isNaN(c)) ? '' : Math.round(c) / 100; }
function cents(v) { return (v === '' || v == null || isNaN(v)) ? undefined : Math.round(Number(v) * 100); }
function finalCents(q) {
  if (q.especial != null) return q.especial;
  return (q.itens || []).reduce(function (a, it) { return a + Math.round((+it.qtd) * (+it.unit)); }, 0);
}
function chaveItem(it) { return it.pn ? 'PN:' + upper(it.pn) : 'NM:' + upper(it.nome); }

/* ---------- conversão linha <-> objeto ---------- */
function limitar(v) { return (typeof v === 'string' && v.length > MAX_TEXTO) ? v.slice(0, MAX_TEXTO) : v; }
function paraLinha(kind, o) {
  var l = paraLinha0(kind, o);
  return kind === 'history' ? l : l.map(limitar);
}
function paraLinha0(kind, o) {
  if (kind === 'clients') {
    var ex = {};
    Object.keys(o).forEach(function (k) { if (KNOWN.clients.indexOf(k) < 0) ex[k] = o[k]; });
    return KNOWN.clients.map(function (k) { return o[k] == null ? '' : o[k]; }).concat([Object.keys(ex).length ? JSON.stringify(ex) : '']);
  }
  if (kind === 'catalog') {
    var ex2 = {};
    Object.keys(o).forEach(function (k) { if (KNOWN.catalog.indexOf(k) < 0) ex2[k] = o[k]; });
    return [chaveItem(o), o.pn || '', o.nome || '', reais(o.unit), reais(o.unitUsd), Object.keys(ex2).length ? JSON.stringify(ex2) : ''];
  }
  if (kind === 'history') {
    return [o.id, o.numero, o.ano, o.data, o.cliente, o.moeda || 'BRL', reais(finalCents(o)), JSON.stringify(o)];
  }
}
function deLinha(kind, r) {
  var o = {};
  if (kind === 'clients') {
    KNOWN.clients.forEach(function (k, i) { if (r[i] !== '' && r[i] != null) o[k] = r[i]; });
    if (o.last != null) o.last = Number(o.last);
    ['cep', 'fone', 'doc'].forEach(function (k) { if (o[k] != null) o[k] = String(o[k]); });
    if (r[11]) { try { var e = JSON.parse(r[11]); Object.keys(e).forEach(function (k) { o[k] = e[k]; }); } catch (x) {} }
    o.nome = String(o.nome || '');
    return o;
  }
  if (kind === 'catalog') {
    o.pn = String(r[1] == null ? '' : r[1]); o.nome = String(r[2] || '');
    var u = cents(r[3]); if (u !== undefined) o.unit = u;
    var d = cents(r[4]); if (d !== undefined) o.unitUsd = d;
    if (r[5]) { try { var e2 = JSON.parse(r[5]); Object.keys(e2).forEach(function (k) { o[k] = e2[k]; }); } catch (x) {} }
    return o;
  }
  if (kind === 'history') { try { return JSON.parse(r[7]); } catch (x) { return null; } }
}
function chaveDe(kind, o) {
  if (kind === 'clients') return upper(o.nome);
  if (kind === 'catalog') return chaveItem(o);
  if (kind === 'history') return String(o.id);
}
function chaveLinha(kind, r) {
  if (kind === 'clients') return upper(r[0]);
  if (kind === 'catalog') return String(r[0]);
  if (kind === 'history') return String(r[0]);
}

/* ---------- leitura/escrita ---------- */
function lerTudo(kind) {
  var sh = aba(kind), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var w = COLS[kind].length;
  var v = sh.getRange(2, 1, n, w).getValues(), res = [];
  for (var i = 0; i < v.length; i++) {
    if (v[i][0] === '' && v[i][1] === '') continue;
    var o = deLinha(kind, v[i]);
    if (o) res.push(o);
  }
  return res;
}
function substituirTudo(kind, lista) {
  var sh = aba(kind), w = COLS[kind].length;
  var ult = sh.getLastRow();
  if (ult > 1) sh.getRange(2, 1, ult - 1, w).clearContent();
  if (!lista.length) return;
  var linhas = lista.map(function (o) { return paraLinha(kind, o); });
  sh.getRange(2, 1, linhas.length, w).setValues(linhas);
}
// junta: clientes e itens recebem só os campos informados; cotações são substituídas
function mesclar(kind, antigo, novo) {
  if (kind === 'history' || !antigo) return novo;
  var o = {};
  Object.keys(antigo).forEach(function (k) { o[k] = antigo[k]; });
  Object.keys(novo).forEach(function (k) { if (novo[k] !== undefined && novo[k] !== null && novo[k] !== '') o[k] = novo[k]; });
  return o;
}
function gravar(kind, lista) {
  if (!lista || !lista.length) return;
  var sh = aba(kind), w = COLS[kind].length, n = sh.getLastRow() - 1;
  var vals = n > 0 ? sh.getRange(2, 1, n, w).getValues() : [];
  var idx = {};
  for (var i = 0; i < vals.length; i++) idx[chaveLinha(kind, vals[i])] = i;
  var novas = [];
  lista.forEach(function (o) {
    var k = chaveDe(kind, o);
    if (!k) return;
    if (idx[k] !== undefined) {
      var ant = deLinha(kind, vals[idx[k]]);
      var m = mesclar(kind, ant, o);
      var lin = paraLinha(kind, m);
      sh.getRange(idx[k] + 2, 1, 1, w).setValues([lin]);
      vals[idx[k]] = lin;
    } else {
      var lin2 = paraLinha(kind, o);
      idx[k] = vals.length + novas.length;
      novas.push(lin2);
      vals.push(lin2);
    }
  });
  if (novas.length) sh.getRange(n + 2, 1, novas.length, w).setValues(novas);
}
function apagar(kind, chave) {
  var sh = aba(kind), n = sh.getLastRow() - 1;
  if (n < 1) return false;
  var v = sh.getRange(2, 1, n, COLS[kind].length).getValues();
  for (var i = 0; i < v.length; i++) {
    if (chaveLinha(kind, v[i]) === (kind === 'clients' ? upper(chave) : String(chave))) { sh.deleteRow(i + 2); return true; }
  }
  return false;
}
function lerContadores() {
  var sh = aba('counters'), n = sh.getLastRow() - 1, c = {};
  if (n < 1) return c;
  sh.getRange(2, 1, n, 2).getValues().forEach(function (r) { if (r[0] !== '') c[String(r[0])] = Number(r[1]) || 0; });
  return c;
}
function gravarContador(ano, valor) {
  var sh = aba('counters'), n = sh.getLastRow() - 1;
  if (n > 0) {
    var v = sh.getRange(2, 1, n, 2).getValues();
    for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(ano)) { sh.getRange(i + 2, 2).setValue(valor); return; }
  }
  sh.appendRow([ano, valor]);
}
function lerSettings() {
  var sh = aba('cfg'), n = sh.getLastRow() - 1;
  if (n < 1) return {};
  var v = sh.getRange(2, 1, n, 2).getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === 'settings') { try { return JSON.parse(v[i][1]); } catch (e) { return {}; } }
  return {};
}
function gravarSettings(s) {
  var sh = aba('cfg'), n = sh.getLastRow() - 1;
  if (n > 0) {
    var v = sh.getRange(2, 1, n, 2).getValues();
    for (var i = 0; i < v.length; i++) if (v[i][0] === 'settings') { sh.getRange(i + 2, 2).setValue(JSON.stringify(s)); return; }
  }
  sh.appendRow(['settings', JSON.stringify(s)]);
}

/* ---------- permissões, auditoria, sessões e login ---------- */
var PERMS = ['cotar', 'historico', 'corrigir', 'cadastros', 'doc', 'apagar', 'exportar', 'lgpd', 'empresa'];
var PRESETS = {
  consulta: { historico: true, cadastros: true },
  vendedor: { cotar: true, historico: true, cadastros: true },
  gerente: { cotar: true, historico: true, corrigir: true, cadastros: true, doc: true, exportar: true, lgpd: true }
};
function permDe(texto, master) {
  var p = {};
  if (master) { PERMS.forEach(function (k) { p[k] = true; }); return p; }
  var j = {};
  try { j = JSON.parse(texto || '{}'); } catch (e) {}
  PERMS.forEach(function (k) { p[k] = j[k] === true; });
  return p;
}
function limparPerm(obj) {
  var p = {};
  PERMS.forEach(function (k) { p[k] = !!(obj && obj[k]); });
  return p;
}
function registrar(usuario, acao, detalhe) {
  try { aba('audit').appendRow([new Date(), usuario, acao, String(detalhe == null ? '' : detalhe).slice(0, 300)]); } catch (e) {}
}
function mascaraDoc(d) {
  var n = String(d || '').replace(/\D/g, '');
  return n.length >= 4 ? '***' + n.slice(-2) : '***';
}
function usuarioPorEmail(email) {
  var v = aba('users').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (norm(v[i][0]) === norm(email)) return { linha: i + 1, d: v[i] };
  return null;
}
function montarUsuario(d) {
  var master = upper(d[3]) === 'MASTER';
  return { email: norm(d[0]), nome: String(d[1] || d[0]), master: master, perm: permDe(d[7], master),
    aceite: String(d[5] || '') === VERSAO_AVISO, trocar: upper(d[6]) === 'SIM' };
}
function revogarSessoes(email) {
  var sh = aba('sessions'), n = sh.getLastRow() - 1;
  if (n < 1) return;
  var v = sh.getRange(2, 1, n, 4).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (norm(v[i][1]) === norm(email)) sh.deleteRow(i + 2);
}
function criarSessao(email) {
  var sh = aba('sessions'), n = sh.getLastRow() - 1, agora = Date.now();
  if (n > 0) { // limpa vencidas
    var v = sh.getRange(2, 1, n, 4).getValues();
    for (var i = v.length - 1; i >= 0; i--) {
      if (agora - Number(v[i][2]) > SESSAO_MAX_MS || agora - Number(v[i][3]) > SESSAO_OCIOSA_MS) sh.deleteRow(i + 2);
    }
  }
  var tok = Utilities.getUuid() + Utilities.getUuid();
  sh.appendRow([sha(tok), norm(email), agora, agora]);
  return tok;
}
function validarSessao(tok) {
  if (!tok) return null;
  var sh = aba('sessions'), n = sh.getLastRow() - 1;
  if (n < 1) return null;
  var h = sha(tok), agora = Date.now();
  var v = sh.getRange(2, 1, n, 4).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) !== h) continue;
    if (agora - Number(v[i][2]) > SESSAO_MAX_MS || agora - Number(v[i][3]) > SESSAO_OCIOSA_MS) { sh.deleteRow(i + 2); return null; }
    var u = usuarioPorEmail(v[i][1]);
    if (!u || upper(u.d[4]) !== 'SIM') return null;
    if (agora - Number(v[i][3]) > 300000) sh.getRange(i + 2, 4).setValue(agora);
    return montarUsuario(u.d);
  }
  return null;
}
function entrar(a) {
  var email = norm(a && a.email), senha = String(a && a.senha || '');
  var cache = CacheService.getScriptCache(), ck = 'falhas_' + sha(email).slice(0, 20);
  var falhas = Number(cache.get(ck) || 0);
  if (falhas >= MAX_FALHAS) { registrar(email, 'login_bloqueado', 'muitas tentativas'); return { erro: 'bloqueado', msg: 'Muitas tentativas erradas. Aguarde 15 minutos e tente de novo.' }; }
  var u = email && senha ? usuarioPorEmail(email) : null;
  if (!u || upper(u.d[4]) !== 'SIM' || !confereSenha(senha, u.d[2])) {
    cache.put(ck, String(falhas + 1), BLOQUEIO_S);
    registrar(email || '(vazio)', 'login_falha', 'tentativa ' + (falhas + 1));
    return { erro: 'auth', msg: 'E-mail ou senha incorretos.' };
  }
  cache.remove(ck);
  var us = montarUsuario(u.d);
  aba('users').getRange(u.linha, 9).setValue(new Date());
  registrar(us.email, 'login_ok', us.master ? 'master' : 'usuário');
  return { ok: true, usuario: us, token: criarSessao(us.email) };
}
function trocarSenhaDe(email, nova, exigirTroca) {
  var u = usuarioPorEmail(email);
  if (!u) return false;
  var sh = aba('users');
  sh.getRange(u.linha, 3).setValue(novoHash(nova));
  sh.getRange(u.linha, 7).setValue(exigirTroca ? 'SIM' : '');
  revogarSessoes(email);
  return true;
}
function senhaTemporaria() {
  var letras = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ', nums = '23456789', s = '';
  var rnd = Utilities.getUuid().replace(/-/g, '');
  for (var i = 0; i < 8; i++) s += letras.charAt(parseInt(rnd.substr(i * 2, 2), 16) % letras.length);
  s += nums.charAt(parseInt(rnd.substr(18, 2), 16) % nums.length) + nums.charAt(parseInt(rnd.substr(20, 2), 16) % nums.length);
  return s;
}

/* ---------- recuperação de senha (código por e-mail) ---------- */
function recuperarPedir(email) {
  email = norm(email);
  var resposta = { ok: true, msg: 'Se este e-mail estiver cadastrado, enviamos um código de 6 números. Ele vale 15 minutos.' };
  if (!email || email.indexOf('@') < 0) return resposta;
  var cache = CacheService.getScriptCache(), k = sha(email).slice(0, 20);
  var pedidos = Number(cache.get('recped_' + k) || 0);
  if (pedidos >= 3) return { ok: false, erro: 'limite', msg: 'Muitos pedidos. Tente de novo em 1 hora.' };
  cache.put('recped_' + k, String(pedidos + 1), 3600);
  var u = usuarioPorEmail(email);
  if (!u || upper(u.d[4]) !== 'SIM') { registrar(email, 'recuperar_pedido_desconhecido', ''); return resposta; }
  var cod = ('000000' + (parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000)).slice(-6);
  cache.put('rec_' + k, JSON.stringify({ h: sha(cod + ':' + email), t: 0 }), 900);
  try {
    MailApp.sendEmail(email, 'Código para redefinir sua senha - Cotações',
      'Seu código para redefinir a senha é: ' + cod + '\n\nEle vale 15 minutos. Se você não pediu, ignore este e-mail e avise o administrador.');
    registrar(email, 'recuperar_codigo_enviado', '');
  } catch (e) {
    registrar(email, 'recuperar_erro_envio', String(e && e.message || e).slice(0, 100));
  }
  return resposta;
}
function recuperarConfirmar(email, codigo, nova) {
  email = norm(email);
  var cache = CacheService.getScriptCache(), k = sha(email).slice(0, 20);
  var raw = cache.get('rec_' + k);
  if (!raw) return { ok: false, erro: 'codigo', msg: 'Código vencido ou inexistente. Peça um novo.' };
  var d = JSON.parse(raw);
  if (d.t >= 5) { cache.remove('rec_' + k); return { ok: false, erro: 'codigo', msg: 'Muitas tentativas. Peça um novo código.' }; }
  if (d.h !== sha(String(codigo).trim() + ':' + email)) {
    d.t++; cache.put('rec_' + k, JSON.stringify(d), 900);
    return { ok: false, erro: 'codigo', msg: 'Código incorreto.' };
  }
  if (!senhaForte(nova)) return { ok: false, erro: 'senha', msg: 'A senha precisa ter pelo menos 8 caracteres, com letras e números.' };
  cache.remove('rec_' + k); cache.remove('falhas_' + k);
  if (!trocarSenhaDe(email, nova, false)) return { ok: false, erro: 'codigo', msg: 'Usuário não encontrado.' };
  registrar(email, 'senha_recuperada', '');
  return { ok: true, msg: 'Senha alterada. Entre com a nova senha.' };
}

/* ---------- entrada ---------- */
function doGet() { return out({ ok: true, msg: 'Servico de cotacoes ativo' }); }

function doPost(e) {
  var lock = LockService.getScriptLock(), travou = false;
  try {
    if (!e || !e.postData || e.postData.length > MAX_BYTES) return out({ ok: false, erro: 'tamanho', msg: 'Envio grande demais.' });
    var req = JSON.parse(e.postData.contents);
    if (req.acao === 'login') {
      var r = entrar(req.auth);
      return out(r.ok ? { ok: true, usuario: r.usuario, token: r.token } : { ok: false, erro: r.erro, msg: r.msg });
    }
    if (req.acao === 'rec_pedir') return out(recuperarPedir((req.dados || {}).email));
    if (req.acao === 'rec_confirmar') { var d0 = req.dados || {}; return out(recuperarConfirmar(d0.email, d0.codigo, d0.nova)); }
    var u = validarSessao(req.token);
    if (!u) return out({ ok: false, erro: 'auth', msg: 'Sessão expirada. Entre novamente.' });
    if (u.trocar && ['senha_trocar', 'logout'].indexOf(req.acao) < 0) return out({ ok: false, erro: 'trocar', msg: 'Você precisa criar uma nova senha.' });
    if (['load', 'aviso', 'logout', 'audit'].indexOf(req.acao) < 0) { lock.waitLock(25000); travou = true; }
    return out(acao(req, u));
  } catch (err) {
    return out({ ok: false, erro: 'erro', msg: 'Erro no servidor.' });
  } finally {
    if (travou) { try { lock.releaseLock(); } catch (x) {} }
  }
}

function negar(msg) { return { ok: false, erro: 'perm', msg: msg || 'Você não tem permissão para isto. Peça ao MASTER.' }; }
function pode(u, k) { return !!(u.master || (u.perm && u.perm[k])); }

function acao(req, u) {
  var p = req.dados || {};
  switch (req.acao) {
    case 'load': {
      var cl = [], cat = [], hist = [];
      if (pode(u, 'cotar') || pode(u, 'cadastros')) { cl = lerTudo('clients'); cat = lerTudo('catalog'); }
      if (pode(u, 'historico') || pode(u, 'corrigir')) hist = lerTudo('history');
      if (!pode(u, 'doc')) cl.forEach(function (c) { if (c.doc) { c.docOculto = mascaraDoc(c.doc); delete c.doc; } });
      return { ok: true, usuario: u, clients: cl, catalog: cat, history: hist, settings: lerSettings(), counters: lerContadores(), agora: Date.now() };
    }
    case 'logout':
      revogarSessoes(u.email); registrar(u.email, 'logout', '');
      return { ok: true };
    case 'aviso': {
      var us = usuarioPorEmail(u.email);
      if (us) { var cel = aba('users').getRange(us.linha, 6); cel.setNumberFormat('@'); cel.setValue(VERSAO_AVISO); SpreadsheetApp.flush(); }
      registrar(u.email, 'aviso_privacidade_aceito', VERSAO_AVISO);
      return { ok: true };
    }
    case 'audit':
      if (!pode(u, 'exportar')) return negar();
      registrar(u.email, String(p.acao || 'evento').slice(0, 40), p.detalhe);
      return { ok: true };
    case 'senha_trocar': {
      var ur = usuarioPorEmail(u.email);
      if (!ur || !confereSenha(String(p.atual || ''), ur.d[2])) return { ok: false, erro: 'senha', msg: 'A senha atual está incorreta.' };
      if (!senhaForte(p.nova)) return { ok: false, erro: 'senha', msg: 'A nova senha precisa ter pelo menos 8 caracteres, com letras e números.' };
      if (String(p.nova) === String(p.atual)) return { ok: false, erro: 'senha', msg: 'A nova senha deve ser diferente da atual.' };
      trocarSenhaDe(u.email, p.nova, false);
      registrar(u.email, 'senha_trocada', '');
      return { ok: true, token: criarSessao(u.email) };
    }
    /* --- gestão de usuários: só o MASTER --- */
    case 'users_list': {
      if (!u.master) return negar('Só o MASTER gerencia usuários.');
      var vu = aba('users').getDataRange().getValues(), lista = [];
      for (var i = 1; i < vu.length; i++) {
        if (!vu[i][0]) continue;
        var m = upper(vu[i][3]) === 'MASTER';
        lista.push({ email: norm(vu[i][0]), nome: String(vu[i][1] || ''), master: m, ativo: upper(vu[i][4]) === 'SIM',
          trocar: upper(vu[i][6]) === 'SIM', perm: permDe(vu[i][7], m), ultimo: vu[i][8] ? new Date(vu[i][8]).getTime() : 0 });
      }
      return { ok: true, users: lista, presets: PRESETS };
    }
    case 'user_save': {
      if (!u.master) return negar('Só o MASTER gerencia usuários.');
      var em = norm(p.email);
      if (!em || em.indexOf('@') < 0) return { ok: false, erro: 'dados', msg: 'E-mail inválido.' };
      var ex = usuarioPorEmail(em), shu = aba('users');
      if (!ex) {
        if (!senhaForte(p.senha)) return { ok: false, erro: 'senha', msg: 'A senha provisória precisa ter 8+ caracteres com letras e números.' };
        shu.appendRow([em, String(p.nome || em).slice(0, 80), novoHash(p.senha), 'USUARIO', p.ativo === false ? 'NAO' : 'SIM', '', 'SIM', JSON.stringify(limparPerm(p.perm)), '']);
        registrar(u.email, 'usuario_criado', em);
        return { ok: true };
      }
      var ehMaster = upper(ex.d[3]) === 'MASTER';
      shu.getRange(ex.linha, 2).setValue(String(p.nome || em).slice(0, 80));
      if (!ehMaster) {
        shu.getRange(ex.linha, 5).setValue(p.ativo === false ? 'NAO' : 'SIM');
        shu.getRange(ex.linha, 8).setValue(JSON.stringify(limparPerm(p.perm)));
        if (p.ativo === false) revogarSessoes(em);
      }
      registrar(u.email, 'usuario_alterado', em + (ehMaster ? ' [master: só nome]' : ''));
      return { ok: true };
    }
    case 'user_reset': {
      if (!u.master) return negar('Só o MASTER gerencia usuários.');
      var er = usuarioPorEmail(p.email);
      if (!er) return { ok: false, msg: 'Usuário não encontrado.' };
      if (upper(er.d[3]) === 'MASTER' && norm(p.email) !== u.email) return negar();
      var tmp = senhaTemporaria();
      trocarSenhaDe(p.email, tmp, true);
      registrar(u.email, 'senha_redefinida_pelo_master', norm(p.email));
      return { ok: true, senha: tmp };
    }
    case 'user_delete': {
      if (!u.master) return negar('Só o MASTER gerencia usuários.');
      var ed = usuarioPorEmail(p.email);
      if (!ed) return { ok: false, msg: 'Usuário não encontrado.' };
      if (upper(ed.d[3]) === 'MASTER') return negar('O MASTER não pode ser apagado.');
      aba('users').deleteRow(ed.linha);
      revogarSessoes(p.email);
      registrar(u.email, 'usuario_apagado', norm(p.email));
      return { ok: true };
    }
    case 'audit_list': {
      if (!u.master) return negar('Só o MASTER vê o registro de acessos.');
      var sa = aba('audit'), na = sa.getLastRow() - 1, lim = Math.min(Number(p.limite) || 100, 300);
      if (na < 1) return { ok: true, itens: [] };
      var ini = Math.max(2, sa.getLastRow() - lim + 1);
      var va = sa.getRange(ini, 1, sa.getLastRow() - ini + 1, 4).getValues().reverse();
      return { ok: true, itens: va.map(function (r) { return { t: new Date(r[0]).getTime(), u: String(r[1]), a: String(r[2]), d: String(r[3]) }; }) };
    }
    /* --- LGPD: titulares --- */
    case 'lgpd_export': {
      if (!pode(u, 'lgpd')) return negar();
      var nm = upper(p.nome);
      var cli = lerTudo('clients').filter(function (c) { return upper(c.nome) === nm; })[0] || null;
      var cots = lerTudo('history').filter(function (h) { return upper(h.cliente) === nm; });
      registrar(u.email, 'lgpd_exportar_titular', 'cliente ' + sha(nm).slice(0, 8) + ', ' + cots.length + ' cotações');
      return { ok: true, cadastro: cli, cotacoes: cots };
    }
    case 'lgpd_apagar': {
      if (!pode(u, 'lgpd')) return negar();
      var nome = upper(p.nome), campos = ['att', 'end', 'cep', 'cidade', 'uf', 'pais', 'fone', 'doc', 'email'];
      var shc = aba('clients'), nc = shc.getLastRow() - 1, achou = false;
      var vc = nc > 0 ? shc.getRange(2, 1, nc, COLS.clients.length).getValues() : [];
      for (var i2 = 0; i2 < vc.length; i2++) {
        if (upper(vc[i2][0]) !== nome) continue;
        achou = true;
        var o = deLinha('clients', vc[i2]);
        if (p.modo === 'tudo') { shc.deleteRow(i2 + 2); }
        else { campos.forEach(function (k) { delete o[k]; }); delete o.atts; delete o.aeronaves; shc.getRange(i2 + 2, 1, 1, COLS.clients.length).setValues([paraLinha('clients', o)]); }
        break;
      }
      var anon = 0;
      if (p.modo === 'tudo') {
        var shh = aba('history'), nh = shh.getLastRow() - 1;
        var vh = nh > 0 ? shh.getRange(2, 1, nh, COLS.history.length).getValues() : [];
        for (var j = 0; j < vh.length; j++) {
          var h = deLinha('history', vh[j]);
          if (h && upper(h.cliente) === nome) { h.cliente = 'CLIENTE REMOVIDO (LGPD)'; h.att = ''; shh.getRange(j + 2, 1, 1, COLS.history.length).setValues([paraLinha('history', h)]); anon++; }
        }
      }
      registrar(u.email, 'lgpd_apagar_' + (p.modo === 'tudo' ? 'tudo' : 'dados'), 'cliente ' + sha(nome).slice(0, 8) + (anon ? ', ' + anon + ' cotações anonimizadas' : ''));
      return { ok: achou, msg: achou ? '' : 'Cliente não encontrado.' };
    }
    /* --- cotações e cadastros --- */
    case 'emit': {
      var q = p.quote;
      if (!q || !q.id) return { ok: false, erro: 'dados', msg: 'Cotação inválida.' };
      if (p.novoNumero ? !pode(u, 'cotar') : !pode(u, 'corrigir')) return negar(p.novoNumero ? 'Você não tem permissão para criar cotações.' : 'Você não tem permissão para corrigir cotações.');
      if (p.client) delete p.client.docOculto;
      var ano = String(q.ano || String(q.data).slice(0, 4));
      if (p.novoNumero) {
        var c = lerContadores();
        var prox = (c[ano] || 0) + 1;
        gravarContador(ano, prox);
        q.numero = prox; q.ano = Number(ano);
      }
      q.criadaEm = q.criadaEm || Date.now();
      q.porUsuario = q.porUsuario || u.email;
      gravar('history', [q]);
      if (p.client) gravar('clients', [p.client]);
      if (p.items && p.items.length) gravar('catalog', p.items);
      registrar(u.email, p.novoNumero ? 'cotacao_emitida' : 'cotacao_corrigida', q.numero + '/' + q.ano);
      return { ok: true, numero: q.numero, ano: q.ano };
    }
    case 'del': {
      if (!pode(u, 'apagar')) return negar('Você não tem permissão para apagar.');
      var okDel = apagar(p.kind, p.chave);
      registrar(u.email, 'apagar_' + p.kind, p.kind === 'clients' ? 'cliente ' + sha(upper(p.chave)).slice(0, 8) : String(p.chave).slice(0, 40));
      return { ok: okDel };
    }
    case 'settings': {
      if (!pode(u, 'empresa')) return negar('Você não tem permissão para alterar os dados da empresa.');
      gravarSettings(p.settings || {});
      registrar(u.email, 'configuracoes_alteradas', '');
      return { ok: true };
    }
    case 'counter': {
      if (!pode(u, 'empresa')) return negar('Você não tem permissão para alterar a numeração.');
      gravarContador(String(p.ano), Number(p.valor) || 0);
      registrar(u.email, 'numeracao_alterada', p.ano + ' = ' + p.valor);
      return { ok: true };
    }
    case 'bulk': {
      if (!u.master) return negar('Só o MASTER faz a carga inicial.');
      substituirTudo('clients', p.clients || []);
      substituirTudo('catalog', p.catalog || []);
      substituirTudo('history', p.history || []);
      gravarSettings(p.settings || {});
      Object.keys(p.counters || {}).forEach(function (a) { gravarContador(a, p.counters[a]); });
      registrar(u.email, 'carga_inicial', (p.clients || []).length + ' clientes, ' + (p.catalog || []).length + ' itens, ' + (p.history || []).length + ' cotações');
      return { ok: true, clientes: (p.clients || []).length, itens: (p.catalog || []).length, cotacoes: (p.history || []).length };
    }
    default:
      return { ok: false, erro: 'acao', msg: 'Ação desconhecida.' };
  }
}
