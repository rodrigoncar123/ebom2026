/**
 * EBOM — API de inscrições (Google Apps Script + Planilha Google)
 * Cole este código em: Planilha > Extensões > Apps Script.
 * Depois: Implantar > Nova implantação > Aplicativo da Web
 *   Executar como: Eu  |  Quem tem acesso: Qualquer pessoa
 */
const CONFIG = {
  ADMIN_KEY: 'TROQUE-ESTA-SENHA',        // senha da área de relatório (admin.html)
  SALT: 'troque-este-texto-secreto-123', // segredo usado no código de verificação
  ANO: 2026,
  ABA: 'Inscricoes'
};
const COLS = ['Protocolo', 'CPF', 'Nome completo', 'Data de nascimento', 'Cidade', 'Inscrito em', 'Código de verificação'];
const TEXTO_RE = /^[\p{L} .'’-]+$/u;

function doGet() { return saida_({ ok: true, msg: 'API EBOM ativa' }); }

function doPost(e) {
  let out;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const req = JSON.parse(e.postData.contents);
    out = rota_(req);
  } catch (err) {
    out = { ok: false, erro: 'Erro no servidor: ' + err };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
  return saida_(out);
}

function saida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function rota_(req) {
  const cpf = String(req.cpf || '').replace(/\D/g, '');
  switch (req.acao) {
    case 'existe':
      return { ok: true, existe: !!achar_(cpf) };
    case 'inscrever':
      return inscrever_(req, cpf);
    case 'listar':
      if (req.chave !== CONFIG.ADMIN_KEY) return { ok: false, erro: 'Senha incorreta.' };
      return { ok: true, registros: todos_() };
    default:
      return { ok: false, erro: 'Ação desconhecida.' };
  }
}

function inscrever_(req, cpf) {
  const nome = String(req.nome || '').replace(/\s+/g, ' ').trim();
  const cidade = String(req.cidade || '').replace(/\s+/g, ' ').trim();
  const nasc = String(req.nascimento || '');
  if (!cpfValido_(cpf)) return { ok: false, erro: 'CPF inválido.' };
  if (nome.length < 5 || nome.indexOf(' ') < 0 || !TEXTO_RE.test(nome)) return { ok: false, erro: 'Nome inválido.' };
  if (cidade.length < 2 || !TEXTO_RE.test(cidade)) return { ok: false, erro: 'Cidade inválida.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(nasc) || nasc > new Date().toISOString().slice(0, 10) || nasc < '1900-01-01') return { ok: false, erro: 'Data de nascimento inválida.' };
  if (!req.consentimento) return { ok: false, erro: 'É necessário autorizar o uso dos dados.' };

  const existente = achar_(cpf);
  if (existente) {
    // só devolve os dados (comprovante) se a data de nascimento conferir
    if (existente.nascimento === nasc) return { ok: true, jaCadastrado: true, registro: existente };
    return { ok: false, codigo: 'CPF_EXISTE', erro: 'Este CPF já está cadastrado. Para abrir o comprovante, informe a data de nascimento usada na inscrição.' };
  }

  const sh = aba_();
  const linha = sh.getLastRow() + 1;
  const protocolo = 'EBOM-' + CONFIG.ANO + '-' + ('0000' + (linha - 1)).slice(-4);
  const criadoEm = new Date().toISOString();
  const codigo = codigo_(protocolo, cpf, criadoEm);
  const faixa = sh.getRange(linha, 1, 1, COLS.length);
  faixa.setNumberFormat('@');
  faixa.setValues([[protocolo, cpf, nome, nasc, cidade, criadoEm, codigo]]);
  return { ok: true, jaCadastrado: false, registro: { protocolo, cpf, nome, nascimento: nasc, cidade, criadoEm, codigo } };
}

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.ABA);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.ABA);
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold').setBackground('#12306b').setFontColor('#ffffff');
    sh.setColumnWidths(1, COLS.length, 160);
  }
  return sh;
}

function todos_() {
  const sh = aba_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, COLS.length).getValues().map(function (v) {
    return { protocolo: String(v[0]), cpf: String(v[1]), nome: String(v[2]), nascimento: String(v[3]).slice(0, 10), cidade: String(v[4]), criadoEm: String(v[5]), codigo: String(v[6]) };
  });
}

function achar_(cpf) {
  if (!cpf) return null;
  const l = todos_();
  for (let i = 0; i < l.length; i++) if (l[i].cpf === cpf) return l[i];
  return null;
}

function codigo_(protocolo, cpf, iso) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, [protocolo, cpf, iso, CONFIG.SALT].join('|'), Utilities.Charset.UTF_8);
  const hex = raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').toUpperCase();
  return hex.slice(0, 5) + '-' + hex.slice(5, 10);
}

function cpfValido_(c) {
  if (c.length !== 11 || /^(\d)\1+$/.test(c)) return false;
  for (let t = 9; t < 11; t++) {
    let s = 0;
    for (let i = 0; i < t; i++) s += Number(c[i]) * (t + 1 - i);
    if (((s * 10) % 11) % 10 !== Number(c[t])) return false;
  }
  return true;
}
