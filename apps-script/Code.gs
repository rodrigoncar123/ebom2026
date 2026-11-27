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
  ABA: 'Inscricoes',
  ENCERRA_EM: '2026-11-27T15:30:00-03:00', // novas inscrições bloqueadas a partir deste horário (Brasília)
  // janelas do registro de presença por QR Code (horário de Brasília)
  DIAS: {
    '27': { ini: '2026-11-27T18:00:00-03:00', fim: '2026-11-27T23:00:00-03:00', col: 8 },
    '28': { ini: '2026-11-28T18:00:00-03:00', fim: '2026-11-28T23:00:00-03:00', col: 9 }
  },
  CERT_LIBERA_EM: '2026-11-28T23:00:00-03:00' // certificado liberado a partir daqui
};
const COLS_PRES = ['Presença 27/11', 'Presença 28/11'];
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
    case 'comprovante':
      return comprovante_(cpf, String(req.nascimento || ''));
    case 'presenca':
      return presenca_(req);
    case 'certificado':
      return certificado_(req);
    case 'listar':
      if (req.chave !== CONFIG.ADMIN_KEY) return { ok: false, erro: 'Senha incorreta.' };
      return { ok: true, registros: todos_() };
    default:
      return { ok: false, erro: 'Ação desconhecida.' };
  }
}

// Abre o comprovante só com CPF + data de nascimento (usado após o encerramento)
function comprovante_(cpf, nasc) {
  if (!cpfValido_(cpf)) return { ok: false, erro: 'CPF inválido.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(nasc)) return { ok: false, erro: 'Data de nascimento inválida.' };
  const r = achar_(cpf);
  if (!r) return { ok: false, codigo: 'NAO_ENCONTRADO', erro: 'Não encontramos inscrição para este CPF. As inscrições foram encerradas.' };
  if (r.nascimento !== nasc) return { ok: false, codigo: 'NASC_DIVERGE', erro: 'Data de nascimento não confere com a inscrição.' };
  return { ok: true, jaCadastrado: true, registro: r };
}

function inscrever_(req, cpf) {
  const nome = titulo_(String(req.nome || '').replace(/\s+/g, ' ').trim());
  const cidade = cidade_(String(req.cidade || '').replace(/\s+/g, ' ').trim());
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

  if (new Date().getTime() >= new Date(CONFIG.ENCERRA_EM).getTime()) return { ok: false, codigo: 'ENCERRADO', erro: 'Inscrições encerradas. Se você já se inscreveu, informe o CPF e a data de nascimento usada para abrir o comprovante.' };

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
  garantirCabecalho_(sh);
  return sh.getRange(2, 1, n - 1, COLS.length + 2).getValues().map(function (v) {
    return { protocolo: String(v[0]), cpf: String(v[1]), nome: String(v[2]), nascimento: String(v[3]).slice(0, 10), cidade: String(v[4]), criadoEm: String(v[5]), codigo: String(v[6]), p27: String(v[7] || ''), p28: String(v[8] || '') };
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

function titulo_(s) {
  const peq = ['da', 'de', 'do', 'das', 'dos', 'e'];
  return s.toLocaleLowerCase('pt-BR').split(' ').map(function (p, i) {
    if (i > 0 && peq.indexOf(p) >= 0) return p;
    return p.replace(/^(.)/u, function (c) { return c.toLocaleUpperCase('pt-BR'); }).replace(/([-'’.])(\p{L})/gu, function (_, a, b) { return a + b.toLocaleUpperCase('pt-BR'); });
  }).join(' ');
}

const CIDADES = ['Valparaíso', 'Guararapes', 'Araçatuba', 'Andradina', 'Birigui', 'Penápolis', 'Bilac', 'Lavínia', 'Barbosa', 'Bento de Abreu', 'Mirandópolis', 'Castilho', 'Clementina', 'Gabriel Monteiro', 'Brejo Alegre', 'Santo Antônio do Aracanguá', 'Promissão', 'Rubiácea', 'Auriflama', 'Sud Mennucci'];

function chave_(s) {
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

// Padroniza a cidade: tira "SP"/"- SP"/"/SP" (até colado, ex.: "valparaisosp"), ignora acento e maiúsculas e usa a grafia oficial da lista
function cidade_(s) {
  s = String(s).replace(/(?:\s*[-\/,]\s*|\s+)\(?sp\)?$/i, '').trim();
  let k = chave_(s);
  for (let i = 0; i < CIDADES.length; i++) if (chave_(CIDADES[i]) === k) return CIDADES[i];
  // "valparaisosp": sufixo SP colado ao nome
  if (/sp$/.test(k)) {
    const k2 = k.slice(0, -2).trim();
    for (let i = 0; i < CIDADES.length; i++) if (chave_(CIDADES[i]) === k2) return CIDADES[i];
  }
  return titulo_(s);
}

// Rodar UMA vez no editor (Executar > corrigirCidades) para padronizar as cidades já gravadas na planilha
function corrigirCidades() {
  const sh = aba_(), n = sh.getLastRow();
  if (n < 2) return 0;
  const rg = sh.getRange(2, 5, n - 1, 1), v = rg.getValues();
  let mudou = 0;
  const novo = v.map(function (r) {
    const c = cidade_(String(r[0]).replace(/\s+/g, ' ').trim());
    if (c !== r[0]) mudou++;
    return [c];
  });
  rg.setValues(novo);
  return mudou;
}

// ---------- presença e certificado ----------
function garantirCabecalho_(sh) {
  const c = sh.getRange(1, COLS.length + 1, 1, 2);
  const v = c.getValues()[0];
  if (v[0] !== COLS_PRES[0] || v[1] !== COLS_PRES[1]) {
    c.setValues([COLS_PRES]).setFontWeight('bold').setBackground('#12306b').setFontColor('#ffffff');
    sh.setColumnWidths(COLS.length + 1, 2, 150);
  }
}

// Localiza o participante pelos 5 primeiros dígitos do CPF (ou CPF completo, se houver outro inscrito com os mesmos 5) + data de nascimento
function buscar_(cpfDigitos, nasc) {
  const d = String(cpfDigitos || '').replace(/\D/g, '');
  if (d.length < 5) return { erro: 'Informe os 5 primeiros dígitos do CPF.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(nasc)) return { erro: 'Informe a data de nascimento.' };
  const lista = todos_();
  const p5 = d.slice(0, 5);
  const iguais = lista.filter(function (r) { return r.cpf.slice(0, 5) === p5; });
  if (iguais.length > 1 && d.length < 11) return { erro: 'Há outro inscrito com os mesmos 5 primeiros dígitos. Informe o CPF completo.', codigo: 'CPF_COMPLETO' };
  for (let i = 0; i < lista.length; i++) {
    const r = lista[i];
    if (r.cpf.indexOf(d) === 0 && r.nascimento === nasc) return { reg: r, linha: i + 2 };
  }
  return { erro: 'Não encontramos uma inscrição com esses dados. Confira o CPF e a data de nascimento.' };
}

function presenca_(req) {
  const dia = String(req.dia || '');
  const cfg = CONFIG.DIAS[dia];
  if (!cfg) return { ok: false, erro: 'Dia inválido.' };
  const agora = new Date().getTime();
  if (agora < new Date(cfg.ini).getTime()) return { ok: false, codigo: 'AINDA_NAO', erro: 'O registro de presença deste dia ainda não abriu.' };
  if (agora > new Date(cfg.fim).getTime()) return { ok: false, codigo: 'ENCERRADO', erro: 'O registro de presença deste dia já foi encerrado.' };
  const b = buscar_(req.cpf, String(req.nascimento || ''));
  if (!b.reg) return { ok: false, codigo: b.codigo, erro: b.erro };
  const sh = aba_();
  garantirCabecalho_(sh);
  const cel = sh.getRange(b.linha, cfg.col);
  const atual = String(cel.getValue() || '');
  if (atual) return { ok: true, jaRegistrado: true, nome: b.reg.nome, dia: dia, hora: atual };
  const hora = new Date().toISOString();
  cel.setNumberFormat('@').setValue(hora);
  return { ok: true, jaRegistrado: false, nome: b.reg.nome, dia: dia, hora: hora };
}

function certificado_(req) {
  if (new Date().getTime() < new Date(CONFIG.CERT_LIBERA_EM).getTime()) return { ok: false, codigo: 'AINDA_NAO', erro: 'Os certificados serão liberados após o encerramento do evento.' };
  const b = buscar_(req.cpf, String(req.nascimento || ''));
  if (!b.reg) return { ok: false, codigo: b.codigo, erro: b.erro };
  const r = b.reg;
  if (!r.p27 || !r.p28) {
    const falta = (!r.p27 && !r.p28) ? 'nos dois dias (27 e 28/11)' : (!r.p27 ? 'no dia 27/11' : 'no dia 28/11');
    return { ok: false, codigo: 'SEM_PRESENCA', erro: 'Não há presença registrada ' + falta + '. O certificado é emitido para quem participou dos dois dias. Procure a organização se acredita que há um engano.' };
  }
  return { ok: true, registro: r };
}
