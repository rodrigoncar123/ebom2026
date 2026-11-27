/* EBOM — funções compartilhadas (validação, API, PDF, Excel) */
(function (root) {
  const CFG = root.EBOM_CONFIG;
  const E = CFG.EVENTO;
  const BLUE = [18, 48, 107], GOLD = [201, 162, 39], GRAY = [100, 110, 125], SOFT = [238, 243, 252];

  /* ---------- utilidades ---------- */
  const onlyDigits = (s) => String(s || "").replace(/\D/g, "");
  const fmtCpf = (s) => {
    const d = onlyDigits(s).slice(0, 11);
    return d.replace(/(\d{3})(\d)/, "$1.$2").replace(/(\d{3})(\d)/, "$1.$2").replace(/(\d{3})(\d{1,2})$/, "$1-$2");
  };
  // Relatório em PDF: só os 5 primeiros dígitos (ex.: 319.69*.***-**)
  const maskCpfRel = (s) => { const d = onlyDigits(s); return d.slice(0, 3) + "." + d.slice(3, 5) + "*.***-**"; };
  const maskCpf = (s) => { const d = onlyDigits(s); return "***." + d.slice(3, 6) + "." + d.slice(6, 9) + "-**"; };

  function validaCpf(c) {
    c = onlyDigits(c);
    if (c.length !== 11 || /^(\d)\1+$/.test(c)) return false;
    for (let t = 9; t < 11; t++) {
      let s = 0;
      for (let i = 0; i < t; i++) s += Number(c[i]) * (t + 1 - i);
      const d = ((s * 10) % 11) % 10;
      if (d !== Number(c[t])) return false;
    }
    return true;
  }

  function validaNascimento(iso) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || "")) return false;
    const [y, m, d] = iso.split("-").map(Number);
    const dt = new Date(y, m - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return false;
    return y >= 1900 && dt <= new Date();
  }

  const NOME_RE = /^[\p{L} .'’-]+$/u;
  const limpa = (s) => String(s || "").replace(/\s+/g, " ").trim();
  const validaNome = (n) => { n = limpa(n); return n.length >= 5 && n.split(" ").length >= 2 && NOME_RE.test(n); };
  const validaCidade = (c) => { c = limpa(c); return c.length >= 2 && NOME_RE.test(c); };

  const fmtData = (iso) => { const p = String(iso || "").slice(0, 10).split("-"); return p.length === 3 ? `${p[2]}/${p[1]}/${p[0]}` : iso; };
  const fmtDataHora = (iso) => {
    try {
      return new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "short", timeZone: "America/Sao_Paulo" })
        .format(new Date(iso)).replace(",", " às");
    } catch (e) { return iso; }
  };

  /* ---------- API (Google Apps Script) ou modo demonstração ---------- */
  function cyrb(str) {
    let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
    for (let i = 0; i < str.length; i++) { const ch = str.charCodeAt(i); h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677); }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16).toUpperCase().padStart(14, "0");
  }

  const Demo = {
    KEY: "ebom_demo_v1",
    load() { try { return JSON.parse(root.localStorage.getItem(this.KEY)) || []; } catch (e) { return []; } },
    save(l) { root.localStorage.setItem(this.KEY, JSON.stringify(l)); },
    handle(req) {
      const lista = this.load();
      const cpf = onlyDigits(req.cpf);
      if (req.acao === "existe") return { ok: true, existe: lista.some((r) => r.cpf === cpf) };
      if (req.acao === "inscrever") {
        if (!validaCpf(cpf)) return { ok: false, erro: "CPF inválido." };
        const ja = lista.find((r) => r.cpf === cpf);
        if (ja) {
          if (ja.nascimento === req.nascimento) return { ok: true, jaCadastrado: true, registro: ja };
          return { ok: false, codigo: "CPF_EXISTE", erro: "Este CPF já está cadastrado. Para abrir o comprovante, informe a data de nascimento usada na inscrição." };
        }
        const criadoEm = new Date().toISOString();
        const protocolo = `EBOM-${E.ano}-${String(lista.length + 1).padStart(4, "0")}`;
        const reg = { protocolo, cpf, nome: limpa(req.nome), nascimento: req.nascimento, cidade: limpa(req.cidade), criadoEm,
          codigo: cyrb(protocolo + cpf + criadoEm).slice(0, 10).replace(/(.{5})(.{5})/, "$1-$2") };
        lista.push(reg); this.save(lista);
        return { ok: true, jaCadastrado: false, registro: reg };
      }
      if (req.acao === "listar") return { ok: true, registros: lista, demo: true };
      return { ok: false, erro: "Ação desconhecida." };
    }
  };

  const Api = {
    demo: !CFG.API_URL,
    async call(payload) {
      if (this.demo) return Demo.handle(payload);
      const r = await fetch(CFG.API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(payload) });
      if (!r.ok) throw new Error("Falha de comunicação com o servidor.");
      return r.json();
    }
  };

  /* ---------- carregamento sob demanda das bibliotecas ---------- */
  const _libs = {};
  function loadLib(src) {
    if (!_libs[src]) _libs[src] = new Promise((res, rej) => {
      const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => rej(new Error("Não foi possível carregar " + src)); document.head.appendChild(s);
    });
    return _libs[src];
  }
  const LIBS = {
    jspdf: "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
    autotable: "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js",
    exceljs: "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"
  };

  /* ---------- logo da igreja (para os PDFs) ---------- */
  let _logo = null;
  function carregaLogo() {
    if (!_logo) _logo = fetch(E.logo).then((r) => r.blob()).then((b) => new Promise((res, rej) => {
      const fr = new FileReader(); fr.onload = () => {
        const im = new Image(); im.onload = () => res({ data: fr.result, w: im.naturalWidth, h: im.naturalHeight }); im.onerror = rej; im.src = fr.result;
      }; fr.onerror = rej; fr.readAsDataURL(b);
    })).catch(() => null);
    return _logo;
  }
  const nomeCompleto = () => `${E.igreja} · ${E.ministerio}`;

  /* ---------- PDF: comprovante de inscrição ---------- */
  async function gerarComprovante(reg) {
    await loadLib(LIBS.jspdf);
    const logo = await carregaLogo();
    const { jsPDF } = root.jspdf;
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    const W = 210, dy = 14; // dy: espaço extra do cabeçalho para a logo

    // cabeçalho
    doc.setFillColor(...BLUE); doc.rect(0, 0, W, 42 + dy, "F");
    doc.setFillColor(...GOLD); doc.rect(0, 42 + dy, W, 2, "F");
    if (logo) { const lh = 20, lw = lh * logo.w / logo.h; doc.addImage(logo.data, "PNG", (W - lw) / 2, 4, lw, lh); }
    doc.setTextColor(255, 255, 255); doc.setFont("helvetica", "bold"); doc.setFontSize(28); doc.text(E.sigla, W / 2, 36, { align: "center" });
    doc.setFont("helvetica", "normal"); doc.setFontSize(12); doc.text(E.nome, W / 2, 43.5, { align: "center" });
    doc.setFontSize(8.5); doc.text(nomeCompleto(), W / 2, 50, { align: "center" });

    // título + protocolo
    doc.setTextColor(...BLUE); doc.setFont("helvetica", "bold"); doc.setFontSize(17);
    doc.text("COMPROVANTE DE INSCRIÇÃO", W / 2, 58 + dy, { align: "center" });
    doc.setFillColor(...SOFT); doc.roundedRect(20, 65 + dy, 170, 18, 3, 3, "F");
    doc.setFontSize(8); doc.setTextColor(...GRAY); doc.setFont("helvetica", "normal");
    doc.text("PROTOCOLO", 28, 72 + dy);
    doc.text("CÓDIGO DE VERIFICAÇÃO", 120, 72 + dy);
    doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.setTextColor(...BLUE);
    doc.text(reg.protocolo, 28, 79 + dy); doc.text(reg.codigo, 120, 79 + dy);

    // dados do participante
    const campo = (rot, val, x, y, w) => {
      doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(...GRAY); doc.text(rot.toUpperCase(), x, y);
      doc.setFont("helvetica", "bold"); doc.setFontSize(12); doc.setTextColor(30, 35, 45);
      doc.text(doc.splitTextToSize(String(val), w)[0], x, y + 6);
    };
    section(doc, "Dados do participante", 96 + dy);
    campo("Nome completo", String(reg.nome).toUpperCase(), 22, 106 + dy, 166);
    campo("CPF", maskCpf(reg.cpf), 22, 120 + dy, 80);
    campo("Data de nascimento", fmtData(reg.nascimento), 110, 120 + dy, 78);
    campo("Cidade", reg.cidade, 22, 134 + dy, 80);
    campo("Inscrição registrada em", fmtDataHora(reg.criadoEm), 110, 134 + dy, 78);

    // evento
    section(doc, "Evento", 152 + dy);
    doc.setFont("helvetica", "italic"); doc.setFontSize(11); doc.setTextColor(30, 35, 45);
    const tema = doc.splitTextToSize(`“${E.tema}” (${E.referencia})`, 166);
    doc.text(tema, 22, 162 + dy);
    let y = 162 + dy + tema.length * 5.5 + 3;
    doc.setFont("helvetica", "normal"); doc.setFontSize(11);
    E.dias.forEach((d) => {
      doc.setFont("helvetica", "bold"); doc.text(`${d.diaSemana}, ${fmtData(d.data)} — ${d.hora.replace(":", "h")}`, 22, y);
      doc.setFont("helvetica", "normal"); doc.text(`Ministração: ${d.pastor}`, 120, y);
      y += 6.5;
    });
    doc.setFontSize(9); doc.setTextColor(...GRAY);
    doc.text(doc.splitTextToSize(`Local: ${nomeCompleto()}`, 166), 22, y); y += 14;

    // declaração
    section(doc, "Declaração", y); y += 10;
    doc.setFont("helvetica", "normal"); doc.setFontSize(10.5); doc.setTextColor(30, 35, 45);
    const decl = `Declaramos que ${String(reg.nome).toUpperCase()}, inscrito(a) no CPF sob o nº ${maskCpf(reg.cpf)}, da cidade de ${reg.cidade}, realizou sua inscrição para participar da ${E.nome} (${E.sigla} ${E.ano}), tendo sido registrada sob o protocolo ${reg.protocolo} em ${fmtDataHora(reg.criadoEm)}.`;
    const linhas = doc.splitTextToSize(decl, 166); doc.text(linhas, 22, y, { lineHeightFactor: 1.5 });
    y += linhas.length * 6.2 + 6;
    doc.setFontSize(9); doc.setTextColor(...GRAY);
    const nota = doc.splitTextToSize("Este comprovante atesta a inscrição no evento e é válido mediante a conferência do protocolo e do código de verificação junto à organização. Apresente-o (impresso ou no celular) na entrada, se solicitado.", 166);
    doc.text(nota, 22, y, { lineHeightFactor: 1.4 });

    // rodapé
    doc.setDrawColor(...GOLD); doc.setLineWidth(0.6); doc.line(20, 270, 190, 270);
    doc.setFontSize(8); doc.setTextColor(...GRAY);
    doc.text(`Documento emitido eletronicamente em ${fmtDataHora(new Date().toISOString())}.`, W / 2, 276, { align: "center" });
    doc.text("Os dados pessoais são usados somente para a organização do evento (Lei nº 13.709/2018 — LGPD).", W / 2, 281, { align: "center" });
    doc.text(nomeCompleto(), W / 2, 286, { align: "center" });
    return doc;
  }

  /* ---------- PDF: certificado de participação (A4 paisagem) ---------- */
  async function carregaImg(src) {
    try {
      const b = await fetch(src).then((r) => { if (!r.ok) throw 0; return r.blob(); });
      return await new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => { const im = new Image(); im.onload = () => res({ data: fr.result, w: im.naturalWidth, h: im.naturalHeight }); im.onerror = rej; im.src = fr.result; }; fr.onerror = rej; fr.readAsDataURL(b); });
    } catch (e) { return null; }
  }
  async function gerarCertificado(reg) {
    await loadLib(LIBS.jspdf);
    const logo = await carregaLogo();
    const { jsPDF } = root.jspdf;
    const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
    const W = 297, H = 210;
    // moldura
    doc.setFillColor(...BLUE); doc.rect(0, 0, W, H, "F");
    doc.setFillColor(255, 255, 255); doc.rect(7, 7, W - 14, H - 14, "F");
    doc.setDrawColor(...GOLD); doc.setLineWidth(1.2); doc.rect(11, 11, W - 22, H - 22);
    doc.setLineWidth(0.3); doc.rect(13.5, 13.5, W - 27, H - 27);
    if (logo) { const lh = 22, lw = lh * logo.w / logo.h; doc.setFillColor(...BLUE); doc.roundedRect((W - lw) / 2 - 8, 16, lw + 16, lh + 6, 3, 3, "F"); doc.addImage(logo.data, "PNG", (W - lw) / 2, 19, lw, lh); }
    doc.setTextColor(...GRAY); doc.setFont("helvetica", "normal"); doc.setFontSize(9.5);
    doc.text(nomeCompleto(), W / 2, 50, { align: "center" });
    doc.setTextColor(...BLUE); doc.setFont("helvetica", "bold"); doc.setFontSize(34);
    doc.text("CERTIFICADO", W / 2, 66, { align: "center" });
    doc.setFillColor(...GOLD); doc.rect(W / 2 - 25, 70, 50, 1.2, "F");
    doc.setFont("helvetica", "normal"); doc.setFontSize(13); doc.setTextColor(30, 35, 45);
    doc.text("Certificamos que", W / 2, 84, { align: "center" });
    doc.setFont("helvetica", "bold"); doc.setFontSize(24); doc.setTextColor(...BLUE);
    doc.text(doc.splitTextToSize(String(reg.nome).toUpperCase(), W - 60)[0], W / 2, 98, { align: "center" });
    doc.setFont("helvetica", "normal"); doc.setFontSize(10); doc.setTextColor(...GRAY);
    doc.text(`CPF ${maskCpf(reg.cpf)}`, W / 2, 105, { align: "center" });
    doc.setFontSize(13); doc.setTextColor(30, 35, 45);
    const pres = E.dias.filter((d, i) => (i === 0 ? reg.p27 : reg.p28));
    const lista = pres.length ? pres : E.dias;
    const dias = lista.map((d) => fmtData(d.data).slice(0, 2)).join(" e ");
    const mes = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"][Number(lista[0].data.slice(5, 7)) - 1];
    const txt = `participou da ${E.nome} (${E.sigla} ${E.ano}), realizada ${lista.length > 1 ? "nos dias" : "no dia"} ${dias} de ${mes} de ${E.ano}, em ${E.igreja} — ${E.ministerio}, ${E.local || ""}. Tema: “${E.tema}” (${E.referencia}).`;
    doc.text(doc.splitTextToSize(txt, W - 70), W / 2, 118, { align: "center", lineHeightFactor: 1.5 });
    // assinaturas
    const ass = (window.EBOM_CONFIG.ASSINATURAS || []);
    const imgs = await Promise.all(ass.map((a) => a.img ? carregaImg(a.img) : null));
    const xs = ass.length === 1 ? [W / 2] : [W / 2 - 55, W / 2 + 55];
    ass.forEach((a, i) => {
      const x = xs[i];
      if (imgs[i]) { const ih = 16, iw = Math.min(50, ih * imgs[i].w / imgs[i].h); doc.addImage(imgs[i].data, "PNG", x - iw / 2, 150 - ih, iw, ih); }
      doc.setDrawColor(30, 35, 45); doc.setLineWidth(0.3); doc.line(x - 38, 152, x + 38, 152);
      doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setTextColor(30, 35, 45); doc.text(a.nome, x, 158, { align: "center" });
      doc.setFont("helvetica", "normal"); doc.setFontSize(9.5); doc.setTextColor(...GRAY); doc.text(a.cargo, x, 163.5, { align: "center" });
    });
    doc.setFontSize(8); doc.setTextColor(...GRAY);
    doc.text(`Protocolo ${reg.protocolo} · Código de verificação ${reg.codigo} · Emitido em ${fmtDataHora(new Date().toISOString())}`, W / 2, 184, { align: "center" });
    doc.text("Certificado emitido eletronicamente conforme a presença registrada no evento.", W / 2, 189, { align: "center" });
    return doc;
  }

  function section(doc, titulo, y) {
    doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setTextColor(...BLUE); doc.text(titulo.toUpperCase(), 22, y);
    doc.setDrawColor(...GOLD); doc.setLineWidth(0.5); doc.line(22, y + 2, 188, y + 2);
  }

  /* ---------- PDF: relatório de inscritos ---------- */
  async function gerarRelatorioPdf(regs) {
    await loadLib(LIBS.jspdf); await loadLib(LIBS.autotable);
    const logo = await carregaLogo();
    const { jsPDF } = root.jspdf;
    const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
    const W = 297;
    doc.setFillColor(...BLUE); doc.rect(0, 0, W, 24, "F"); doc.setFillColor(...GOLD); doc.rect(0, 24, W, 1.5, "F");
    if (logo) { const lh = 18, lw = lh * logo.w / logo.h; doc.addImage(logo.data, "PNG", W - 12 - lw, 3, lw, lh); }
    doc.setTextColor(255, 255, 255); doc.setFont("helvetica", "bold"); doc.setFontSize(16);
    doc.text(`${E.sigla} ${E.ano} — Relatório de inscritos`, 12, 12);
    doc.setFont("helvetica", "normal"); doc.setFontSize(9);
    doc.text(`${E.nome} · ${nomeCompleto()}`, 12, 19);
    doc.text(`Emitido em ${fmtDataHora(new Date().toISOString())}`, 12, 22.5);

    const porCidade = resumoCidades(regs);
    doc.setTextColor(30, 35, 45); doc.setFont("helvetica", "bold"); doc.setFontSize(11);
    doc.text(`Total de inscritos: ${regs.length}`, 12, 34);
    doc.setFont("helvetica", "normal"); doc.setFontSize(9);
    const txt = doc.splitTextToSize("Por cidade: " + porCidade.map(([c, n]) => `${c} (${n})`).join(" · "), W - 24);
    doc.text(txt, 12, 40);

    doc.autoTable({
      startY: 40 + txt.length * 4.5 + 3,
      head: [["Nº", "Protocolo", "Nome completo", "CPF", "Nascimento", "Cidade", "Inscrito em", "27/11", "28/11"]],
      body: regs.map((r, i) => [i + 1, r.protocolo, r.nome, maskCpfRel(r.cpf), fmtData(r.nascimento), r.cidade, fmtDataHora(r.criadoEm), pres(r.p27), pres(r.p28)]),
      styles: { fontSize: 8.5, cellPadding: 2 },
      headStyles: { fillColor: BLUE, textColor: 255 },
      alternateRowStyles: { fillColor: SOFT },
      margin: { left: 12, right: 12 },
      didDrawPage: () => {
        const n = doc.internal.getNumberOfPages();
        doc.setFontSize(8); doc.setTextColor(...GRAY);
        doc.text(`Página ${n}`, W / 2, 205, { align: "center" });
      }
    });
    return doc;
  }
  const pres = (iso) => { if (!iso) return ""; const d = new Date(iso); return isNaN(d) ? "" : "✓ " + d.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit", timeZone: "America/Sao_Paulo" }); };
  function resumoCidades(regs) {
    const m = {};
    const chave = (x) => x.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/(?:\s*[-\/,]\s*|\s+)\(?sp\)?$/i, "").replace(/(valparaiso)sp$/, "$1").trim();
    const nomes = {};
    regs.forEach((r) => {
      const c0 = limpa(r.cidade) || "—", k = chave(c0);
      if (!nomes[k]) nomes[k] = c0.replace(/(?:\s*[-\/,]\s*|\s+)\(?sp\)?$/i, "").replace(/(valparaíso)sp$/i, "$1").trim() || c0;
      const c = nomes[k]; m[c] = (m[c] || 0) + 1;
    });
    return Object.entries(m).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "pt-BR"));
  }

  /* ---------- Excel: relatório de inscritos (formato tabela) ---------- */
  async function gerarRelatorioXlsx(regs) {
    await loadLib(LIBS.exceljs);
    const wb = new root.ExcelJS.Workbook();
    wb.creator = E.igreja; wb.created = new Date();

    const ws = wb.addWorksheet("Inscritos", { views: [{ state: "frozen", ySplit: 1 }] });
    ws.addTable({
      name: "Inscritos", ref: "A1", headerRow: true,
      style: { theme: "TableStyleMedium2", showRowStripes: true },
      columns: ["Nº", "Protocolo", "Nome completo", "CPF", "Data de nascimento", "Cidade", "Inscrito em", "Presença 27/11", "Presença 28/11", "Código de verificação"].map((n) => ({ name: n, filterButton: true })),
      rows: regs.map((r, i) => [i + 1, r.protocolo, r.nome, fmtCpf(r.cpf), fmtData(r.nascimento), r.cidade, fmtDataHora(r.criadoEm), pres(r.p27), pres(r.p28), r.codigo])
    });
    [6, 18, 40, 17, 18, 22, 22, 16, 16, 22].forEach((w, i) => (ws.getColumn(i + 1).width = w));

    const rs = wb.addWorksheet("Resumo");
    rs.getCell("A1").value = `${E.sigla} ${E.ano} — ${E.nome}`; rs.getCell("A1").font = { bold: true, size: 14, color: { argb: "FF12306B" } };
    rs.getCell("A2").value = "Total de inscritos"; rs.getCell("B2").value = regs.length; rs.getCell("A2").font = { bold: true }; rs.getCell("B2").font = { bold: true };
    rs.addTable({
      name: "PorCidade", ref: "A4", headerRow: true, style: { theme: "TableStyleMedium2", showRowStripes: true },
      columns: [{ name: "Cidade", filterButton: true }, { name: "Inscritos", filterButton: true }],
      rows: resumoCidades(regs).map(([c, n]) => [c, n])
    });
    rs.getColumn(1).width = 30; rs.getColumn(2).width = 14;

    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  function baixar(blob, nome) {
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = nome;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }

  root.EBOM = { CFG, E, Api, onlyDigits, fmtCpf, maskCpf, validaCpf, validaNascimento, validaNome, validaCidade, limpa,
    fmtData, fmtDataHora, gerarComprovante, gerarCertificado, gerarRelatorioPdf, gerarRelatorioXlsx, resumoCidades, baixar, loadLib, LIBS };
})(typeof window !== "undefined" ? window : globalThis);
