/* =====================================================================
   EBOM — Configurações do evento
   Para ativar o cadastro REAL (todos os celulares gravando no mesmo lugar),
   cole abaixo a URL do Google Apps Script (veja o README.md).
   Enquanto API_URL estiver vazio, o site funciona em MODO DEMONSTRAÇÃO
   (dados salvos só no próprio aparelho — apenas para testes).
   ===================================================================== */
window.EBOM_CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbzu33sLgVIWuR36b7NGlnmUOiXOFdjGxwm5IVLHUkybfcyU3eMjegqL-zNw6f1sJf5x/exec",

  // Encerramento das inscrições (horário de Brasília). O servidor também bloqueia após este horário.
  ENCERRA_EM: "", // inscrições reabertas em 27/11 (vazio = sem prazo)

  // Registro de presença por QR Code (horário de Brasília) — o servidor aplica as mesmas janelas
  PRESENCA: {
    "27": { ini: "2026-11-27T18:00:00-03:00", fim: "2026-11-27T23:00:00-03:00" },
    "28": { ini: "2026-11-28T18:00:00-03:00", fim: "2026-11-28T23:00:00-03:00" }
  },
  // Certificado liberado a partir deste horário (quem participou dos dois dias)
  CERT_LIBERA_EM: "2026-11-28T23:00:00-03:00",
  ASSINATURAS: [
    { nome: "Pr. Edvaldo Natalino Pedro", cargo: "Pastor", img: "" },
    { nome: "Felipe Custódio", cargo: "Coordenador da EBD", img: "assets/assinatura-felipe.png" }
  ],

  EVENTO: {
    sigla: "EBOM",
    nome: "Escola Bíblica de Obreiros e Membros",
    ano: 2026,
    igreja: "Assembleia de Deus Brás Valparaíso",
    ministerio: "Ministério de Madureira - Campo de Guararapes",
    logo: "assets/logo-ad.png",
    local: "Valparaíso - SP",
    tema: "Porque pela graça sois salvos, por meio da fé; e isso não vem de vós; é dom de Deus.",
    referencia: "Efésios 2:8",
    referenciaUrl: "https://www.bible.com/pt/bible/212/EPH.2.8.ARC",
    dias: [
      { data: "2026-11-27", diaSemana: "Sexta-feira", hora: "19:00", pastor: "Pr. Arnaldo", foto: "assets/pr-adauto.jpg" },
      { data: "2026-11-28", diaSemana: "Sábado",      hora: "19:00", pastor: "Pr. Adauto",  foto: "assets/pr-arnaldo.jpg" }
    ]
  }
};
