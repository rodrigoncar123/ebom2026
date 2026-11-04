/* =====================================================================
   EBOM — Configurações do evento
   Para ativar o cadastro REAL (todos os celulares gravando no mesmo lugar),
   cole abaixo a URL do Google Apps Script (veja o README.md).
   Enquanto API_URL estiver vazio, o site funciona em MODO DEMONSTRAÇÃO
   (dados salvos só no próprio aparelho — apenas para testes).
   ===================================================================== */
window.EBOM_CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbzu33sLgVIWuR36b7NGlnmUOiXOFdjGxwm5IVLHUkybfcyU3eMjegqL-zNw6f1sJf5x/exec",

  EVENTO: {
    sigla: "EBOM",
    nome: "Escola Bíblica de Obreiros e Membros",
    ano: 2026,
    igreja: "Assembleia de Deus Brás Valparaíso",
    ministerio: "Ministério de Madureira",
    local: "Valparaíso - SP",
    tema: "Porque pela graça sois salvos, por meio da fé; e isso não vem de vós; é dom de Deus.",
    referencia: "Efésios 2:8",
    referenciaUrl: "https://www.bible.com/pt/bible/212/EPH.2.8.ARC",
    dias: [
      { data: "2026-11-27", diaSemana: "Sexta-feira", hora: "19:00", pastor: "Pr. Arnaldo", foto: "assets/pr-arnaldo.jpg" },
      { data: "2026-11-28", diaSemana: "Sábado",      hora: "19:00", pastor: "Pr. Adauto",  foto: "assets/pr-adauto.jpg" }
    ]
  }
};
