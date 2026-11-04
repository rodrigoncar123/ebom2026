# EBOM 2026 — Sistema de inscrição

Site estático (HTML/CSS/JS) para GitHub Pages + banco de dados em **Planilha Google** (via Apps Script).

## Arquivos

| Arquivo | Função |
|---|---|
| `index.html` | Página de inscrição (celular), comprovante em PDF |
| `admin.html` | Relatório de inscritos (senha) — baixa Excel e PDF |
| `config.js` | Dados do evento + `API_URL` |
| `common.js`, `style.css` | Lógica e visual |
| `assets/` | Fotos dos pastores |
| `apps-script/Code.gs` | API que grava na planilha |

## 1. Ativar o cadastro real (Planilha Google) — 5 min

1. Crie uma planilha em sheets.google.com (ex.: "EBOM 2026 – Inscritos").
2. **Extensões → Apps Script**. Apague o conteúdo e cole `apps-script/Code.gs`.
3. No topo, troque `ADMIN_KEY` (senha do relatório) e `SALT` (texto secreto qualquer).
4. **Implantar → Nova implantação → Aplicativo da Web** · Executar como: *Eu* · Quem tem acesso: *Qualquer pessoa* → Implantar → autorize → **copie a URL** (termina em `/exec`).
5. Em `config.js`, cole a URL em `API_URL: "..."`.

> Ao alterar o `Code.gs` depois, use *Implantar → Gerenciar implantações → Editar → Nova versão*.

## 2. Publicar no GitHub Pages

1. Crie um repositório (ex.: `ebom`) e envie todos os arquivos desta pasta.
2. **Settings → Pages → Branch: `main` / pasta `/ (root)`** → Save.
3. O link fica `https://SEU-USUARIO.github.io/ebom/` — compartilhe no WhatsApp.
   Relatório: `.../ebom/admin.html`.

## Regras do sistema

- CPF validado (dígitos verificadores); um CPF = uma inscrição.
- CPF repetido: avisa "já inscrito" e abre o comprovante **se a data de nascimento conferir** (protege os dados de terceiros).
- Comprovante PDF: protocolo `EBOM-2026-0001`, código de verificação, CPF parcialmente oculto, data/hora do registro.
- Dados pessoais ficam só na Planilha Google (não no GitHub). Consentimento LGPD obrigatório.
- Sem `API_URL`, roda em **modo demonstração** (dados só no aparelho).
