# AFRO RUN — Fumaça 2026

Site estático pronto para GitHub Pages: `index.html`, `style.css`, `script.js`, `logo.png`, `favicon.png` e `Code.gs`.

## Publicar no GitHub Pages
1. Coloque os arquivos na raiz do repositório `secretariapindobacu/afrorunfumaca`.
2. No GitHub: Settings → Pages → Deploy from a branch; selecione branch `main` e pasta `/ (root)`.
3. Configure a Planilha Google e o Apps Script conforme os passos comentados no `Code.gs`.
4. Copie a URL publicada do Web App para a constante `SCRIPT_URL` no início de `script.js`.
5. Commit e push. O formulário só envia quando SCRIPT_URL estiver configurada.

## Observações
- Os campos contato, redes e apoiadores estão marcados `[PREENCHER]`.
- A validação de CPF é feita no front-end; valide também no servidor antes de usar os dados em contexto oficial.
- O número de inscrição é sequencial conforme as linhas existentes na planilha.
- O compartilhamento via WhatsApp compartilha texto; o PNG personalizado pode ser baixado pelo botão correspondente.
- A logo original está incluída como `logo.png`; favicon também está incluído.

## Pacote reenviado com ajustes
- O número de peito é desenhado por completo antes de habilitar a confirmação e o download.
- O Apps Script usa bloqueio (LockService) para reduzir risco de números duplicados em envios simultâneos.
- A URL `SCRIPT_URL` permanece como placeholder até que o Web App seja publicado.

## Versão com menu e regulamento
- Menu no topo: links no desktop e menu sanfona (botão ☰) em tela cheia no celular. Ele destaca a seção que está na tela.
- Nova seção 03 / REGULAMENTO, com 10 artigos em abas que abrem e fecham, botão "Abrir tudo" e botão "Imprimir" (imprime só o regulamento).
- No formulário entrou o aceite obrigatório "Li e aceito o regulamento", gravado na coluna "Aceite do regulamento" da planilha (Code.gs atualizado).

## Regulamento oficial aplicado
- Texto do regulamento oficial da Comissão Organizadora (7 capítulos, 19 artigos), na íntegra.
- Inscrições a partir de 12/11/2026 (quinta-feira). Atenção: o PDF do regulamento diz 16/11; o site usa 12/11, a data confirmada pelo Ricardo.
- Local: Praça dos Potes (Gameleira do Sr. Nequinho), com concentração às 15h15 (Art. 2º e 11).
- Pontos de água no km 1,5 e na chegada (Art. 4º) marcados na trilha do percurso.
- Menores de 16 anos: o formulário pede o nome e a autorização do responsável (Art. 6º), salvos na planilha.
