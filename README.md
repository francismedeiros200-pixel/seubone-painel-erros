# Painel de Erros — SeuBoné

CRM leve para registrar erros de pedidos, visualizar indicadores e auditar casos.
Os dados vivem numa planilha do Google; um Web App do Google Apps Script faz a
ponte entre a planilha e o painel.

```
┌─────────────┐   registra/audita (POST)   ┌──────────────┐   grava    ┌───────────┐
│  index.html │ ─────────────────────────▶ │  Apps Script │ ─────────▶ │  Planilha │
│  (o painel) │ ◀───────────────────────── │  (Codigo.gs) │ ◀───────── │  Google   │
└─────────────┘        lê dados (GET)       └──────────────┘    lê      └───────────┘
```

## Arquivos

| Arquivo       | O que é                                                             |
|---------------|--------------------------------------------------------------------|
| `index.html`  | O painel inteiro (front-end). É o que a equipe abre.               |
| `Codigo.gs`   | O backend (Apps Script) que lê e escreve na planilha.             |
| `README.md`   | Este guia.                                                          |

- **Planilha:** `Cadastro de erros` (Google Sheets)
- **API atual (Web App /exec):** já configurada em `index.html` → `CONFIG.API_URL`

---

## O que já funciona hoje

- ✅ O painel **lê os dados reais** da planilha (346 registros na última checagem).
- ✅ Botão **"Registrar novo erro"** (formulário completo) grava na planilha.
- ✅ Tela **"Casos / Auditoria"** completa os casos pendentes.

---

## 1) Publicar o backend (Apps Script)

> ⚠️ **Segurança:** a leitura já funciona. Só troque o backend seguindo os passos
> abaixo, que evitam quebrar o que está no ar.

1. Abra a planilha → menu **Extensões → Apps Script**.
2. **Antes de mudar qualquer coisa**, copie o conteúdo atual do editor e guarde
   num arquivo de backup (ex: `Codigo_ANTIGO.gs`).
3. Cole o conteúdo de `Codigo.gs` (este repositório) no editor.
4. Se a sua aba **não** se chama `Cadastro de erros`, ajuste a linha
   `var SHEET_NAME = '...'` no topo do arquivo.
5. **Confira o mapeamento** antes de confiar: no editor, selecione a função
   `verColunas` e clique em **▶ Executar**. Em *Registros de execução* aparece
   como cada coluna foi mapeada. Verifique especialmente:
   - `tipoResolucao` deve apontar para a coluna **"Solução"**;
   - `data` deve apontar para a **1ª coluna** (a que guarda a data).
6. Publique: **Implantar → Gerenciar implantações →** (lápis para editar a
   implantação existente) **→ Nova versão → Implantar**.
   - Mantendo a **mesma implantação**, a URL `/exec` **não muda** e o painel
     continua funcionando sem editar nada.
   - Acesso: **"Qualquer pessoa"** (é o que permite o painel ler sem login).

### Testando a leitura sem risco
Abra a URL `/exec` no navegador. Deve aparecer algo como
`{"ok":true,"rows":[ ... ]}`. Compare com o backup antigo se tiver dúvida.

---

## 2) Publicar o painel no GitHub Pages (link fixo para a equipe)

1. Crie uma conta em <https://github.com> (se ainda não tiver).
2. Crie um repositório novo, ex: `painel-erros`.
3. Faça upload do `index.html` (botão **Add file → Upload files**).
4. Vá em **Settings → Pages → Build and deployment**:
   - **Source:** Deploy from a branch
   - **Branch:** `main` / `/root` → **Save**
5. Em ~1 minuto o painel fica no ar em:
   `https://SEU-USUARIO.github.io/painel-erros/`
6. Compartilhe esse link com a equipe. Toda vez que você subir um `index.html`
   novo, todos veem a versão atualizada.

> Alternativas ao GitHub Pages: Netlify Drop (arrastar a pasta em
> <https://app.netlify.com/drop>) ou Cloudflare Pages. Qualquer hospedagem de
> site estático serve — o painel é um único arquivo.

---

## 3) Como usar (equipe)

- **Registrar erro:** botão **"+ Novo caso"** no topo. Campos obrigatórios: ID
  da venda, nome do card, descrição e quem cadastrou. A **classificação**
  (setor, custo, resolução…) é opcional: se preencher, o caso já entra
  auditado; se não, entra como **pendente** e alguém completa depois.
- **Auditar:** tela **"Casos / Auditoria"** → clicar num caso pendente.
- **Atualizar dados:** botão **⟳** no topo (recarrega da planilha).

---

## Detalhes técnicos

- O mapeamento de colunas é por **nome de cabeçalho** (tolerante a acento e
  maiúscula), então reordenar colunas na planilha não quebra o painel.
- A escrita usa `fetch(..., { mode: 'no-cors' })`: o navegador grava mas não
  consegue ler a resposta, então o painel faz um *refresh* em segundo plano
  ~1,8s depois para confirmar.
- Quirks do mapeamento (confirmados por engenharia reversa da planilha atual):
  - `tipoResolucao` ← coluna **"Solução"**
  - `data` ← **1ª coluna**
  - `linkPedido` ← 1ª URL encontrada dentro da **descrição**
