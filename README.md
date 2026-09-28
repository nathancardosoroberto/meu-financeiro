# Meu Financeiro — dashboard

Dashboard que lê a planilha `planilha.xlsx` e mostra quanto ainda dá para gastar no mês, a reserva, a estimativa do mês, as dívidas e os gráficos, com filtros.

## Publicar no GitHub (uma vez)
1. Crie um repositório em github.com → **New repository** (ex.: `meu-financeiro`).
2. **Add file → Upload files**: arraste TODOS os arquivos desta pasta (inclusive a pasta `vendor`).
3. Arraste também a sua planilha, com o nome exato **`planilha.xlsx`**. Clique em **Commit changes**.
4. **Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: `main` / `(root)` → Save**.
5. Em 1–2 minutos o site fica em `https://SEU-USUARIO.github.io/meu-financeiro/`.

## Atualizar
Salve a planilha no Excel, e no GitHub faça **Add file → Upload files** com o arquivo `planilha.xlsx` (mesmo nome, ele substitui o antigo) → **Commit**. O dashboard lê a versão nova sozinho em ~1 minuto (clique em **Recarregar** se a página já estiver aberta).

## Alternativa: Google Sheets (atualiza sem subir arquivo)
Se você edita a planilha no Google Sheets: **Arquivo → Compartilhar → Publicar na Web → Documento inteiro → Microsoft Excel (.xlsx) → Publicar**. Copie o link e cole em `config.js` no lugar de `planilha.xlsx`. Se o navegador bloquear o link, volte para a opção do arquivo.

## Privacidade
O GitHub Pages é sempre público: quem tiver o endereço do site consegue ver os números (e o `planilha.xlsx` fica visível num repositório público). Se não quiser isso, não envie a planilha: publique só os arquivos do dashboard e use o botão **Carregar planilha** — o arquivo é lido só no seu navegador e fica guardado nele até você carregar outro.
