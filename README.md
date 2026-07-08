# Newsletter Econômica

Projeto para geração automática de uma newsletter econômica institucional em HTML, com dados separados da apresentação e preparada para evoluir com coleta via Python e IA.

## Estrutura

```text
data/newsletter_exemplo.json       Dados da edição
templates/newsletter_web.html      Template Jinja2 da versão web/preview
static/css/newsletter_web.css      Estilos da versão web/preview
output/newsletter_preview.html     HTML gerado
output/mais_noticias.html          Página complementar de notícias
output/relatorios.html             Página complementar de relatórios em PDF
docs/                              Arquivos prontos para GitHub Pages
render_newsletter.py               Script de renderização
requirements.txt                   Dependências Python
```

## Instalação

Crie e ative um ambiente virtual, se desejar, e instale as dependências:

```powershell
pip install -r requirements.txt
```

## Como gerar

```powershell
python render_newsletter.py
```

O resultado será salvo em:

```text
output/newsletter_preview.html
output/mais_noticias.html
output/relatorios.html
output/banco_de_dados.html
docs/index.html
docs/mais_noticias.html
docs/relatorios.html
docs/banco_de_dados.html
docs/static/
```

A pasta `docs/` é a versão pronta para publicação no GitHub Pages. O script copia os arquivos estáticos e ajusta os caminhos do CSS automaticamente.

Para gerar e tentar publicar no GitHub automaticamente, use:

```powershell
python render_newsletter.py --publish
```

Esse comando faz `git add docs`, cria um commit e executa `git push`. Ele só funciona se esta pasta for um clone Git do repositório do GitHub e se o login do Git já estiver configurado no computador.

## Onde alterar os dados

Edite o arquivo:

```text
data/newsletter_exemplo.json
```

Nele ficam título, subtítulo, data da edição, indicadores econômicos, destaques, notícias, relatórios e matérias da FGV IBRE.

Use `noticias` para a página principal e `mais_noticias` para a página complementar acessada pelo botão “Mais notícias”.

Use `relatorios` para a página complementar de PDFs, acessada pela seta no card “Relatórios e indicadores”.

Os relatórios dessa página são preenchidos automaticamente a partir da Carta de Conjuntura do Ipea. O script busca a matéria mais recente de cada categoria e captura o link “Acesse o texto completo”, que aponta para o PDF:

- Visão Geral;
- Atividade Econômica;
- Inflação;
- Moeda e Crédito;
- Finanças Públicas.

## Indicadores de mercado

Ao gerar a newsletter, o script tenta atualizar automaticamente:

- Dólar (`USD-BRL`);
- Ibovespa (`^BVSP`);
- Bitcoin em reais (`BTC-BRL`).
- Meta Selic Copom, pela série 432 do SGS/Banco Central do Brasil.
- IPCA, INCC e IGP-M pela planilha `INDICES_COMPLETO.xlsx`, aba `VISÃO GERAL`.

A consulta usa os endpoints v2 da brapi.dev:

- `/api/v2/currency?currency=USD-BRL`;
- `/api/v2/stocks/quote?symbols=^BVSP`;
- `/api/v2/crypto?coin=BTC&currency=BRL`.

Se a API exigir token, defina a variável de ambiente antes de rodar:

```powershell
$env:BRAPI_TOKEN="seu_token"
python render_newsletter.py
```

Também é possível adicionar ao arquivo local `env`:

```text
BRAPI_TOKEN=seu_token
```

O script também aceita o nome `BRAPI_API_TOKEN`, caso você prefira seguir o padrão usado nos exemplos da brapi:

```text
BRAPI_API_TOKEN=seu_token
```

O token é enviado apenas no header `Authorization` da chamada feita pelo Python e nunca vai para o HTML. A Selic usa o endpoint público do Banco Central. Se algum endpoint falhar, o card recebe fallback amigável com `N/D`, classe `indicator--neutra` e `Atualização indisponível`.

Para `Dólar` e `Bitcoin`, se a brapi retornar bloqueio de permissão (`403`), o script tenta uma segunda fonte pública pela AwesomeAPI antes de usar o fallback.

Os índices de preços são lidos deste arquivo local:

```text
C:\Users\SAMSUNG\OneDrive - Sindicato da Ind da Const Civl do Estado de SP\BI\Índices de Preços\INDICES_COMPLETO.xlsx
```

Na aba `VISÃO GERAL`, o card usa `12 meses (%)` como valor principal e `Variação mensal (%)` como linha menor. Os cards de indicadores que não vêm das APIs ou dessa planilha foram removidos da faixa.

## Notícias do Blog do IBRE

O bloco `FGV IBRE | Últimas notícias` é preenchido automaticamente com as três primeiras matérias da página inicial do Blog do IBRE:

```text
https://blogdoibre.fgv.br/
```

O scraper procura links com caminho `/posts/`, ignora links de navegação e “Leia mais”, e mantém os três primeiros títulos encontrados. Se a página estiver indisponível ou a estrutura mudar, o script mantém os links de fallback em `data/newsletter_exemplo.json`.

## Validações

Antes de renderizar, o script verifica se:

- o título não está vazio;
- existe ao menos uma notícia;
- cada notícia tem título, resumo, URL, data e tópico;
- nenhuma notícia usa `url` igual a `#`;
- cada indicador tem nome e valor;
- a direção do indicador é `alta`, `queda` ou `neutra`.

Se houver erro, a geração é interrompida e as mensagens aparecem no terminal.

## Versão web e futura versão e-mail

A versão atual é uma newsletter web/preview, com HTML semântico e CSS moderno. Ela prioriza leitura, responsividade e organização visual.

No futuro, a versão específica para e-mail poderá ser adicionada em arquivos separados:

```text
templates/newsletter_email.html
static/css/newsletter_email.css
```

Essa separação evita misturar o layout moderno de preview com as restrições técnicas de clientes de e-mail.
