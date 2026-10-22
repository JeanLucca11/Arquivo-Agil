# Arquivo Ágil

Aplicação autoral por **Jean Lucca Dias** para organizar o planilhamento de documentos físicos. Esta versão independente funciona no navegador e pode ser hospedada no **Cloudflare Pages**, sem serviços do ChatGPT em execução.

## O que esta atualização resolve

O envio antigo ao GitHub continha apenas configurações na raiz. Faltavam as pastas `app`, `public` e `scripts` referenciadas por aqueles arquivos. Esta atualização repõe a aplicação completa e substitui o empacotamento anterior por React + Vite.

## Funcionalidades

- Identidade autoral azul e branca; modos claro e escuro.
- Importação local de `.xlsx` de até 50 MB em Web Worker.
- Consulta por matrícula com índice em memória, sem limite de 2.000 registros.
- Reconhecimento dos cabeçalhos de matrícula e nome nas primeiras 50 linhas.
- Preservação de zeros à esquerda presentes como texto ou formato numérico de zeros.
- Consolidação de linhas idênticas; bloqueio de matrículas com nomes conflitantes.
- Organização em caixas com etiqueta, setor Saúde Ocupacional e data de temporalidade.
- Conferência e remoção de registros; bloqueio opcional de duplicidades na caixa.
- Excel `.xlsx` com tabela nativa, filtros e linhas alternadas.
- Histórico local de até 1.000 resumos e relatórios com filtro por temporalidade.
- Demonstração opcional com oito pessoas fictícias e indicadores separados.
- Confirmação antes de descartar registros não exportados.

## Publicar gratuitamente pelo navegador

Não é necessário instalar Node.js no computador de trabalho. A Cloudflare instala as dependências e compila o código em seus próprios servidores.

1. Coloque **todos os arquivos deste projeto**, incluindo as pastas `app` e `public`, no repositório GitHub.
2. Entre em https://dash.cloudflare.com/ e crie ou use sua conta gratuita.
3. Abra **Workers & Pages → Create application → Pages → Connect to Git**. A opção também pode aparecer como **Import an existing Git repository**.
4. Escolha **GitHub** e autorize apenas o repositório deste projeto.
5. Selecione `JeanLucca11/Arquivo-gil` e clique em **Begin setup**.
6. Configure:

| Campo | Valor |
|---|---|
| Project name | `arquivo-agil-jean`, se disponível |
| Production branch | `main`, ou a branch principal do repositório |
| Framework preset | `React (Vite)`, `Vite` ou `None` com os campos abaixo |
| Build command | `npm run build` |
| Build output directory | `dist` |
| Root directory | Vazio, se `package.json` estiver na raiz |
| Environment variables | Nenhuma obrigatória para a aplicação |

7. Clique em **Save and Deploy**.
8. Aguarde a confirmação de publicação e abra o endereço `*.pages.dev` fornecido pela Cloudflare.
9. Quando novas alterações forem enviadas à branch de produção, a integração poderá publicar automaticamente uma nova versão.

O endereço gratuito exato depende do nome disponível escolhido na Cloudflare. Nenhum endereço externo é criado apenas por adicionar estes arquivos ao GitHub.

Documentação: https://developers.cloudflare.com/pages/framework-guides/deploy-a-react-site/

## Conferência após publicar

1. Abra **Iniciar demonstração**.
2. Mantenha a etiqueta `DEMO-001` e inicie as consultas.
3. Consulte `DEMO001`, `DEMO002` e `DEMO003`.
4. Revise os cinco campos e gere o Excel.
5. Confira o download e abra **Caixas processadas** e **Relatórios**.
6. Teste um período sem resultados e retorne ao período completo.
7. Alterne o tema e recarregue para conferir a preferência salva.
8. Saia da demonstração e importe uma planilha autorizada para uso real.

## Domínio próprio

O subdomínio `*.pages.dev` é fornecido pela hospedagem. Um domínio registrado em seu nome, como `.com.br`, tem custo separado e precisa estar disponível. Após registrar um domínio, use **Custom domains → Set up a domain** no projeto Pages.

## Dados e acesso

- A planilha é lida na memória do navegador; o código da aplicação não envia a base para um servidor.
- A base e os registros em andamento desaparecem ao recarregar a página.
- Apenas resumos operacionais, tema e preferência de duplicidades são salvos localmente.
- O histórico não é compartilhado entre computadores e não é transferido do antigo endereço do ChatGPT para o novo.
- Não há sincronização automática com SOC, SharePoint ou Power Automate nem OCR por câmera.
- A matrícula precisa identificar uma pessoa. Conflitos precisam ser resolvidos na origem.
- Zeros ou precisão numérica já perdidos no Excel não são recuperáveis automaticamente.
- A restrição de acesso do Sites do ChatGPT não existe nesta versão independente. O site publicado no Pages será acessível conforme a configuração de acesso da nova hospedagem; repositório privado não protege o site publicado.
- Esta é a versão autoral de portfólio, sem marca Globo. A versão corporativa permanece separada.
- Nunca inclua o headcount real, credenciais ou arquivos com dados de funcionários no repositório.

## Desenvolvimento e testes

Requer Node.js 22.13 ou superior. Estes comandos são opcionais para desenvolvimento e não precisam ser executados no computador de trabalho ao usar a publicação pelo GitHub:

```sh
npm ci
npm run dev
npm test
```

`npm run build` verifica TypeScript e gera a pasta estática `dist`.
`npm test` executa o build e os testes de importação, conflitos, zeros, histórico, arquivo Excel e empacotamento independente. O teste de volume usa uma planilha sintética com 70.000 registros; o desempenho do arquivo real depende do conteúdo e do computador.

Na conferência pelo navegador, foram testados: consulta de três matrículas da demonstração, revisão dos cinco campos, geração do Excel, indicadores, filtro de temporalidade sem resultados e modo escuro.

O projeto usa React, TypeScript, Vite, ExcelJS e Lucide. Não há backend nem credenciais obrigatórias para o fluxo atual.
