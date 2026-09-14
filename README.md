# Arquivo Ágil

Aplicação autoral criada por **Jean Lucca Dias** para acelerar o planilhamento de documentos físicos. O sistema importa uma planilha exportada do SOC, permite localizar funcionários pela matrícula e gera um arquivo `.xlsx` pronto para envio e armazenamento.

## Demonstração

[Acesse a versão publicada do Arquivo Ágil](https://arquivo-agil-prototipo.jean230jls.chatgpt.site)

## Funcionalidades

- Importação local da base de funcionários em `.xlsx`, `.xls` ou `.csv`
- Identificação automática das colunas de matrícula, nome e local de trabalho
- Busca real por matrícula, sem funcionários fictícios no código
- Organização por número da etiqueta, data e setor Saúde Ocupacional
- Conferência e remoção de registros antes da exportação
- Excel com tabela nativa, cabeçalho, filtros, faixas alternadas e colunas ajustadas
- Histórico de caixas processadas sem armazenar nomes ou matrículas
- Relatórios de produtividade e documentos por período
- Consulta da base carregada e configurações operacionais
- Prevenção opcional de matrículas duplicadas
- Interface responsiva com modos claro e escuro

## Privacidade

A planilha SOC é processada somente no navegador. A aplicação não envia essa base a um servidor e não a mantém após recarregar a página. Por isso, a base deve ser importada novamente a cada nova sessão. Não publique planilhas reais ou dados pessoais no GitHub.

O histórico salvo no navegador contém apenas etiqueta, data, quantidade de documentos e horário de exportação.

## Arquivo exportado

A planilha final mantém os cinco campos do processo:

1. Número da etiqueta
2. Nome do funcionário
3. Matrícula
4. Setor
5. Data

## Tecnologias

- React 19 e TypeScript
- Vinext/Next.js
- ExcelJS
- Lucide Icons
- CSS responsivo com identidade visual autoral

## Executar localmente

Requisitos: Node.js 22.13 ou superior.

```bash
npm install
npm run dev
```

Depois, acesse o endereço exibido no terminal e importe uma planilha exportada do SOC para iniciar os testes.

## Autor

Conceito, produto, UX e desenvolvimento por **Jean Lucca Dias**, como projeto de portfólio.
