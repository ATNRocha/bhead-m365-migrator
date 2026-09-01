# Changelog

## V0.4.5 — Formula Compatibility Engine

- Adicionado analisador de compatibilidade de fórmulas do XLSX exportado pelo Google Sheets.
- Detecta fórmulas matriciais exportadas como `<f t="array">`, inclusive matrizes de uma única célula.
- Converte `IFS(...)` para `IF(...)` aninhado no arquivo OOXML, aumentando a compatibilidade com Excel/Microsoft 365 e evitando dependência de `IFS` no cliente.
- Normaliza fórmulas matriciais de uma única célula quando o conteúdo é escalar (`IFS` convertido ou `XLOOKUP` escalar).
- Preserva fórmulas genuinamente matriciais de múltiplas células para revisão, sem tentar adivinhar sua semântica.
- Detecta `#REF!` já existente nas fórmulas do XLSX de origem e sinaliza `REVISAR`.
- Detecta erros visíveis no HTML exportado do Google, como `#REF!`, `#N/A`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NOME?` e `#NUM!`.
- Cruza valores em cache do XLSX com valores exibidos no HTML para apontar divergências de origem.
- Habilita recálculo completo do workbook no Microsoft 365 (`calcMode=auto`, `fullCalcOnLoad=1`, `forceFullCalc=1`).
- Mantido o fluxo obrigatório **ANALISAR → CONVERTER PARA M365**.
- Mantida a preservação do nome original do XLSX no download.
- Mantida a invalidação automática da análise quando XLSX, HTML/ZIP ou modo de migração são alterados.

## V0.4.4

- O arquivo XLSX convertido preserva o nome original enviado pelo usuário no modo híbrido.
- Removido o sufixo automático `_M365` do nome do arquivo de saída.
- No modo HTML/ZIP homologado, o nome do XLSX é derivado do pacote de origem e remove sufixos comuns como `(HTML)`.
- Mantido o fluxo obrigatório **ANALISAR → CONVERTER PARA M365**.
- Mantida a invalidação automática da análise quando os arquivos ou o modo de migração são alterados.

## V0.4.2

- Conversão bloqueada até análise válida dos arquivos atuais.
- Remoção do botão Limpar Sessão; limpeza automática do estado e temporários.
