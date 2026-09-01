# BHead M365 Migrator Web V0.4.5

Versão Web do conversor Google Sheets → Microsoft 365 Online, com **Formula Compatibility Engine**.

## Objetivo

Executar o motor do BHead M365 Migrator em um host e permitir que os usuários utilizem o sistema pelo navegador, preservando estrutura, formatação e fórmulas na migração para Excel Online / Microsoft 365.

## Modos

- **NOVO MODELO — XLSX + HTML/ZIP**: recomendado para planilhas novas. O XLSX é a fonte estrutural (fórmulas, validações, referências) e o HTML é a fonte visual e de valores de referência do Google.
- **MODELO HOMOLOGADO — somente HTML/ZIP**: fluxo rápido para modelos já conhecidos pelo perfil.

## Formula Compatibility Engine

No modo híbrido, a V0.4.5:

- analisa fórmulas diretamente no OOXML do XLSX;
- identifica fórmulas matriciais exportadas pelo Google;
- converte `IFS(...)` para `IF(...)` aninhado;
- normaliza fórmulas matriciais de uma única célula quando reconhecidas como escalares;
- detecta `#REF!` e outros erros já existentes na origem;
- utiliza os valores exibidos no HTML como referência adicional de diagnóstico;
- marca o workbook para recálculo completo ao abrir no Microsoft 365.

O sistema não inventa referências perdidas. Quando o arquivo de origem contém `#REF!`, a análise retorna **REVISAR**, mas a conversão pode prosseguir para que os demais recursos compatíveis sejam normalizados.

## Fluxo obrigatório

1. Selecione XLSX + HTML/ZIP no modo Novo Modelo, ou HTML/ZIP no modo Homologado.
2. Clique em **ANALISAR**.
3. O botão **CONVERTER PARA M365** só é habilitado depois de uma análise válida dos arquivos atuais.
4. Alterar XLSX, HTML/ZIP ou modo invalida automaticamente a análise anterior.
5. `INCOMPATÍVEL` bloqueia a conversão; `REVISAR` permite converter com ressalvas exibidas/registradas.

## Nome do arquivo convertido

O modo híbrido preserva o nome original do XLSX.

Exemplo:

- Entrada: `planilha_exemplo.xlsx`
- Saída: `planilha_exemplo.xlsx`

O sufixo `_M365` não é acrescentado.

## Gerenciamento da sessão

Não existe botão manual de limpeza. Ao trocar entradas ou modo, a análise e os downloads anteriores são invalidados e os temporários da migração anterior são removidos. Temporários antigos são eliminados automaticamente após 6 horas em execuções posteriores.

## Dependências

- Python 3.11+
- Streamlit 1.x
- XlsxWriter 3.x
- Biblioteca padrão do Python

## Linux / Ubuntu

Exemplo de execução:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m streamlit run app_web.py --server.address=127.0.0.1 --server.port=8501 --server.headless=true
```

Em produção, recomenda-se executar via `systemd`, reverse proxy/tunnel e manter a porta Streamlit restrita a localhost.

## Destino

Arquivos `.xlsx` gerados para homologação e uso em Excel Online / OneDrive / SharePoint / Microsoft 365.
