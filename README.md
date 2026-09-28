# Demonstração — Emissão de NFS-e em Salvador

> Projeto de portfólio de **Victória Pedrosa**. **Demonstração** de emissão de NFS-e em Salvador — versão com dados fictícios (nomes, CNPJs, e-mails e IDs internos substituídos).

## Problema de negócio
A emissão de NFS-e em Salvador para muitos clientes era manual, repetitiva e sujeita a erro.

## Antes x depois
| | Antes | Depois |
|---|---|---|
| Como é feito | Equipe emitia nota por nota no portal da prefeitura. | Painel web (Apps Script) recebe as solicitações; o robô Python emite no portal, resolve o captcha e devolve o PDF ao Drive. |

## Ganho
- Emissão em lote, sem digitação no portal.

## Tecnologias
APIs REST, Gatilhos agendados, Gemini API, Google Apps Script, Google Drive, Google Sheets, HTML/JavaScript, OCR de captcha, Python, SQLite, Selenium, Web App (HtmlService), pandas

## Arquivos
- `Codigo_AppsScript.gs`
- `Index_AppsScript.html`
- `exemplo_nfse.html`
- `prefeituradesalvador.py`
- `requirements.txt`
- `robo_emissor.py`
- `tarefa.xml`
- `versaoapsheet.py`

## Como rodar
1. `pip install -r requirements.txt`
2. Copie `.env.exemplo` para `.env` e preencha os caminhos.
3. Execute o script principal.

## Autora
Victória Pedrosa — Product Owner do Time de IA, automação de processos contábeis e fiscais.
