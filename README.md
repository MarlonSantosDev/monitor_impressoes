# Monitor de Impressões Windows

Monitora as filas de impressão do Windows e registra cada job em um Excel do dia (usuário, documento, data/hora, páginas, impressora, etc.). Distribuição via **um único .exe** — na máquina de uso **não é preciso instalar Python**.

**Plataforma:** somente **Windows**. Não funciona em Linux/macOS.

| Item | Descrição |
|------|------------|
| Entrada | Filas de impressão (API Windows: `win32print`) |
| Saída | `log_impressoes_DDMMYYYY.xlsx` na pasta do .exe + pasta `arquivos/` |
| Dependências (runtime) | Nenhuma na máquina destino; o .exe é autocontido |
| Dependências (build) | Python 3.8+, `requirements-build.txt` (pywin32, openpyxl, PyInstaller) |

---

## Como criar o .exe

Na pasta do projeto (com `app.py`, `requirements.txt`, `requirements-build.txt`, `build.bat` e `build.ps1`):

1. Tenha **Python 3.8+** instalado no PATH ([python.org](https://www.python.org/downloads/)).
2. **Duplo clique em `build.bat`** (ou no PowerShell: `.\build.ps1`).
3. Aguarde o build. O executável será gerado em **`dist\MonitorImpressoes.exe`**.

O script `build.ps1` cria/usa um ambiente virtual, instala as dependências (incluindo PyInstaller) e gera o .exe. Não é preciso rodar nenhum outro instalador antes.

---

## Como usar o .exe

1. Copie **`dist\MonitorImpressoes.exe`** para o PC ou servidor onde quer rodar o monitor.
2. Execute o .exe (duplo clique ou pelo prompt).
3. O Excel do dia fica **na raiz** (mesma pasta do .exe): **`log_impressoes_DDMMYYYY.xlsx`** (ex.: `log_impressoes_20022026.xlsx`). A pasta **`arquivos/`** também é criada na raiz.
4. Para parar: **Ctrl+C** na janela do monitor.

Não é necessário instalar Python na máquina onde o .exe roda.

---

## Arquivos do projeto (só para gerar o .exe)

| Arquivo | Uso |
|---------|-----|
| `app.py` | Código do monitor. |
| `requirements.txt` | Dependências do app (`pywin32`, `openpyxl`). |
| `requirements-build.txt` | Dependências do app + PyInstaller (para o build). |
| `build.bat` | Entrada: duplo clique para gerar o .exe (chama `build.ps1`). |
| `build.ps1` | Cria venv (se não existir), instala deps e roda PyInstaller. |

Detalhes para **Windows Server** (copiar e rodar o .exe, tarefa agendada): **`instalacao.md`**.

---

## Requisitos (só na máquina onde você *gera* o .exe)

| Item | Requisito |
|------|-----------|
| SO | Windows |
| Python | 3.8+ (para rodar o build) |
| Dependências | Instaladas automaticamente pelo `build.ps1` via `requirements-build.txt` |

---

## Estrutura (após rodar o .exe)

O Excel do dia fica **na raiz** (mesma pasta do .exe). A pasta **arquivos/** é criada na raiz.

```
<pasta do exe>/
  MonitorImpressoes.exe
  log_impressoes_20022026.xlsx     # um arquivo por dia (DDMMYYYY)
  log_impressoes_19022026.xlsx
  arquivos/
```

- **Um arquivo por dia** na raiz: `log_impressoes_DDMMYYYY.xlsx`; retenção de **2 dias** (arquivos mais antigos são removidos automaticamente).
- Aba: **Impressões**.

Enquanto o monitor estiver rodando, evite deixar o arquivo do dia aberto no Excel para não dar erro de permissão.

---

## Dados coletados (colunas do Excel)

Cada linha registra um job de impressão com as colunas abaixo (ordem fixa):

| Coluna | Descrição |
|--------|-----------|
| ID_Job | ID do job na fila do Windows |
| Usuario | Usuário que enviou o job (pUserName) |
| Data_Hora | Data/hora de submissão do job (YYYY-MM-DD HH:MM:SS) |
| Arquivo | Nome do documento (pDocument) |
| Paginas | Total de páginas (TotalPages) |
| Impressora | Nome da impressora |
| Tamanho_Bytes | Tamanho do job em bytes (Size) |
| IP | IP da máquina onde o monitor está rodando |
| Local_Arquivo | Caminho completo do arquivo .xlsx (para referência) |

Fonte dos dados: API Windows de impressão (`win32print`: EnumPrinters, EnumJobs nível 2).

---

## Regras e comportamento

| Regra | Valor / Comportamento |
|-------|------------------------|
| Intervalo de verificação | 2 segundos entre cada varredura das filas |
| Retenção de logs | 2 dias; arquivos `log_impressoes_*.xlsx` mais antigos são removidos automaticamente |
| Cache de jobs processados | Jobs são guardados em memória para não duplicar; entradas com mais de 24 h são removidas |
| Limpeza de logs antigos | Na inicialização e depois uma vez a cada 24 h |
| Duplicação | Evitada por chave única: `Nome da impressora` + `JobId` |
| Jobs em spooling | Se páginas = 0 e status = spooling, o job é ignorado nesse ciclo (tentado no próximo) |
| Impressoras inacessíveis | Geram apenas aviso no console; o monitor continua com as demais |

---

## Solução de problemas

| Problema | Causa provável | Solução |
|----------|----------------|---------|
| "Arquivo Excel aberto por outro programa" | O arquivo do dia está aberto no Excel. | Feche o arquivo. |
| Nenhuma impressão no log | Sem permissão para ver as filas. | Execute o .exe como **Administrador** (ou conta com permissão nas impressoras). |
| Avisos de impressora inacessível | Impressora de rede offline. | Normal; o monitor continua. |
| Build falha (PyInstaller) | Python ou PATH incorreto. | Verifique `python --version` e use a mesma pasta do projeto ao rodar `build.bat`. |
