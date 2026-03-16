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
- **Pasta `arquivos/`**: se a opção de cópia do spool estiver ativa, aqui ficam cópias dos arquivos de spool (`.spl`) de cada job; a coluna **Local_Arquivo** do Excel contém o caminho completo desses arquivos.

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
| Local_Arquivo | Caminho completo do arquivo de cópia do spool em `arquivos/` (ex.: `...\arquivos\42_Impressora_Doc.spl`) |

Fonte dos dados: API Windows de impressão (`win32print`: EnumPrinters, EnumJobs nível 2).

---

## Regras e comportamento

| Regra | Valor / Comportamento |
|-------|------------------------|
| Intervalo de verificação | 0,1 segundo entre cada varredura das filas (configurável via `INTERVALO_SEGUNDOS` em `app.py`) |
| Retenção de logs | 2 dias; arquivos `log_impressoes_*.xlsx` mais antigos são removidos automaticamente |
| Cache de jobs processados | Jobs são guardados em memória para não duplicar; entradas com mais de 24 h são removidas |
| Limpeza de logs antigos | Na inicialização e depois uma vez a cada 24 h |
| Duplicação | Evitada por chave única: `Nome da impressora` + `JobId` |
| Jobs com 0 páginas | Registrados assim mesmo (evita perder jobs que saem da fila antes de preencher TotalPages) |
| Impressoras inacessíveis | Geram apenas aviso no console; o monitor continua com as demais |
| Cópia do spool | Se ativada (`SALVAR_COPIA_SPL = True` em `app.py`), grava em `arquivos/` uma cópia do arquivo de spool (`.spl`) e dos metadados (`.shd`) de cada job; pode exigir execução como Administrador |
| Visualização do impresso | Se `CONVERTER_EMF_PARA_IMAGEM = True` (padrão), renderiza cada página do SPL para imagens `.bmp` ao lado do `.spl` em `arquivos/` (ex.: `42_Impressora_Doc_p001.bmp`); funciona para jobs EMF (maioria dos drivers Windows/GDI); jobs RAW/PostScript/PCL não são convertidos (o `.spl` ainda é salvo normalmente) |

---

## Solução de problemas

| Problema | Causa provável | Solução |
|----------|----------------|---------|
| "Arquivo Excel aberto por outro programa" | O arquivo do dia está aberto no Excel. | Feche o arquivo. |
| Nenhuma impressão no log | Sem permissão para ver as filas. | Execute o .exe como **Administrador** (ou conta com permissão nas impressoras). |
| Avisos de impressora inacessível | Impressora de rede offline. | Normal; o monitor continua. |
| Build falha (PyInstaller) | Python ou PATH incorreto. | Verifique `python --version` e use a mesma pasta do projeto ao rodar `build.bat`. |
| Imagens BMP não geradas | Job em formato RAW (PostScript/PCL) não suportado para conversão; ou SPL removido antes da captura; ou sem permissão de Administrador. | Normal para impressoras PS/PCL; o `.spl` ainda é copiado. Para desativar tentativas, defina `CONVERTER_EMF_PARA_IMAGEM = False` em `app.py`. |
| "a linha será gravada na próxima tentativa" | Arquivo Excel do dia estava aberto no momento da impressão. | Feche o Excel; o registro é salvo automaticamente no próximo ciclo. |
