# Instalação e uso no Windows Server (servidor de impressão)

Como configurar e usar o **Monitor de Impressões** no **Windows Server** que gerencia as impressoras (Print Server), usando **apenas o .exe** (sem instalar Python no servidor).

---

## Instalação rápida

1. **Gere o .exe** na sua máquina de desenvolvimento (duplo clique em **`build.bat`** na pasta do projeto). O arquivo estará em **`dist\MonitorImpressoes.exe`**.

2. **Copie só o arquivo** `MonitorImpressoes.exe` para o servidor (ex.: `C:\MonitorImpressao\MonitorImpressoes.exe`).

3. **No servidor:** duplo clique no .exe ou execute pelo prompt. Não é preciso instalar Python.

4. O Excel do dia fica **na raiz**: **`log_impressoes_DDMMYYYY.xlsx`** (ex.: `log_impressoes_20022026.xlsx`). A pasta **`arquivos/`** também é criada na raiz.

---

## 1. Pré-requisitos no servidor

| Item | Detalhe |
|------|--------|
| Sistema | Windows Server (com função de impressão) |
| Função | Servidor com impressoras compartilhadas (Print Server) |
| Python | **Não é necessário** — use apenas o .exe |
| Permissões | Executar o .exe com usuário que possa listar jobs das filas (ex.: Administrador) |

---

## 2. Onde colocar o .exe

Recomendação: pasta dedicada, por exemplo:

```
C:\MonitorImpressao\
  MonitorImpressoes.exe
```

O arquivo do dia (**log_impressoes_DDMMYYYY.xlsx**) é criado na mesma pasta do .exe. A pasta **arquivos/** é criada automaticamente na raiz.

---

## 3. Como utilizar no dia a dia

### 3.1 Execução manual

Duplo clique em **MonitorImpressoes.exe** ou, no prompt:

```powershell
cd C:\MonitorImpressao
.\MonitorImpressoes.exe
```

Para parar: **Ctrl+C**.

Não deixe o arquivo Excel do dia aberto no Excel enquanto o monitor estiver rodando.

### 3.2 Execução em segundo plano (Tarefa agendada)

Para o monitor iniciar com o servidor:

1. Abra o **Agendador de Tarefas** (Taskschd.msc).
2. Criar Tarefa…
3. **Gatilho:** "Ao iniciar" ou "Quando o usuário fizer logon" (conta com permissão nas impressoras).
4. **Ação:** Iniciar um programa.
   - **Programa/script:** `C:\MonitorImpressao\MonitorImpressoes.exe`
   - **Argumentos:** (deixe em branco)
   - **Iniciar em:** `C:\MonitorImpressao`
5. Em Configurações: marque "Executar tarefa o mais rápido possível após uma inicialização agendada ser perdida" e "Se a tarefa falhar, reiniciar a cada...".

---

## 4. Onde ficam os logs

| O que | Onde |
|-------|------|
| Arquivo do dia | Na raiz: `C:\MonitorImpressao\log_impressoes_DDMMYYYY.xlsx` (ex.: `log_impressoes_20022026.xlsx`) |
| Pasta | `C:\MonitorImpressao\arquivos\` (criada automaticamente) |
| Aba | **Impressões** |
| Colunas | ID_Job, Usuario, Data_Hora, Arquivo, Paginas, Impressora, Tamanho_Bytes, IP, Local_Arquivo (detalhes no README) |
| Retenção | Até **2 dias**; arquivos mais antigos são removidos automaticamente |

---

## 5. Resumo do fluxo no Windows Server

1. **Servidor** = Print Server (impressoras instaladas/compartilhadas).
2. **MonitorImpressoes.exe** roda no servidor e usa a API do Windows para listar impressoras e jobs.
3. Cada novo job é registrado em uma linha do Excel do dia.
4. Logs antigos (> 2 dias) são removidos automaticamente.

---

## 6. Problemas comuns

| Problema | Causa provável | O que fazer |
|----------|----------------|-------------|
| "Arquivo Excel aberto por outro programa" | O `.xlsx` do dia está aberto. | Fechar o arquivo. |
| Nenhuma impressão no log | Sem permissão nas filas. | Executar o .exe como **Administrador**. |
| Avisos "[Aviso] Impressora 'X': ..." | Impressora de rede inacessível. | Normal; o monitor continua. |
| .exe para sozinho | Disco cheio, permissão, etc. | Ver mensagem no console; usar Tarefa Agendada com "reiniciar se falhar". |
