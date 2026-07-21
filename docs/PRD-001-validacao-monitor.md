# PRD-001 — Validação e conformidade do Monitor de Impressões

> Status: concluido

## Objetivo

Confirmar conformidade do monitor com README/instalacao.md (filas → Excel + `arquivos/` + preview EMF) e corrigir divergências, com evidências na máquina de desenvolvimento Windows.

## Registro

| ID | Tipo | Descrição | Status | Check / evidência | Depende de |
|----|------|-----------|--------|-------------------|------------|
| L-001 | Lacuna | Ambiente dev local; escopo corrigir tudo | concluido | Decisão usuário | Nenhuma |
| DEC-001 | Decisão | `Tamanho_Bytes` int; `Tamanho_Legivel`; `Local_Arquivo`=.spl; `Local_Preview`=BMPs | concluido | app.py CABECALHO | L-001 |
| RF-001 | RF | EnumPrinters/EnumJobs nível 2 | concluido | app.py | Nenhuma |
| RF-002 | RF | Chave única job, fila pendente Excel | concluido | app.py | RF-001 |
| RF-003 | RF | Excel diário, retenção 2 dias | concluido | app.py | RF-002 |
| RF-004 | RF | Cópia spool | concluido | app.py | RF-001 |
| RF-005 | RF | EMF → BMP | concluido | app.py | RF-004 |
| RF-006 | RF | erro.log | concluido | app.py | Nenhuma |
| AC-001 | AC | Linha Excel com usuário/documento | parcial | EnumPrinters OK; impressão real requer job na fila (ver T-002) | RF-001 |
| AC-002 | AC | .spl e .bmp com admin | bloqueado | BLOQ-001: spool inacessível sem admin neste host | RF-004 |
| AC-003 | AC | Colunas pós-DEC-001 | concluido | tests/test_excel_schema.py | DEC-001 |
| AC-004 | AC | pytest + compileall | concluido | 13 passed, compileall OK | T-004 |
| AC-005 | AC | build exe | concluido | dist/MonitorImpressoes.exe + smoke 5s | T-005 |
| BLOQ-001 | Bloqueio | Spool `System32\spool\PRINTERS` negado sem elevação | aberto | `_verificar_acesso_spool()` → False; AC-002 após rodar como Admin | — |
| T-001 | T | Este PRD | concluido | docs/PRD-001 | Nenhuma |
| T-002 | T | Testes manuais | concluido | Diagnóstico abaixo | T-001 |
| T-003 | T | DEC-001 código | concluido | app.py | T-002 |
| T-004 | T | pytest | concluido | pytest -q | T-001 |
| T-005 | T | build smoke | concluido | build.ps1 | T-001 |
| T-006 | T | Docs + PII | concluido | README, instalacao.md | T-003 |

## Histórico de implementação

| Data | ID | Solução (resumo) | Autor | Evidência |
|------|-----|------------------|-------|-----------|
| 2026-07-21 | T-001 | PRD-001 criado | agente | docs/PRD-001-validacao-monitor.md |
| 2026-07-21 | DEC-001 / T-003 | Colunas Excel alinhadas; linha grava bytes + legível + spl + previews | agente | app.py, test_excel_schema.py |
| 2026-07-21 | T-004 | tests/ + requirements-dev.txt; 13 testes | agente | pytest -q |
| 2026-07-21 | T-003 | `_verificar_acesso_spool` não propaga pywintypes.error | agente | exe smoke pós-rebuild |
| 2026-07-21 | T-005 | build.ps1 OK | agente | dist/MonitorImpressoes.exe |
| 2026-07-21 | T-006 | README (colunas, PII, pytest); instalacao colunas | agente | diff docs |
| 2026-07-21 | T-002 | 1 impressora (Microsoft Print to PDF); spool sem admin | agente | script diagnóstico |

## Evidências (checks)

### T-002 — Matriz manual (dev local)

- [x] Diagnóstico: **1** impressora (`Microsoft Print to PDF`)
- [x] Spool: **inacessível** sem admin (`Spool acessivel: False`) → BLOQ-001
- [ ] Impressão teste → linha Excel (não automatizada: `Out-Printer` para PDF não gerou log no ciclo)
- [ ] Admin → `.spl`/`.bmp` (pendente execução elevada pelo operador)
- [ ] Excel aberto → pendente (pendente teste manual)

**Recomendação operador:** executar `MonitorImpressoes.exe` **como Administrador**, imprimir uma página de teste e validar AC-001/AC-002.

### T-004

```
pytest -q  → 13 passed
```

### T-005

```
python -m compileall -q app.py
.\build.ps1  → dist\MonitorImpressoes.exe
Smoke: processo ativo 5s sem crash (pós-correção spool)
```
