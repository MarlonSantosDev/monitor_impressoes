"""
Monitor de filas de impressão Windows.

Registra cada job em um Excel do dia na raiz (pasta onde o .exe roda):
  log_impressoes_DDMMYYYY.xlsx  (ex.: log_impressoes_20022026.xlsx)
Cria a pasta arquivos/ e, se SALVAR_COPIA_SPL = True, grava lá uma cópia do
arquivo de spool (.SPL) do job — não é o documento original (PDF/DOC), e sim
o formato que o Windows envia para a impressora. Pode exigir permissões de admin.
Requer Windows (pywin32) e openpyxl.
Erros e falhas de execução são registrados em erro.log na mesma pasta do .exe.
Se não coletar dados nos testes, defina DEBUG = True para ver impressoras e jobs na fila.
"""
import logging
import os
import re
import shutil
import sys
import threading
import time
import traceback
from datetime import datetime, timedelta

import win32con
import win32file
import win32print
import win32timezone  # Usado por pywin32 ao acessar job['Submitted']; necessário no .exe
from openpyxl import Workbook, load_workbook

# --- CONFIGURAÇÃO ---
# Base: pasta onde o .exe é executado (ou do script)
if getattr(sys, "frozen", False):
    PASTA_SCRIPT = os.path.dirname(sys.executable)
else:
    PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
# Saída: Excel do dia na raiz (log_impressoes_DDMMYYYY.xlsx) + pasta arquivos/
PASTA_ARQUIVOS = os.path.join(PASTA_SCRIPT, "arquivos")
ARQUIVO_LOG_ERRO = os.path.join(PASTA_SCRIPT, "erro.log")
NOME_ABA = "Impressões"
CABECALHO = [
    "Usuario", "Data_Hora", "Arquivo", "Paginas", "Impressora", "Tamanho", "Local_Arquivo"
]
DIAS_RETENCAO = 2       # Remover log_impressoes_*.xlsx com mais de 2 dias
INTERVALO_SEGUNDOS = 0.5  # Varredura das filas (menor = mais chance de pegar jobs rápidos)
CACHE_JOB_HORAS = 24    # Limpar do cache jobs processados há mais de 24 h
DEBUG = False           # True: mostra quantas impressoras/jobs por ciclo
SALVAR_COPIA_SPL = True  # Copiar arquivo de spool do job para pasta arquivos/ (pode exigir admin)

# Definido na inicialização: True se o processo tem acesso à pasta System32\spool\PRINTERS
_spool_acessivel: bool | None = None


def _verificar_acesso_spool() -> bool:
    """
    Verifica se o processo tem acesso à pasta do spool (leitura).
    Retorna True se conseguir abrir/listar; False caso contrário.
    Evita tentativas e avisos repetidos quando não há permissão de admin.
    """
    spool_dir = os.path.join(
        os.environ.get("SystemRoot", "C:\\Windows"), "System32", "spool", "PRINTERS"
    )
    try:
        h = win32file.CreateFile(
            spool_dir,
            win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_FLAG_BACKUP_SEMANTICS,
            None,
        )
        win32file.CloseHandle(h)
        return True
    except OSError:
        return False


def configurar_log_erro():
    """Configura o módulo logging para gravar em erro.log (append) com data/hora e ponto do código."""
    log = logging.getLogger("monitor_impressoes")
    log.setLevel(logging.DEBUG)
    if log.handlers:
        return log
    try:
        fh = logging.FileHandler(ARQUIVO_LOG_ERRO, mode="a", encoding="utf-8")
    except OSError:
        return log
    fh.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh.setFormatter(fmt)
    log.addHandler(fh)
    return log


def log_erro(ponto: str, mensagem: str, exc: BaseException | None = None):
    """Registra em erro.log: ponto do código, mensagem e, se houver, exceção com traceback."""
    log = configurar_log_erro()
    texto = f"[{ponto}] {mensagem}"
    if exc is not None:
        log.error("%s | %s: %s\n%s", texto, type(exc).__name__, exc, traceback.format_exc())
    else:
        log.error(texto)


def log_aviso(ponto: str, mensagem: str):
    """Registra aviso em erro.log (falha de execução sem exceção)."""
    configurar_log_erro().warning("[%s] %s", ponto, mensagem)


def _bytes_para_kb_mb_gb(size_bytes: int) -> str:
    """Converte tamanho em bytes para string legível em KB, MB ou GB."""
    try:
        n = int(size_bytes)
    except (TypeError, ValueError):
        return "0 KB"
    if n < 0:
        return "0 KB"
    if n < 1024:
        return f"{n / 1024:.2f} KB"
    if n < 1024 * 1024:
        return f"{n / 1024:.2f} KB"
    if n < 1024**3:
        return f"{n / (1024**2):.2f} MB"
    return f"{n / (1024**3):.2f} GB"


def _sanitizar_nome_arquivo(nome: str, max_len: int = 80) -> str:
    """Remove caracteres inválidos para nome de arquivo Windows."""
    invalidos = r'<>:"/\\|?*'
    for c in invalidos:
        nome = nome.replace(c, "_")
    nome = nome.strip(". ") or "documento"
    return nome[:max_len] if len(nome) > max_len else nome


def caminho_arquivo_copia(job_id: int, nome_documento: str, impressora: str) -> str:
    """
    Retorna o caminho completo (local do arquivo) onde a cópia do spool será salva em arquivos/.
    Esse valor deve ser usado na coluna Local_Arquivo do Excel.
    """
    base = _sanitizar_nome_arquivo(nome_documento)
    impressora_safe = _sanitizar_nome_arquivo(impressora, 40)
    dest_nome = f"{job_id}_{impressora_safe}_{base}.spl"
    return os.path.join(PASTA_ARQUIVOS, dest_nome)


def copiar_spool_para_arquivos(job_id: int, dest_path: str, max_tentativas: int = 5) -> bool:
    """
    Tenta copiar o arquivo de spool do job (System32\\spool\\PRINTERS) para dest_path.
    Faz até max_tentativas com pequenas pausas (o .SPL pode aparecer um pouco depois do job).
    Retorna True se copiou, False se falhou (ex.: sem permissão, arquivo já removido).
    O arquivo é .SPL (formato spool Windows), não o documento original.
    """
    spool_dir = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "System32", "spool", "PRINTERS")
    nome_spl = f"{job_id:05d}.SPL"
    src = os.path.join(spool_dir, nome_spl)
    for _ in range(max_tentativas):
        if os.path.isfile(src):
            try:
                shutil.copy2(src, dest_path)
                return True
            except (OSError, PermissionError):
                pass
        time.sleep(0.25)
    log_aviso(
        "copiar_spool_para_arquivos",
        f"Falha após {max_tentativas} tentativas: job_id={job_id} src={src} dest={dest_path}",
    )
    return False


def caminho_log_do_dia(data=None):
    """Retorna o caminho do Excel do dia na raiz: log_impressoes_DDMMYYYY.xlsx"""
    if data is None:
        data = datetime.now()
    if isinstance(data, datetime):
        data_str = data.strftime("%d%m%Y")  # ex.: 20022026
    else:
        data_str = datetime.strptime(str(data), "%Y-%m-%d").strftime("%d%m%Y")
    nome = f"log_impressoes_{data_str}.xlsx"
    return os.path.join(PASTA_SCRIPT, nome)


def iniciar_log():
    """Cria a pasta arquivos/ na raiz e executa limpeza de logs antigos."""
    global _spool_acessivel
    _spool_acessivel = _verificar_acesso_spool()
    os.makedirs(PASTA_ARQUIVOS, exist_ok=True)
    print(f"Log do dia na raiz: log_impressoes_DDMMYYYY.xlsx")
    print(f"Pasta criada: arquivos/")
    if SALVAR_COPIA_SPL and not _spool_acessivel:
        print("[Aviso] Cópia do spool desativada (sem permissão em System32\\spool\\PRINTERS). Execute como Administrador para salvar cópias.")
    limpar_logs_antigos()


def limpar_logs_antigos():
    """Remove arquivos log_impressoes_DDMMYYYY.xlsx da raiz com mais de DIAS_RETENCAO dias."""
    limite = datetime.now() - timedelta(days=DIAS_RETENCAO)
    padrao = re.compile(r"^log_impressoes_(\d{8})\.xlsx$")  # DDMMYYYY
    try:
        for nome in os.listdir(PASTA_SCRIPT):
            m = padrao.match(nome)
            if not m:
                continue
            try:
                d = datetime.strptime(m.group(1), "%d%m%Y")
                if d < limite:
                    path = os.path.join(PASTA_SCRIPT, nome)
                    os.remove(path)
                    print(f"Removido log antigo (>{DIAS_RETENCAO} dias): {nome}")
            except (ValueError, OSError):
                pass
    except OSError:
        pass


def watch_spool_directory(spool_copies: dict):
    """
    Thread daemon: monitora C:\\Windows\\System32\\spool\\PRINTERS\\ via ReadDirectoryChangesW.
    Ao detectar a criação de um .SPL, copia imediatamente para arquivos/_temp_{job_id}.spl
    e registra em spool_copies[job_id] = dest_path.
    Requer permissões de administrador.
    Nota: o .SPL é formato Windows EMF/RAW, não o documento original (PDF/DOCX).
    """
    spool_dir = os.path.join(
        os.environ.get("SystemRoot", "C:\\Windows"), "System32", "spool", "PRINTERS"
    )
    try:
        h_dir = win32file.CreateFile(
            spool_dir,
            win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_FLAG_BACKUP_SEMANTICS,
            None,
        )
    except OSError as e:
        log_erro("watch_spool_directory.inicio", f"Não foi possível abrir pasta do spool: {spool_dir}", e)
        print(f"[Aviso] Watcher de spool não iniciado: {e}")
        return

    print(f"[Spool] Assistindo: {spool_dir}")
    while True:
        try:
            results = win32file.ReadDirectoryChangesW(
                h_dir,
                1024,
                False,  # não recursivo
                win32con.FILE_NOTIFY_CHANGE_FILE_NAME,
                None,
                None,
            )
            for action, filename in results:
                if action != 1:  # 1 = FILE_ACTION_ADDED
                    continue
                if not filename.upper().endswith(".SPL"):
                    continue
                src = os.path.join(spool_dir, filename)
                try:
                    job_id = int(os.path.splitext(filename)[0])
                except ValueError:
                    continue
                os.makedirs(PASTA_ARQUIVOS, exist_ok=True)
                dest = os.path.join(PASTA_ARQUIVOS, f"_temp_{job_id}.spl")
                # Retenta a cópia: o .SPL pode estar sendo escrito e bloqueado no primeiro instante
                copiou = False
                for _ in range(5):
                    try:
                        time.sleep(0.15)
                        shutil.copy2(src, dest)
                        spool_copies[job_id] = dest
                        copiou = True
                        if DEBUG:
                            print(f"[Spool] Capturado: {filename} → {dest}")
                        break
                    except (OSError, PermissionError):
                        pass
                if not copiou:
                    log_aviso("watch_spool_directory.copia", f"Cópia do spool falhou após 5 tentativas: job_id={job_id} arquivo={filename}")
        except Exception as e:
            log_erro("watch_spool_directory.loop", "Erro ao processar alterações do spool", e)
            if DEBUG:
                print(f"[Spool] Erro no watcher: {e}")
            time.sleep(1)


def monitorar_impressoes():
    global _spool_acessivel
    if _spool_acessivel is None:
        _spool_acessivel = _verificar_acesso_spool()
    print(f"Monitorando impressões... (Pressione Ctrl+C para parar)")

    jobs_processados = {}
    ultima_limpeza = datetime.now()

    # Dict compartilhado: job_id → caminho do SPL copiado pela thread watcher (só inicia se houver acesso ao spool)
    spool_copies: dict = {}
    if SALVAR_COPIA_SPL and _spool_acessivel:
        t = threading.Thread(target=watch_spool_directory, args=(spool_copies,), daemon=True)
        t.start()

    # Local + conexões de rede (documentação Windows: cobre impressoras disponíveis)
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    primeiro_ciclo = True

    while True:
        try:
            printers = win32print.EnumPrinters(flags)
            if DEBUG or primeiro_ciclo:
                print(f"[Diagnóstico] Impressoras encontradas: {len(printers)}")
                for p in printers:
                    print(f"  - {p[2]}")
                primeiro_ciclo = False

            total_jobs_ciclo = 0
            for printer in printers:
                printer_name = printer[2]
                p_handle = None

                try:
                    p_handle = win32print.OpenPrinter(printer_name)

                    # Nível 2 traz detalhes do dono e nome do documento
                    jobs = win32print.EnumJobs(p_handle, 0, -1, 2)
                    total_jobs_ciclo += len(jobs)
                    if DEBUG and jobs:
                        print(f"[Diagnóstico] Impressora '{printer_name}': {len(jobs)} job(s) na fila")

                    for job in jobs:
                        job_id = job['JobId']
                        # Chave única: Nome da Impressora + ID do Job
                        job_unique_key = f"{printer_name}_{job_id}"
                        
                        # Se já processamos este job, ignora
                        if job_unique_key in jobs_processados:
                            continue

                        # --- EXTRAÇÃO DE DADOS ---
                        usuario = job.get('pUserName', 'Sistema/Desconhecido')
                        # Nome do arquivo impresso: pywin32 pode usar 'pDocument' ou 'Document'
                        documento = (
                            job.get('pDocument') or job.get('Document') or job.get('pDocName')
                            or ''
                        )
                        if isinstance(documento, str):
                            documento = documento.strip()
                        if not documento:
                            documento = f"Sem Nome (job {job_id})"
                        else:
                            documento = str(documento)
                        paginas = job.get('TotalPages', 0)
                        tamanho = job.get('Size', 0)  # Tamanho em bytes
                        
                        # Tenta pegar a data de submissão original do job
                        try:
                            data_raw = job.get('Submitted')
                            if data_raw is None:
                                raise ValueError("Submitted ausente")
                            data_hora = f"{data_raw.year}-{data_raw.month:02d}-{data_raw.day:02d} {data_raw.hour:02d}:{data_raw.minute:02d}:{data_raw.second:02d}"
                        except (AttributeError, TypeError, ValueError):
                            data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        # Não filtrar por 0 páginas/spooling: muitos drivers mantêm 0 até o job
                        # terminar; ao terminar o job sai da fila e perdemos o registro. Melhor
                        # registrar mesmo com 0 páginas/bytes do que não registrar.

                        # Local do arquivo: watcher já copiou o SPL; fallback para cópia direta (só se tiver acesso ao spool)
                        local_arquivo = ""
                        if SALVAR_COPIA_SPL and _spool_acessivel:
                            os.makedirs(PASTA_ARQUIVOS, exist_ok=True)  # garante pasta se foi removida
                            dest_path = caminho_arquivo_copia(job_id, documento, printer_name)
                            if job_id in spool_copies:
                                temp = spool_copies.pop(job_id)
                                try:
                                    if os.path.isfile(temp):
                                        os.rename(temp, dest_path)
                                        local_arquivo = dest_path
                                except OSError:
                                    if os.path.isfile(temp):
                                        local_arquivo = temp  # fallback: registra o _temp_*.spl no Excel
                            if not local_arquivo:  # fallback: tenta cópia direta com retries
                                if copiar_spool_para_arquivos(job_id, dest_path):
                                    local_arquivo = dest_path
                            if not local_arquivo:
                                log_aviso(
                                    "monitorar_impressoes.copia_spool",
                                    f"Cópia do spool não salva: job_id={job_id} documento={documento!r} impressora={printer_name!r}",
                                )
                                print(f"      [Aviso] Cópia do spool não salva (arquivo não encontrado ou sem permissão em System32\\spool\\PRINTERS)")

                        # --- SALVAR NO EXCEL (raiz: log_impressoes_DDMMYYYY.xlsx) ---
                        arquivo_hoje = caminho_log_do_dia()
                        try:
                            if not os.path.exists(arquivo_hoje):
                                wb = Workbook()
                                ws = wb.active
                                ws.title = NOME_ABA
                                ws.append(CABECALHO)
                            else:
                                wb = load_workbook(arquivo_hoje)
                                if NOME_ABA in wb.sheetnames:
                                    ws = wb[NOME_ABA]
                                else:
                                    ws = wb.create_sheet(NOME_ABA)
                                    ws.append(CABECALHO)
                            ws.append([
                                usuario, data_hora, documento, paginas, printer_name,
                                _bytes_para_kb_mb_gb(tamanho),
                                local_arquivo
                            ])
                            wb.save(arquivo_hoje)

                            print(f"[NOVO] {data_hora} | {usuario} | {documento} ({paginas} pgs)")
                            if SALVAR_COPIA_SPL and os.path.isfile(local_arquivo):
                                print(f"      Cópia spool salva: {local_arquivo}")

                            jobs_processados[job_unique_key] = time.time()

                        except PermissionError as e_perm:
                            log_erro("monitorar_impressoes.excel", f"Arquivo Excel aberto por outro programa: {arquivo_hoje}", e_perm)
                            print("Erro: Arquivo Excel aberto por outro programa. Feche o arquivo e tente novamente.")
                        except Exception as e_excel:
                            log_erro("monitorar_impressoes.excel", f"Falha ao salvar/abrir Excel: {arquivo_hoje}", e_excel)
                            print(f"Erro ao salvar Excel ({type(e_excel).__name__}): {e_excel}")

                except OSError as e:
                    log_aviso("monitorar_impressoes.impressora", f"Impressora inacessível: {printer_name!r} | {e}")
                    # Impressora inacessível (rede, permissão, etc.) — não interrompe o monitoramento
                    print(f"[Aviso] Impressora '{printer_name}': {e}")
                finally:
                    if p_handle:
                        win32print.ClosePrinter(p_handle)

            if DEBUG and total_jobs_ciclo > 0:
                print(f"[Diagnóstico] Total de jobs na fila neste ciclo: {total_jobs_ciclo}")

            # Limpeza de memória: remove jobs do cache com mais de 24h
            agora = time.time()
            chaves_para_remover = [k for k, v in jobs_processados.items() if agora - v > CACHE_JOB_HORAS * 3600]
            for k in chaves_para_remover:
                del jobs_processados[k]

            # Limpeza de logs antigos: uma vez por dia, remove arquivos com mais de DIAS_RETENCAO
            if datetime.now() - ultima_limpeza > timedelta(days=1):
                limpar_logs_antigos()
                ultima_limpeza = datetime.now()

        except KeyboardInterrupt:
            raise  # Ctrl+C — encerra normalmente
        except Exception as e:
            log_erro("monitorar_impressoes.loop", "Erro no loop de monitoramento", e)
            print(f"Erro no loop de monitoramento: {type(e).__name__}: {e}")
        
        time.sleep(INTERVALO_SEGUNDOS)

if __name__ == "__main__":
    def _excepthook(tipo, valor, tb):
        log = configurar_log_erro()
        log.critical(
            "[main] Exceção não tratada: %s: %s\n%s",
            tipo.__name__, valor, "".join(traceback.format_exception(tipo, valor, tb)),
        )
        sys.__excepthook__(tipo, valor, tb)

    sys.excepthook = _excepthook
    try:
        iniciar_log()
        monitorar_impressoes()
    except KeyboardInterrupt:
        print("\nMonitoramento encerrado pelo usuário.")
    except Exception as e:
        log_erro("main", "Encerramento por exceção", e)
        raise
