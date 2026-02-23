"""
Monitor de filas de impressão Windows.

Registra cada job em um Excel do dia na raiz (pasta onde o .exe roda):
  log_impressoes_DDMMYYYY.xlsx  (ex.: log_impressoes_20022026.xlsx)
Cria a pasta arquivos/ e, se SALVAR_COPIA_SPL = True, grava lá uma cópia do
arquivo de spool (.SPL) do job — não é o documento original (PDF/DOC), e sim
o formato que o Windows envia para a impressora. Pode exigir permissões de admin.
Requer Windows (pywin32) e openpyxl.
Se não coletar dados nos testes, defina DEBUG = True para ver impressoras e jobs na fila.
"""
import os
import re
import shutil
import sys
import threading
import time
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
NOME_ABA = "Impressões"
CABECALHO = [
    "ID_Job", "Usuario", "Data_Hora", "Arquivo", "Paginas", "Impressora", "Tamanho_Bytes",
    "Local_Arquivo"
]
DIAS_RETENCAO = 2       # Remover log_impressoes_*.xlsx com mais de 2 dias
INTERVALO_SEGUNDOS = 0.5  # Varredura das filas (menor = mais chance de pegar jobs rápidos)
CACHE_JOB_HORAS = 24    # Limpar do cache jobs processados há mais de 24 h
DEBUG = False           # True: mostra quantas impressoras/jobs por ciclo
SALVAR_COPIA_SPL = True  # Copiar arquivo de spool do job para pasta arquivos/ (pode exigir admin)


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


def copiar_spool_para_arquivos(job_id: int, dest_path: str) -> bool:
    """
    Tenta copiar o arquivo de spool do job (System32\\spool\\PRINTERS) para dest_path.
    Retorna True se copiou, False se falhou (ex.: sem permissão, pooling ativo).
    O arquivo é .SPL (formato spool Windows), não o documento original.
    """
    spool_dir = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "System32", "spool", "PRINTERS")
    nome_spl = f"{job_id:05d}.SPL"
    src = os.path.join(spool_dir, nome_spl)
    if not os.path.isfile(src):
        return False
    try:
        shutil.copy2(src, dest_path)
        return True
    except (OSError, PermissionError):
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
    os.makedirs(PASTA_ARQUIVOS, exist_ok=True)
    print(f"Log do dia na raiz: log_impressoes_DDMMYYYY.xlsx")
    print(f"Pasta criada: arquivos/")
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
                try:
                    shutil.copy2(src, dest)
                    spool_copies[job_id] = dest
                    if DEBUG:
                        print(f"[Spool] Capturado: {filename} → {dest}")
                except (OSError, PermissionError):
                    pass  # Arquivo ainda bloqueado ou já deletado; fallback no loop principal
        except Exception as e:
            if DEBUG:
                print(f"[Spool] Erro no watcher: {e}")
            time.sleep(1)


def monitorar_impressoes():
    print(f"Monitorando impressões... (Pressione Ctrl+C para parar)")

    jobs_processados = {}
    ultima_limpeza = datetime.now()

    # Dict compartilhado: job_id → caminho do SPL copiado pela thread watcher
    spool_copies: dict = {}
    if SALVAR_COPIA_SPL:
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
                        documento = job.get('pDocument', 'Sem Nome')
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

                        # Local do arquivo: watcher já copiou o SPL; fallback para cópia direta
                        local_arquivo = ""
                        if SALVAR_COPIA_SPL:
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
                            if not local_arquivo:  # fallback: tenta cópia direta
                                if copiar_spool_para_arquivos(job_id, dest_path):
                                    local_arquivo = dest_path

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
                                job_id, usuario, data_hora, documento, paginas, printer_name, tamanho,
                                local_arquivo
                            ])
                            wb.save(arquivo_hoje)

                            print(f"[NOVO] {data_hora} | {usuario} | {documento} ({paginas} pgs)")
                            if SALVAR_COPIA_SPL and os.path.isfile(local_arquivo):
                                print(f"      Cópia spool salva: {local_arquivo}")

                            jobs_processados[job_unique_key] = time.time()

                        except PermissionError:
                            print("Erro: Arquivo Excel aberto por outro programa. Feche o arquivo e tente novamente.")
                        except Exception as e_excel:
                            print(f"Erro ao salvar Excel ({type(e_excel).__name__}): {e_excel}")

                except OSError as e:
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
            print(f"Erro no loop de monitoramento: {type(e).__name__}: {e}")
        
        time.sleep(INTERVALO_SEGUNDOS)

if __name__ == "__main__":
    try:
        iniciar_log()
        monitorar_impressoes()
    except KeyboardInterrupt:
        print("\nMonitoramento encerrado pelo usuário.")
