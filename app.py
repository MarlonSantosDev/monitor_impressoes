"""
Monitor de filas de impressão Windows.

Registra cada job em um Excel do dia na raiz (pasta onde o .exe roda):
  log_impressoes_DDMMYYYY.xlsx  (ex.: log_impressoes_20022026.xlsx)
Também cria a pasta arquivos/ no mesmo diretório.
Requer Windows (pywin32) e openpyxl.
"""
import os
import re
import socket
import sys
import time
from datetime import datetime, timedelta

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
    "IP", "Local_Arquivo"
]
DIAS_RETENCAO = 2       # Remover log_impressoes_*.xlsx com mais de 2 dias
INTERVALO_SEGUNDOS = 2  # Varredura das filas a cada 2 segundos
CACHE_JOB_HORAS = 24    # Limpar do cache jobs processados há mais de 24 h


def obter_ip_local():
    """Obtém o IP local da máquina (evita 127.0.0.1 quando possível)."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.settimeout(0.5)
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except (OSError, socket.error):
        try:
            return socket.gethostbyname(socket.gethostname())
        except socket.gaierror:
            return "N/A"


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


def monitorar_impressoes():
    print(f"Monitorando impressões... (Pressione Ctrl+C para parar)")
    ip_local = obter_ip_local()
    print(f"IP desta máquina: {ip_local}")

    jobs_processados = {}
    ultima_limpeza = datetime.now()

    while True:
        try:
            # Lista impressoras locais e de rede instaladas
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            
            for printer in printers:
                printer_name = printer[2]
                p_handle = None
                
                try:
                    p_handle = win32print.OpenPrinter(printer_name)
                    
                    # Nível 2 traz detalhes do dono e nome do documento
                    jobs = win32print.EnumJobs(p_handle, 0, -1, 2)
                    
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
                        tamanho = job.get('Size', 0) # Tamanho em bytes
                        
                        # Tenta pegar a data de submissão original do job
                        try:
                            data_raw = job.get('Submitted')
                            if data_raw is None:
                                raise ValueError("Submitted ausente")
                            data_hora = f"{data_raw.year}-{data_raw.month:02d}-{data_raw.day:02d} {data_raw.hour:02d}:{data_raw.minute:02d}:{data_raw.second:02d}"
                        except (AttributeError, TypeError, ValueError):
                            data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        # Filtro de Qualidade: 
                        # Às vezes o job aparece com 0 bytes/páginas enquanto é criado (Spooling).
                        # Se for 0, esperamos o próximo ciclo para ver se atualiza, a menos que esteja imprimindo.
                        status = job.get('Status', 0)
                        if paginas == 0 and status & win32print.JOB_STATUS_SPOOLING:
                            continue # Pula e tenta pegar no próximo loop com os dados completos

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
                                ws = wb[NOME_ABA] if NOME_ABA in wb.sheetnames else wb.active
                                if ws.title != NOME_ABA:
                                    ws.title = NOME_ABA
                            ws.append([
                                job_id, usuario, data_hora, documento, paginas, printer_name, tamanho,
                                ip_local, arquivo_hoje
                            ])
                            wb.save(arquivo_hoje)

                            print(f"[NOVO] {data_hora} | {usuario} | {documento} ({paginas} pgs) | IP: {ip_local}")

                            jobs_processados[job_unique_key] = time.time()

                        except PermissionError:
                            print("Erro: Arquivo Excel aberto por outro programa. Feche o arquivo e tente novamente.")

                except OSError as e:
                    # Impressora inacessível (rede, permissão, etc.) — não interrompe o monitoramento
                    print(f"[Aviso] Impressora '{printer_name}': {e}")
                finally:
                    if p_handle:
                        win32print.ClosePrinter(p_handle)
            
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
