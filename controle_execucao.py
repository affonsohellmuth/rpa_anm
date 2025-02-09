import os
import subprocess
import json
import psutil
import sys
import uuid

# Nome dos arquivos utilizados para controle
PASTA_SALVAMENTO = os.path.join(os.path.expanduser("~"), "Documents/Checkpoint ANM")
os.makedirs(PASTA_SALVAMENTO, exist_ok=True)


def carregar_progresso(processo_id):
    """Carrega o progresso salvo da instância especificada. Retorna None se estiver vazio ou corrompido."""
    nome_arquivo = os.path.join(PASTA_SALVAMENTO, f"progresso_{processo_id}.json")

    if os.path.exists(nome_arquivo):
        try:
            if os.path.getsize(nome_arquivo) == 0:  # Arquivo existe, mas está vazio
                return None
            with open(nome_arquivo, "r") as arquivo:
                return json.load(arquivo)
        except json.JSONDecodeError:
            print(
                f"\n⚠️ Erro ao carregar progresso da instância {processo_id}: arquivo corrompido. Iniciando do zero...")
            return None
    return None


def salvar_progresso(progresso, processo_id):
    """Salva o progresso da execução em um arquivo JSON separado para cada processo."""
    nome_arquivo = os.path.join(PASTA_SALVAMENTO, f"progresso_{processo_id}.json")
    with open(nome_arquivo, "w") as arquivo:
        json.dump(progresso, arquivo, indent=4)


def gerar_processo_id():
    """Gera um ID único para cada nova instância."""
    return str(uuid.uuid4())[:8]  # Usa apenas os primeiros 8 caracteres do UUID


def iniciar_execucao(nova_execucao=False):
    """Inicia uma nova execução ou continua de onde parou."""
    processo_id = gerar_processo_id()  # Sempre gera um novo ID

    if nova_execucao:
        progresso = carregar_progresso(processo_id)
        if progresso:
            print("\n⚠️ Um progresso anterior foi detectado.")
            confirmacao = input("Deseja realmente iniciar do zero? (s/n): ").strip().lower()
            if confirmacao == "s":
                os.remove(os.path.join(PASTA_SALVAMENTO, f"progresso_{processo_id}.json"))
                print("\n🗑️ Progresso anterior apagado. Iniciando nova execução...")
            else:
                print("\n🔄 Mantendo progresso anterior. Continuando execução...")

    print(f"\n🔄 Iniciando processo {processo_id}...")

    # Opções para abrir em novo terminal no Windows/Linux/Mac
    if sys.platform == "win32":
        processo = subprocess.Popen(["python", "processo.py", processo_id], creationflags=subprocess.CREATE_NEW_CONSOLE)
    else:
        processo = subprocess.Popen(["python", "processo.py", processo_id], stdout=subprocess.DEVNULL,
                                    stderr=subprocess.DEVNULL)

    # Salvar PID do processo
    nome_pid = os.path.join(PASTA_SALVAMENTO, f"execucao_{processo_id}.pid")
    with open(nome_pid, "w") as f:
        f.write(str(processo.pid))

    print(f"✅ Processo {processo_id} iniciado com PID {processo.pid}")


def continuar_execucao():
    """Continua todas as execuções que possuírem progresso salvo automaticamente."""
    arquivos_progresso = [f for f in os.listdir(PASTA_SALVAMENTO) if f.startswith("progresso_") and f.endswith(".json")]

    if not arquivos_progresso:
        print("\n⚠️ Nenhum progresso encontrado. Iniciando uma nova execução...")
        iniciar_execucao(nova_execucao=False)
        return

    print("\n🔄 Retomando todas as execuções pendentes...")

    for arquivo in arquivos_progresso:
        processo_id = arquivo.replace("progresso_", "").replace(".json", "")
        print(f"➡️ Retomando instância {processo_id}...")

        pid_file = os.path.join(PASTA_SALVAMENTO, f"execucao_{processo_id}.pid")

        if os.path.exists(pid_file):
            print(f"⚠️ Instância {processo_id} já está em execução. Pulando...")
        else:
            # Passa o caminho do arquivo de progresso como argumento para o processo
            arquivo_progresso = os.path.join(PASTA_SALVAMENTO, f"progresso_{processo_id}.json")
            print(
                f"📄 Arquivo de progresso encontrado para {processo_id}. Iniciando a execução a partir do progresso...")

            # Inicia o processo e salva o PID
            processo = subprocess.Popen(
                ["python", "processo.py", processo_id, arquivo_progresso]
            )

            with open(pid_file, "w") as pid_f:
                pid_f.write(str(processo.pid))  # Armazena o PID no arquivo

    print("\n✅ Todas as execuções disponíveis foram retomadas.")

def matar_processo(pid):
    """Encerra o processo e todos os seus subprocessos."""
    try:
        processo = psutil.Process(pid)

        # Encerra todos os subprocessos primeiro
        for subproc in processo.children(recursive=True):
            subproc.terminate()  # Tenta encerrar de forma segura

        # Agora encerra o processo principal
        processo.terminate()

        # Aguarda encerramento e força caso necessário
        processo.wait(timeout=5)
    except psutil.NoSuchProcess:
        print(f"⚠️ Processo {pid} já estava encerrado.")
    except psutil.TimeoutExpired:
        print(f"⏳ Processo {pid} não respondeu, forçando encerramento...")
        processo.kill()  # Força encerramento se não parar


def parar_execucao():
    """Para todas as execuções em andamento."""
    arquivos_pid = [f for f in os.listdir(PASTA_SALVAMENTO) if f.startswith("execucao_") and f.endswith(".pid")]

    if not arquivos_pid:
        print("\n⚠️ Nenhum processo em execução encontrado.")
        return

    print("\n⏹️ Interrompendo todas as execuções...")

    for arquivo in arquivos_pid:
        nome_pid = os.path.join(PASTA_SALVAMENTO, arquivo)
        with open(nome_pid, "r") as f:
            pid = int(f.read())

        print(f"🔪 Encerrando processo {pid}...")
        matar_processo(pid)  # Usa a nova função para garantir que o processo seja encerrado

        # Remove o arquivo de PID após garantir que o processo foi finalizado
        os.remove(nome_pid)

    print("\n✅ Todas as execuções foram interrompidas.")
