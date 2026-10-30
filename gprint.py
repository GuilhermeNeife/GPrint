from builtins import super
import os
import sys
import shutil
import time
import math
import json
import queue
import tempfile
import subprocess
import threading
import logging
import io
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, END, colorchooser
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style
from tkinterdnd2 import DND_FILES, TkinterDnD
from datetime import datetime
from typing import List
import fitz  
from PyPDF2 import PdfMerger
from pikepdf import Pdf, Rectangle
from reportlab.lib.pagesizes import A3, A4
from reportlab.pdfgen import canvas
from PIL import Image
import re
import sqlite3
import openpyxl
from openpyxl import Workbook, load_workbook
import uuid
from openpyxl.utils import get_column_letter
import random
import ctypes
import ipaddress
import hmac
from tkinter import simpledialog

# ============================================================================
# GPRINT - CONFIGURACAO CENTRAL DA PRODUCAO
# Ajuste SOMENTE este diretorio para localizar ambos os bancos.
# Nao crie a pasta da rede automaticamente se o servidor estiver offline.
# ============================================================================
GPRINT_DIRETORIO_DADOS = r"\\192.168.200.253\producao\____REPOSITORIO GERAL POD\UICLAP\LOGS"
GPRINT_BANCO_HISTORICO = os.path.join(GPRINT_DIRETORIO_DADOS, "impressoes.db")
GPRINT_BANCO_EQUIPAMENTOS = os.path.join(GPRINT_DIRETORIO_DADOS, "equipamentos.db")

# Senha de acesso a Configuracoes > Equipamentos.
# ALTERE antes de compilar o EXE. Nunca use a senha real em um GitHub publico.
# Protege a interface do cadastro; nao substitui permissoes da pasta de rede.
GPRINT_SENHA_ADMIN_EQUIPAMENTOS = "ALTERE_ESTA_SENHA_ANTES_DE_USAR"


# ============================================================================
# GPRINT - HISTORICO EM LOTES / FILA LOCAL DURAVEL
# Nao modifica imposicao de PDFs, montagem ou comandos LPR.
# Cada evento e salvo PRIMEIRO no computador e transmitido em lotes ao final.
# ============================================================================
import socket
from contextlib import contextmanager

# O arquivo .lock continua compativel com as estacoes que usam versoes antigas.
# A nova versao tenta uma gravacao CENTRAL por lote, nao por copia/PDF.
GPRINT_ESPERA_TRAVA_SEGUNDOS = 20
GPRINT_TAMANHO_LOTE_BANCO = 250
GPRINT_AVISO_INTERVALO_SEGUNDOS = 120
GPRINT_FILA_LOCAL = os.path.join(
    os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'),
    'GPrint', 'historico_pendente.db'
)
_GPRINT_ULTIMOS_AVISOS = {}
_GPRINT_AVISOS_MUTEX = threading.Lock()
_GPRINT_SINCRONIZACAO_MUTEX = threading.Lock()


class GPrintBancoIndisponivel(Exception):
    """Diretorio da rede ou arquivo do banco central indisponivel."""


class GPrintBancoOcupado(Exception):
    """Outra estacao manteve a trava alem da espera."""


def gprint_checar_banco_central(caminho):
    """Nao cria o banco central se a rede estiver fora do ar."""
    if not os.path.isdir(os.path.dirname(caminho)) or not os.path.isfile(caminho):
        raise GPrintBancoIndisponivel('Banco central ou pasta de producao inacessivel')


def gprint_avisar_operador(root, tipo, detalhe=''):
    """Avisos de operacao sem nomes de arquivos, caminhos ou termos tecnicos."""
    # 'trava' e lentidao temporaria nao devem gerar uma mensagem para producao.
    if tipo not in ('rede', 'local'):
        return
    agora = time.monotonic()
    with _GPRINT_AVISOS_MUTEX:
        ultimo = _GPRINT_ULTIMOS_AVISOS.get(tipo, -1e10)
        if agora - ultimo < GPRINT_AVISO_INTERVALO_SEGUNDOS:
            return
        _GPRINT_ULTIMOS_AVISOS[tipo] = agora

    if tipo == 'rede':
        titulo = 'GPrint - HISTORICO INDISPONIVEL'
        mensagem = (
            'Não foi possível acessar o banco de dados da produção.\n\n'
            'Verifique se a pasta "produção" está online e acessível.\n'
            'Se o problema continuar, INFORME O GUILHERME.\n\n'
            'Os registros desta estação permanecerão guardados localmente '
            'e serão reenviados na próxima sincronização com a rede disponível.'
        )
    else:
        titulo = 'GPrint - FALHA AO GUARDAR HISTORICO'
        mensagem = (
            'Não foi possível guardar o histórico neste computador.\n'
            'INFORME O GUILHERME imediatamente.\n\n'
            'O processamento pode ter ocorrido sem registro completo.'
        )

    def mostrar():
        try:
            messagebox.showwarning(titulo, mensagem, parent=root)
        except Exception:
            logging.exception('Falha ao apresentar mensagem do banco')

    try:
        if root is not None:
            root.after(0, mostrar)
        else:
            logging.warning('%s: %s', titulo, mensagem)
    except Exception:
        logging.exception('Falha ao agendar aviso de banco')


@contextmanager
def gprint_lock_banco(caminho_lock, timeout=GPRINT_ESPERA_TRAVA_SEGUNDOS):
    """Trava compativel entre estacoes, liberada somente por quem a criou."""
    identidade = f'{socket.gethostname()}|{os.getpid()}|{uuid.uuid4().hex}'
    inicio = time.monotonic()
    while True:
        try:
            fd = os.open(caminho_lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                with os.fdopen(fd, 'w', encoding='utf-8') as arquivo:
                    arquivo.write(identidade)
                    arquivo.flush()
            except Exception:
                try:
                    os.remove(caminho_lock)
                except OSError:
                    pass
                raise
            break
        except FileExistsError:
            if time.monotonic() - inicio >= timeout:
                raise GPrintBancoOcupado('Banco ocupado por outra estacao')
            time.sleep(random.uniform(0.15, 0.35))
    try:
        yield
    finally:
        try:
            with open(caminho_lock, 'r', encoding='utf-8') as arquivo:
                atual = arquivo.read()
            if atual == identidade:
                os.remove(caminho_lock)
            else:
                logging.error('Trava do banco pertence a outra estacao; nao removida')
        except FileNotFoundError:
            logging.warning('Trava do banco nao estava presente ao liberar')
        except OSError:
            logging.exception('Nao foi possivel liberar trava do banco')


_GPRINT_CRIAR_LOGS = """
    CREATE TABLE IF NOT EXISTS logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data_hora TEXT, lote TEXT, pilha TEXT, arquivo TEXT,
        maquina TEXT, ip TEXT, formato TEXT, modo TEXT,
        quantidade INTEGER, status TEXT, erro TEXT, copia TEXT,
        usuario TEXT, montagem_ok INTEGER DEFAULT 0,
        envio_ok INTEGER DEFAULT 0
    )
"""
_GPRINT_INSERIR_LOG = """
    INSERT INTO logs
    (data_hora, lote, pilha, arquivo, maquina, ip, formato, modo,
     quantidade, status, erro, copia, usuario, montagem_ok, envio_ok)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""


def gprint_guardar_evento_local(dados, caminho_local=GPRINT_FILA_LOCAL):
    """Commit local antes de retornar; sobrevive ao encerramento do GPrint."""
    pasta = os.path.dirname(caminho_local)
    os.makedirs(pasta, exist_ok=True)
    identificador = uuid.uuid4().hex
    instante = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with sqlite3.connect(caminho_local, timeout=20) as conexao:
        # Local apenas: WAL nao e ativado no banco central em rede.
        conexao.execute('PRAGMA busy_timeout=20000')
        conexao.execute("""
            CREATE TABLE IF NOT EXISTS fila_eventos (
                evento_id TEXT PRIMARY KEY,
                data_hora TEXT NOT NULL,
                dados_json TEXT NOT NULL
            )
        """)
        conexao.execute("""
            INSERT INTO fila_eventos (evento_id, data_hora, dados_json)
            VALUES (?, ?, ?)
        """, (identificador, instante, json.dumps(dados, ensure_ascii=False)))
    return identificador


def gprint_contar_pendentes(caminho_local=GPRINT_FILA_LOCAL):
    if not os.path.isfile(caminho_local):
        return 0
    with sqlite3.connect(caminho_local, timeout=10) as conexao:
        try:
            return conexao.execute('SELECT COUNT(*) FROM fila_eventos').fetchone()[0]
        except sqlite3.OperationalError as e:
            if 'no such table' in str(e).lower():
                return 0
            raise


def gprint_ler_pendentes(limite=GPRINT_TAMANHO_LOTE_BANCO,
                        caminho_local=GPRINT_FILA_LOCAL):
    if not os.path.isfile(caminho_local):
        return []
    with sqlite3.connect(caminho_local, timeout=10) as conexao:
        try:
            return conexao.execute("""
                SELECT evento_id, data_hora, dados_json FROM fila_eventos
                ORDER BY rowid LIMIT ?
            """, (limite,)).fetchall()
        except sqlite3.OperationalError as e:
            if 'no such table' in str(e).lower():
                return []
            raise


def gprint_gravar_lote_central(caminho_central, eventos):
    """Uma trava / uma transacao por lote, compatível com a tabela existente.

    gprint_ids_recebidos impede duplicacoes se a rede falhar DEPOIS do commit
    mas ANTES de apagar a fila local. Nao altera colunas da tabela logs.
    """
    if not eventos:
        return 0
    gprint_checar_banco_central(caminho_central)
    with gprint_lock_banco(caminho_central + '.lock'):
        conexao = sqlite3.connect(caminho_central, timeout=15)
        try:
            conexao.execute('PRAGMA busy_timeout=15000')
            conexao.execute('BEGIN IMMEDIATE')
            conexao.execute(_GPRINT_CRIAR_LOGS)
            conexao.execute("""
                CREATE TABLE IF NOT EXISTS gprint_ids_recebidos (
                    evento_id TEXT PRIMARY KEY,
                    data_hora TEXT NOT NULL
                )
            """)
            novos = 0
            for evento_id, instante, dados_json in eventos:
                dados = json.loads(dados_json)
                cursor = conexao.execute("""
                    INSERT OR IGNORE INTO gprint_ids_recebidos
                    (evento_id, data_hora) VALUES (?, ?)
                """, (evento_id, instante))
                if cursor.rowcount != 1:
                    continue
                valores = (
                    instante, dados.get('lote', ''), dados.get('pilha', ''),
                    dados.get('arquivo', ''), dados.get('maquina', ''),
                    dados.get('ip', ''), dados.get('formato', ''),
                    dados.get('modo', ''), dados.get('quantidade', 0),
                    dados.get('status', ''), dados.get('erro', ''),
                    dados.get('copia', ''), dados.get('usuario', ''),
                    dados.get('montagem_ok', 0), dados.get('envio_ok', 0)
                )
                conexao.execute(_GPRINT_INSERIR_LOG, valores)
                novos += 1
            conexao.commit()
            return novos
        except Exception:
            conexao.rollback()
            raise
        finally:
            conexao.close()


def gprint_remover_enviados_local(ids, caminho_local=GPRINT_FILA_LOCAL):
    if not ids:
        return
    with sqlite3.connect(caminho_local, timeout=20) as conexao:
        conexao.executemany('DELETE FROM fila_eventos WHERE evento_id=?', [(i,) for i in ids])


def gprint_sincronizar_pendentes(central, root=None, emitir_log=None,
                                caminho_local=GPRINT_FILA_LOCAL):
    """Tenta sincronizar a fila; nunca reenvia PDFs ou comandos LPR."""
    if emitir_log is None:
        emitir_log = lambda mensagem: logging.info('%s', mensagem)
    if not _GPRINT_SINCRONIZACAO_MUTEX.acquire(blocking=False):
        return 0  # Outra sincronizacao desta instancia esta em andamento.
    enviados = 0
    try:
        while True:
            eventos = gprint_ler_pendentes(caminho_local=caminho_local)
            if not eventos:
                if enviados:
                    emitir_log(f'Historico sincronizado: {enviados} registro(s).')
                return enviados
            try:
                gprint_gravar_lote_central(central, eventos)
            except GPrintBancoOcupado:
                # Nao incomoda operadores com arquivos .lock ou detalhes de banco.
                emitir_log('Historico aguardando disponibilidade do banco central. Registros protegidos neste computador.')
                return enviados
            except (GPrintBancoIndisponivel, OSError) as e:
                logging.warning('Historico: falha de rede ao sincronizar: %s', e)
                gprint_avisar_operador(root, 'rede')
                emitir_log('Historico aguardando rede de producao. Registros protegidos neste computador.')
                return enviados
            except sqlite3.OperationalError as e:
                # SQLite ocupado NAO equivale a rede offline.
                logging.warning('Historico: banco central temporariamente ocupado/erro: %s', e)
                emitir_log('Historico ainda nao sincronizado. Registros protegidos neste computador.')
                return enviados
            except Exception as e:
                logging.exception('Erro inesperado ao sincronizar historico: %s', e)
                emitir_log('Historico aguardando verificacao. Registros protegidos neste computador.')
                return enviados
            # Somente depois de o commit central confirmar e que removemos da fila.
            # Se a limpeza falhar, a mesma ID sera ignorada na proxima tentativa.
            ids = [evento[0] for evento in eventos]
            try:
                gprint_remover_enviados_local(ids, caminho_local=caminho_local)
            except Exception:
                logging.exception('Central confirmou registros, mas fila nao foi limpa')
                return enviados
            enviados += len(ids)
    except Exception as e:
        logging.exception('Nao foi possivel consultar fila local de historico: %s', e)
        gprint_avisar_operador(root, 'local')
        return enviados
    finally:
        _GPRINT_SINCRONIZACAO_MUTEX.release()


def gprint_registrar_log_seguro(dados, central, local, root, emitir_log):
    """Compativel com chamadas anteriores, mas SEM acessar a rede por evento."""
    try:
        gprint_guardar_evento_local(dados)
        return True
    except Exception:
        logging.exception('ERRO: falha ao registrar evento na fila local')
        emitir_log('ATENCAO: registro do historico nao foi guardado. Informar Guilherme.')
        gprint_avisar_operador(root, 'local')
        return False


def gprint_verificar_banco_no_inicio(caminho, root):
    """Verificacao e recuperacao de eventos anteriores, sem travar a interface."""
    try:
        gprint_checar_banco_central(caminho)
    except (OSError, GPrintBancoIndisponivel):
        gprint_avisar_operador(root, 'rede')
        return
    if gprint_contar_pendentes():
        gprint_sincronizar_pendentes(caminho, root=root)

# ============================================================================

# ============================================================================

# LIMITE DE ARQUIVOS TXT NA ABA MIOLO:
# 0 = sem limite (quantos arquivos TXT desejar)
# 1, 2, 3... = quantidade maxima de TXT por processamento
MAX_ARQUIVOS_TXT_MIOLO = 2

DATA_LIMITE = "29/07/2027"
PASTA_BASE = r"\\192.168.200.253\producao\____REPOSITORIO GERAL POD\UICLAP\LOTES"
PASTA_TRABALHO = os.path.join(os.environ.get('TEMP', os.path.expanduser("~")), "UICLAP_TEMP")
CONFIG_FILE = os.path.expanduser("~/uiclap_config.json")
MM = 2.83465 # Fator de conversão mm para pontos
# Configurar logging para console
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
# Garantir que a pasta de trabalho exista
if not os.path.exists(PASTA_TRABALHO):
    os.makedirs(PASTA_TRABALHO, exist_ok=True)
# Temas claros
TEMAS_CLAROS = [
    "cosmo",
    "flatly",
    "journal",
    "litera",
    "lumen",
    "minty",
    "pulse",
    "sandstone",
    "united",
    "yeti",
    "morph"
]
# Temas escuros
TEMAS_ESCUROS = [
    "darkly",
    "superhero",
    "solar",
    "cyborg",
    "vapor"
]
DEFAULT_CONFIG = {
    'cor_fosco': (0, 0, 1),
    'cor_brilho': (1, 0.7, 0.1),
    'fonte': 'helv',
    'tamanho_fonte': 11,
    'deslocamento_padrao': -30,
    'deslocamento_capa_28x56': 92,
    'deslocamento_x': 70 * 2.83465,
    'tema': 'darkly',
        # NOVO
    "sangria_padrao_miolo": 3
}
def obter_sangria_padrao_miolo():
    """Retorna o ajuste atual de sangria, em milímetros, para uma execução de miolo.

    O valor é capturado uma vez no começo do lote. Os algoritmos de imposição
    permanecem inalterados, inclusive com o valor original de 3 mm por padrão.
    """
    valor = CONFIG.get("sangria_padrao_miolo", DEFAULT_CONFIG["sangria_padrao_miolo"])
    try:
        valor = float(str(valor).replace(",", "."))
        if not math.isfinite(valor) or not (0 <= valor <= 20):
            raise ValueError("Sangria fora da faixa de 0 a 20 mm")
        return valor
    except (ValueError, TypeError):
        logging.warning("Sangria inválida nas configurações (%r). Usando 3 mm.", valor)
        return float(DEFAULT_CONFIG["sangria_padrao_miolo"])


# ============================================================================
# GPRINT - MONITOR VISUAL LEVE (CAPA E MIOLO)
# Apenas interface: nao interfere na montagem, nos PDFs nem no envio LPR.
# A leitura de RAM usa API nativa do Windows, sem dependencia adicional.
# ============================================================================
GPRINT_RAM_INTERVALO_MS = 2000   # Consultar a cada 2 segundos (baixo custo)
GPRINT_RAM_AMARELO_ACIMA = 70.0 # Amarelo acima de 70% ate 85%
GPRINT_RAM_VERMELHO_ACIMA = 85.0 # Vermelho SOMENTE acima de 85%


class _GPrintMemoryStatusEx(ctypes.Structure):
    """Estrutura WinAPI para obter o percentual de uso da RAM fisica total."""
    _fields_ = [
        ('dwLength', ctypes.c_ulong),
        ('dwMemoryLoad', ctypes.c_ulong),
        ('ullTotalPhys', ctypes.c_ulonglong),
        ('ullAvailPhys', ctypes.c_ulonglong),
        ('ullTotalPageFile', ctypes.c_ulonglong),
        ('ullAvailPageFile', ctypes.c_ulonglong),
        ('ullTotalVirtual', ctypes.c_ulonglong),
        ('ullAvailVirtual', ctypes.c_ulonglong),
        ('ullAvailExtendedVirtual', ctypes.c_ulonglong),
    ]


def gprint_percentual_ram_pc():
    """Uso percentual de RAM total no Windows; None se indisponivel.

    Nao inclui leituras por processo; e somente a RAM fisica do computador.
    No Windows, evita instalar psutil no executavel PyInstaller.
    """
    try:
        if os.name == 'nt':
            status = _GPrintMemoryStatusEx()
            status.dwLength = ctypes.sizeof(_GPrintMemoryStatusEx)
            if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                return None
            # Usar totais fisicos para permitir precisao decimal.
            if status.ullTotalPhys <= 0:
                return None
            em_uso = status.ullTotalPhys - status.ullAvailPhys
            return max(0.0, min(100.0, em_uso * 100.0 / status.ullTotalPhys))
        # Compatibilidade opcional de testes fora do Windows.
        import psutil
        return float(psutil.virtual_memory().percent)
    except (OSError, ImportError, AttributeError, ValueError):
        return None


def gprint_cor_ram(percentual):
    """Semaforo: verde ate 70%, amarelo acima de 70%, vermelho acima de 85%."""
    if percentual is None:
        return 'secondary'
    if percentual > GPRINT_RAM_VERMELHO_ACIMA:
        return 'danger'
    if percentual > GPRINT_RAM_AMARELO_ACIMA:
        return 'warning'
    return 'success'


def gprint_tempo_minutos(segundos):
    """Tempo decorrido em minutos e segundos, inclusive para mais de uma hora."""
    total = max(0, int(segundos))
    minutos, segundos_restantes = divmod(total, 60)
    return f'{minutos:02d}min {segundos_restantes:02d}s'


class GPrintIndicadoresAndamento:
    """Indicadores comuns nas abas Capa/Miolo, com timers do Tkinter.

    Nao cria threads, nao consulta a rede e nao altera a rotina de processamento.
    A RAM continua visivel em repouso; o cronometro roda apenas durante o trabalho.
    """
    def __init__(self, quadro, linha_ram=1):
        self.quadro = quadro
        self.inicio = None
        self._ram_after = None
        self._tempo_after = None
        self._cor_atual = None

        ttk.Label(quadro, text='RAM do computador:').grid(
            row=linha_ram, column=0, padx=5, pady=4, sticky='w')
        self.barra_ram = ttk.Progressbar(
            quadro, mode='determinate', maximum=100, bootstyle='success')
        self.barra_ram.grid(
            row=linha_ram, column=1, sticky='ew', padx=5, pady=4)
        self.valor_ram = ttk.Label(quadro, text='--%  |  aguardando leitura', anchor='e')
        self.valor_ram.grid(row=linha_ram, column=2, padx=5, pady=4, sticky='e')

        ttk.Label(quadro, text='Tempo decorrido:').grid(
            row=linha_ram + 1, column=0, padx=5, pady=4, sticky='w')
        self.valor_tempo = ttk.Label(quadro, text='00min 00s', anchor='e')
        self.valor_tempo.grid(
            row=linha_ram + 1, column=2, padx=5, pady=4, sticky='e')

        quadro.bind('<Destroy>', self._ao_destruir, add='+')
        # Nao bloquear a abertura consultando RAM durante a construcao da tela.
        self._ram_after = quadro.after(600, self._atualizar_ram)

    def _atualizar_ram(self):
        self._ram_after = None
        if not self.quadro.winfo_exists():
            return
        percentual = gprint_percentual_ram_pc()
        cor = gprint_cor_ram(percentual)
        if percentual is None:
            self.barra_ram.configure(value=0)
            self.valor_ram.configure(text='RAM indisponivel')
        else:
            self.barra_ram.configure(value=percentual)
            estado = ('ALTA' if cor == 'danger' else
                      'ATENCAO' if cor == 'warning' else 'NORMAL')
            self.valor_ram.configure(text=f'{percentual:.1f}%  |  {estado}')
        # Aplicar cor apenas ao mudar de faixa: evita reconfiguracoes desnecessarias.
        if cor != self._cor_atual:
            self.barra_ram.configure(bootstyle=cor)
            self.valor_ram.configure(bootstyle=cor)
            self._cor_atual = cor
        self._ram_after = self.quadro.after(GPRINT_RAM_INTERVALO_MS, self._atualizar_ram)

    def iniciar(self):
        self.inicio = time.monotonic()
        if self._tempo_after is not None:
            self.quadro.after_cancel(self._tempo_after)
            self._tempo_after = None
        self._atualizar_tempo()

    def _atualizar_tempo(self):
        self._tempo_after = None
        if self.inicio is None or not self.quadro.winfo_exists():
            return
        self.valor_tempo.configure(
            text=gprint_tempo_minutos(time.monotonic() - self.inicio))
        self._tempo_after = self.quadro.after(1000, self._atualizar_tempo)

    def finalizar(self):
        if self.inicio is not None:
            self.valor_tempo.configure(
                text=gprint_tempo_minutos(time.monotonic() - self.inicio))
        self.inicio = None
        if self._tempo_after is not None:
            self.quadro.after_cancel(self._tempo_after)
            self._tempo_after = None

    def _ao_destruir(self, evento):
        if evento.widget is not self.quadro:
            return
        for nome in ('_ram_after', '_tempo_after'):
            identificador = getattr(self, nome)
            if identificador is not None:
                try:
                    self.quadro.after_cancel(identificador)
                except tk.TclError:
                    pass
                setattr(self, nome, None)



# ============================================================================
# CADASTRO CENTRAL DE EQUIPAMENTOS E FILAS (SQLite independente do historico)
# ============================================================================
class GPrintCadastroErro(Exception):
    pass


class GPrintEquipamentosRepository:
    '''Cadastro central compartilhado pelas estacoes. Sem acesso a impressoras.'''
    def __init__(self, db_path=GPRINT_BANCO_EQUIPAMENTOS):
        self.db_path = db_path

    @contextmanager
    def _conectar(self):
        pasta = os.path.dirname(self.db_path)
        if not os.path.isdir(pasta):
            raise GPrintCadastroErro('A pasta producao esta indisponivel na rede.')
        conexao = None
        try:
            conexao = sqlite3.connect(self.db_path, timeout=15)
            conexao.row_factory = sqlite3.Row
            conexao.execute('PRAGMA foreign_keys = ON')
            conexao.execute('PRAGMA busy_timeout = 15000')
            yield conexao
        except (sqlite3.Error, OSError) as exc:
            raise GPrintCadastroErro('Nao foi possivel acessar equipamentos.db: ' + str(exc)) from exc
        finally:
            if conexao is not None:
                conexao.close()

    def preparar_banco(self):
        '''Cria apenas equipamentos.db se a pasta central ja estiver online.'''
        with self._conectar() as conn:
            with conn:
                conn.execute('''CREATE TABLE IF NOT EXISTS equipamentos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    ip TEXT NOT NULL,
                    tipo TEXT NOT NULL CHECK(tipo IN ('CAPA', 'MIOLO', 'AMBOS')),
                    ativo INTEGER NOT NULL DEFAULT 1,
                    criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )''')
                conn.execute('''CREATE TABLE IF NOT EXISTS filas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    equipamento_id INTEGER NOT NULL REFERENCES equipamentos(id) ON DELETE CASCADE,
                    nome TEXT NOT NULL,
                    descricao TEXT NOT NULL DEFAULT '',
                    ignorar_folha_obra INTEGER NOT NULL DEFAULT 0,
                    ativo INTEGER NOT NULL DEFAULT 1,
                    UNIQUE(equipamento_id, nome)
                )''')
                conn.execute('CREATE INDEX IF NOT EXISTS idx_filas_equipamento ON filas(equipamento_id, ativo)')

    def listar_equipamentos(self, tipo=None, incluir_inativos=False):
        sql = 'SELECT id,nome,ip,tipo,ativo FROM equipamentos WHERE 1=1'
        args = []
        if not incluir_inativos:
            sql += ' AND ativo=1'
        if tipo:
            sql += ' AND tipo IN (?,\'AMBOS\')'
            args.append(tipo)
        sql += ' ORDER BY nome COLLATE NOCASE'
        with self._conectar() as conn:
            return [dict(r) for r in conn.execute(sql, args)]

    def listar_filas(self, equipamento_id, incluir_inativas=False):
        sql = '''SELECT id,equipamento_id,nome,descricao,ignorar_folha_obra,ativo
                 FROM filas WHERE equipamento_id=?'''
        if not incluir_inativas:
            sql += ' AND ativo=1'
        sql += ' ORDER BY nome COLLATE NOCASE'
        with self._conectar() as conn:
            return [dict(r) for r in conn.execute(sql, (equipamento_id,))]

    def salvar_equipamento(self, nome, ip, tipo, equipamento_id=None):
        nome, ip, tipo = nome.strip(), ip.strip(), tipo.strip().upper()
        if not nome or tipo not in ('CAPA', 'MIOLO', 'AMBOS'):
            raise ValueError('Informe nome e tipo (CAPA, MIOLO ou AMBOS).')
        if any(ch in nome for ch in '<>:"/\\|?*') or nome.endswith('.'):
            raise ValueError('O nome do equipamento contem caracteres invalidos para o Windows.')
        try:
            ipaddress.IPv4Address(ip)
        except ValueError as exc:
            raise ValueError('Informe um endereco IP valido para o equipamento.') from exc
        try:
            with self._conectar() as conn:
                with conn:
                    if equipamento_id:
                        conn.execute('UPDATE equipamentos SET nome=?,ip=?,tipo=? WHERE id=?',
                                     (nome,ip,tipo,equipamento_id))
                        return equipamento_id
                    cur = conn.execute('INSERT INTO equipamentos(nome,ip,tipo) VALUES(?,?,?)',
                                       (nome,ip,tipo))
                    return cur.lastrowid
        except GPrintCadastroErro:
            raise

    def alterar_ativo_equipamento(self, equipamento_id, ativo):
        with self._conectar() as conn:
            with conn:
                conn.execute('UPDATE equipamentos SET ativo=? WHERE id=?', (int(bool(ativo)),equipamento_id))

    def salvar_fila(self, equipamento_id, nome, descricao='', ignorar_folha_obra=False, fila_id=None):
        nome = nome.strip()
        if not nome:
            raise ValueError('Informe o nome EXATO da fila de impressao.')
        # A flag e armazenada como metadado; NAO acrescenta comandos LPR.
        try:
            with self._conectar() as conn:
                with conn:
                    if fila_id:
                        conn.execute('''UPDATE filas SET nome=?, descricao=?, ignorar_folha_obra=?
                                        WHERE id=? AND equipamento_id=?''',
                                     (nome,descricao.strip(),int(bool(ignorar_folha_obra)),fila_id,equipamento_id))
                        return fila_id
                    cur = conn.execute('''INSERT INTO filas(equipamento_id,nome,descricao,ignorar_folha_obra)
                                          VALUES(?,?,?,?)''',
                                       (equipamento_id,nome,descricao.strip(),int(bool(ignorar_folha_obra))))
                    return cur.lastrowid
        except GPrintCadastroErro:
            raise

    def alterar_ativo_fila(self, fila_id, ativo):
        with self._conectar() as conn:
            with conn:
                conn.execute('UPDATE filas SET ativo=? WHERE id=?', (int(bool(ativo)),fila_id))

    def cadastrar_filas_em_lote(self, equipamento_id, nomes):
        """Cadastra filas novas numa unica transacao. Nao altera filas preexistentes."""
        unicos = list(dict.fromkeys(nome.strip() for nome in nomes if nome.strip()))
        if not unicos:
            raise ValueError('Cole ao menos um nome de fila.')
        if len(unicos) > 500:
            raise ValueError('Cadastre no maximo 500 filas por operacao.')
        with self._conectar() as conn:
            with conn:
                existe = conn.execute('SELECT 1 FROM equipamentos WHERE id=?',
                                      (equipamento_id,)).fetchone()
                if not existe:
                    raise ValueError('Equipamento nao encontrado no banco.')
                inseridos = 0
                for nome in unicos:
                    resultado = conn.execute("INSERT OR IGNORE INTO filas (equipamento_id, nome, descricao, ignorar_folha_obra) VALUES (?, ?, '', 0)", (equipamento_id, nome))
                    inseridos += max(resultado.rowcount, 0)
                return inseridos, len(unicos) - inseridos

    def excluir_fila(self, equipamento_id, fila_id):
        """Remove apenas uma fila do equipamento selecionado, nunca historico."""
        with self._conectar() as conn:
            with conn:
                return conn.execute('DELETE FROM filas WHERE id=? AND equipamento_id=?',
                                    (fila_id, equipamento_id)).rowcount

    def validar_filas(self, equipamento_id, filas):
        '''TXT determina nome da fila; cadastro apenas valida, nunca renomeia.'''
        disponiveis = {r['nome'] for r in self.listar_filas(equipamento_id)}
        return sorted({x for x in filas if x and x not in disponiveis})


def gprint_nome_maquina(equipamento):
    return f"{equipamento['nome']} - {equipamento['ip']}"


class GPrintGerenciadorFilas(tk.Toplevel):
    """Janela de filas de UMA maquina, visivel somente apos senha administrativa."""
    def __init__(self, cadastro, equipamento):
        super().__init__(cadastro)
        self.cadastro = cadastro
        self.repo = cadastro.app.repositorio_equipamentos
        self.equipamento = equipamento.copy()
        self.filas = {}
        self.fila_id = None
        self.title('GPrint - Filas de ' + equipamento['nome'])
        self.geometry('980x650')
        self.minsize(730, 470)
        self.transient(cadastro.winfo_toplevel())
        self.protocol('WM_DELETE_WINDOW', self.fechar)
        self._montar()
        self.atualizar()
        self.grab_set()
        self.focus_set()

    def _montar(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        ttk.Label(self, text='Filas de impressao — ' + self.equipamento['nome'],
                  font=('Segoe UI', 15, 'bold')).grid(row=0, column=0,
                  sticky='w', padx=15, pady=(12, 3))
        ttk.Label(self,
                  text='Cadastro para uso futuro. Nao altera montagens, TXT, regras LPR ou a aba Refazer.',
                  bootstyle=SECONDARY).grid(row=1, column=0, sticky='w', padx=15, pady=(0, 8))

        abas = ttk.Notebook(self)
        abas.grid(row=2, column=0, sticky='nsew', padx=12, pady=5)
        tab_unica = ttk.Frame(abas, padding=10)
        tab_lote = ttk.Frame(abas, padding=10)
        abas.add(tab_unica, text='Consultar / Cadastrar / Editar')
        abas.add(tab_lote, text='Cadastrar varias filas')
        self._aba_unica(tab_unica)
        self._aba_lote(tab_lote)

        rodape=ttk.Frame(self, padding=(12, 5))
        rodape.grid(row=3, column=0, sticky='ew')
        self.contador=ttk.Label(rodape, text='', bootstyle=SECONDARY)
        self.contador.pack(side=tk.LEFT)
        ttk.Button(rodape, text='Fechar', command=self.fechar).pack(side=tk.RIGHT)

    def _aba_unica(self, parent):
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(4, weight=1)
        self.nome_var=tk.StringVar()
        self.descricao_var=tk.StringVar()
        self.ignorar_var=tk.BooleanVar(value=False)
        ttk.Label(parent, text='Nome EXATO da fila LPR:').grid(row=0,column=0,sticky='w',padx=4,pady=5)
        ttk.Entry(parent,textvariable=self.nome_var).grid(row=0,column=1,columnspan=3,
                                                         sticky='ew',padx=4,pady=5)
        ttk.Label(parent,text='Descricao:').grid(row=1,column=0,sticky='w',padx=4,pady=5)
        ttk.Entry(parent,textvariable=self.descricao_var).grid(row=1,column=1,columnspan=3,
                                                              sticky='ew',padx=4,pady=5)
        ttk.Checkbutton(parent,text='Ignorar folha-de-obra (somente informativo)',
                        variable=self.ignorar_var).grid(row=2,column=0,columnspan=4,
                                                       sticky='w',pady=6)
        acoes=ttk.Frame(parent)
        acoes.grid(row=3,column=0,columnspan=4,sticky='ew',pady=(4,10))
        ttk.Button(acoes,text='Nova fila',command=self.nova).pack(side=tk.LEFT,padx=3)
        ttk.Button(acoes,text='Salvar fila',command=self.salvar,bootstyle=SUCCESS).pack(side=tk.LEFT,padx=3)
        ttk.Button(acoes,text='Ativar / Desativar',command=self.alternar).pack(side=tk.LEFT,padx=3)
        ttk.Button(acoes,text='Excluir fila',command=self.excluir,bootstyle=DANGER).pack(side=tk.LEFT,padx=3)
        ttk.Button(acoes,text='Atualizar lista',command=self.atualizar).pack(side=tk.RIGHT,padx=3)

        lista=ttk.Frame(parent)
        lista.grid(row=4,column=0,columnspan=4,sticky='nsew')
        lista.rowconfigure(0,weight=1)
        lista.columnconfigure(0,weight=1)
        self.tabela=ttk.Treeview(lista,columns=('nome','descricao','ignorar','ativo'),
                                 show='headings',selectmode='browse')
        for col,titulo,largura in [
                ('nome','Nome da fila',270),('descricao','Descricao',260),
                ('ignorar','Ignorar folha',110),('ativo','Situacao',90)]:
            self.tabela.heading(col,text=titulo)
            self.tabela.column(col,width=largura,anchor='w')
        self.tabela.grid(row=0,column=0,sticky='nsew')
        rolagem=ttk.Scrollbar(lista,orient=tk.VERTICAL,command=self.tabela.yview)
        rolagem.grid(row=0,column=1,sticky='ns')
        self.tabela.configure(yscrollcommand=rolagem.set)
        self.tabela.bind('<<TreeviewSelect>>',self.selecionar)

    def _aba_lote(self, parent):
        parent.columnconfigure(0,weight=1)
        parent.rowconfigure(1,weight=1)
        ttk.Label(parent,text='Cole um nome de fila por linha. Filas que ja existem serao ignoradas.').grid(
            row=0,column=0,sticky='w',pady=(4,10))
        self.texto_lote=tk.Text(parent,wrap='none',height=14,font=('Consolas',10))
        self.texto_lote.grid(row=1,column=0,sticky='nsew')
        ttk.Button(parent,text='Cadastrar filas da lista',bootstyle=SUCCESS,
                   command=self.importar).grid(row=2,column=0,sticky='e',pady=12)

    def atualizar(self):
        try:
            resultados=self.repo.listar_filas(self.equipamento['id'],incluir_inativas=True)
            self.filas={str(r['id']):r for r in resultados}
            anterior=str(self.fila_id) if self.fila_id else None
            antigos=self.tabela.get_children()
            if antigos:self.tabela.delete(*antigos)
            for r in resultados:
                self.tabela.insert('','end',iid=str(r['id']),values=(
                    r['nome'],r['descricao'],
                    'Sim' if r['ignorar_folha_obra'] else 'Nao',
                    'Ativa' if r['ativo'] else 'Inativa'))
            self.contador.config(text=f'{len(resultados)} filas cadastradas neste equipamento')
            if anterior in self.filas:
                self.tabela.selection_set(anterior)
                self.tabela.see(anterior)
                self.selecionar()
        except Exception as exc:
            messagebox.showerror('GPrint - Filas', f'Nao foi possivel consultar as filas.\n{exc}',parent=self)

    def nova(self):
        self.fila_id=None
        self.nome_var.set('')
        self.descricao_var.set('')
        self.ignorar_var.set(False)
        self.tabela.selection_remove(self.tabela.selection())

    def selecionar(self, _event=None):
        selecionados=self.tabela.selection()
        if not selecionados:return
        fila=self.filas.get(selecionados[0])
        if not fila:return
        self.fila_id=fila['id']
        self.nome_var.set(fila['nome'])
        self.descricao_var.set(fila['descricao'])
        self.ignorar_var.set(bool(fila['ignorar_folha_obra']))

    def salvar(self):
        try:
            self.fila_id=self.repo.salvar_fila(
                self.equipamento['id'],self.nome_var.get(),self.descricao_var.get(),
                self.ignorar_var.get(),self.fila_id)
            self.atualizar()
            messagebox.showinfo('GPrint - Filas','Fila salva com sucesso.',parent=self)
        except Exception as exc:
            messagebox.showerror('GPrint - Filas',str(exc),parent=self)

    def alternar(self):
        fila=self.filas.get(str(self.fila_id))
        if fila is None:
            messagebox.showwarning('GPrint - Filas','Selecione uma fila.',parent=self)
            return
        try:
            self.repo.alterar_ativo_fila(fila['id'],not bool(fila['ativo']))
            self.atualizar()
        except Exception as exc:
            messagebox.showerror('GPrint - Filas',str(exc),parent=self)

    def excluir(self):
        fila=self.filas.get(str(self.fila_id))
        if fila is None:
            messagebox.showwarning('GPrint - Filas','Selecione uma fila.',parent=self)
            return
        if not messagebox.askyesno('Confirmar exclusao',
                f"Excluir definitivamente a fila '{fila['nome']}' de {self.equipamento['nome']}?",
                parent=self):return
        try:
            self.repo.excluir_fila(self.equipamento['id'],fila['id'])
            self.nova()
            self.atualizar()
        except Exception as exc:
            messagebox.showerror('GPrint - Filas',str(exc),parent=self)

    def importar(self):
        nomes=self.texto_lote.get('1.0','end').splitlines()
        try:
            inseridos,ignorados=self.repo.cadastrar_filas_em_lote(self.equipamento['id'],nomes)
            self.atualizar()
            messagebox.showinfo('GPrint - Filas',
                f'{inseridos} fila(s) cadastrada(s).\n{ignorados} ja existente(s) ou duplicada(s).',parent=self)
        except Exception as exc:
            messagebox.showerror('GPrint - Filas',str(exc),parent=self)

    def fechar(self):
        try:self.grab_release()
        except tk.TclError:pass
        self.destroy()


class GPrintCadastroEquipamentosTab(ttk.Frame):
    """Equipamentos administrados dentro das Configuracoes protegidas por senha."""
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app=app
        self.equipamento_id=None
        self.equipamentos={}
        self._janela_filas=None
        self._montar_tela()

    def _montar_tela(self):
        self.columnconfigure(0,weight=1)
        self.rowconfigure(1,weight=1)
        formulario=ttk.LabelFrame(self,text='Cadastro central de equipamentos',padding=12)
        formulario.grid(row=0,column=0,sticky='ew',padx=10,pady=8)
        for coluna in (1,3,5):formulario.columnconfigure(coluna,weight=1)
        self.nome_var=tk.StringVar()
        self.ip_var=tk.StringVar()
        self.tipo_var=tk.StringVar(value='MIOLO')
        ttk.Label(formulario,text='Nome:').grid(row=0,column=0,sticky='w',padx=5)
        ttk.Entry(formulario,textvariable=self.nome_var).grid(row=0,column=1,sticky='ew',padx=5)
        ttk.Label(formulario,text='IP:').grid(row=0,column=2,sticky='w',padx=5)
        ttk.Entry(formulario,textvariable=self.ip_var).grid(row=0,column=3,sticky='ew',padx=5)
        ttk.Label(formulario,text='Uso:').grid(row=0,column=4,sticky='w',padx=5)
        ttk.Combobox(formulario,textvariable=self.tipo_var,
                     values=['MIOLO','CAPA','AMBOS'],state='readonly',width=12
                     ).grid(row=0,column=5,sticky='ew',padx=5)
        acoes=ttk.Frame(formulario)
        acoes.grid(row=1,column=0,columnspan=6,sticky='ew',pady=(12,0))
        ttk.Button(acoes,text='Novo equipamento',command=self.novo_equipamento).pack(side=tk.LEFT,padx=3)
        ttk.Button(acoes,text='Salvar equipamento',command=self.salvar_equipamento,
                   bootstyle=SUCCESS).pack(side=tk.LEFT,padx=3)
        ttk.Button(acoes,text='Ativar / Desativar',command=self.alternar_equipamento).pack(side=tk.LEFT,padx=3)
        ttk.Button(acoes,text='Gerenciar filas da maquina selecionada',
                   command=self.gerenciar_filas,bootstyle=PRIMARY).pack(side=tk.LEFT,padx=12)
        ttk.Button(acoes,text='Atualizar',command=self.atualizar).pack(side=tk.RIGHT,padx=3)

        painel=ttk.LabelFrame(self,text='Equipamentos cadastrados (duplo clique para ver as filas)',padding=8)
        painel.grid(row=1,column=0,sticky='nsew',padx=10,pady=5)
        painel.columnconfigure(0,weight=1)
        painel.rowconfigure(0,weight=1)
        self.lista_equip=ttk.Treeview(painel,columns=('nome','ip','tipo','ativo'),
                                      show='headings',selectmode='browse',height=15)
        for col,titulo,largura in [('nome','Equipamento',250),('ip','Endereco IP',155),
                                   ('tipo','Finalidade',130),('ativo','Situacao',100)]:
            self.lista_equip.heading(col,text=titulo)
            self.lista_equip.column(col,width=largura,anchor='w')
        self.lista_equip.grid(row=0,column=0,sticky='nsew')
        barra=ttk.Scrollbar(painel,orient=tk.VERTICAL,command=self.lista_equip.yview)
        barra.grid(row=0,column=1,sticky='ns')
        self.lista_equip.configure(yscrollcommand=barra.set)
        self.lista_equip.bind('<<TreeviewSelect>>',self.selecionar_equipamento)
        self.lista_equip.bind('<Double-1>',self._abrir_por_clique)
        ttk.Label(self,
                  text='As filas sao cadastradas separadamente por equipamento. '
                       'Nao interferem no processamento atual de Capa, Miolo ou Refazer.',
                  bootstyle=SECONDARY).grid(row=2,column=0,sticky='w',padx=15,pady=8)

    def _erro(self, exc):
        messagebox.showerror('GPrint - Equipamentos',
                             f'Nao foi possivel concluir a operacao.\n{exc}',parent=self)

    def atualizar(self):
        try:
            selecionado=str(self.equipamento_id) if self.equipamento_id else None
            itens=self.app.repositorio_equipamentos.listar_equipamentos(incluir_inativos=True)
            self.equipamentos={str(r['id']):r for r in itens}
            antigos=self.lista_equip.get_children()
            if antigos:self.lista_equip.delete(*antigos)
            for r in itens:
                self.lista_equip.insert('','end',iid=str(r['id']),values=(
                    r['nome'],r['ip'],r['tipo'],
                    'Ativo' if r['ativo'] else 'Inativo'))
            if selecionado in self.equipamentos:
                self.lista_equip.selection_set(selecionado)
                self.selecionar_equipamento()
            else:
                self.equipamento_id=None
            self.app.atualizar_listas_equipamentos()
        except Exception as exc:self._erro(exc)

    def novo_equipamento(self):
        self.equipamento_id=None
        self.nome_var.set('')
        self.ip_var.set('')
        self.tipo_var.set('MIOLO')
        selecionados=self.lista_equip.selection()
        if selecionados:self.lista_equip.selection_remove(*selecionados)

    def selecionar_equipamento(self,_event=None):
        selecao=self.lista_equip.selection()
        if not selecao:return
        row=self.equipamentos.get(selecao[0])
        if row is None:return
        self.equipamento_id=row['id']
        self.nome_var.set(row['nome'])
        self.ip_var.set(row['ip'])
        self.tipo_var.set(row['tipo'])

    def salvar_equipamento(self):
        try:
            self.equipamento_id=self.app.repositorio_equipamentos.salvar_equipamento(
                self.nome_var.get(),self.ip_var.get(),self.tipo_var.get(),self.equipamento_id)
            self.atualizar()
            messagebox.showinfo('GPrint','Equipamento salvo no cadastro central.',parent=self)
        except Exception as exc:self._erro(exc)

    def alternar_equipamento(self):
        row=self.equipamentos.get(str(self.equipamento_id))
        if not row:return
        ativo=not bool(row['ativo'])
        if not messagebox.askyesno('Confirmar',
                f"{'Ativar' if ativo else 'Desativar'} {row['nome']} em todas as estacoes?",
                parent=self):return
        try:
            self.app.repositorio_equipamentos.alterar_ativo_equipamento(row['id'],ativo)
            self.atualizar()
        except Exception as exc:self._erro(exc)

    def _abrir_por_clique(self,event):
        linha=self.lista_equip.identify_row(event.y)
        if linha:
            self.lista_equip.selection_set(linha)
            self.selecionar_equipamento()
            self.gerenciar_filas()

    def gerenciar_filas(self):
        row=self.equipamentos.get(str(self.equipamento_id))
        if row is None:
            messagebox.showwarning('GPrint - Filas',
                                   'Selecione um equipamento primeiro.',parent=self)
            return
        if self._janela_filas is not None and self._janela_filas.winfo_exists():
            self._janela_filas.lift()
            self._janela_filas.focus_set()
            return
        self._janela_filas=GPrintGerenciadorFilas(self,row)

class ConfigTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._admin_liberado = False
        self._cadastro_widget = None
        self._criar_configuracoes_internas()
        self.setup_ui()

    def _criar_configuracoes_internas(self):
        """Preferencias locais abertas; edicao de equipamentos somente com senha."""
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.abas_config = ttk.Notebook(self)
        self.abas_config.grid(row=0, column=0, sticky='nsew')
        self.tab_preferencias = ttk.Frame(self.abas_config)
        self.tab_equipamentos = ttk.Frame(self.abas_config)
        self.abas_config.add(self.tab_preferencias, text='Preferencias pessoais')
        self.abas_config.add(self.tab_equipamentos, text='Equipamentos e filas (restrito)')
        self.tab_equipamentos.columnconfigure(0, weight=1)
        self.tab_equipamentos.rowconfigure(1, weight=1)

        self._acesso_equipamentos = ttk.Frame(self.tab_equipamentos, padding=22)
        self._acesso_equipamentos.grid(row=1, column=0, sticky='nsew')
        ttk.Label(self._acesso_equipamentos, text='Cadastro central de equipamentos',
                  font=('Segoe UI', 15, 'bold')).pack(anchor='w', pady=(0, 10))
        ttk.Label(self._acesso_equipamentos,
                  text='Apenas pessoas autorizadas podem alterar impressoras, IPs e filas.\n'
                       'As preferencias pessoais estao disponiveis sem senha.',
                  justify='left').pack(anchor='w', pady=(0, 15))
        ttk.Button(self._acesso_equipamentos, text='Desbloquear cadastro',
                   command=self.solicitar_acesso_equipamentos,
                   bootstyle=PRIMARY).pack(anchor='w')

        self._barra_administrativa = ttk.Frame(self.tab_equipamentos, padding=(12, 6))
        ttk.Label(self._barra_administrativa,
                  text='Administracao central — alteracoes afetam todas as estacoes',
                  bootstyle=SECONDARY).pack(side=tk.LEFT)
        ttk.Button(self._barra_administrativa, text='Bloquear cadastro',
                   command=self.bloquear_cadastro, bootstyle=WARNING).pack(side=tk.RIGHT)
        self.abas_config.bind('<<NotebookTabChanged>>', self._ao_mudar_configuracao, add='+')

    def _ao_mudar_configuracao(self, _event=None):
        if self.abas_config.select() == str(self.tab_equipamentos):
            if self._admin_liberado:
                if self._cadastro_widget is not None:
                    self._cadastro_widget.atualizar()
            else:
                self.after_idle(self.solicitar_acesso_equipamentos)

    def mostrar_equipamentos(self):
        """Abre a subaba; solicita senha antes de exibir dados editaveis."""
        self.abas_config.select(self.tab_equipamentos)

    def solicitar_acesso_equipamentos(self):
        if self._admin_liberado:
            return True
        if (self.abas_config.select() != str(self.tab_equipamentos)
                or self.app.notebook.select() != str(self)):
            return False
        if not GPRINT_SENHA_ADMIN_EQUIPAMENTOS.strip():
            messagebox.showwarning('GPrint - Cadastro restrito',
                 'Configure GPRINT_SENHA_ADMIN_EQUIPAMENTOS no codigo antes de usar.',
                 parent=self)
            return False
        senha = simpledialog.askstring('GPrint - Acesso administrativo',
                'Informe a senha para cadastrar ou editar equipamentos e filas:',
                show='*', parent=self)
        if senha is None:
            return False
        if not hmac.compare_digest(senha, GPRINT_SENHA_ADMIN_EQUIPAMENTOS):
            messagebox.showerror('GPrint - Acesso negado',
                 'Senha incorreta. O cadastro permanece bloqueado.', parent=self)
            return False
        self._admin_liberado = True
        self._acesso_equipamentos.grid_remove()
        self._barra_administrativa.grid(row=0, column=0, sticky='ew')
        if self._cadastro_widget is None:
            self._cadastro_widget = GPrintCadastroEquipamentosTab(self.tab_equipamentos, self.app)
        self._cadastro_widget.grid(row=1, column=0, sticky='nsew')
        self._cadastro_widget.atualizar()
        return True

    def bloquear_cadastro(self):
        """Destroi o formulario, sem apagar dados do banco."""
        self._admin_liberado = False
        if self._cadastro_widget is not None:
            self._cadastro_widget.destroy()
            self._cadastro_widget = None
        self._barra_administrativa.grid_remove()
        self._acesso_equipamentos.grid(row=1, column=0, sticky='nsew')

    def setup_ui(self):
        # Preferencias continuam individuais em cada estacao.
        self.tab_preferencias.columnconfigure(0, weight=1)
        self.tab_preferencias.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.tab_preferencias)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Canvas + Scrollbar (mantido para suportar muitas opções)
        canvas = tk.Canvas(main_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Configura o grid interno do scrollable_frame
        scrollable_frame.columnconfigure(0, weight=1)

        # ================= SEÇÃO: Configurações de Miolo =================
        miolo_frame = ttk.LabelFrame(
            scrollable_frame,
            text="⚙️ Configurações de Miolo",
            padding=10,
            bootstyle=PRIMARY
        )
        miolo_frame.grid(row=0, column=0, sticky="ew", pady=5)
        miolo_frame.columnconfigure(1, weight=1)

        ttk.Label(miolo_frame, text="Sangria Padrão (mm):", width=25).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sangria_padrao_var = tk.DoubleVar(value=3.0)
        self.sangria_padrao_spin = ttk.Spinbox(
            miolo_frame,
            from_=0, to=20, increment=0.5,
            width=10,
            textvariable=self.sangria_padrao_var,
            bootstyle=PRIMARY
        )
        self.sangria_padrao_spin.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # ================= SEÇÃO: Configurações de Cores =================
        colors_frame = ttk.LabelFrame(
            scrollable_frame,
            text="🎨 Configurações de Cores",
            padding=10,
            bootstyle=PRIMARY
        )
        colors_frame.grid(row=1, column=0, sticky="ew", pady=5)
        colors_frame.columnconfigure(1, weight=1)

        # Cor Fosco
        ttk.Label(colors_frame, text="Cor Fosco:", width=20).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.cor_fosco_btn = ttk.Button(
            colors_frame, text="Selecionar",
            command=lambda: self.selecionar_cor('cor_fosco'),
            bootstyle=INFO
        )
        self.cor_fosco_btn.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.cor_fosco_preview = tk.Frame(colors_frame, width=50, height=25, relief=tk.SOLID, bg="#0000FF")
        self.cor_fosco_preview.grid(row=0, column=2, padx=5, pady=5)

        # Cor Brilho
        ttk.Label(colors_frame, text="Cor Brilho:", width=20).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.cor_brilho_btn = ttk.Button(
            colors_frame, text="Selecionar",
            command=lambda: self.selecionar_cor('cor_brilho'),
            bootstyle=INFO
        )
        self.cor_brilho_btn.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.cor_brilho_preview = tk.Frame(colors_frame, width=50, height=25, relief=tk.SOLID, bg="#FFB300")
        self.cor_brilho_preview.grid(row=1, column=2, padx=5, pady=5)

        # ================= SEÇÃO: Configurações de Fonte =================
        font_frame = ttk.LabelFrame(
            scrollable_frame,
            text="🔤 Configurações de Fonte",
            padding=10,
            bootstyle=PRIMARY
        )
        font_frame.grid(row=2, column=0, sticky="ew", pady=5)
        font_frame.columnconfigure(1, weight=1)

        ttk.Label(font_frame, text="Nome da Fonte:", width=20).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.fonte_entry = ttk.Entry(font_frame, bootstyle=PRIMARY)
        self.fonte_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(font_frame, text="Tamanho da Fonte:", width=20).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.tamanho_fonte_spin = ttk.Spinbox(
            font_frame, from_=6, to=24, increment=1,
            bootstyle=PRIMARY
        )
        self.tamanho_fonte_spin.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # ================= SEÇÃO: Configurações de Deslocamento =================
        offset_frame = ttk.LabelFrame(
            scrollable_frame,
            text="📐 Configurações de Deslocamento",
            padding=10,
            bootstyle=PRIMARY
        )
        offset_frame.grid(row=3, column=0, sticky="ew", pady=5)
        offset_frame.columnconfigure(1, weight=1)

        ttk.Label(offset_frame, text="Deslocamento Padrão (Y):", width=25).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.deslocamento_padrao_spin = ttk.Spinbox(
            offset_frame, from_=-100, to=100, increment=1,
            bootstyle=PRIMARY
        )
        self.deslocamento_padrao_spin.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(offset_frame, text="Deslocamento 28x56 (Y):", width=25).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.deslocamento_capa_spin = ttk.Spinbox(
            offset_frame, from_=-100, to=100, increment=1,
            bootstyle=PRIMARY
        )
        self.deslocamento_capa_spin.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(offset_frame, text="Deslocamento X:", width=25).grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.deslocamento_x_spin = ttk.Spinbox(
            offset_frame, from_=0, to=500, increment=1,
            bootstyle=PRIMARY
        )
        self.deslocamento_x_spin.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # ================= BOTÕES DE AÇÃO =================
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.grid(row=4, column=0, sticky="ew", pady=10)
        btn_frame.columnconfigure(0, weight=1)

        ttk.Button(
            btn_frame, text="💾 Salvar Configurações",
            command=self.salvar_configuracoes,
            bootstyle=SUCCESS, width=20
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            btn_frame, text="🔄 Restaurar Padrões",
            command=self.restaurar_padroes,
            bootstyle=WARNING, width=20
        ).pack(side=tk.LEFT, padx=5)

        # Carrega as configurações salvas
        self.carregar_configuracoes()
    def atualizar_previews(self):
        """Reaplica as cores salvas nos previews (chamado após troca de tema)"""
        if hasattr(self, 'atualizar_preview_cor'):
            # Pega as cores atuais do CONFIG (já salvas)
            cor_fosco = CONFIG.get('cor_fosco', (0, 0, 1))    # azul padrão
            cor_brilho = CONFIG.get('cor_brilho', (1, 0.7, 0)) # laranja
            self.atualizar_preview_cor('cor_fosco', cor_fosco)
            self.atualizar_preview_cor('cor_brilho', cor_brilho)
    def carregar_configuracoes(self):
        global CONFIG
        
        # Cores
        self.atualizar_preview_cor('cor_fosco', CONFIG['cor_fosco'])
        self.atualizar_preview_cor('cor_brilho', CONFIG['cor_brilho'])
        
        # Fonte
        self.fonte_entry.delete(0, END)
        self.fonte_entry.insert(0, CONFIG['fonte'])
        
        self.tamanho_fonte_spin.delete(0, END)
        self.tamanho_fonte_spin.insert(0, str(CONFIG['tamanho_fonte']))
        
        # Deslocamentos
        self.deslocamento_padrao_spin.delete(0, END)
        self.deslocamento_padrao_spin.insert(0, str(CONFIG['deslocamento_padrao']))
        
        self.deslocamento_capa_spin.delete(0, END)
        self.deslocamento_capa_spin.insert(0, str(CONFIG['deslocamento_capa_28x56']))
        
        self.deslocamento_x_spin.delete(0, END)
        self.deslocamento_x_spin.insert(0, str(CONFIG['deslocamento_x']))   
        self.sangria_padrao_spin.delete(0, "end")
        self.sangria_padrao_spin.insert(
            0,
            CONFIG.get("sangria_padrao_miolo", 3)
)
    def selecionar_cor(self, tipo_cor):
        cor_atual = CONFIG[tipo_cor]
        cor_rgb = "#%02x%02x%02x" % (
            int(cor_atual[0] * 255),
            int(cor_atual[1] * 255),
            int(cor_atual[2] * 255)
        )
        
        cor = colorchooser.askcolor(
            initialcolor=cor_rgb,
            title=f"Selecione a cor {tipo_cor}"
        )
        
        if cor[0]:
            r, g, b = cor[0]
            nova_cor = (r/255, g/255, b/255)
            CONFIG[tipo_cor] = nova_cor
            self.atualizar_preview_cor(tipo_cor, nova_cor)   
    def atualizar_preview_cor(self, tipo_cor, cor_rgb):
        cor_hex = "#%02x%02x%02x" % (
            int(cor_rgb[0] * 255),
            int(cor_rgb[1] * 255),
            int(cor_rgb[2] * 255)
        )
        
        if tipo_cor == 'cor_fosco':
            self.cor_fosco_preview.config(bg=cor_hex)
        else:
            self.cor_brilho_preview.config(bg=cor_hex)   
    def salvar_configuracoes(self):
        global CONFIG
        
        try:
            CONFIG['fonte'] = self.fonte_entry.get()
            CONFIG['tamanho_fonte'] = int(self.tamanho_fonte_spin.get())
            
            CONFIG['deslocamento_padrao'] = int(self.deslocamento_padrao_spin.get())
            CONFIG['deslocamento_capa_28x56'] = int(self.deslocamento_capa_spin.get())
            CONFIG['deslocamento_x'] = float(self.deslocamento_x_spin.get())

            # Sangria de miolo: salva em milímetros, inclusive valores decimais.
            sangria_texto = self.sangria_padrao_spin.get().strip().replace(",", ".")
            nova_sangria = float(sangria_texto)
            if not math.isfinite(nova_sangria) or not (0 <= nova_sangria <= 20):
                raise ValueError("Sangria deve estar entre 0 e 20 mm")
            CONFIG['sangria_padrao_miolo'] = nova_sangria

            if salvar_configuracoes(CONFIG):
                messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
            else:
                messagebox.showerror("Erro", "Não foi possível salvar as configurações.")
                
        except ValueError as e:
            messagebox.showerror("Erro", f"Valor inválido: {e}")   
    def restaurar_padroes(self):
        global CONFIG
        
        if messagebox.askyesno(
            "Confirmar",
            "Deseja restaurar TODAS as configurações para os valores padrão?"
        ):
            CONFIG = DEFAULT_CONFIG.copy()
            self.carregar_configuracoes()
            salvar_configuracoes(CONFIG)
            messagebox.showinfo("Sucesso", "Configurações padrão restauradas!")           
class MioloTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.processos = []
        self.arquivos_txt = []  # arquivos selecionados para a proxima execucao
        self.processando = False
        self.thread = None
        self.log_queue = queue.Queue()
        self.impressora_var = tk.StringVar()
        self.status_var = tk.StringVar(value="✅ Pronto para processar")
        
        # Caminho do banco SQLite
        self.log_db_path = GPRINT_BANCO_HISTORICO
        self.log_db_local_path = os.path.join(os.path.expanduser("~"), "Desktop", "impressoes_fallback.db")
        
        self.setup_ui()
        self.atualizar_log_queue()
        logging.info("Sistema iniciado")

    # ------------------------------------------------------------
    # Mecanismo de trava para escrita concorrente no SQLite
    # ------------------------------------------------------------
    def _adquirir_trava(self, caminho_trava, timeout=5):
        inicio = time.time()
        while True:
            try:
                fd = os.open(caminho_trava, os.O_CREAT | os.O_EXCL | os.O_RDWR)
                os.close(fd)
                return True
            except FileExistsError:
                if time.time() - inicio > timeout:
                    return False
                time.sleep(random.uniform(0.1, 0.4))
            except Exception:
                return False

    def _liberar_trava(self, caminho_trava):
        try:
            os.remove(caminho_trava)
        except:
            pass

    # ------------------------------------------------------------
    # REGISTRO DE LOG NO BANCO SQLite (sem total_paginas)
    # ------------------------------------------------------------
    def registrar_log_sqlite(self, dados):
        return gprint_registrar_log_seguro(
            dados, self.log_db_path, self.log_db_local_path,
            self.app.root, lambda msg: self.log(msg, "ERRO" if "CRITICO" in msg else "INFO")
        )

    # ------------------------------------------------------------
    # Interface gráfica (setup_ui) - layout ajustado
    # ------------------------------------------------------------
    def setup_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")

        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(3, weight=2)

        # ================= CONFIG =================
        config_frame = ttk.LabelFrame(main_frame, text="⚙️ Configurações de Impressão", bootstyle=(PRIMARY))
        config_frame.grid(row=0, column=0, sticky="ew", pady=5)

        # Agora teremos 3 colunas: label, IP, toggle (e checkboxes ocupam as 3)
        config_frame.columnconfigure(0, weight=0)  # labels
        config_frame.columnconfigure(1, weight=1)  # IP
        config_frame.columnconfigure(2, weight=1)  # toggle

        ttk.Label(config_frame, text="Impressora:").grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Maquinas compativeis sao carregadas de equipamentos.db.
        impressoras = []

        self.combo_impressora = ttk.Combobox(
            config_frame,
            textvariable=self.impressora_var,
            values=impressoras,
            state="readonly"
        )
        self.combo_impressora.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        ttk.Label(config_frame, text="IP selecionado:").grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.ip_label = ttk.Label(
            config_frame,
            text="---",
            font=("Consolas", 10, "bold"),
            bootstyle=(PRIMARY)
        )
        self.ip_label.grid(row=1, column=1, padx=5, pady=2, sticky="w")

        self.combo_impressora.bind('<<ComboboxSelected>>', self.atualizar_ip_label)

        # Toggle de impressão automática agora na MESMA LINHA do IP (coluna 2)
        self.impressao_auto_var = tk.BooleanVar(value=True)
        self.toggle_impressao = ttk.Checkbutton(
            config_frame,
            text="enviar para impressora automaticamente",
            variable=self.impressao_auto_var,
            bootstyle="success-round-toggle"
        )
        self.toggle_impressao.grid(
            row=1, column=2,
            padx=5, pady=2,
            sticky="w"
        )

        # =========================================
        # CHECKBOX COMPRIMIR PDF
        # =========================================
        self.comprimir_pdf_var = tk.BooleanVar(value=False)
        self.checkbox_comprimir = ttk.Checkbutton(
            config_frame,
            text="comprimir PDF antes do envio (Deixar selecionado somente em casos extremos)",
            variable=self.comprimir_pdf_var,
            bootstyle="info-square-toggle"
        )
        self.checkbox_comprimir.grid(
            row=2, column=0, columnspan=3,
            padx=5, pady=5, sticky="w"
        )

        # =========================================
        # NOVO CHECKBOX "Enviar apenas uma cópia"
        # =========================================
        self.enviar_uma_copia_var = tk.BooleanVar(value=False)
        self.chk_enviar_uma = ttk.Checkbutton(
            config_frame,
            text="Enviar apenas uma cópia (ignora regra padrão de divisão)",
            variable=self.enviar_uma_copia_var,
            bootstyle="info-square-toggle"
        )
        self.chk_enviar_uma.grid(
            row=3, column=0, columnspan=3,
            padx=5, pady=5, sticky="w"
        )

        # ================= ARQUIVO =================
        file_frame = ttk.LabelFrame(
            main_frame,
            text="📂 Arquivo de Lote (TXT)",
            bootstyle=(PRIMARY)
        )

        file_frame.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=5
        )

        file_frame.columnconfigure(1, weight=1)

        self.btn_select = ttk.Button(
            file_frame,
            text="Selecionar TXT(s)",
            command=self.selecionar_txt,
            bootstyle=(PRIMARY)
        )

        self.btn_select.grid(
            row=0,
            column=0,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.file_label = ttk.Label(
            file_frame,
            text="Nenhum arquivo selecionado",
            anchor="w",
            padding=10,
            bootstyle="secondary"
        )

        self.file_label.grid(
            row=0,
            column=1,
            padx=5,
            pady=5,
            sticky="ew"
        )

        # ================= TABELA =================
        preview_frame = ttk.LabelFrame(main_frame, text="📋 Processos", bootstyle=(PRIMARY))
        preview_frame.grid(row=2, column=0, sticky="nsew", pady=5)

        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        columns = ('TXT', 'Linha', 'Arquivo', 'Formato', 'Acabamento', 'Modo', 'QTD', 'Total')

        self.tree = ttk.Treeview(preview_frame, columns=columns, show='headings')

        for col in columns:
            self.tree.heading(col, text=col)

        self.tree.column('TXT', width=50, anchor="center")
        self.tree.column('Linha', width=50, anchor="center")
        self.tree.column('Arquivo', width=250)
        self.tree.column('Formato', width=80, anchor="center")
        self.tree.column('Acabamento', width=100, anchor="center")
        self.tree.column('Modo', width=140)
        self.tree.column('QTD', width=70, anchor="center")
        self.tree.column('Total', width=80, anchor="center")

        scrollbar = ttk.Scrollbar(preview_frame, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # ================= LOG =================
        log_frame = ttk.LabelFrame(main_frame, text="📝 Log de Processamento", bootstyle=(PRIMARY))
        log_frame.grid(row=3, column=0, sticky="nsew", pady=5)

        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            wrap='word',
            state='disabled',
            font=('Consolas', 9),
            bg='#1e1e1e',
            fg='#dcdcdc',
            insertbackground='white',
            padx=10,
            pady=10
        )

        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")

        # ================= ANDAMENTO =================
        progress_frame = ttk.LabelFrame(main_frame, text="⏳ Andamento", bootstyle=(SUCCESS))
        progress_frame.grid(row=4, column=0, sticky="ew", pady=5)
        progress_frame.columnconfigure(0, weight=0)
        progress_frame.columnconfigure(1, weight=1)
        progress_frame.columnconfigure(2, weight=0)

        ttk.Label(progress_frame, text="Andamento:").grid(
            row=0, column=0, padx=5, pady=5, sticky='w')
        self.progress = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            bootstyle=SUCCESS
        )
        self.progress.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        self.progress_label = ttk.Label(progress_frame, text="0/0")
        self.progress_label.grid(row=0, column=2, padx=5, pady=5, sticky='e')
        self.indicadores_miolo = GPrintIndicadoresAndamento(progress_frame)

        # ================= STATUS =================
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=5, column=0, sticky="ew", pady=5)

        status_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10, "bold"),
            bootstyle="success"
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.btn_processar = ttk.Button(
            status_frame,
            text="▶ Processar e Enviar",
            command=self.iniciar_processamento,
            bootstyle="success",
            width=22
        )
        self.btn_processar.grid(row=0, column=1, sticky="e")

    def atualizar_equipamentos(self, equipamentos):
        anterior = self.impressora_var.get()
        self.equipamentos_disponiveis = {gprint_nome_maquina(x): x for x in equipamentos}
        nomes = list(self.equipamentos_disponiveis)
        self.combo_impressora.configure(values=nomes)
        if anterior in self.equipamentos_disponiveis:
            self.impressora_var.set(anterior)
        elif len(nomes) == 1:
            self.impressora_var.set(nomes[0])
        else:
            self.impressora_var.set('')
        self.atualizar_ip_label()

    def atualizar_ip_label(self, event=None):
        equipamento = getattr(self, 'equipamentos_disponiveis', {}).get(self.impressora_var.get())
        self.ip_label.config(text=equipamento['ip'] if equipamento else '---')

    def atualizar_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.config(state='normal')
                self.log_text.insert(tk.END, f"{msg}\n")
                self.log_text.see(tk.END)
                self.log_text.config(state='disabled')
        except queue.Empty:
            pass
        finally:
            self.after(100, self.atualizar_log_queue)

    def log(self, mensagem, tipo="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        cores = {"INFO": "", "SUCESSO": "✅ ", "ERRO": "❌ ", "ENVIO": "🚀 "}
        prefixo = cores.get(tipo, "")
        msg_formatada = f"[{timestamp}] {prefixo}{mensagem}"
        self.log_queue.put(msg_formatada)
        
        if tipo == "ERRO":
            logging.error(mensagem)
        elif tipo == "SUCESSO":
            logging.info(f"✅ {mensagem}")
        elif tipo == "ENVIO":
            logging.info(f"📤 {mensagem}")
        else:
            logging.info(mensagem)

    def iniciar_processamento(self):
        equipamento = getattr(self, 'equipamentos_disponiveis', {}).get(self.impressora_var.get())
        if not equipamento:
            messagebox.showwarning('GPrint - Miolo', 'Selecione um equipamento de Miolo cadastrado em Configuracoes > Equipamentos.')
            return
        # Congela o destino desta execucao; mudancas no cadastro nao desviam envios em curso.
        self.equipamento_execucao = equipamento.copy()
        # As filas do cadastro sao informativas para Capa/Miolo.
        # O TXT determina a fila LPR como na logica operacional original.
        if not self.processos:
            messagebox.showwarning("Aviso", "Carregue um arquivo TXT primeiro.")
            return
        
        if self.processando:
            messagebox.showwarning("Aviso", "Já existe um processamento em andamento.")
            return
        
        self.processando = True
        self.btn_processar.config(state='disabled', text='⏳ Processando...')
        self.btn_select.config(state='disabled')  # nao alterar a fila durante o processamento
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.processos)
        self.indicadores_miolo.iniciar()

        for child in self.winfo_children():
            if isinstance(child, ttk.Frame):
                for frame in child.winfo_children():
                    if isinstance(frame, ttk.LabelFrame):
                        for btn in frame.winfo_children():
                            if isinstance(btn, ttk.Button) and btn['text'] == '🔄 Cancelar':
                                btn.config(state='normal')
        
        self.thread = threading.Thread(target=self._executar, daemon=True)
        self.thread.start()

    def enviar_por_lpr(self, caminho_arquivo, ip, formato, job_name, copia_numero, total_copias):
        try:
            self.log(f"📤 Enviando cópia {copia_numero}/{total_copias}...", "ENVIO")
            self.log(f"   IP: {ip}, Fila: {formato}, Job: {job_name}", "ENVIO")
            
            pasta_temp = tempfile.gettempdir()
            nome_temp = f"{uuid.uuid4()}.pdf"
            arquivo_temp = os.path.join(pasta_temp, nome_temp)

            shutil.copy2(caminho_arquivo, arquivo_temp)
            time.sleep(1)

            comando = [
                "lpr",
                "-S", ip,
                "-P", formato,
                "-J", job_name,
                arquivo_temp
            ]

            resultado = subprocess.run(
                comando,
                capture_output=True,
                text=True,
                timeout=120
            )

            try:
                os.remove(arquivo_temp)
            except:
                pass

            print(f"\n--- LPR DEBUG (cópia {copia_numero}) ---")
            print(f"Arquivo original: {caminho_arquivo}")
            print(f"Arquivo temp: {arquivo_temp}")
            print(f"Comando: {resultado.args}")
            print(f"Return code: {resultado.returncode}")
            print(f"STDOUT: {resultado.stdout}")
            print(f"STDERR: {resultado.stderr}")
            print(f"-----------------\n")

            if resultado.stdout:
                self.log(f"   Saída: {resultado.stdout}", "ENVIO")
            if resultado.stderr:
                self.log(f"   Erro: {resultado.stderr}", "ERRO")

            if resultado.returncode == 0:
                self.log(f"✅ Cópia {copia_numero} enviada com sucesso", "SUCESSO")
                return True
            else:
                self.log(f"❌ Falha no envio da cópia {copia_numero} (código {resultado.returncode})", "ERRO")
                return False

        except subprocess.TimeoutExpired:
            self.log(f"❌ Timeout no envio da cópia {copia_numero} (120s)", "ERRO")
            return False
        except Exception as e:
            self.log(f"❌ Exceção no envio: {str(e)}", "ERRO")
            print(f"Erro detalhado: {e}")
            return False

    def _executar(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        output_dir = os.path.join(desktop, "PDFs_PROCESSADOS")
        os.makedirs(output_dir, exist_ok=True)
        # Cada execução fica em uma pasta independente (evita sobrescrever
        # PDFs de outra pilha/execução e mantém a raiz organizada).
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pasta_lote = os.path.join(output_dir, f"MIOLO_{timestamp}_{uuid.uuid4().hex[:6]}")
        os.makedirs(pasta_lote, exist_ok=False)
        # Congela o parâmetro no começo do lote: mudar a configuração durante
        # a execução não produz PDFs com sangrias diferentes na mesma pilha.
        sangria_mm = obter_sangria_padrao_miolo()
        self.log(f"📐 Sangria padrão de miolo nesta execução: {sangria_mm:g} mm", "INFO")

        success_count = 0
        error_count = 0
        envio_count = 0
        envio_erro_count = 0
        salvos_count = 0
        # Arquivos de entrada ficam no TEMP do usuário, jamais na pasta dos
        # PDFs finalizados. O nome exclusivo também evita colisão de sessões.
        sessao_temp = tempfile.mkdtemp(prefix="sessao_miolo_", dir=PASTA_TRABALHO)

        impressora_selecionada = self.impressora_var.get()
        # IP vem do cadastro; nome da maquina pode conter hifens.
        equipamento_execucao = getattr(self, 'equipamento_execucao', None)
        if not equipamento_execucao:
            self.log('Equipamento nao encontrado no cadastro; processamento cancelado.', 'ERRO')
            self.processando = False
            self.after(0, lambda: self.btn_processar.configure(state='normal', text='▶ Processar e Enviar'))
            return
        ip_extraido = equipamento_execucao['ip']
        
        # ===== EXTRAÇÃO DE LOTE (CAMINHO COMPLETO) E PILHA (NOME DO TXT) =====
        if hasattr(self, 'arquivo_txt') and self.arquivo_txt:
            lote = os.path.dirname(self.arquivo_txt)
            pilha = os.path.splitext(os.path.basename(self.arquivo_txt))[0]
        else:
            lote = "desconhecido"
            pilha = "desconhecido"

        # ===== NOME DO USUÁRIO DO WINDOWS =====
        try:
            usuario = os.getlogin()
        except:
            usuario = os.environ.get('USERNAME', 'desconhecido')

        self.log(f"\n{'='*50}", "INFO")
        self.log(f"🚀 INICIANDO PROCESSAMENTO", "INFO")
        self.log(f"👤 Usuário: {usuario}", "INFO")
        self.log(f"📁 Lote (pasta): {lote}", "INFO")
        self.log(f"📄 Pilha (arquivo TXT): {pilha}", "INFO")
        self.log(f"📠 Impressora: {impressora_selecionada}", "INFO")
        self.log(f"📠 IP: {ip_extraido}", "INFO")
        self.log(f"📁 Pasta de saída: {pasta_lote}", "INFO")
        self.log(f"{'='*50}", "INFO")

        # Pasta para salvamento (envio desligado)
        nome_pasta_salvos = None
        if hasattr(self, 'nome_base_txt') and self.nome_base_txt:
            nome_pasta_salvos = self.nome_base_txt
        elif hasattr(self, 'nome_arquivo_txt') and self.nome_arquivo_txt:
            nome_pasta_salvos = os.path.splitext(self.nome_arquivo_txt)[0]
        elif hasattr(self, 'file_path') and self.file_path:
            nome_pasta_salvos = os.path.splitext(os.path.basename(self.file_path))[0]
        else:
            if self.processos:
                primeiro_nome = self.processos[0].get('NOMEFIC', 'desconhecido')
                nome_pasta_salvos = f"{primeiro_nome}_{timestamp}"
            else:
                nome_pasta_salvos = f"PDFs_SALVOS_{timestamp}"
        
        nome_pasta_salvos = re.sub(r'[<>:"/\\|?*]', '_', nome_pasta_salvos)
        # A pasta COPIAS_SALVAS só é criada se houver salvamento manual.
        # No modo de envio automático não cria diretórios vazios.
        pasta_salvos = os.path.join(pasta_lote, "COPIAS_SALVAS")
        self.log(f"📁 PDFs finalizados desta pilha ({nome_pasta_salvos}): {pasta_lote}", "INFO")

        varios_txt = len(getattr(self, 'arquivos_txt', [])) > 1
        try:
            for i, p in enumerate(self.processos):
                if not self.processando:
                    self.log("⚠️ Processamento cancelado pelo usuário", "INFO")
                    break
                
                try:
                    # Cada processo mantem a origem para log, historico e saida.
                    # A imposicao e o LPR abaixo nao sao modificados.
                    origem_txt = p.get("TXT_ORIGEM", getattr(self, "arquivo_txt", ""))
                    lote = os.path.dirname(origem_txt) if origem_txt else "desconhecido"
                    pilha = os.path.splitext(os.path.basename(origem_txt))[0] if origem_txt else "desconhecido"
                    numero_txt = p.get("TXT_NUMERO", 1)
                    # Com varios TXT, isola os PDFs por arquivo de entrada:
                    # nomes NOMEFIC iguais nao se sobrescrevem.
                    pasta_saida_txt = os.path.join(pasta_lote, f"TXT_{numero_txt:02d}") if varios_txt else pasta_lote
                    os.makedirs(pasta_saida_txt, exist_ok=True)
                    pasta_salvos = os.path.join(pasta_saida_txt, "COPIAS_SALVAS")
                    self.progress['value'] = i
                    self.progress_label.config(text=f"{i+1}/{len(self.processos)}")
                
                    self.log(f"\n--- TXT {numero_txt}: {os.path.basename(origem_txt)} | L{p['LINHA']}: {p['ARQUIVO']} ---", "INFO")
                    self.log(f"📄 Formato: {p['FORMATO']}", "INFO")
                    self.log(f"🔧 Modo: {p['MODO']}", "INFO")
                    self.log(f"📊 QTD para envio: {p['QTD']}, QTD Total: {p['QTD_total']}", "INFO")
                
                    origem = os.path.normpath(os.path.join(p["CAMINHO"], p["ARQUIVO"]))
                    if not os.path.exists(origem):
                        self.log(f"Arquivo não encontrado: {origem}", "ERRO")
                        error_count += 1
                        continue
                
                    local = os.path.join(sessao_temp, f"TXT{numero_txt:02d}_L{p['LINHA']}_{p['ARQUIVO']}")
                    shutil.copy2(origem, local)
                    self.log(f"📋 Arquivo copiado para temporário", "INFO")
                
                    output_path = os.path.join(pasta_saida_txt, f"{p['NOMEFIC']}.pdf")
                    w_pt, h_pt = p["LARGURA_MM"] * MM, p["ALTURA_MM"] * MM
                
                    self.log(f"⚙️ Gerando PDF montado...", "INFO")
                
                    if p["MODO"] == "dois_na_folha":
                        dois_na_folha(local, output_path, w_pt, h_pt, espacamento_mm=p["ESPACAMENTO"], sangria_mm=sangria_mm)
                    elif p["MODO"] == "grampo":
                        grampo(local, output_path, w_pt, h_pt, sangria_mm=sangria_mm)
                    elif p["MODO"] == "Corta_empilha":
                        Corta_empilha(local, output_path, w_pt, h_pt, sangria_mm=sangria_mm, comprimir_entrada=self.comprimir_pdf_var.get())
                    elif p["MODO"] == "corta_empilha_rotacionado":
                        corta_empilha_rotacionado(local, output_path, w_pt, h_pt, sangria_mm=sangria_mm, comprimir_entrada=self.comprimir_pdf_var.get())
                    elif p["MODO"] == "dois_na_folha_rotacionado":
                        dois_na_folha_rotacionado(local, output_path, w_pt, h_pt, espacamento_mm=p["ESPACAMENTO"], sangria_mm=sangria_mm)
                    elif p["MODO"] == "um_na_folha_atexco":
                        um_na_folha_atexco(local, output_path, w_pt, h_pt, sangria_mm=sangria_mm, comprimir_entrada=self.comprimir_pdf_var.get())
                    else:
                        um_na_folha(local, output_path, w_pt, h_pt, sangria_mm=sangria_mm)
                
                    self.log(f"✅ PDF gerado: {os.path.basename(output_path)}", "SUCESSO")
                
                    # Quantidade de envios conforme o modo de montagem
                    if p["MODO"] == "dois_na_folha" or p["MODO"] == "um_na_folha_atexco" and p["QTD"] >= 10:
                        qtd_envios = 1
                        self.log(f"🔢 Modo 1 envio/: {p['QTD']} exemplares -> 1 envio", "INFO")
                    else:
                        qtd_envios = p["QTD"]
                        self.log(f"🔢 Modo {p['MODO']}: {p['QTD']} exemplares -> {qtd_envios} envios", "INFO")

                    # Forçar apenas uma cópia caso a opção esteja marcada
                    if self.enviar_uma_copia_var.get():
                        self.log("📌 Opção 'Enviar apenas uma cópia' ativada. Enviando 1 cópia.", "INFO")
                        qtd_envios = 1
                    
                    formato_base = p["FORMATO"]
               

                    if "RICOH" in impressora_selecionada:
                        job_name = f"{p['IDENTIFIC']}_{p['NOMEFIC']}"
                    elif "BALTORO" in impressora_selecionada:
                        job_name = f"{p['NOMEFIC']}_{p['IDENTIFIC']}"
                    else:
                        job_name = p["NOMEFIC"]
                

                    # ===== REGISTRO DE MONTAGEM COM SUCESSO =====
                    dados_montagem = {
                        'lote': lote,
                        'pilha': pilha,
                        'arquivo': os.path.basename(output_path),
                        'maquina': impressora_selecionada,
                        'ip': ip_extraido,
                        'formato': p['FORMATO'],
                        'modo': p['MODO'],
                        # Um PDF final preparado. As copias enviadas sao
                        # registradas individualmente na etapa de envio.
                        'quantidade': 1,
                        'status': 'Montagem OK',
                        'erro': '',
                        'copia': '0/0',
                        'usuario': usuario,
                        'montagem_ok': 1,
                        'envio_ok': 0
                    }
                    self.registrar_log_sqlite(dados_montagem)
                
                    # Loop de envios (ou salvamentos)
                    for envio_idx in range(qtd_envios):
                        if not self.processando:
                            self.log("⚠️ Cancelando envios...", "INFO")
                            break

                        copia_num = envio_idx + 1

                        # Aguarda arquivo liberar
                        timeout_arquivo = 10
                        inicio = time.time()
                        while True:
                            try:
                                with open(output_path, "rb") as f:
                                    f.read(1)
                                break
                            except Exception:
                                if time.time() - inicio > timeout_arquivo:
                                    raise Exception("Arquivo não liberou para leitura")
                                time.sleep(0.5)

                        if self.impressao_auto_var.get():
                            self.log("🖨️ Impressão automática ativada", "INFO")
                            sucesso = self.enviar_por_lpr(
                                output_path,
                                ip_extraido,
                                formato_base,
                                job_name,
                                copia_num,
                                qtd_envios
                            )
                            if sucesso:
                                envio_count += 1
                                status = "Enviado"
                                erro = ""
                                envio_ok_val = 1
                            else:
                                envio_erro_count += 1
                                status = "Falha no envio"
                                erro = "Falha no envio via LPR"
                                envio_ok_val = 0
                        else:
                            # MODO DESLIGADO: SALVA O PDF
                            self.log(f"💾 Envio desligado. Salvando cópia {copia_num}/{qtd_envios} em {pasta_salvos}", "INFO")
                            os.makedirs(pasta_salvos, exist_ok=True)
                            nome_base_pdf = os.path.splitext(os.path.basename(output_path))[0]
                            nome_copia = f"{nome_base_pdf}_c{copia_num}.pdf"
                            destino = os.path.join(pasta_salvos, nome_copia)
                            try:
                                shutil.copy2(output_path, destino)
                                self.log(f"✅ PDF salvo: {destino}", "SUCESSO")
                                salvos_count += 1
                                sucesso = True
                                status = "Salvo"
                                erro = ""
                            except Exception as e:
                                self.log(f"❌ Erro ao salvar: {e}", "ERRO")
                                sucesso = False
                                status = "Falha ao salvar"
                                erro = str(e)
                            envio_ok_val = 0   # salvar não conta como envio real

                        # ===== REGISTRO DE ENVIO/SALVAMENTO =====
                        dados_log = {
                            'lote': lote,
                            'pilha': pilha,
                            'arquivo': os.path.basename(output_path),
                            'maquina': impressora_selecionada,
                            'ip': ip_extraido,
                            'formato': p['FORMATO'],
                            'modo': p['MODO'],
                            # Cada linha representa UM envio LPR (ou uma
                            # copia salva manualmente), nunca o total do grupo.
                            # Quando um unico job encapsula varios exemplares,
                            # o total fisico deve ser apurado separadamente.
                            'quantidade': 1,
                            'status': status,
                            'erro': erro,
                            'copia': f"{copia_num}/{qtd_envios}",
                            'usuario': usuario,
                            'montagem_ok': 0,
                            'envio_ok': envio_ok_val
                        }
                        self.registrar_log_sqlite(dados_log)

                        time.sleep(1.5)
                
                    success_count += 1
               
                except Exception as e:
                    self.log(f"❌ Erro: {str(e)}", "ERRO")
                    error_count += 1
                    logging.error(f"Erro detalhado: {e}", exc_info=True)
                
                    # ===== REGISTRO DE FALHA NA MONTAGEM =====
                    dados_erro_montagem = {
                        'lote': lote,
                        'pilha': pilha,
                        'arquivo': p.get('NOMEFIC', 'desconhecido'),
                        'maquina': impressora_selecionada,
                        'ip': ip_extraido,
                        'formato': p.get('FORMATO', ''),
                        'modo': p.get('MODO', ''),
                        'quantidade': 0,
                        'status': 'Montagem Erro',
                        'erro': str(e),
                        'copia': '0/0',
                        'usuario': usuario,
                        'montagem_ok': 0,
                        'envio_ok': 0
                    }
                    self.registrar_log_sqlite(dados_erro_montagem)

        finally:
            # Só remove os insumos temporários desta execução; preserva os PDFs gerados.
            try:
                shutil.rmtree(sessao_temp)
                self.log("🧹 Temporários do miolo removidos", "INFO")
            except OSError as erro_limpeza:
                self.log(f"⚠️ Não foi possível remover temporários: {erro_limpeza}", "INFO")

        
        # Registros foram guardados localmente durante o processamento.
        # Uma unica sincronizacao de lotes, sem entrar no banco por PDF/copia.
        gprint_sincronizar_pendentes(
            self.log_db_path, root=self.app.root,
            emitir_log=lambda msg: self.log(msg, "INFO")
        )

        self.log(f"\n{'='*50}", "INFO")
        self.log(f"📊 RESUMO FINAL", "INFO")
        self.log(f"{'='*50}", "INFO")
        self.log(f"✅ Montagens com sucesso: {success_count}", "SUCESSO")
        self.log(f"❌ Montagens com erro: {error_count}", "ERRO" if error_count > 0 else "INFO")
        self.log(f"🚀 Envios com sucesso: {envio_count}", "SUCESSO")
        self.log(f"❌ Envios com erro: {envio_erro_count}", "ERRO" if envio_erro_count > 0 else "INFO")
        self.log(f"💾 PDFs salvos (envio desligado): {salvos_count}", "INFO")
        self.log(f"{'='*50}", "INFO")
        
        self.status_var.set(f"Concluído: {success_count} OK, {error_count} Erros | Enviados: {envio_count} | Salvos: {salvos_count}")
        
        self.processando = False
        self.after(0, self.indicadores_miolo.finalizar)
        self.btn_processar.config(state='normal', text='▶ Processar e Enviar')
        self.btn_select.config(state='normal')
        # Exibe o andamento concluido ate a proxima execucao.
        self.progress['value'] = len(self.processos)
        self.progress_label.config(text=f"{len(self.processos)}/{len(self.processos)}")
        
        for child in self.winfo_children():
            if isinstance(child, ttk.Frame):
                for frame in child.winfo_children():
                    if isinstance(frame, ttk.LabelFrame):
                        for btn in frame.winfo_children():
                            if isinstance(btn, ttk.Button) and btn['text'] == '🔄 Cancelar':
                                btn.config(state='disabled')
        
        messagebox.showinfo("Concluído", 
                        f"Processamento Finalizado!\n\n"
                        f"✅ Montagens: {success_count}\n"
                        f"❌ Erros: {error_count}\n"
                        f"🚀 Envios: {envio_count}\n"
                        f"💾 PDFs salvos: {salvos_count}\n"
                        f"📁 Pasta: {pasta_lote}\n"
                        f"📂 Salvos em: {(pasta_lote if varios_txt else pasta_salvos) if salvos_count else 'Nenhuma cópia manual'}")


    def selecionar_txt(self):
        """Seleciona ate MAX_ARQUIVOS_TXT_MIOLO arquivos (0 = ilimitado)."""
        if self.processando:
            messagebox.showwarning("Miolo em processamento", "Aguarde o processamento atual terminar.")
            return
        selecionados = filedialog.askopenfilenames(
            title="Selecionar arquivos TXT de miolo",
            filetypes=[("Arquivos TXT", "*.txt")]
        )
        if not selecionados:
            return  # Cancelar nao altera a selecao anterior
        arquivos = list(dict.fromkeys(os.path.normpath(p) for p in selecionados))
        if MAX_ARQUIVOS_TXT_MIOLO > 0 and len(arquivos) > MAX_ARQUIVOS_TXT_MIOLO:
            messagebox.showwarning(
                "Limite de arquivos TXT",
                f"Selecione no maximo {MAX_ARQUIVOS_TXT_MIOLO} arquivo(s) TXT por vez.\n"
                f"Voce selecionou {len(arquivos)}.\n\n"
                "Altere MAX_ARQUIVOS_TXT_MIOLO no inicio do codigo para mudar o limite "
                "(0 = ilimitado)."
            )
            return
        self.ler_arquivo_txt_miolo(arquivos)

    def ler_arquivo_txt_miolo(self, arquivos=None):
        """Carrega todos os TXT mantendo cada linha vinculada ao seu TXT de origem."""
        try:
            if self.processando:
                raise RuntimeError("Aguarde o processamento atual terminar.")
            if arquivos is None:
                arquivos = getattr(self, "arquivos_txt", []) or (
                    [self.arquivo_txt] if getattr(self, "arquivo_txt", None) else []
                )
            arquivos = list(dict.fromkeys(os.path.normpath(p) for p in arquivos))
            if not arquivos:
                return False
            if MAX_ARQUIVOS_TXT_MIOLO > 0 and len(arquivos) > MAX_ARQUIVOS_TXT_MIOLO:
                raise ValueError(f"O limite atual e de {MAX_ARQUIVOS_TXT_MIOLO} arquivos TXT.")

            novos_processos = []
            # Somente substitui a fila atual quando TODOS os arquivos forem lidos.
            for numero_txt, caminho_txt in enumerate(arquivos, start=1):
                lines = None
                encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'latin-1']
                for enc in encodings:
                    try:
                        with open(caminho_txt, 'r', encoding=enc) as f:
                            lines = f.readlines()
                        self.log(f"TXT {numero_txt}: {os.path.basename(caminho_txt)} (encoding: {enc})", "INFO")
                        break
                    except UnicodeDecodeError:
                        continue
                if lines is None:
                    raise ValueError(f"Nao foi possivel identificar a codificacao do TXT: {caminho_txt}")

                for idx, linha in enumerate(lines):
                    linha = linha.strip()
                    if not linha or linha.startswith('#'):
                        continue

                    campos = linha.split('|')
                    if len(campos) < 13:
                        continue

                    cod_acab = campos[9].strip()
                    mapa_acab = {
                        "1": "Corta e Empilha / 2 na folha",
                        "2": "Grampo",
                        "3": "1 na folha",
                        "4": "Corta e Empilha Rotacionado"
                    }
                    desc_acab = mapa_acab.get(cod_acab, "Desconhecido")

                    qtd = int(campos[2].strip()) if campos[2].strip().isdigit() else 1
                    qtd_total_str = campos[10].strip() if len(campos) > 10 else "0"
                    try:
                        qtd_total = int(qtd_total_str)
                    except:
                        qtd_total = 0

                    processo = {
                        "LINHA": idx + 1,
                        "CAMINHO": os.path.normpath(campos[0].strip().replace('\ufeff', '').replace('/', '\\')),
                        "ARQUIVO": campos[1].strip(),
                        "QTD": qtd,
                        "CODIGOACAB": cod_acab,
                        "QTD_total": qtd_total,
                        "NOMEFIC": campos[11].strip(),
                        "FORMATO": campos[12].strip(),
                        "IDENTIFIC": campos[13].strip() if len(campos) > 13 else ""
                    }

                    if cod_acab == "1":
                        modo = "corta_empilha_rotacionado" if qtd < 10 else "dois_na_folha"
                    elif cod_acab == "2":
                        modo = "grampo"
                    elif cod_acab == "3":
                        modo = "um_na_folha"
                    elif cod_acab == "4":
                        modo = "corta_empilha_rotacionado" if qtd < 10 else "dois_na_folha"
                    elif cod_acab == "5":
                        modo = "um_na_folha_atexco"
                    elif cod_acab == "6":
                        modo = "dois_na_folha_rotacionado"
                    else:
                        modo = "Corta_empilha"

                    partes = processo["FORMATO"].split('_')
                    if len(partes) >= 2:
                        formato_str = partes[1]
                    else:
                        formato_str = partes[0]

                    formato_limpo = formato_str.lower().replace(',', '.')
                    try:
                        if 'x' in formato_limpo:
                            dim_parts = formato_limpo.split('x')
                        elif 'X' in formato_limpo:
                            dim_parts = formato_limpo.lower().split('x')
                        else:
                            raise ValueError("Formato não contém 'x'")

                        dim1 = float(dim_parts[0])
                        dim2 = float(dim_parts[1])

                        if cod_acab == "3":
                            largura_mm = min(dim1, dim2) * 10
                            altura_mm = max(dim1, dim2) * 10
                            espacamento = 0
                        else:
                            largura_mm = max(dim1, dim2) * 10
                            altura_mm = min(dim1, dim2) * 10
                            if (int(round(dim1)) == 22 and int(round(dim2)) == 32) or \
                               (int(round(dim1)) == 32 and int(round(dim2)) == 22):
                                espacamento = 3
                            else:
                                espacamento = 0

                        processo.update({
                            "LARGURA_MM": largura_mm,
                            "ALTURA_MM": altura_mm,
                            "ESPACAMENTO": espacamento,
                            "MODO": modo,
                            "DESC_ACAB": desc_acab
                        })

                    except Exception as e:
                        self.log(f"Erro ao converter formato da linha {idx+1}: {e} - usando valores padrão", "AVISO")
                        processo.update({
                            "LARGURA_MM": 480,
                            "ALTURA_MM": 320,
                            "ESPACAMENTO": 0,
                            "MODO": modo,
                            "DESC_ACAB": desc_acab
                        })

                    processo["TXT_ORIGEM"] = caminho_txt
                    processo["TXT_NUMERO"] = numero_txt
                    novos_processos.append(processo)


            # Confirmacao atomica da nova fila de trabalhos: nao deixa selecao parcial.
            self.processos = novos_processos
            self.arquivos_txt = arquivos
            self.arquivo_txt = arquivos[0]  # compatibilidade com fluxo de TXT unico
            self.nome_base_txt = os.path.splitext(os.path.basename(arquivos[0]))[0]

            for item in self.tree.get_children():
                self.tree.delete(item)
            for processo in self.processos:
                self.tree.insert('', 'end', values=(
                    processo['TXT_NUMERO'],
                    f"L{processo['LINHA']}",
                    os.path.basename(processo['ARQUIVO']),
                    processo['FORMATO'].split('_')[1].upper() if len(processo['FORMATO'].split('_')) >= 2 else processo['FORMATO'].upper(),
                    processo['CODIGOACAB'],
                    processo['MODO'],
                    processo['QTD'],
                    processo['QTD_total']
                ))
            quantidade = len(arquivos)
            self.file_label.config(text=(
                os.path.basename(arquivos[0]) if quantidade == 1
                else f"{quantidade} TXT selecionados | {len(self.processos)} processos"
            ))
            self.status_var.set(f"✅ {len(self.processos)} processos de {quantidade} TXT carregados")
            self.log(f"✅ {quantidade} TXT carregado(s), total {len(self.processos)} processos", "SUCESSO")
            for numero, caminho in enumerate(arquivos, start=1):
                self.log(f"   TXT {numero}: {caminho}", "INFO")
            return True
        except Exception as e:
            messagebox.showerror("Erro ao carregar TXT", f"Nao foi possivel carregar todos os TXT:\n{e}")
            self.log(f"ERRO ao carregar TXT: {e}", "ERRO")
            return False

def get_border_pixmap(page, rect, zoom=2.0):
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, clip=rect)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img
def espelhar_borda(img, flip_horizontal=False, flip_vertical=False):
    if flip_horizontal:
        img = img.transpose(Image.FLIP_LEFT_RIGHT)
    if flip_vertical:
        img = img.transpose(Image.FLIP_TOP_BOTTOM)
    return img
def rotacionar_imagem(img, angle):
    """Rotaciona a imagem PIL por um ângulo múltiplo de 90°."""
    if angle == 90:
        return img.transpose(Image.ROTATE_90)
    elif angle == 180:
        return img.transpose(Image.ROTATE_180)
    elif angle == 270:
        return img.transpose(Image.ROTATE_270)
    return img
def inserir_imagem_na_pagina(page, img, rect):
    with io.BytesIO() as output:
        img.save(output, format='PNG')
        data = output.getvalue()
    page.insert_image(rect, stream=data)
def desenhar_marcas(page, trim, tam, gap=2*MM, lado=None):
    x0, y0, x1, y1 = trim
    s = page.new_shape()
    if lado in (None, 'esquerda'):
        s.draw_line((x0, y0 - tam), (x0, y0 - gap))
        s.draw_line((x0 - tam, y0), (x0 - gap, y0))
        s.draw_line((x0, y1 + tam), (x0, y1 + gap))
        s.draw_line((x0 - tam, y1), (x0 - gap, y1))
    if lado in (None, 'direita'):
        s.draw_line((x1, y0 - tam), (x1, y0 - gap))
        s.draw_line((x1 + tam, y0), (x1 + gap, y0))
        s.draw_line((x1, y1 + tam), (x1, y1 + gap))
        s.draw_line((x1 + tam, y1), (x1 + gap, y1))
    s.finish(width=0.5)
    s.commit()
def desenhar_marca_central_vertical(page, x, y0, y1, tam=5*MM, gap=2*MM):
    s = page.new_shape()
    s.draw_line((x, y0 - tam), (x, y0 - gap))
    s.draw_line((x, y1 + tam), (x, y1 + gap))
    s.finish(width=0.5)
    s.commit()
def aplicar_sangria_espelhada_externa(page, nova, trim_original, destino, sangria_pt, bordas, rotate=0):
    """
    Extensão de sangria com espelhamento externo, com suporte a rotação (0, 90, 180, 270).
    Para rotate=180, as bordas e cantos são mapeados de forma invertida.
    """
    zoom = 2.0

    # Mapeamento de bordas para rotação 180°
    mapa_bordas = {}
    if rotate == 180:
        mapa_bordas = {'top': 'bottom', 'bottom': 'top', 'left': 'right', 'right': 'left'}

    # Processa cada borda solicitada
    for borda_destino in bordas:
        # Determina qual borda da página original deve ser usada
        borda_orig = mapa_bordas.get(borda_destino, borda_destino)

        # Define retângulo de extração e flips de espelhamento conforme a borda original
        if borda_orig == 'top':
            rect = fitz.Rect(trim_original.x0, trim_original.y0,
                             trim_original.x1, trim_original.y0 + sangria_pt)
            flip_v, flip_h = True, False
        elif borda_orig == 'bottom':
            rect = fitz.Rect(trim_original.x0, trim_original.y1 - sangria_pt,
                             trim_original.x1, trim_original.y1)
            flip_v, flip_h = True, False
        elif borda_orig == 'left':
            rect = fitz.Rect(trim_original.x0, trim_original.y0,
                             trim_original.x0 + sangria_pt, trim_original.y1)
            flip_v, flip_h = False, True
        elif borda_orig == 'right':
            rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y0,
                             trim_original.x1, trim_original.y1)
            flip_v, flip_h = False, True
        else:
            continue

        img = get_border_pixmap(page, rect, zoom)

        # Aplica rotação se necessário (depois da extração)
        if rotate != 0:
            img = rotacionar_imagem(img, rotate)

        # Espelha conforme a borda original
        img = espelhar_borda(img, flip_horizontal=flip_h, flip_vertical=flip_v)
        img = img.resize((int(rect.width), int(rect.height)), Image.Resampling.LANCZOS)

        # Define o retângulo de destino na página nova (já considera a orientação final)
        if borda_destino == 'top':
            dest_rect = fitz.Rect(destino.x0, destino.y0 - sangria_pt, destino.x1, destino.y0)
        elif borda_destino == 'bottom':
            dest_rect = fitz.Rect(destino.x0, destino.y1, destino.x1, destino.y1 + sangria_pt)
        elif borda_destino == 'left':
            dest_rect = fitz.Rect(destino.x0 - sangria_pt, destino.y0, destino.x0, destino.y1)
        elif borda_destino == 'right':
            dest_rect = fitz.Rect(destino.x1, destino.y0, destino.x1 + sangria_pt, destino.y1)
        else:
            continue

        inserir_imagem_na_pagina(nova, img, dest_rect)

    # --- Cantos ---
    mapa_cantos = {}
    if rotate == 180:
        mapa_cantos = {
            ('top', 'left'): ('bottom', 'right'),
            ('top', 'right'): ('bottom', 'left'),
            ('bottom', 'left'): ('top', 'right'),
            ('bottom', 'right'): ('top', 'left')
        }

    combinacoes = [('top','left'), ('top','right'), ('bottom','left'), ('bottom','right')]
    for borda1, borda2 in combinacoes:
        if borda1 in bordas and borda2 in bordas:
            b1_orig, b2_orig = mapa_cantos.get((borda1, borda2), (borda1, borda2))

            if b1_orig == 'top' and b2_orig == 'left':
                rect = fitz.Rect(trim_original.x0, trim_original.y0,
                                 trim_original.x0 + sangria_pt, trim_original.y0 + sangria_pt)
            elif b1_orig == 'top' and b2_orig == 'right':
                rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y0,
                                 trim_original.x1, trim_original.y0 + sangria_pt)
            elif b1_orig == 'bottom' and b2_orig == 'left':
                rect = fitz.Rect(trim_original.x0, trim_original.y1 - sangria_pt,
                                 trim_original.x0 + sangria_pt, trim_original.y1)
            elif b1_orig == 'bottom' and b2_orig == 'right':
                rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y1 - sangria_pt,
                                 trim_original.x1, trim_original.y1)
            else:
                continue

            img = get_border_pixmap(page, rect, zoom)
            if rotate != 0:
                img = rotacionar_imagem(img, rotate)
            # Cantos sempre usam ambos os flips
            img = espelhar_borda(img, flip_horizontal=True, flip_vertical=True)
            img = img.resize((int(sangria_pt), int(sangria_pt)), Image.Resampling.LANCZOS)

            if borda1 == 'top' and borda2 == 'left':
                dest_rect = fitz.Rect(destino.x0 - sangria_pt, destino.y0 - sangria_pt, destino.x0, destino.y0)
            elif borda1 == 'top' and borda2 == 'right':
                dest_rect = fitz.Rect(destino.x1, destino.y0 - sangria_pt, destino.x1 + sangria_pt, destino.y0)
            elif borda1 == 'bottom' and borda2 == 'left':
                dest_rect = fitz.Rect(destino.x0 - sangria_pt, destino.y1, destino.x0, destino.y1 + sangria_pt)
            elif borda1 == 'bottom' and borda2 == 'right':
                dest_rect = fitz.Rect(destino.x1, destino.y1, destino.x1 + sangria_pt, destino.y1 + sangria_pt)
            else:
                continue

            inserir_imagem_na_pagina(nova, img, dest_rect)
def pagina_tem_conteudo(page):
    if page.get_contents(): return True
    if page.get_text("text").strip(): return True
    if page.get_images(): return True
    return False
def obter_area_corte_segura(page, w_ref, h_ref):
    if page.trimbox and page.trimbox.width > 0 and page.trimbox.height > 0: return page.trimbox
    rect = page.rect
    if rect.width > 0 and rect.height > 0: return rect
    return fitz.Rect(0, 0, w_ref, h_ref)
def calcular_y_com_sangria(altura_pt, h, sangria_pt):
    if h + 2 * sangria_pt > altura_pt:
        return (altura_pt - h) / 2, False
    else:
        y_center = (altura_pt - h) / 2
        y = max(sangria_pt, min(y_center, altura_pt - h - sangria_pt))
        return y, True
def ajustar_para_media_box(rect, largura, altura):
    x0 = max(0, min(rect.x0, largura - 1))
    y0 = max(0, min(rect.y0, altura - 1))
    x1 = max(x0 + 1, min(rect.x1, largura))
    y1 = max(y0 + 1, min(rect.y1, altura))
    return fitz.Rect(x0, y0, x1, y1)
def corta_empilha_rotacionado(input_path, output_path, output_width, output_height,sangria_mm=3, add_marcas=True,texto_offset_normal_mm=0.5, texto_offset_180_mm=0.5,comprimir_entrada=False):

    # Se solicitado, comprime o PDF de entrada num arquivo temporário
    if comprimir_entrada:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_input = tmp.name
        comprimir_pdf(input_path, tmp_input)
        doc = fitz.open(tmp_input)
    else:
        doc = fitz.open(input_path)

    if len(doc) == 0:
        print("Erro: O documento PDF está vazio.")
        doc.close()
        return False

    nome_arquivo = os.path.basename(input_path)

    # ------------------------------------------------------------
    # 1. Determinar tamanho padronizado para páginas em branco
    # ------------------------------------------------------------
    def _tamanho_padronizado(doc):
        larguras = []
        alturas = []
        for i in range(len(doc)):
            page = doc[i]
            # Tenta obter trim box já definida, ou media box
            trim = page.trimbox
            if trim.is_infinite or trim.is_empty:
                trim = page.mediabox
            larguras.append(trim.width)
            alturas.append(trim.height)
        if not larguras:
            return 210 * MM, 297 * MM  # fallback A4
        # Usa mediana para evitar outliers
        largura_pad = sorted(larguras)[len(larguras)//2]
        altura_pad = sorted(alturas)[len(alturas)//2]
        return largura_pad, altura_pad

    largura_padrao, altura_padrao = _tamanho_padronizado(doc)

    # ------------------------------------------------------------
    # 2. Ajustar total de páginas para múltiplo de 4
    # ------------------------------------------------------------
    total_original = len(doc)
    if total_original % 2 != 0:
        doc.new_page(width=largura_padrao, height=altura_padrao)
        # Define trimbox igual à mediabox
        doc[-1].set_trimbox(doc[-1].mediabox)
        print(f"Total ímpar ({total_original}). Adicionada 1 página em branco (tamanho padronizado).")

    if len(doc) % 4 != 0:
        for _ in range(2):
            doc.new_page(width=largura_padrao, height=altura_padrao)
            doc[-1].set_trimbox(doc[-1].mediabox)
        print(f"Total não múltiplo de 4. Adicionadas 2 páginas em branco (tamanho padronizado).")

    out = fitz.open()
    sangria_pt = sangria_mm * MM
    offset_normal_pt = texto_offset_normal_mm * MM
    offset_180_pt = texto_offset_180_mm * MM
    espaco_horizontal_pt = 0 * MM

    total = len(doc)

    ref_page = doc.load_page(0)
    if ref_page is None:
        ref_page = doc[0]
    if ref_page is None:
        print("Erro: não foi possível carregar a primeira página.")
        doc.close()
        out.close()
        return False

    trim_ref = obter_area_corte_segura(ref_page, 210 * MM, 297 * MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    # =========================================================================
    # Função de identificação (mantida)
    # =========================================================================
    def _adicionar_identificacao(nova_page, texto_esq, texto_dir, destino_esq, destino_dir, rot_esq, rot_dir):
        font_size = 5
        altura_texto = 4 * MM
        margem_min = 0.5 * MM

        if texto_esq:
            if rot_esq == 180:
                y0 = destino_esq.y1 + offset_180_pt
                y1 = y0 + altura_texto
                if y1 > nova_page.rect.height - margem_min:
                    y1 = nova_page.rect.height - margem_min
                    y0 = y1 - altura_texto
                rect = fitz.Rect(destino_esq.x0, y0, destino_esq.x1, y1)
                nova_page.insert_textbox(rect, texto_esq, fontsize=font_size, fontname="helv",
                                         align=fitz.TEXT_ALIGN_CENTER, rotate=180)
            else:
                y1 = destino_esq.y0 - offset_normal_pt
                y0 = y1 - altura_texto
                if y0 < margem_min:
                    y0 = margem_min
                    y1 = y0 + altura_texto
                rect = fitz.Rect(destino_esq.x0, y0, destino_esq.x1, y1)
                nova_page.insert_textbox(rect, texto_esq, fontsize=font_size, fontname="helv",
                                         align=fitz.TEXT_ALIGN_CENTER, rotate=0)

        if texto_dir:
            if rot_dir == 180:
                y0 = destino_dir.y1 + offset_180_pt
                y1 = y0 + altura_texto
                if y1 > nova_page.rect.height - margem_min:
                    y1 = nova_page.rect.height - margem_min
                    y0 = y1 - altura_texto
                rect = fitz.Rect(destino_dir.x0, y0, destino_dir.x1, y1)
                nova_page.insert_textbox(rect, texto_dir, fontsize=font_size, fontname="helv",
                                         align=fitz.TEXT_ALIGN_CENTER, rotate=180)
            else:
                y1 = destino_dir.y0 - offset_normal_pt
                y0 = y1 - altura_texto
                if y0 < margem_min:
                    y0 = margem_min
                    y1 = y0 + altura_texto
                rect = fitz.Rect(destino_dir.x0, y0, destino_dir.x1, y1)
                nova_page.insert_textbox(rect, texto_dir, fontsize=font_size, fontname="helv",
                                         align=fitz.TEXT_ALIGN_CENTER, rotate=0)

    # =========================================================================
    # Função que monta UMA face
    # =========================================================================
    def _montar_uma_face(idx_esq, idx_dir, rot_esq, rot_dir, esq_superior, lado):
        page_esq = doc[idx_esq]
        page_dir = doc[idx_dir]

        trim_esq = obter_area_corte_segura(page_esq, w_ref, h_ref)
        w_esq, h_esq = trim_esq.width, trim_esq.height
        trim_dir = obter_area_corte_segura(page_dir, w_ref, h_ref)
        w_dir, h_dir = trim_dir.width, trim_dir.height

        MARGEM_DESEJADA = 5 * MM
        MARGEM_MINIMA = 5 * MM
        ALTURA_TEXTO = 4 * MM

        if esq_superior:
            rot_sup = rot_esq
            rot_inf = rot_dir
            h_sup, h_inf = h_esq, h_dir
        else:
            rot_sup = rot_dir
            rot_inf = rot_esq
            h_sup, h_inf = h_dir, h_esq

        precisa_acima = (rot_sup == 0)
        precisa_abaixo = (rot_inf == 180)

        margem_superior_necessaria = MARGEM_DESEJADA
        margem_inferior_necessaria = MARGEM_DESEJADA

        if precisa_acima:
            margem_superior_necessaria = max(MARGEM_DESEJADA, offset_normal_pt + ALTURA_TEXTO)
        if precisa_abaixo:
            margem_inferior_necessaria = max(MARGEM_DESEJADA, offset_180_pt + ALTURA_TEXTO)

        altura_total_paginas = h_sup + h_inf
        espaco_disponivel = output_height - altura_total_paginas

        margem_superior = margem_superior_necessaria
        margem_inferior = margem_inferior_necessaria
        soma_margens = margem_superior + margem_inferior

        if soma_margens > espaco_disponivel:
            fator = espaco_disponivel / soma_margens if soma_margens > 0 else 1
            margem_superior = max(MARGEM_MINIMA, margem_superior * fator)
            margem_inferior = max(MARGEM_MINIMA, margem_inferior * fator)
            if margem_superior + margem_inferior > espaco_disponivel:
                margem_superior = MARGEM_MINIMA
                margem_inferior = max(MARGEM_MINIMA, espaco_disponivel - margem_superior)
            if espaco_disponivel < 0:
                print(f"Aviso: páginas maiores que a folha. Sobreposição inevitável.")

        y_sup = margem_superior
        y_inf = output_height - margem_inferior - h_inf

        if esq_superior:
            y_esq, y_dir = y_sup, y_inf
        else:
            y_esq, y_dir = y_inf, y_sup

        # CENTRALIZAÇÃO HORIZONTAL (CORRIGIDA)
        largura_total = w_esq + w_dir + espaco_horizontal_pt + 2 * sangria_pt
        x_inicio = (output_width - largura_total) / 2
        x1 = x_inicio + sangria_pt

        if largura_total > output_width:
            print(f"Aviso: largura total ({largura_total/MM:.2f} mm) excede a folha ({output_width/MM:.2f} mm). Centralizado, mas haverá corte nas laterais.")

        destino_esq = fitz.Rect(x1, y_esq, x1 + w_esq, y_esq + h_esq)
        destino_dir = fitz.Rect(x1 + w_esq + espaco_horizontal_pt, y_dir,
                                x1 + w_esq + espaco_horizontal_pt + w_dir, y_dir + h_dir)

        mediabox = fitz.Rect(0, 0, output_width, output_height)
        destino_esq = destino_esq.intersect(mediabox)
        destino_dir = destino_dir.intersect(mediabox)

        nova = out.new_page(width=output_width, height=output_height)

        # 1. Conteúdo das páginas
        if pagina_tem_conteudo(page_esq):
            nova.show_pdf_page(destino_esq, doc, idx_esq, clip=trim_esq, rotate=rot_esq)
        if pagina_tem_conteudo(page_dir):
            nova.show_pdf_page(destino_dir, doc, idx_dir, clip=trim_dir, rotate=rot_dir)

        # 2. Sangria (com definição correta de cabe_esq/cabe_dir)
        if sangria_pt > 0:
            cabe_esq = (destino_esq.x0 >= sangria_pt and destino_esq.x1 <= output_width - sangria_pt)
            cabe_dir = (destino_dir.x0 >= sangria_pt and destino_dir.x1 <= output_width - sangria_pt)

            if cabe_esq and pagina_tem_conteudo(page_esq):
                aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino_esq, sangria_pt,
                                                bordas=['top', 'bottom', 'left'], rotate=rot_esq)
            if cabe_dir and pagina_tem_conteudo(page_dir):
                aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino_dir, sangria_pt,
                                                bordas=['top', 'bottom', 'right'], rotate=rot_dir)

            bleed_total = fitz.Rect(
                destino_esq.x0 - sangria_pt,
                min(destino_esq.y0, destino_dir.y0) - sangria_pt,
                destino_dir.x1 + sangria_pt,
                max(destino_esq.y1, destino_dir.y1) + sangria_pt
            ).intersect(mediabox)
            if not bleed_total.is_empty and bleed_total.get_area() > 0:
                nova.set_bleedbox(bleed_total)

        # 3. Marcas de corte e dobra
        if add_marcas:
            desenhar_marcas(nova, destino_esq, tam=5*MM, gap=2*MM, lado='esquerda')
            desenhar_marcas(nova, destino_dir, tam=5*MM, gap=2*MM, lado='direita')
            x_centro = destino_esq.x1
            desenhar_marca_central_vertical(nova, x_centro, destino_esq.y0, destino_esq.y1,
                                            tam=5*MM, gap=2*MM)

        # 4. Trimbox
        trim_union = fitz.Rect(
            destino_esq.x0, min(destino_esq.y0, destino_dir.y0),
            destino_dir.x1, max(destino_esq.y1, destino_dir.y1)
        ).intersect(mediabox)
        if trim_union.is_empty or trim_union.get_area() == 0:
            trim_union = mediabox
        nova.set_trimbox(trim_union)

        # 5. Textos de identificação
        texto_esq = f"{idx_esq + 1}/{total} {nome_arquivo}"
        texto_dir = f"{idx_dir + 1}/{total} {nome_arquivo}"
        _adicionar_identificacao(nova, texto_esq, texto_dir, destino_esq, destino_dir, rot_esq, rot_dir)

    # =========================================================================
    # MONTAGEM DO CADERNO ÚNICO
    # =========================================================================
    metade = total // 2
    num_folhas = total // 4

    for i_folha in range(num_folhas):
        p_esq_frente = metade + 2 * i_folha + 1
        p_dir_frente = 2 * i_folha + 1
        p_esq_verso  = 2 * i_folha + 2
        p_dir_verso  = metade + 2 * i_folha + 2

        _montar_uma_face(p_esq_frente - 1, p_dir_frente - 1,
                         rot_esq=180, rot_dir=0, esq_superior=True, lado="FRENTE")
        _montar_uma_face(p_esq_verso - 1, p_dir_verso - 1,
                         rot_esq=0, rot_dir=180, esq_superior=False, lado="VERSO")

    out.save(output_path)
    doc.close()
    out.close()

    # Remove o arquivo temporário da compressão, se existir
    if comprimir_entrada:
        try:
            os.unlink(tmp_input)
        except:
            pass
        print("PDF comprimido usado como entrada e depois removido.")

    print(f"Montagem concluída: {num_folhas} folhas geradas (caderno único).")
    return True
def dois_na_folha_rotacionado(input_path, output_path, output_width, output_height, espacamento_mm=0, sangria_mm=3, add_marcas=True):
    """
    Montagem 2 na folha com a segunda página rotacionada 180 graus.
    """
    doc = fitz.open(input_path)
    out = fitz.open()
    sangria_pt = sangria_mm * MM
    espaco_pt = espacamento_mm * MM
    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    for page in doc:
        trim_original = obter_area_corte_segura(page, w_ref, h_ref)
        w, h = trim_original.width, trim_original.height

        nova = out.new_page(width=output_width, height=output_height)

        largura_total = 2 * w + espaco_pt + 2 * sangria_pt
        x_inicio = (output_width - largura_total) / 2

        if sangria_pt > 0:
            y, cabe = calcular_y_com_sangria(output_height, h, sangria_pt)
            sangria_aplicar = cabe
        else:
            y = (output_height - h) / 2
            sangria_aplicar = False

        x1 = x_inicio + sangria_pt
        destino1 = fitz.Rect(x1, y, x1 + w, y + h)
        destino1 = ajustar_para_media_box(destino1, output_width, output_height)

        x2 = x1 + w + espaco_pt
        destino2 = fitz.Rect(x2, y, x2 + w, y + h)
        destino2 = ajustar_para_media_box(destino2, output_width, output_height)

        if pagina_tem_conteudo(page):
            # Página esquerda normal, direita rotacionada 180°
            nova.show_pdf_page(destino1, doc, page.number, clip=trim_original, rotate=0)
            nova.show_pdf_page(destino2, doc, page.number, clip=trim_original, rotate=180)

        nova.set_trimbox(destino1)

        if sangria_pt > 0 and sangria_aplicar and pagina_tem_conteudo(page):
            aplicar_sangria_espelhada_externa(page, nova, trim_original, destino1, sangria_pt,
                                              bordas=['top', 'bottom', 'left'])
            aplicar_sangria_espelhada_externa(page, nova, trim_original, destino2, sangria_pt,
                                              bordas=['top', 'bottom', 'right'])

            bleed_total = fitz.Rect(
                destino1.x0 - sangria_pt,
                destino1.y0 - sangria_pt,
                destino2.x1 + sangria_pt,
                destino1.y1 + sangria_pt
            )
            bleed_total = ajustar_para_media_box(bleed_total, output_width, output_height)
            if bleed_total.width > 0 and bleed_total.height > 0:
                nova.set_bleedbox(bleed_total)

        if add_marcas:
            if espaco_pt > 0:
                desenhar_marcas(nova, destino1, tam=5*MM, gap=2*MM, lado='esquerda')
                desenhar_marca_central_vertical(nova, destino1.x1, destino1.y0, destino1.y1, tam=5*MM, gap=2*MM)
                desenhar_marcas(nova, destino2, tam=5*MM, gap=2*MM, lado='direita')
                desenhar_marca_central_vertical(nova, destino2.x0, destino2.y0, destino2.y1, tam=5*MM, gap=2*MM)
            else:
                desenhar_marcas(nova, destino1, tam=5*MM, gap=2*MM, lado='esquerda')
                desenhar_marcas(nova, destino2, tam=5*MM, gap=2*MM, lado='direita')
                x_centro = destino1.x1
                desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1, tam=5*MM, gap=2*MM)

    out.save(output_path)
    doc.close()
    out.close()
    return True
def um_na_folha(input_path, output_path, output_width, output_height,
                sangria_mm=3, add_marcas=True,
                add_identificacao=True,
                texto_offset_mm=0.5,
                comprimir_entrada=False):

    # Compressão opcional
    tmp_input = None
    if comprimir_entrada:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_input = tmp.name
        comprimir_pdf(input_path, tmp_input)
        doc = fitz.open(tmp_input)
    else:
        doc = fitz.open(input_path)

    if len(doc) == 0:
        print("Erro: O documento PDF está vazio.")
        doc.close()
        return False

    out = fitz.open()
    sangria_pt = sangria_mm * MM
    offset_pt = texto_offset_mm * MM
    total = len(doc)
    nome_arquivo = os.path.basename(input_path)

    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    # ---------- Identificação local (só rot=0) ----------
    def _identificar(nova_page, texto, destino):
        if not texto:
            return
        font_size = 5
        altura_texto = 4 * MM
        margem_min = 0.5 * MM

        y1 = destino.y0 - offset_pt
        y0 = y1 - altura_texto
        if y0 < margem_min:
            y0 = margem_min
            y1 = y0 + altura_texto

        rect = fitz.Rect(destino.x0, y0, destino.x1, y1)
        nova_page.insert_textbox(rect, texto, fontsize=font_size, fontname="helv",
                                 align=fitz.TEXT_ALIGN_CENTER, rotate=0)
    # ---------------------------------------------------

    for page in doc:
        trim_original = obter_area_corte_segura(page, w_ref, h_ref)
        w, h = trim_original.width, trim_original.height

        nova = out.new_page(width=output_width, height=output_height)

        x = (output_width - w) / 2
        if sangria_pt > 0:
            y, cabe = calcular_y_com_sangria(output_height, h, sangria_pt)
            sangria_aplicar = cabe
        else:
            y = (output_height - h) / 2
            sangria_aplicar = False

        destino = fitz.Rect(x, y, x + w, y + h)
        destino = ajustar_para_media_box(destino, output_width, output_height)

        tem_conteudo = pagina_tem_conteudo(page)

        if tem_conteudo:
            nova.show_pdf_page(destino, doc, page.number, clip=trim_original)

        nova.set_trimbox(destino)

        if sangria_pt > 0 and sangria_aplicar and tem_conteudo:
            bleed = fitz.Rect(
                destino.x0 - sangria_pt,
                destino.y0 - sangria_pt,
                destino.x1 + sangria_pt,
                destino.y1 + sangria_pt
            )
            bleed = ajustar_para_media_box(bleed, output_width, output_height)
            if bleed.width > 0 and bleed.height > 0:
                nova.set_bleedbox(bleed)
            aplicar_sangria_espelhada_externa(page, nova, trim_original, destino,
                                              sangria_pt,
                                              bordas=['top', 'bottom', 'left', 'right'])

        if add_marcas:
            desenhar_marcas(nova, destino, tam=5*MM, gap=2*MM)

        if add_identificacao and tem_conteudo:
            texto = f"{page.number + 1}/{total} {nome_arquivo}"
            _identificar(nova, texto, destino)
            
            


    out.save(output_path)
    doc.close()
    out.close()

    if comprimir_entrada and tmp_input:
        try:
            os.unlink(tmp_input)
        except Exception:
            pass

    return True
def um_na_folha_atexco(input_path, output_path, output_width, output_height,
                       sangria_mm=3, add_marcas=True,
                       add_identificacao=True,
                       texto_offset_mm=0.5,
                       comprimir_entrada=False):

    # ===== PASSO 1: forçar montagem em PÉ (portrait) =====
    if output_width > output_height:
        output_width, output_height = output_height, output_width
    # =====================================================

    # Compressão opcional
    tmp_input = None
    if comprimir_entrada:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_input = tmp.name
        comprimir_pdf(input_path, tmp_input)
        doc = fitz.open(tmp_input)
    else:
        doc = fitz.open(input_path)

    if len(doc) == 0:
        print("Erro: O documento PDF está vazio.")
        doc.close()
        return False

    out = fitz.open()
    sangria_pt = sangria_mm * MM
    offset_pt = texto_offset_mm * MM
    total = len(doc)
    nome_arquivo = os.path.basename(input_path)

    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    def _identificar(nova_page, texto, destino):
        if not texto:
            return
        font_size = 5
        altura_texto = 4 * MM
        margem_min = 0.5 * MM

        y1 = destino.y0 - offset_pt
        y0 = y1 - altura_texto
        if y0 < margem_min:
            y0 = margem_min
            y1 = y0 + altura_texto

        rect = fitz.Rect(destino.x0, y0, destino.x1, y1)
        nova_page.insert_textbox(rect, texto, fontsize=font_size, fontname="helv",
                                 align=fitz.TEXT_ALIGN_CENTER, rotate=0)

    for page in doc:
        trim_original = obter_area_corte_segura(page, w_ref, h_ref)
        w, h = trim_original.width, trim_original.height

        nova = out.new_page(width=output_width, height=output_height)

        x = (output_width - w) / 2
        if sangria_pt > 0:
            y, cabe = calcular_y_com_sangria(output_height, h, sangria_pt)
            sangria_aplicar = cabe
        else:
            y = (output_height - h) / 2
            sangria_aplicar = False

        destino = fitz.Rect(x, y, x + w, y + h)
        destino = ajustar_para_media_box(destino, output_width, output_height)

        tem_conteudo = pagina_tem_conteudo(page)

        if tem_conteudo:
            nova.show_pdf_page(destino, doc, page.number, clip=trim_original)

        nova.set_trimbox(destino)

        if sangria_pt > 0 and sangria_aplicar and tem_conteudo:
            bleed = fitz.Rect(
                destino.x0 - sangria_pt,
                destino.y0 - sangria_pt,
                destino.x1 + sangria_pt,
                destino.y1 + sangria_pt
            )
            bleed = ajustar_para_media_box(bleed, output_width, output_height)
            if bleed.width > 0 and bleed.height > 0:
                nova.set_bleedbox(bleed)
            aplicar_sangria_espelhada_externa(page, nova, trim_original, destino,
                                              sangria_pt,
                                              bordas=['top', 'bottom', 'left', 'right'])

        if add_marcas:
            desenhar_marcas(nova, destino, tam=5*MM, gap=2*MM)

        if add_identificacao and tem_conteudo:
            texto = f"{page.number + 1}/{total} {nome_arquivo}"
            _identificar(nova, texto, destino)

    # ===== PASSO 2 + 3: frentes 90°, versos 270° =====
    for i, p in enumerate(out):
        if i % 2 == 0:      # páginas 1, 3, 5... → 90°
            p.set_rotation(90)
        else:               # páginas 2, 4, 6... → 270° (180° a mais)
            p.set_rotation(270)
    # ==================================================

    out.save(output_path)
    doc.close()
    out.close()

    if comprimir_entrada and tmp_input:
        try:
            os.unlink(tmp_input)
        except Exception:
            pass

    return True
def dois_na_folha(input_path, output_path, output_width, output_height,
                  espacamento_mm=0, sangria_mm=3, add_marcas=True,
                  add_identificacao=True, texto_offset_mm=0.5):
    doc = fitz.open(input_path)
    out = fitz.open()
    sangria_pt = sangria_mm * MM
    espaco_pt = espacamento_mm * MM
    offset_pt = texto_offset_mm * MM
    total = len(doc)
    nome_arquivo = os.path.basename(input_path)

    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    # ---------- Identificação local ----------
    def _identificar(nova_page, texto, destino):
        if not texto:
            return
        font_size = 5
        altura_texto = 4 * MM
        margem_min = 0.5 * MM

        y1 = destino.y0 - offset_pt
        y0 = y1 - altura_texto
        if y0 < margem_min:
            y0 = margem_min
            y1 = y0 + altura_texto

        rect = fitz.Rect(destino.x0, y0, destino.x1, y1)
        nova_page.insert_textbox(rect, texto, fontsize=font_size, fontname="helv",
                                 align=fitz.TEXT_ALIGN_CENTER, rotate=0)
    # ------------------------------------------

    for page in doc:
        trim_original = obter_area_corte_segura(page, w_ref, h_ref)
        w, h = trim_original.width, trim_original.height

        nova = out.new_page(width=output_width, height=output_height)

        largura_total = 2 * w + espaco_pt + 2 * sangria_pt
        x_inicio = (output_width - largura_total) / 2

        if sangria_pt > 0:
            y, cabe = calcular_y_com_sangria(output_height, h, sangria_pt)
            sangria_aplicar = cabe
        else:
            y = (output_height - h) / 2
            sangria_aplicar = False

        x1 = x_inicio + sangria_pt
        destino1 = fitz.Rect(x1, y, x1 + w, y + h)
        destino1 = ajustar_para_media_box(destino1, output_width, output_height)

        x2 = x1 + w + espaco_pt
        destino2 = fitz.Rect(x2, y, x2 + w, y + h)
        destino2 = ajustar_para_media_box(destino2, output_width, output_height)

        tem_conteudo = pagina_tem_conteudo(page)

        if tem_conteudo:
            nova.show_pdf_page(destino1, doc, page.number, clip=trim_original)
            nova.show_pdf_page(destino2, doc, page.number, clip=trim_original)

        nova.set_trimbox(destino1)

        if sangria_pt > 0 and sangria_aplicar and tem_conteudo:
            aplicar_sangria_espelhada_externa(page, nova, trim_original, destino1, sangria_pt,
                                              bordas=['top', 'bottom', 'left'])
            aplicar_sangria_espelhada_externa(page, nova, trim_original, destino2, sangria_pt,
                                              bordas=['top', 'bottom', 'right'])

            bleed_total = fitz.Rect(
                destino1.x0 - sangria_pt,
                destino1.y0 - sangria_pt,
                destino2.x1 + sangria_pt,
                destino1.y1 + sangria_pt
            )
            bleed_total = ajustar_para_media_box(bleed_total, output_width, output_height)
            if bleed_total.width > 0 and bleed_total.height > 0:
                nova.set_bleedbox(bleed_total)

        if add_marcas:
            if espaco_pt > 0:
                desenhar_marcas(nova, destino1, tam=5*MM, gap=2*MM, lado='esquerda')
                desenhar_marca_central_vertical(nova, destino1.x1, destino1.y0, destino1.y1, tam=5*MM, gap=2*MM)
                desenhar_marcas(nova, destino2, tam=5*MM, gap=2*MM, lado='direita')
                desenhar_marca_central_vertical(nova, destino2.x0, destino2.y0, destino2.y1, tam=5*MM, gap=2*MM)
            else:
                desenhar_marcas(nova, destino1, tam=5*MM, gap=2*MM, lado='esquerda')
                desenhar_marcas(nova, destino2, tam=5*MM, gap=2*MM, lado='direita')
                x_centro = destino1.x1
                desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1, tam=5*MM, gap=2*MM)

        if add_identificacao and tem_conteudo:
            texto = f"{page.number + 1}/{total} {nome_arquivo}"
            _identificar(nova, texto, destino1)
            _identificar(nova, texto, destino2)

    out.save(output_path)
    doc.close()
    out.close()
    return True
# Caminho do Ghostscript
GS_PATH = r"C:\Ghostscript\gswin64c.exe"
def comprimir_pdf(entrada, saida):
    """Comprime um PDF usando Ghostscript."""
    if not os.path.exists(GS_PATH):
        raise FileNotFoundError(f"Ghostscript não encontrado em: {GS_PATH}")
    comando = [
        GS_PATH,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/ebook",
        "-dDownsampleColorImages=true",
        "-dDownsampleGrayImages=true",
        "-dDownsampleMonoImages=true",
        "-dColorImageResolution=150",
        "-dGrayImageResolution=150",
        "-dMonoImageResolution=150",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={saida}",
        entrada
    ]
    try:
        subprocess.run(comando, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Erro ao comprimir: {e}")
        # Fallback: copia o original
        import shutil
        shutil.copy(entrada, saida)
def Corta_empilha(input_path, output_path, output_width, output_height,
                  sangria_mm=3, add_marcas=True,
                  add_identificacao=True,
                  texto_offset_normal_mm=0.5,
                  texto_offset_180_mm=0.5,
                  comprimir_entrada=True):
    """
    Montagem corta-empilha com caderno único.
    Se comprimir_entrada=True, o PDF original é comprimido antes da imposição.
    """

    # Se solicitado, comprime o PDF de entrada num arquivo temporário
    tmp_input = None
    if comprimir_entrada:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_input = tmp.name
        comprimir_pdf(input_path, tmp_input)
        doc = fitz.open(tmp_input)
    else:
        doc = fitz.open(input_path)

    nome_arquivo = os.path.basename(input_path)

    # ---------- AJUSTE PARA MÚLTIPLO DE 4 ----------
    total_original = len(doc)
    try:
        ref = doc.load_page(0) or doc[0]
        w_ref, h_ref = ref.rect.width, ref.rect.height
    except:
        w_ref, h_ref = 210 * MM, 297 * MM

    if total_original % 2 != 0:
        doc.new_page(width=w_ref, height=h_ref)
    while len(doc) % 4 != 0:
        doc.new_page(width=w_ref, height=h_ref)

    total = len(doc)
    metade = total // 2
    num_folhas = total // 4
    # ------------------------------------------------

    out = fitz.open()
    sangria_pt = sangria_mm * MM
    offset_normal_pt = texto_offset_normal_mm * MM
    offset_180_pt = texto_offset_180_mm * MM

    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210 * MM, 297 * MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    # ---------- IDENTIFICAÇÃO LOCAL ----------
    def _identificar(nova_page, texto, destino, rot):
        if not texto:
            return
        font_size = 5
        altura_texto = 4 * MM
        margem_min = 0.5 * MM

        if rot == 180:
            y0 = destino.y1 + offset_180_pt
            y1 = y0 + altura_texto
            if y1 > nova_page.rect.height - margem_min:
                y1 = nova_page.rect.height - margem_min
                y0 = y1 - altura_texto
            rect = fitz.Rect(destino.x0, y0, destino.x1, y1)
            nova_page.insert_textbox(rect, texto, fontsize=font_size,
                                     fontname="helv",
                                     align=fitz.TEXT_ALIGN_CENTER, rotate=180)
        else:
            y1 = destino.y0 - offset_normal_pt
            y0 = y1 - altura_texto
            if y0 < margem_min:
                y0 = margem_min
                y1 = y0 + altura_texto
            rect = fitz.Rect(destino.x0, y0, destino.x1, y1)
            nova_page.insert_textbox(rect, texto, fontsize=font_size,
                                     fontname="helv",
                                     align=fitz.TEXT_ALIGN_CENTER, rotate=0)
    # ----------------------------------------

    # ---------- FUNÇÃO AUXILIAR (uma face da folha) ----------
    def _criar_pagina(idx_esq, idx_dir, rot_esq, rot_dir, esq_superior):
        page_esq = doc[idx_esq] if idx_esq is not None else None
        page_dir = doc[idx_dir] if idx_dir is not None else None

        # Dimensões
        if page_esq:
            trim_esq = obter_area_corte_segura(page_esq, w_ref, h_ref)
            w_esq, h_esq = trim_esq.width, trim_esq.height
        else:
            trim_esq = None
            w_esq, h_esq = w_ref, h_ref
        if page_dir:
            trim_dir = obter_area_corte_segura(page_dir, w_ref, h_ref)
            w_dir, h_dir = trim_dir.width, trim_dir.height
        else:
            trim_dir = None
            w_dir, h_dir = w_ref, h_ref

        # Quem fica em cima?
        if esq_superior:
            h_sup, h_inf = h_esq, h_dir
        else:
            h_sup, h_inf = h_dir, h_esq

        nova = out.new_page(width=output_width, height=output_height)

        # Posições Y (com sangria)
        if sangria_pt > 0:
            y_sup, cabe_sup = calcular_y_com_sangria(output_height, h_sup, sangria_pt)
            y_inf, cabe_inf = calcular_y_com_sangria(output_height, h_inf, sangria_pt)
            sangria_aplicar_sup = cabe_sup
            sangria_aplicar_inf = cabe_inf
        else:
            y_sup = (output_height - h_sup) / 2
            y_inf = (output_height - h_inf) / 2
            sangria_aplicar_sup = sangria_aplicar_inf = False

        if esq_superior:
            y_esq, y_dir = y_sup, y_inf
            sangria_aplicar_esq = sangria_aplicar_sup and (page_esq is not None)
            sangria_aplicar_dir = sangria_aplicar_inf and (page_dir is not None)
        else:
            y_esq, y_dir = y_inf, y_sup
            sangria_aplicar_esq = sangria_aplicar_inf and (page_esq is not None)
            sangria_aplicar_dir = sangria_aplicar_sup and (page_dir is not None)

        # Posições X
        largura_total = w_esq + w_dir + 2 * sangria_pt
        x_inicio = (output_width - largura_total) / 2
        x1 = x_inicio + sangria_pt
        destino1 = fitz.Rect(x1, y_esq, x1 + w_esq, y_esq + h_esq)
        destino1 = ajustar_para_media_box(destino1, output_width, output_height)
        x2 = x1 + w_esq
        destino2 = fitz.Rect(x2, y_dir, x2 + w_dir, y_dir + h_dir)
        destino2 = ajustar_para_media_box(destino2, output_width, output_height)

        tem_esq = page_esq is not None and pagina_tem_conteudo(page_esq)
        tem_dir = page_dir is not None and pagina_tem_conteudo(page_dir)

        # Inserir páginas com rotação
        if tem_esq:
            nova.show_pdf_page(destino1, doc, idx_esq, clip=trim_esq, rotate=rot_esq)
        if tem_dir:
            nova.show_pdf_page(destino2, doc, idx_dir, clip=trim_dir, rotate=rot_dir)

        # Trimbox
        trim_union = fitz.Rect(destino1.x0, min(destino1.y0, destino2.y0),
                               destino2.x1, max(destino1.y1, destino2.y1))
        trim_union = ajustar_para_media_box(trim_union, output_width, output_height)
        nova.set_trimbox(trim_union)

        # Sangria
        if sangria_pt > 0:
            if tem_esq and sangria_aplicar_esq:
                aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1,
                                                  sangria_pt,
                                                  bordas=['top', 'bottom', 'left'],
                                                  rotate=rot_esq)
            if tem_dir and sangria_aplicar_dir:
                aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2,
                                                  sangria_pt,
                                                  bordas=['top', 'bottom', 'right'],
                                                  rotate=rot_dir)

            bleed_total = fitz.Rect(
                destino1.x0 - sangria_pt,
                min(destino1.y0, destino2.y0) - sangria_pt,
                destino2.x1 + sangria_pt,
                max(destino1.y1, destino2.y1) + sangria_pt
            )
            media_box = fitz.Rect(0, 0, output_width, output_height)
            bleed_total = bleed_total.intersect(media_box)

            if not bleed_total.is_empty and bleed_total.get_area() > 0:
                nova.set_bleedbox(bleed_total)

        # Marcas
        if add_marcas:
            desenhar_marcas(nova, destino1, tam=5 * MM, gap=2 * MM, lado='esquerda')
            desenhar_marcas(nova, destino2, tam=5 * MM, gap=2 * MM, lado='direita')
            x_centro = destino1.x1
            desenhar_marca_central_vertical(nova, x_centro,
                                            destino1.y0, destino1.y1,
                                            tam=5 * MM, gap=2 * MM)

        # Identificação (só onde há conteúdo)
        if add_identificacao:
            if tem_esq:
                _identificar(nova, f"{idx_esq + 1}/{total} {nome_arquivo}",
                             destino1, rot_esq)
            if tem_dir:
                _identificar(nova, f"{idx_dir + 1}/{total} {nome_arquivo}",
                             destino2, rot_dir)
    # ------------------------------------------------------------

    # MONTAGEM DAS FOLHAS
    for i_folha in range(num_folhas):
        p_esq_frente = metade + 2 * i_folha + 1
        p_dir_frente = 2 * i_folha + 1
        p_esq_verso  = 2 * i_folha + 2
        p_dir_verso  = metade + 2 * i_folha + 2

        # FRENTE
        _criar_pagina(p_esq_frente - 1, p_dir_frente - 1,
                      rot_esq=180, rot_dir=0, esq_superior=True)
        # VERSO
        _criar_pagina(p_esq_verso - 1, p_dir_verso - 1,
                      rot_esq=0, rot_dir=180, esq_superior=False)

    out.save(output_path)
    doc.close()
    out.close()

    # Remove o temporário da compressão, se criado
    if comprimir_entrada and tmp_input:
        try:
            os.unlink(tmp_input)
        except:
            pass
        print("PDF comprimido usado como entrada e depois removido.")

    print("Montagem concluída.")
    return True
def grampo(input_path, output_path, output_width, output_height,
           sangria_mm=3, add_marcas=True,
           add_identificacao=True,
           texto_offset_mm=0.5,
           comprimir_entrada=False):

    # Compressão opcional
    tmp_input = None
    if comprimir_entrada:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_input = tmp.name
        comprimir_pdf(input_path, tmp_input)
        doc = fitz.open(tmp_input)
    else:
        doc = fitz.open(input_path)

    if len(doc) == 0:
        print("Erro: O documento PDF está vazio.")
        doc.close()
        return False

    out = fitz.open()
    sangria_pt = sangria_mm * MM
    offset_pt = texto_offset_mm * MM
    total = len(doc)
    num_folhas = (total + 1) // 2
    nome_arquivo = os.path.basename(input_path)

    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    # ---------- Identificação local (só rot=0, acima do trim) ----------
    def _identificar(nova_page, texto, destino):
        if not texto:
            return
        font_size = 5
        altura_texto = 4 * MM
        margem_min = 0.5 * MM

        y1 = destino.y0 - offset_pt
        y0 = y1 - altura_texto
        if y0 < margem_min:
            y0 = margem_min
            y1 = y0 + altura_texto

        rect = fitz.Rect(destino.x0, y0, destino.x1, y1)
        nova_page.insert_textbox(rect, texto, fontsize=font_size, fontname="helv",
                                 align=fitz.TEXT_ALIGN_CENTER, rotate=0)
    # -------------------------------------------------------------------

    for i in range(num_folhas):
        if i % 2 == 0:
            idx_esq_base1 = total - i
            idx_dir_base1 = i + 1
        else:
            idx_esq_base1 = i + 1
            idx_dir_base1 = total - i

        idx_esq = idx_esq_base1 - 1 if 1 <= idx_esq_base1 <= total else None
        idx_dir = idx_dir_base1 - 1 if 1 <= idx_dir_base1 <= total else None

        page_esq = doc[idx_esq] if idx_esq is not None else None
        page_dir = doc[idx_dir] if idx_dir is not None else None

        if page_esq:
            trim_esq = obter_area_corte_segura(page_esq, w_ref, h_ref)
            w_esq, h_esq = trim_esq.width, trim_esq.height
        else:
            trim_esq = None
            w_esq, h_esq = w_ref, h_ref

        if page_dir:
            trim_dir = obter_area_corte_segura(page_dir, w_ref, h_ref)
            w_dir, h_dir = trim_dir.width, trim_dir.height
        else:
            trim_dir = None
            w_dir, h_dir = w_ref, h_ref

        nova = out.new_page(width=output_width, height=output_height)

        if sangria_pt > 0:
            y_esq, cabe_esq = calcular_y_com_sangria(output_height, h_esq, sangria_pt)
            y_dir, cabe_dir = calcular_y_com_sangria(output_height, h_dir, sangria_pt)
            sangria_aplicar_esq = cabe_esq and (page_esq is not None)
            sangria_aplicar_dir = cabe_dir and (page_dir is not None)
        else:
            y_esq = (output_height - h_esq) / 2
            y_dir = (output_height - h_dir) / 2
            sangria_aplicar_esq = sangria_aplicar_dir = False

        largura_total = w_esq + w_dir + 2 * sangria_pt
        x_inicio = (output_width - largura_total) / 2

        x1 = x_inicio + sangria_pt
        destino1 = fitz.Rect(x1, y_esq, x1 + w_esq, y_esq + h_esq)
        destino1 = ajustar_para_media_box(destino1, output_width, output_height)

        x2 = x1 + w_esq
        destino2 = fitz.Rect(x2, y_dir, x2 + w_dir, y_dir + h_dir)
        destino2 = ajustar_para_media_box(destino2, output_width, output_height)

        tem_esq = page_esq is not None and pagina_tem_conteudo(page_esq)
        tem_dir = page_dir is not None and pagina_tem_conteudo(page_dir)

        if tem_esq:
            nova.show_pdf_page(destino1, doc, idx_esq, clip=trim_esq)
        if tem_dir:
            nova.show_pdf_page(destino2, doc, idx_dir, clip=trim_dir)

        nova.set_trimbox(destino1)

        if sangria_pt > 0:
            if tem_esq and sangria_aplicar_esq:
                aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1,
                                                  sangria_pt,
                                                  bordas=['top', 'bottom', 'left'])
            if tem_dir and sangria_aplicar_dir:
                aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2,
                                                  sangria_pt,
                                                  bordas=['top', 'bottom', 'right'])

            bleed_total = fitz.Rect(
                destino1.x0 - sangria_pt,
                min(destino1.y0, destino2.y0) - sangria_pt,
                destino2.x1 + sangria_pt,
                max(destino1.y1, destino2.y1) + sangria_pt
            )
            bleed_total = ajustar_para_media_box(bleed_total, output_width, output_height)
            if bleed_total.width > 0 and bleed_total.height > 0:
                nova.set_bleedbox(bleed_total)

        if add_marcas:
            desenhar_marcas(nova, destino1, tam=5*MM, gap=2*MM, lado='esquerda')
            desenhar_marcas(nova, destino2, tam=5*MM, gap=2*MM, lado='direita')
            x_centro = destino1.x1
            desenhar_marca_central_vertical(nova, x_centro,
                                            destino1.y0, destino1.y1,
                                            tam=5*MM, gap=2*MM)

        # Identificação — só onde há conteúdo
        if add_identificacao:
            if tem_esq:
                _identificar(nova, f"{idx_esq + 1}/{total} {nome_arquivo}", destino1)
            if tem_dir:
                _identificar(nova, f"{idx_dir + 1}/{total} {nome_arquivo}", destino2)

    out.save(output_path)
    doc.close()
    out.close()

    if comprimir_entrada and tmp_input:
        try:
            os.unlink(tmp_input)
        except Exception:
            pass

    return True
def gprint_reservar_destino_capa(pasta_destino, identific, arquivo_txt):
    """Reserva um nome para o PDF final sem sobrescrever arquivos existentes.

    Mantem o nome original IDENTIFIC.pdf, acrescentando _1, _2, _3...
    somente quando ja houver um PDF com o mesmo nome na pasta.
    A reserva O_EXCL evita que duas estacoes escolham o mesmo destino.
    O arquivo_txt permanece na assinatura para compatibilidade com o fluxo.
    """
    nome_base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(identific)).strip(' .') or 'sem_nome'

    # Guarda margem para extensao e sufixo, considerando caminhos do Windows.
    limite = min(190, max(32, 240 - len(os.path.abspath(pasta_destino))))
    nome_base = nome_base[:limite]

    for numero in range(0, 10001):
        sufixo = '' if numero == 0 else f'_{numero}'
        nome = f"{nome_base[:limite-len(sufixo)]}{sufixo}.pdf"
        caminho = os.path.join(pasta_destino, nome)
        try:
            descritor = os.open(caminho, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o666)
        except FileExistsError:
            continue
        else:
            os.close(descritor)
            return caminho
    raise RuntimeError('Nao foi possivel reservar um nome livre para a capa mesclada')


class Application:
    def __init__(self, root):
        self.root = root
        self.arquivos_selecionados = []
        self.imprimir_automaticamente = tk.BooleanVar(value=True)
        self.tema_atual = tk.StringVar(value=CONFIG.get('tema', 'darkly'))
        self.style = Style(theme=self.tema_atual.get())
        self.repositorio_equipamentos = GPrintEquipamentosRepository()
        self.equipamentos_cache = []
        # Ícone do aplicativo
        # "Q:\___LISTA DE PRIORIDADES\GUILHERME\GPrint.ico"
        # Caminhos dos arquivos visuais
        self.icon_path = r"Q:\___LISTA DE PRIORIDADES\GUILHERME\GPrint.ico"
        self.logo_path = r"Q:\___LISTA DE PRIORIDADES\GUILHERME\GPrint.png"

        # Ícone da janela
        try:
            self.root.iconbitmap(self.icon_path)
        except (OSError, tk.TclError) as e:
            logging.warning("Icone da rede indisponivel: %s", e)

        
        # Controle de progresso
        self.processamento_em_andamento = False
        self.tempo_inicio = 0
        self.log_queue = queue.Queue()
        
        self.setup_ui()
        self.root.after(600, self.iniciar_cadastro_equipamentos)
        self.setup_drag_and_drop()
        self.atualizar_log_queue()

        # Pasta temporária para envio LPR
        self.pasta_envio_curto = os.path.join(tempfile.gettempdir(), "UICLAP_LPR")
        os.makedirs(self.pasta_envio_curto, exist_ok=True)

        # Caminhos do banco SQLite (mesmo usado pelo Miolo)
        self.log_db_path = GPRINT_BANCO_HISTORICO
        self.log_db_local_path = os.path.join(os.path.expanduser("~"), "Desktop", "impressoes_fallback.db")

        # Alertar ao abrir, sem travar a interface quando a rede estiver lenta.
        self.root.after(900, lambda: threading.Thread(
            target=gprint_verificar_banco_no_inicio,
            args=(self.log_db_path, self.root), daemon=True
        ).start())

    # ==================== MÉTODOS DE LOG NO SQLITE ====================
    def _adquirir_trava(self, lock_file, timeout):
        inicio = time.time()
        while time.time() - inicio < timeout:
            try:
                with open(lock_file, 'x') as f:
                    f.write(str(os.getpid()))
                return True
            except FileExistsError:
                time.sleep(0.1)
        return False

    def _liberar_trava(self, lock_file):
        try:
            os.remove(lock_file)
        except:
            pass

    def registrar_log_sqlite(self, dados):
        return gprint_registrar_log_seguro(
            dados, self.log_db_path, self.log_db_local_path,
            self.root, self.adicionar_log
        )

    # ==================== INTERFACE GRÁFICA ====================
    def setup_ui(self):
        self.root.title("GPrint")
        self.root.geometry("900x700")
        self.root.state('zoomed')
        self.center_window()
        
        self.criar_menu_superior()
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        # Aba Principal (Capa)
        self.aba_principal = ttk.Frame(self.notebook)
        self.notebook.add(self.aba_principal, text="Capa")
        self.criar_interface_principal(self.aba_principal)
        
        # Aba Miolo
        self.aba_miolo = MioloTab(self.notebook, self)
        self.notebook.add(self.aba_miolo, text="Miolo")

        # Aba Histórico
        self.aba_historico = HistoricoTab(self.notebook, self)
        self.notebook.add(self.aba_historico, text="Histórico")
        
        # Aba ENVIO
        self.aba_hotfolder = HotfolderTab(self.notebook, self)
        self.notebook.add(self.aba_hotfolder, text="Hotfolder")

        # Aba Refazer
        self.aba_refazer = RefazerTab(self.notebook, self)
        self.notebook.add(self.aba_refazer, text="Refazer")
        
        # Aba Configurações
        self.aba_config = ConfigTab(self.notebook, self)
        self.notebook.add(self.aba_config, text="Configurações")
        self.notebook.bind('<<NotebookTabChanged>>', self._ao_mudar_aba_configuracoes, add='+')
 
    def iniciar_cadastro_equipamentos(self):
        """Inicializacao em thread para nao prender a janela se a rede cair."""
        def tarefa():
            try:
                self.repositorio_equipamentos.preparar_banco()
                equipamentos = self.repositorio_equipamentos.listar_equipamentos()
                self.root.after(0, lambda: self._cadastro_carregado(equipamentos))
            except Exception as exc:
                logging.warning('Cadastro central indisponivel: %s', exc)
                self.root.after(0, lambda: messagebox.showwarning(
                    'GPrint - Equipamentos indisponiveis',
                    'Nao foi possivel acessar o cadastro central de equipamentos.\n'
                    'Verifique se a pasta producao esta online.\n'
                    'Se persistir, informe o Guilherme.'))
        threading.Thread(target=tarefa, daemon=True).start()

    def _cadastro_carregado(self, equipamentos):
        """Cadastro inicial pede acesso restrito; uso normal nao pede senha."""
        self.equipamentos_cache = equipamentos
        self.atualizar_listas_equipamentos()
        if not equipamentos:
            self.notebook.select(self.aba_config)
            messagebox.showinfo('Configuracao inicial do GPrint',
                'Nenhum equipamento ativo foi encontrado.\n\n'
                'Acesse Configuracoes > Equipamentos e filas, entre com a senha '
                'administrativa e cadastre as impressoras e suas filas.\n'
                'O cadastro sera compartilhado pelas estacoes.', parent=self.root)
            self.aba_config.mostrar_equipamentos()

    def _ao_mudar_aba_configuracoes(self, _event=None):
        aba_atual = self.notebook.select()
        if aba_atual != str(self.aba_config):
            # Sair de Configuracoes revoga o acesso; novo acesso pede senha.
            self.aba_config.bloquear_cadastro()
        if aba_atual in (str(self.aba_principal),str(self.aba_miolo),str(self.aba_hotfolder)):
            try:
                self.atualizar_listas_equipamentos()
            except Exception:
                pass

    def atualizar_listas_equipamentos(self):
        try:
            todos = self.repositorio_equipamentos.listar_equipamentos()
            self.equipamentos_cache = todos
            capas = [x for x in todos if x['tipo'] in ('CAPA','AMBOS')]
            miolos = [x for x in todos if x['tipo'] in ('MIOLO','AMBOS')]
            self.equipamentos_capa = {gprint_nome_maquina(x):x for x in capas}
            anterior = self.equipamento_capa_var.get()
            nomes = list(self.equipamentos_capa)
            self.combo_equipamento_capa.configure(values=nomes)
            if anterior in nomes:
                self.equipamento_capa_var.set(anterior)
            elif len(nomes) == 1:
                self.equipamento_capa_var.set(nomes[0])
            else:
                self.equipamento_capa_var.set('')
            self.aba_miolo.atualizar_equipamentos(miolos)
            self.aba_hotfolder.atualizar_equipamentos(todos)
            if self.notebook.select() == str(self.aba_hotfolder) and self.aba_hotfolder.machine_var.get():
                self.aba_hotfolder.on_machine_change()
            if hasattr(self, 'aba_historico'):
                anterior_maquina = self.aba_historico.maquina_combo.get()
                maquinas_historico = ['Todas'] + [gprint_nome_maquina(x) for x in todos]
                self.aba_historico.maquina_combo.configure(values=maquinas_historico)
                self.aba_historico.maquina_combo.set(
                    anterior_maquina if anterior_maquina in maquinas_historico else 'Todas')
            return todos
        except Exception as exc:
            logging.warning('Cadastro de equipamentos indisponivel: %s', exc)
            return []

    def criar_interface_principal(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        
        main_frame = ttk.Frame(parent, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=0)
        main_frame.rowconfigure(1, weight=0)
        main_frame.rowconfigure(2, weight=3)
        main_frame.rowconfigure(3, weight=1)
        main_frame.rowconfigure(4, weight=0)
        main_frame.rowconfigure(5, weight=0)
        
        # ================= CONFIGURAÇÕES =================
        config_frame = ttk.LabelFrame(main_frame, text="⚙️ Configurações de Processamento", bootstyle=PRIMARY)
        config_frame.grid(row=0, column=0, sticky="ew", pady=5)
        config_frame.columnconfigure(0, weight=1)
        
        self.checkbox_imprimir = ttk.Checkbutton(
            config_frame,
            text="Enviar para impressora automaticamente após processar",
            variable=self.imprimir_automaticamente,
            bootstyle="success-round-toggle"
        )
        self.checkbox_imprimir.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.equipamento_capa_var = tk.StringVar()
        self.equipamentos_capa = {}
        ttk.Label(config_frame,text='Equipamento de Capa:').grid(row=1,column=0,sticky='w',padx=5)
        self.combo_equipamento_capa = ttk.Combobox(config_frame,
            textvariable=self.equipamento_capa_var,values=[],state='readonly',width=40)
        self.combo_equipamento_capa.grid(row=2,column=0,sticky='w',padx=5,pady=(0,6))
        ttk.Label(config_frame,text='Fila de impressao: identificada automaticamente em cada TXT e validada no cadastro.',
                  bootstyle=SECONDARY).grid(row=3,column=0,sticky='w',padx=5,pady=(0,6))
        
        # ================= SELEÇÃO DE ARQUIVOS =================
        file_frame = ttk.LabelFrame(main_frame, text="📂 Arquivos TXT (Capa)", bootstyle=PRIMARY)
        file_frame.grid(row=1, column=0, sticky="ew", pady=5)
        file_frame.columnconfigure(1, weight=1)
        
        self.btn_select = ttk.Button(file_frame, text="Selecionar TXTs", command=self.selecionar_arquivos_manual, bootstyle=PRIMARY)
        self.btn_select.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.file_label = ttk.Label(file_frame, text="Nenhum arquivo selecionado", anchor="w", padding=10, bootstyle="secondary")
        self.file_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # ================= LISTA DE ARQUIVOS =================
        lista_frame = ttk.LabelFrame(main_frame, text="📋 Arquivos Carregados", bootstyle=PRIMARY)
        lista_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        lista_frame.columnconfigure(0, weight=1)
        lista_frame.rowconfigure(0, weight=1)
        
        self.tree_arquivos = ttk.Treeview(lista_frame, columns=('Arquivo',), show='headings', height=12)
        self.tree_arquivos.heading('Arquivo', text='Arquivos Selecionados')
        self.tree_arquivos.column('Arquivo', width=400)
        
        scrollbar = ttk.Scrollbar(lista_frame, command=self.tree_arquivos.yview)
        self.tree_arquivos.configure(yscrollcommand=scrollbar.set)
        self.tree_arquivos.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # ================= LOG =================
        log_frame = ttk.LabelFrame(main_frame, text="📝 Log de Processamento", bootstyle=PRIMARY)
        log_frame.grid(row=3, column=0, sticky="nsew", pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, wrap='word', state='disabled', font=('Consolas', 9),
                                bg='#1e1e1e', fg='#dcdcdc', insertbackground='white',
                                padx=10, pady=10, height=6)
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")
        
        # ================= PROGRESSO =================
        progress_frame = ttk.LabelFrame(main_frame, text="⏳ Andamento", bootstyle=SUCCESS)
        progress_frame.grid(row=4, column=0, sticky="ew", pady=5)
        progress_frame.columnconfigure(0, weight=0)
        progress_frame.columnconfigure(1, weight=1)
        progress_frame.columnconfigure(2, weight=0)

        ttk.Label(progress_frame, text="Andamento:").grid(
            row=0, column=0, padx=5, pady=5, sticky='w')
        self.progress_bar = ttk.Progressbar(
            progress_frame, mode='determinate', bootstyle="success-striped")
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        self.progress_percent_label = ttk.Label(
            progress_frame, text="0%", font=('Helvetica', 9, 'bold'))
        self.progress_percent_label.grid(row=0, column=2, padx=5, pady=5, sticky='e')
        self.indicadores_capa = GPrintIndicadoresAndamento(progress_frame)
        
        # ================= STATUS E AÇÃO =================
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=5, column=0, sticky="ew", pady=5)
        status_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, text="✅ Pronto para processar", font=("Segoe UI", 10, "bold"), bootstyle="success")
        self.status_label.grid(row=0, column=0, sticky="w")
        
        self.btn_processar = ttk.Button(status_frame, text="▶ Processar e Enviar", command=self.iniciar_processamento, bootstyle="success", width=22)
        self.btn_processar.grid(row=0, column=1, sticky="e")

    # ==================== MÉTODOS DE LOG NA INTERFACE ====================
    def atualizar_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.configure(state='normal')
                self.log_text.insert(tk.END, msg + '\n')
                self.log_text.see(tk.END)
                self.log_text.configure(state='disabled')
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.atualizar_log_queue)
    
    def adicionar_log(self, mensagem):
        self.log_queue.put(mensagem)
    
    # ==================== SELEÇÃO DE ARQUIVOS ====================
    def selecionar_arquivos_manual(self):
        files = filedialog.askopenfilenames(
            title="Selecione os arquivos TXT",
            filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        if files:
            self.processar_arquivos_selecionados(files)
    
    def processar_arquivos_selecionados(self, file_paths: List[str]):
        txt_files = [f for f in file_paths if f.lower().endswith('.txt')]
        if not txt_files:
            messagebox.showwarning("Aviso", "Nenhum arquivo .txt válido foi selecionado!")
            return
        
        self.arquivos_selecionados = txt_files
        self.file_label.config(text=f"{len(txt_files)} arquivo(s) selecionado(s)")
        self.status_label.config(text=f"📁 {len(txt_files)} arquivo(s) carregado(s)")
        
        for item in self.tree_arquivos.get_children():
            self.tree_arquivos.delete(item)
        
        for arquivo in txt_files:
            self.tree_arquivos.insert('', tk.END, values=(os.path.basename(arquivo),))
    
    # ==================== PROCESSAMENTO ====================
    def iniciar_processamento(self):
        if not verificar_data_limite():
            return
        if not self.arquivos_selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um arquivo para processar!")
            return
        equipamento = self.equipamentos_capa.get(self.equipamento_capa_var.get())
        if not equipamento:
            messagebox.showwarning('GPrint - Capa', 'Selecione um equipamento de Capa cadastrado em Configuracoes > Equipamentos.')
            return
        self.maquina_capa_execucao = gprint_nome_maquina(equipamento)
        self.ip_capa_execucao = equipamento['ip']
        # O cadastro de filas nao valida nem altera os envios de Capa.
        # A FILACOLOR do TXT continua sendo usada sem alteracao.
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        
        self.progress_bar['value'] = 0
        self.progress_percent_label.config(text="0%")
        self.status_label.config(text="⏳ Processando...", bootstyle="warning")
        self.btn_processar.config(state=DISABLED)
        self.btn_select.config(state=DISABLED)
        
        self.processamento_em_andamento = True
        self.indicadores_capa.iniciar()
        
        thread = threading.Thread(target=self.processar_arquivos, daemon=True)
        thread.start()
    
    def atualizar_timer_ui(self):
        """Compatibilidade: o cronometro e atualizado por indicadores_capa."""
        if self.processamento_em_andamento:
            self.indicadores_capa._atualizar_tempo()
    
    def atualizar_progresso_ui(self, percentual, texto):
        self.progress_bar['value'] = percentual
        self.progress_percent_label.config(text=f"{int(percentual)}%")
        self.adicionar_log(texto)
        self.root.update_idletasks()
    
    def finalizar_processamento(self, sucesso=True):
        self.processamento_em_andamento = False
        self.indicadores_capa.finalizar()
        self.btn_processar.config(state=NORMAL)
        self.btn_select.config(state=NORMAL)
        if sucesso:
            self.status_label.config(text="✅ Processamento concluído!", bootstyle="success")
            self.progress_bar['value'] = 100
            self.progress_percent_label.config(text="100%")
        else:
            self.status_label.config(text="❌ Erro no processamento", bootstyle="danger")
    
    def processar_arquivos(self):
        # Uma pasta unica por execucao, isolada de outras estacoes/processamentos.
        pasta_temp = tempfile.mkdtemp(prefix="capa_", dir=PASTA_TRABALHO)
        
        try:
            self.adicionar_log("Iniciando processamento...")
            self.root.after(0, lambda: self.atualizar_progresso_ui(0, "Preparando processamento..."))
            
            total_arquivos = 0
            for txt in self.arquivos_selecionados:
                total_arquivos += len(ler_arquivo_txt(txt))
            
            arquivos_processados = 0
            pdfs_para_mesclar = {}
            # Cada TXT deve gerar sua propria mesclagem, mesmo com IDENTIFIC igual.
            totais_montagem = {}
            
            # Prepara usuário uma vez
            try:
                usuario = os.getlogin()
            except:
                usuario = 'desconhecido'
            
            # Constantes para a Capa
            MAQUINA_CAPA = self.maquina_capa_execucao
            IP_CAPA = self.ip_capa_execucao
            
            for arquivo_txt in self.arquivos_selecionados:
                self.adicionar_log(f"Lendo arquivo: {os.path.basename(arquivo_txt)}")
                dados = ler_arquivo_txt(arquivo_txt)
                pilha_txt = os.path.splitext(os.path.basename(arquivo_txt))[0]
                
                for linha in dados:
                    try:
                        caminho_origem = os.path.join(linha["CAMINHO"], linha["ARQUIVO"])
                        if not os.path.exists(caminho_origem):
                            self.adicionar_log(f"⚠️ Arquivo não encontrado: {caminho_origem}")
                            continue
                        
                        quantidade = linha["QTD"]
                        identific = linha["IDENTIFIC"]
                        nomefic = linha["NOMEFIC"]
                        filacolor = linha["FILACOLOR"]
                        
                        nome_arquivo, extensao = os.path.splitext(linha["ARQUIVO"])
                        # Subpasta exclusiva por item evita substituir PDFs de entrada
                        # e mantém o nome interno esperado pela marcacao original.
                        pasta_item = tempfile.mkdtemp(prefix="item_", dir=pasta_temp)
                        temp_file = os.path.join(pasta_item, f"{nomefic}_temp{extensao}")
                        shutil.copy(caminho_origem, temp_file)
                        
                        if temp_file.lower().endswith(".pdf"):
                            self.adicionar_log(f"Adicionando marca d'água em {linha['ARQUIVO']}")
                            pdf_marcado = adicionar_marca_dagua(
                                temp_file,
                                linha["LAMINAÇÃO"], 
                                linha["CODIGOACAB"], 
                                identific,
                                linha["FILACOLOR"]
                            )
                            if pdf_marcado:
                                chave = (os.path.abspath(arquivo_txt), identific, filacolor)
                                if chave not in pdfs_para_mesclar:
                                    pdfs_para_mesclar[chave] = []
                                    totais_montagem[chave] = 0
                                pdfs_para_mesclar[chave].extend([pdf_marcado] * quantidade)
                                totais_montagem[chave] += quantidade
                                
                                # Um registro por arquivo de capa preparado.
                                # A quantidade indica quantas capas deste
                                # arquivo entraram na mesclagem, nao envios LPR.
                                dados_montagem = {
                                    'lote': os.path.dirname(arquivo_txt),
                                    'pilha': pilha_txt,
                                    'arquivo': os.path.basename(linha["ARQUIVO"]),
                                    'maquina': MAQUINA_CAPA,
                                    'ip': IP_CAPA,
                                    'formato': filacolor,
                                    'modo': 'capa_marcacao',
                                    'quantidade': quantidade,
                                    'status': 'Montagem OK',
                                    'erro': '',
                                    'copia': '',
                                    'usuario': usuario,
                                    'montagem_ok': 1,
                                    'envio_ok': 0
                                }
                                self.registrar_log_sqlite(dados_montagem)
                        
                        arquivos_processados += 1
                        percentual = (arquivos_processados / total_arquivos) * 80
                        texto = f"Processando {linha['ARQUIVO']}... ({arquivos_processados}/{total_arquivos})"
                        self.root.after(0, lambda p=percentual, t=texto: self.atualizar_progresso_ui(p, t))
                    
                    except Exception as e:
                        self.adicionar_log(f"❌ Erro ao processar {linha['ARQUIVO']}: {e}")
                        continue
            
            total_identifics = len(pdfs_para_mesclar)
            identifics_processados = 0
            for chave, lista_pdfs in pdfs_para_mesclar.items():
                arquivo_txt, identific, filacolor = chave
                pasta_destino = os.path.dirname(arquivo_txt)
                pilha_txt = os.path.splitext(os.path.basename(arquivo_txt))[0]
                total_montagens = totais_montagem.get(chave, 0)
                
                percentual = 80 + ((identifics_processados / total_identifics) * 20)
                texto = f"Agrupando PDFs... ({identifics_processados + 1}/{total_identifics})"
                self.root.after(0, lambda p=percentual, t=texto: self.atualizar_progresso_ui(p, t))

                # O arquivo intermediario permanece em TEMP; somente o PDF final
                # e salvo junto ao TXT de origem, com nome exclusivo.
                nome_temp = os.path.join(pasta_temp, f"mescla_{uuid.uuid4().hex}.pdf")
                self.mesclar_pdfs(lista_pdfs, nome_temp, filacolor)
                if not os.path.isfile(nome_temp) or os.path.getsize(nome_temp) == 0:
                    raise RuntimeError(f"Mesclagem nao gerou PDF valido para {pilha_txt}: {identific}")

                nome_final = gprint_reservar_destino_capa(pasta_destino, identific, arquivo_txt)
                try:
                    shutil.move(nome_temp, nome_final)
                except Exception:
                    # Remove apenas a reserva desta tentativa; nunca outros PDFs.
                    try:
                        os.remove(nome_final)
                    except OSError:
                        pass
                    raise
                self.adicionar_log(f"✅ Arquivo gerado para TXT {pilha_txt}: {nome_final}")

                # A mesclagem final e uma etapa separada do preparo das capas
                # e do envio LPR. Nao altera o arquivo nem a rotina de envio.
                self.registrar_log_sqlite({
                    'lote': pasta_destino,
                    'pilha': pilha_txt,
                    'arquivo': os.path.basename(nome_final),
                    'maquina': MAQUINA_CAPA,
                    'ip': IP_CAPA,
                    'formato': filacolor,
                    'modo': 'mesclagem',
                    'quantidade': total_montagens,
                    'status': 'Mesclagem OK',
                    'erro': '',
                    'copia': '',
                    'usuario': usuario,
                    # Mesclagem nao e outra capa montada; evita dupla contagem
                    # em consultas antigas que somam montagem_ok.
                    'montagem_ok': 0,
                    'envio_ok': 0
                })

                if self.imprimir_automaticamente.get():
                    sucesso = self.enviar_para_impressora(
                        nome_final,
                        f"{identific}.pdf",  # Mantem o titulo do trabalho LPR original.
                        filacolor
                    )
                    
                    # ===== REGISTRO DO ENVIO (SQLite) – QUANTIDADE = TOTAL DE MONTAGENS =====
                    status_envio = 'Enviado' if sucesso else 'Falha no envio'
                    dados_envio = {
                        'lote': pasta_destino,
                        'pilha': pilha_txt,
                        'arquivo': os.path.basename(nome_final),
                        'maquina': MAQUINA_CAPA,
                        'ip': IP_CAPA,
                        'formato': filacolor,
                        'modo': 'envio',
                        'quantidade': total_montagens,   # <-- TOTAL DE CAPAS MONTADAS
                        'status': status_envio,
                        'erro': '' if sucesso else 'Falha no LPR',
                        'copia': '1/1',
                        'usuario': usuario,
                        'montagem_ok': 0,
                        'envio_ok': 1 if sucesso else 0
                    }
                    self.registrar_log_sqlite(dados_envio)
                else:
                    dados_nao_enviado = {
                        'lote': pasta_destino,
                        'pilha': pilha_txt,
                        'arquivo': os.path.basename(nome_final),
                        'maquina': MAQUINA_CAPA,
                        'ip': IP_CAPA,
                        'formato': filacolor,
                        'modo': 'envio',
                        'quantidade': total_montagens,
                        'status': 'Não enviado (manual)',
                        'erro': '',
                        'copia': '1/1',
                        'usuario': usuario,
                        'montagem_ok': 0,
                        'envio_ok': 0
                    }
                    self.registrar_log_sqlite(dados_nao_enviado)
                
                identifics_processados += 1
                time.sleep(1)
            
            self.root.after(0, lambda: self.atualizar_progresso_ui(100, "Processamento concluído!"))
            self.adicionar_log("🎉 Todos os arquivos foram processados com sucesso!")
            time.sleep(0.5)
            self.root.after(0, lambda: self.finalizar_processamento(True))
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", "Processamento dos arquivos selecionados concluído!"))
        
        except Exception as e:
            self.adicionar_log(f"❌ ERRO CRÍTICO: {e}")
            logging.error(f"Erro durante o processamento: {e}")
            self.root.after(0, lambda: self.finalizar_processamento(False))
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {e}"))
        finally:
            # Historico acumulado localmente durante a execucao:
            # sincroniza so ao concluir o grupo inteiro, inclusive se houve erro.
            gprint_sincronizar_pendentes(
                self.log_db_path, root=self.root,
                emitir_log=self.adicionar_log
            )
            # Limpa apenas a pasta desta execucao, inclusive subpastas de itens.
            # Repetir a tentativa atende a bloqueios de arquivo momentaneos no Windows.
            for tentativa in range(3):
                try:
                    shutil.rmtree(pasta_temp)
                    self.adicionar_log(f"🧹 Temporarios da capa removidos: {pasta_temp}")
                    break
                except FileNotFoundError:
                    break
                except OSError as erro_limpeza:
                    if tentativa == 2:
                        self.adicionar_log(f"⚠️ Temporarios nao removidos: {pasta_temp} ({erro_limpeza})")
                    else:
                        time.sleep(0.5)

    # ==================== MESCLAGEM E IMPRESSÃO ====================
    def mesclar_pdfs(self, lista_pdfs, nome_saida, filacolor):
        try:
            merger = PdfMerger()
            for pdf in lista_pdfs:
                merger.append(pdf)
            merger.write(nome_saida)
            merger.close()
            self.adicionar_log(f"📄 PDF mesclado: {os.path.basename(nome_saida)}")
        except Exception as e:
            self.adicionar_log(f"❌ Erro ao mesclar PDFs: {e}")
            messagebox.showerror("Erro", f"Falha ao mesclar arquivos: {e}")
    
    def enviar_para_impressora(self, arquivo_pdf, nome_arquivo, filacolor):
        temp_arquivo = None
        try:
            nome_simples = f"job_{uuid.uuid4().hex[:8]}.pdf"
            temp_arquivo = os.path.join(self.pasta_envio_curto, nome_simples)
            shutil.copy2(arquivo_pdf, temp_arquivo)
            temp_normalizado = temp_arquivo.replace('\\', '/')
            
            titulo = nome_arquivo.replace(filacolor, "").strip()
            comando = [
                "lpr", "-S", self.ip_capa_execucao, "-P", filacolor,
                "-J", titulo,
                temp_normalizado
            ]
            self.adicionar_log(f"🖨️ Enviando: {' '.join(comando)}")
            resultado = subprocess.run(comando, capture_output=True, text=True, check=False)
            
            if resultado.returncode == 0:
                self.adicionar_log(f"✅ Envio concluído: {nome_arquivo}")
                return True
            else:
                erro_real = resultado.stderr.strip() or resultado.stdout.strip()
                self.adicionar_log(f"❌ Falha (código {resultado.returncode}): {erro_real}")
                messagebox.showerror("Erro", f"Falha ao enviar para {filacolor}.\nErro do LPR: {erro_real}")
                return False
        except FileNotFoundError:
            self.adicionar_log("❌ Comando 'lpr' não encontrado.")
            messagebox.showerror("Erro", "Serviço LPR não instalado.")
            return False
        except Exception as e:
            self.adicionar_log(f"❌ Exceção: {e}")
            return False
        finally:
            if temp_arquivo and os.path.exists(temp_arquivo):
                try:
                    os.remove(temp_arquivo)
                except:
                    pass

    def setup_drag_and_drop(self):
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_files_dropped)
    
    def on_files_dropped(self, event):
        files = self.root.tk.splitlist(event.data)
        txt_files = [f for f in files if f.lower().endswith('.txt')]
        if not txt_files:
            messagebox.showwarning("Aviso", "Nenhum arquivo .txt válido foi encontrado!")
            return
        self.processar_arquivos_selecionados(txt_files)
    
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 2
        self.root.geometry(f'+{x}+{y}')
    
    def criar_menu_superior(self):
        menubar = tk.Menu(self.root)

        # =========================
        # MENU TEMAS
        # =========================
        menu_tema = tk.Menu(menubar, tearoff=0)

        # Submenu claros
        menu_claro = tk.Menu(menu_tema, tearoff=0)

        for tema in TEMAS_CLAROS:
            menu_claro.add_radiobutton(
                label=tema.capitalize(),
                variable=self.tema_atual,
                value=tema,
                command=self.mudar_tema
            )

        # Submenu escuros
        menu_escuro = tk.Menu(menu_tema, tearoff=0)

        for tema in TEMAS_ESCUROS:
            menu_escuro.add_radiobutton(
                label=tema.capitalize(),
                variable=self.tema_atual,
                value=tema,
                command=self.mudar_tema
            )

        # Adiciona submenus
        menu_tema.add_cascade(label="☀ Temas Claros", menu=menu_claro)
        menu_tema.add_separator()
        menu_tema.add_cascade(label="🌙 Temas Escuros", menu=menu_escuro)

        menubar.add_cascade(label="Temas", menu=menu_tema)

        # =========================
        # MENU AJUDA
        # =========================
        menu_ajuda = tk.Menu(menubar, tearoff=0)
        menu_ajuda.add_command(label="Sobre", command=self.mostrar_sobre)

        menubar.add_cascade(label="Ajuda", menu=menu_ajuda)

        self.root.config(menu=menubar)
    
    def mudar_tema(self):
        self.style.theme_use(self.tema_atual.get())
        CONFIG['tema'] = self.tema_atual.get()
        salvar_configuracoes(CONFIG)
        self.aplicar_cores_tema()
        messagebox.showinfo("Tema Alterado", f"Tema mudado para: {self.tema_atual.get().capitalize()}")
    
    def aplicar_cores_tema(self):
        # Atualiza os previews de cor da aba Configurações
        if hasattr(self, 'aba_config') and hasattr(self.aba_config, 'atualizar_previews'):
            self.aba_config.atualizar_previews()
    
    def mostrar_sobre(self):
        import tkinter as tk

        about_window = tk.Toplevel(self.root)
        about_window.title("Sobre o GPrint")
        about_window.geometry("450x350")
        about_window.resizable(False, False)
        about_window.configure(bg="#0f172a")

        # Centralizar
        about_window.update_idletasks()
        w, h = 450, 350
        x = (about_window.winfo_screenwidth() // 2) - (w // 2)
        y = (about_window.winfo_screenheight() // 2) - (h // 2)
        about_window.geometry(f"{w}x{h}+{x}+{y}")

        main = tk.Frame(about_window, bg="#0f172a", padx=30, pady=20)
        main.pack(fill="both", expand=True)

        tk.Label(main, text="GPrint", font=("Segoe UI", 18, "bold"),
                bg="#0f172a", fg="#f8fafc").pack(pady=(10, 0))

        versao_badge = tk.Label(main, text="  v5.0  ", font=("Segoe UI", 10, "bold"),
                                bg="#3b82f6", fg="white", padx=8, pady=2)
        versao_badge.pack(pady=(4, 12))

        tk.Frame(main, height=1, bg="#334155").pack(fill="x", pady=10)

        tk.Label(main, text="Automação de processamento e envio\nde arquivos para POD com eficiência.",
                font=("Segoe UI", 10), bg="#0f172a", fg="#cbd5e1", justify="center").pack(pady=5)

        dev_frame = tk.Frame(main, bg="#0f172a")
        dev_frame.pack(pady=20)
        tk.Label(dev_frame, text="Desenvolvido por", font=("Segoe UI", 9),
                bg="#0f172a", fg="#94a3b8").pack()
        tk.Label(dev_frame, text="Guilherme Neife", font=("Segoe UI", 13, "bold"),
                bg="#0f172a", fg="#facc15").pack()

        tk.Label(main, text="© 2026 • Todos os direitos reservados", font=("Segoe UI", 7),
                bg="#0f172a", fg="#64748b").pack(side="bottom", pady=(5, 0))

        btn_fechar = tk.Button(main, text="✕  Fechar", font=("Segoe UI", 10, "bold"),
                            bg="#1e293b", fg="#f8fafc", activebackground="#334155",
                            activeforeground="white", relief="flat", bd=0, padx=20, pady=6,
                            cursor="hand2", command=about_window.destroy)
        btn_fechar.pack(pady=(10, 5))

        def on_enter(e): btn_fechar.config(bg="#334155")
        def on_leave(e): btn_fechar.config(bg="#1e293b")
        btn_fechar.bind("<Enter>", on_enter)
        btn_fechar.bind("<Leave>", on_leave)
def carregar_configuracoes():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                # Converter cores de string para tupla
                if isinstance(config.get('cor_fosco'), str):
                    config['cor_fosco'] = eval(config['cor_fosco'])
                if isinstance(config.get('cor_brilho'), str):
                    config['cor_brilho'] = eval(config['cor_brilho'])
                return config
    except Exception as e:
        logging.error(f"Erro ao carregar configurações: {e}")
    
    return DEFAULT_CONFIG.copy()
def salvar_configuracoes(config):
    try:
        config_to_save = config.copy()
        config_to_save['cor_fosco'] = str(config['cor_fosco'])
        config_to_save['cor_brilho'] = str(config['cor_brilho'])
        
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config_to_save, f, indent=4)
        return True
    except Exception as e:
        logging.error(f"Erro ao salvar configurações: {e}")
        return False
def verificar_data_limite():
    try:
        data_limite = datetime.strptime(DATA_LIMITE, "%d/%m/%Y")
        data_atual = datetime.now()
        if data_atual > data_limite:
            messagebox.showerror("Licença Expirada", "O limite de uso deste programa foi excedido.")
            return False
        return True
    except Exception as e:
        logging.error(f"Erro ao verificar data limite: {e}")
        return True
def ler_arquivo_txt(caminho_txt):
    dados = []
    try:
        # utf-8-sig remove BOM automaticamente
        with open(caminho_txt, "r", encoding="utf-8-sig") as arquivo:
            for idx, linha in enumerate(arquivo, 1):
                campos = linha.strip().split("|")
                if len(campos) >= 14:
                    try:
                        # --- Normalização do CAMINHO ---
                        caminho_bruto = campos[0].strip()
                        # Remove qualquer BOM residual
                        caminho_bruto = caminho_bruto.replace('\ufeff', '').strip()
                        # Substitui todas as barras normais por invertidas (padrão Windows)
                        caminho_bruto = caminho_bruto.replace('/', '\\')
                        # Normaliza (remove duplicidades, resolve .. etc)
                        caminho = os.path.normpath(caminho_bruto)
                        # --- Fim da normalização ---

                        dados.append({
                            "CAMINHO": caminho,          # agora normalizado
                            "ARQUIVO": campos[1].strip(),
                            "QTD": int(campos[2].strip()) if campos[2].strip().isdigit() else 1,
                            "LAMINAÇÃO": campos[4].strip(),
                            "FORMATO CAPA": campos[12].strip(),
                            "CODIGOACAB": campos[9].strip(),
                            "IDENTIFIC": campos[13].strip(),
                            "FILACOLOR": campos[12].strip(),
                            "NOMEFIC": campos[11].strip()
                        })
                    except (ValueError, IndexError) as e:
                        logging.warning(f"Linha {idx} mal formatada no arquivo {caminho_txt}: {linha}")
        return dados
    except Exception as e:
        logging.error(f"Erro ao ler arquivo {caminho_txt}: {e}")
        messagebox.showerror("Erro", f"Não foi possível ler o arquivo {os.path.basename(caminho_txt)}")
        return []
def adicionar_marca_dagua(pdf_path, laminacao, codigoacab, identific, filacolor):
    try:
        doc = fitz.open(pdf_path)
        nome_arquivo = os.path.splitext(os.path.basename(pdf_path))[0]

        cor_marca = CONFIG['cor_fosco'] if laminacao.upper() == "FOSCO" else CONFIG['cor_brilho']

        # ===== RECUOS PADRÃO POR FORMATO (em cm) =====
        recuos_cm = {
            "Capas_28X56": 0.5,   # 5 mm
            "Capas_32X60": 0.4,   # 4 mm
            "Capas_32X66": 0.4,   # 4 mm
            "Capas_32X48": 0.4    # 4 mm
        }

        # Obtém o recuo para o formato desejado (fallback 0.4 se não mapeado)
        recuo_topo_cm = recuos_cm.get(filacolor, 0.4)
        recuo_topo_pt = recuo_topo_cm * 28.3465

        # Altura estimada da caixa de texto da marca
        altura_caixa_texto = 40

        # ===== MAPEAMENTO DOS FORMATOS (largura_cm, altura_cm) =====
        formatos_cm = {
            "Capas_28X56": (56, 28),
            "Capas_32X60": (60, 32),
            "Capas_32X66": (66, 32),
            "Capas_32X48": (48, 32)
        }

        # Tolerância para considerar formato já pronto (em cm)
        tolerancia_cm = 0.5

        novo_doc = fitz.open()

        for pagina in doc:
            largura_pt = pagina.rect.width
            altura_pt = pagina.rect.height
            largura_cm = largura_pt / 28.3465
            altura_cm = altura_pt / 28.3465

            # ===== VERIFICA SE A PÁGINA JÁ ESTÁ EM UM FORMATO CONHECIDO =====
            formato_atual = None
            for nome, (larg_ref, alt_ref) in formatos_cm.items():
                if (abs(largura_cm - larg_ref) <= tolerancia_cm and
                    abs(altura_cm - alt_ref) <= tolerancia_cm):
                    formato_atual = nome
                    break

            if formato_atual is not None:
                # ========== CASO 1: PDF JÁ EM FORMATO FINAL ==========
                nova_pagina = novo_doc.new_page(width=largura_pt, height=altura_pt)
                nova_pagina.show_pdf_page(fitz.Rect(0, 0, largura_pt, altura_pt),
                                          doc, pagina.number)

                # Se for Capas_28X56 e a página não estiver já em 56x28, cortar para 56x28
                if filacolor == "Capas_28X56" and formato_atual != "Capas_28X56":
                    corte_lateral_cm = 2
                    corte_topo_cm = 4
                    corte_lateral_pt = corte_lateral_cm * 28.3465
                    corte_topo_pt = corte_topo_cm * 28.3465
                    crop_rect = fitz.Rect(corte_lateral_pt, corte_topo_pt,
                                          largura_pt - corte_lateral_pt, altura_pt)
                    nova_pagina.set_cropbox(crop_rect)

                # Área visível após o corte (ou página inteira)
                visivel = nova_pagina.rect

                # Posição da marca: recuo padrão do topo da área visível
                x_marca = visivel.x0 + 50 + CONFIG.get('deslocamento_x', 0)
                largura_marca = visivel.width - 100
                y0_marca = visivel.y0 + recuo_topo_pt
                y1_marca = y0_marca + altura_caixa_texto

                posicao_marca = fitz.Rect(x_marca, y0_marca, x_marca + largura_marca, y1_marca)
                nova_pagina.insert_textbox(
                    posicao_marca,
                    f"{nome_arquivo} - {identific}",
                    fontsize=CONFIG['tamanho_fonte'],
                    color=cor_marca,
                    fontname=CONFIG['fonte'],
                    align=1
                )

            else:
                # ========== CASO 2: PDF EM FORMATO ORIGINAL ==========
                if filacolor in formatos_cm:
                    largura_dest_cm, altura_dest_cm = formatos_cm[filacolor]
                else:
                    largura_dest_cm, altura_dest_cm = 60, 32  # fallback

                largura_dest_pt = largura_dest_cm * 28.3465
                altura_dest_pt = altura_dest_cm * 28.3465

                nova_pagina = novo_doc.new_page(width=largura_dest_pt, height=altura_dest_pt)

                # Conteúdo original centralizado na parte inferior
                x0 = (largura_dest_pt - largura_pt) / 2
                y0 = altura_dest_pt - altura_pt
                ret_arte = fitz.Rect(x0, y0, x0 + largura_pt, y0 + altura_pt)
                nova_pagina.show_pdf_page(ret_arte, doc, pagina.number)

                # Marca d'água no topo da nova página, usando o mesmo recuo padrão
                x_marca = 50 + CONFIG.get('deslocamento_x', 0)
                largura_marca = largura_dest_pt - 100
                y0_marca = recuo_topo_pt
                y1_marca = y0_marca + altura_caixa_texto

                posicao_marca = fitz.Rect(x_marca, y0_marca, x_marca + largura_marca, y1_marca)
                nova_pagina.insert_textbox(
                    posicao_marca,
                    f"{nome_arquivo} - {identific}",
                    fontsize=CONFIG['tamanho_fonte'],
                    color=cor_marca,
                    fontname=CONFIG['fonte'],
                    align=1
                )

        # Salva o documento final
        novo_caminho = pdf_path.replace(".pdf", "_marcado.pdf")
        novo_doc.save(novo_caminho)
        novo_doc.close()
        doc.close()
        return novo_caminho

    except Exception as e:
        logging.error(f"Erro ao adicionar marca d'água em {pdf_path}: {e}")
        messagebox.showerror("Erro", f"Falha ao marcar o arquivo {os.path.basename(pdf_path)}")
        return None
# --- INICIALIZAÇÃO ---
CONFIG = carregar_configuracoes()
# Equipamentos e filas são administrados exclusivamente em equipamentos.db.
class HotfolderTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.temp_dir = tempfile.mkdtemp(prefix="hotfolders_")
        self.file_paths = []
        self.sending = False
        self.cancel_send = False

        self.setup_ui()
        self.bind('<Destroy>', self._on_destroy)

    def setup_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")

        main_frame.columnconfigure(0, weight=1)
        # Apenas o log expande verticalmente (agora na linha 4)
        main_frame.rowconfigure(4, weight=1)

        # ================= MÁQUINA =================
        machine_line = ttk.Frame(main_frame)
        machine_line.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        machine_line.columnconfigure(2, weight=1)

        ttk.Label(machine_line, text="Máquina:", width=10).grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.machine_var = tk.StringVar()
        self.machine_combo = ttk.Combobox(
            machine_line,
            textvariable=self.machine_var,
            values=[],
            state="readonly",
            width=18,
            bootstyle=(PRIMARY)
        )
        self.machine_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.machine_combo.bind("<<ComboboxSelected>>", self.on_machine_change)

        self.ip_label = ttk.Label(
            machine_line,
            text="IP: ---",
            font=("Consolas", 10, "bold"),
            bootstyle=(PRIMARY)
        )
        self.ip_label.grid(row=0, column=2, padx=15, pady=5, sticky="w")

        # ================= ARQUIVOS PDF =================
        files_frame = ttk.LabelFrame(main_frame, text="📄 Arquivos PDF", bootstyle=(PRIMARY))
        files_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        files_frame.columnconfigure(0, weight=1)

        self.drop_area = tk.Label(
            files_frame,
            text="Arraste arquivos PDF para cá",
            bg="#e8e8e8",
            height=3,
            relief=tk.FLAT,
            font=("Arial", 10)
        )
        self.drop_area.grid(row=0, column=0, sticky="ew", pady=5)
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.on_drop)

        list_frame = ttk.Frame(files_frame)
        list_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        files_frame.rowconfigure(1, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_frame, height=6, selectmode=tk.EXTENDED)
        scroll_list = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll_list.set)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scroll_list.grid(row=0, column=1, sticky="ns")
        self.listbox.bind("<Double-1>", self.add_files_manually)

        btn_frame = ttk.Frame(files_frame)
        btn_frame.grid(row=2, column=0, sticky="ew", pady=5)
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        btn_frame.columnconfigure(3, weight=1)

        ttk.Button(
            btn_frame,
            text="❌ Remover selecionados",
            command=self.delete_selected,
            bootstyle=(DANGER)
        ).grid(row=0, column=0, padx=2, sticky="ew")

        ttk.Button(
            btn_frame,
            text="🧹 Limpar todos",
            command=self.clear_all,
            bootstyle=(WARNING)
        ).grid(row=0, column=1, padx=2, sticky="ew")

        self.send_btn = ttk.Button(
            btn_frame,
            text="📤 Enviar",
            command=self.start_send,
            bootstyle=(SUCCESS)
        )
        self.send_btn.grid(row=0, column=2, padx=2, sticky="ew")

        self.cancel_btn = ttk.Button(
            btn_frame,
            text="❌ Cancelar",
            command=self.cancel,
            state=tk.DISABLED,
            bootstyle=(INFO)
        )
        self.cancel_btn.grid(row=0, column=3, padx=2, sticky="ew")

        self.counter_label = ttk.Label(files_frame, text="0 arquivo(s)")
        self.counter_label.grid(row=3, column=0, sticky="w", pady=2)

        # ================= CONFIGURAÇÃO DE IMPRESSÃO =================
        print_frame = ttk.LabelFrame(
            main_frame,
            text="⚙️ Configuração de impressão e identificação",
            bootstyle=(PRIMARY)
        )
        print_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        print_frame.columnconfigure(1, weight=1)
        print_frame.columnconfigure(3, weight=1)
        print_frame.columnconfigure(5, weight=1)

        ttk.Label(print_frame, text="Fila cadastrada:", width=15).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.format_cb = ttk.Combobox(print_frame, state="readonly", width=25, bootstyle=(PRIMARY))
        self.format_cb.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.format_cb.bind("<<ComboboxSelected>>", self.on_format_change)

        ttk.Label(print_frame, text="Descricao:", width=18).grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.paper_cb = ttk.Combobox(print_frame, state="disabled", width=25, bootstyle=(PRIMARY))
        self.paper_cb.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(print_frame, text="Identificação (opcional):", width=20).grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.id_entry = ttk.Entry(print_frame, width=30, bootstyle=(PRIMARY))
        self.id_entry.grid(row=0, column=5, padx=5, pady=5, sticky="w")
        self._setup_placeholder()

        # ================= LOG DE EVENTOS =================
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de eventos", bootstyle=(PRIMARY))
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            height=8,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg='#1e1e1e',
            fg='#dcdcdc',
            insertbackground='white',
            padx=10,
            pady=10
        )
        scroll_log = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll_log.set)

        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll_log.grid(row=0, column=1, sticky="ns")

        # ================= ANDAMENTO =================
        progress_frame = ttk.LabelFrame(main_frame, text="⏳ Andamento", bootstyle=(SUCCESS))
        progress_frame.grid(row=4, column=0, sticky="ew", pady=(0, 5))
        progress_frame.columnconfigure(0, weight=1)

        self.current_file_label = ttk.Label(progress_frame, text="")
        self.current_file_label.grid(row=0, column=0, sticky="w", padx=5, pady=(5, 0))

        bar_frame = ttk.Frame(progress_frame)
        bar_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))
        bar_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            bar_frame,
            mode='determinate',
            bootstyle="success-striped"
        )
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.percent_label = ttk.Label(bar_frame, text="0%")
        self.percent_label.grid(row=0, column=1)

        # Inicializa
        # Aguarda o cadastro central; nao utiliza lista interna de IPs.
        self.reset_progress()
    
    def _add_tooltip(self, widget, text):
        def enter(e):
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{widget.winfo_rootx()+10}+{widget.winfo_rooty()+20}")
            label = tk.Label(self.tooltip, text=text, background="#ffffe0", relief=tk.SOLID, borderwidth=1)
            label.pack()
        def leave(e):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def _setup_placeholder(self):
        self.placeholder = "ex: cliente_lote"
        self.id_entry.insert(0, self.placeholder)
        self.id_entry.configure(foreground="gray")
        self.id_entry.bind("<FocusIn>", self._on_focus_in)
        self.id_entry.bind("<FocusOut>", self._on_focus_out)

    def _on_focus_in(self, event):
        if self.id_entry.get() == self.placeholder:
            self.id_entry.delete(0, tk.END)
            self.id_entry.configure(foreground="black")

    def _on_focus_out(self, event):
        if not self.id_entry.get():
            self.id_entry.insert(0, self.placeholder)
            self.id_entry.configure(foreground="gray")

    def _get_optional(self):
        text = self.id_entry.get().strip()
        if text == self.placeholder or not text:
            return ""
        return text

    def log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def update_counter(self):
        count = len(self.file_paths)
        self.counter_label.config(text=f"{count} arquivo(s)" if count != 1 else "1 arquivo")
        if count == 0:
            self.reset_progress()

    def reset_progress(self):
        self.progress['value'] = 0
        self.progress['maximum'] = 1
        self.percent_label.config(text="0%")
        self.current_file_label.config(text="Nenhum envio em andamento")

    def atualizar_equipamentos(self, equipamentos):
        anterior=self.machine_var.get()
        self.equipamentos_disponiveis={r['nome']:r for r in equipamentos}
        nomes=list(self.equipamentos_disponiveis)
        self.machine_combo.configure(values=nomes)
        self.machine_var.set(anterior if anterior in nomes else (nomes[0] if len(nomes)==1 else ''))
        # A consulta das filas fica sob demanda, para evitar rede na abertura.
        maquina=self.equipamentos_disponiveis.get(self.machine_var.get())
        self.ip_label.configure(text=f"IP: {maquina['ip']}" if maquina else 'IP: ---')
        self.format_cb.configure(values=[])
        self.format_cb.set('')
        self.paper_cb.set('')

    def on_machine_change(self, event=None):
        maquina=self.equipamentos_disponiveis.get(self.machine_var.get()) if hasattr(self,'equipamentos_disponiveis') else None
        if not maquina:
            self.ip_label.configure(text='IP: ---')
            self.format_cb.configure(values=[])
            self.format_cb.set('')
            self.paper_cb.set('')
            return
        self.ip_label.configure(text=f"IP: {maquina['ip']}")
        try:
            filas=self.app.repositorio_equipamentos.listar_filas(maquina['id'])
        except Exception as exc:
            self.log(f'Cadastro de filas indisponivel: {exc}')
            filas=[]
        self.filas_disponiveis={r['nome']:r for r in filas}
        self.format_cb.configure(values=list(self.filas_disponiveis))
        self.format_cb.set('')
        self.paper_cb.set('')

    def on_format_change(self, event=None):
        fila=getattr(self,'filas_disponiveis',{}).get(self.format_cb.get())
        self.paper_cb.set(fila['descricao'] if fila else '')

    def on_drop(self, event):
        files = re.findall(r'\{.*?\}|\S+', event.data)
        added = 0
        for f in files:
            f = f.strip('{}')
            f = os.path.normpath(f)
            if f.lower().endswith(".pdf") and os.path.exists(f) and f not in self.file_paths:
                self.file_paths.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))
                added += 1
        if added:
            self.update_counter()
            self.log(f"+{added} arquivo(s) via drag & drop")

    def add_files_manually(self, event=None):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        added = 0
        for f in files:
            f = os.path.normpath(f)
            if f not in self.file_paths:
                self.file_paths.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))
                added += 1
        if added:
            self.update_counter()
            self.log(f"+{added} arquivo(s) adicionado(s) manualmente")

    def delete_selected(self):
        selected = self.listbox.curselection()
        for idx in reversed(selected):
            self.listbox.delete(idx)
            del self.file_paths[idx]
        self.update_counter()
        self.log(f"{len(selected)} arquivo(s) removido(s)")

    def clear_all(self):
        self.listbox.delete(0, tk.END)
        self.file_paths.clear()
        self.update_counter()
        self.log("Lista limpa")

    def start_send(self):
        machine=self.machine_var.get()
        maquina=getattr(self,'equipamentos_disponiveis',{}).get(machine)
        if not maquina:
            messagebox.showwarning('GPrint','Selecione um equipamento cadastrado.')
            return
        fila=self.format_cb.get()
        if fila not in getattr(self,'filas_disponiveis',{}):
            messagebox.showwarning('GPrint','Selecione uma fila cadastrada para este equipamento.')
            return
        if not self.file_paths:
            messagebox.showwarning('Aviso','Nenhum arquivo para enviar.')
            return
        if self.sending:
            messagebox.showinfo('Aguarde','Envio ja em andamento.')
            return
        total=len(self.file_paths)
        self.progress['maximum']=total
        self.progress['value']=0
        self.percent_label.config(text='0%')
        self.current_file_label.config(text='Preparando envio...')
        self.sending=True
        self.cancel_send=False
        self.send_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        # O envio usa o nome EXATO de fila cadastrado, sem alterar os argumentos LPR.
        self.equipamento_envio_atual=maquina.copy()
        thread=threading.Thread(target=self._send_thread,args=(machine,fila,None),daemon=True)
        thread.start()

    def cancel(self):
        self.cancel_send = True
        self.log("Cancelamento solicitado...")

    def _send_thread(self, machine, formato, papel):
        config = self.equipamento_envio_atual
        ip = config["ip"]
        optional = self._get_optional()
        current_date = datetime.now().strftime("%d-%m-%Y")
        total = len(self.file_paths)
        success = 0

        queue_name = formato  # Nome exato cadastrado na tabela filas.
        self.log(f"Fila selecionada: {queue_name}")

        for idx, orig_path in enumerate(self.file_paths, 1):
            if self.cancel_send:
                break

            file_basename = os.path.basename(orig_path)
            self.after(0, lambda v=idx-1: self.progress.configure(value=v))
            percent_antes = int((idx-1) / total * 100)
            self.after(0, lambda p=percent_antes: self.percent_label.config(text=f"{p}%"))
            self.after(0, lambda i=idx, t=total, nome=file_basename:
                       self.current_file_label.config(text=f"Enviando ({i}/{t}): {nome}"))

            base_name = file_basename[:40]
            fictitious = f"{base_name}-{current_date}-{optional}".strip("-")
            temp_path = os.path.join(self.temp_dir, f"{machine}_{idx}_{os.getpid()}.pdf")
            try:
                shutil.copy2(orig_path, temp_path)
            except Exception as e:
                self.log(f"ERRO cópia: {e}")
                continue

            cmd = f'lpr -S {ip} -P "{queue_name}" -J "{fictitious}" "{temp_path}"'
            self.log(f"Enviando: {fictitious} para {ip} (fila: {queue_name})")
            try:
                result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=90)
                if result.returncode == 0:
                    success += 1
                    self.log(f"✓ OK: {fictitious}")
                else:
                    self.log(f"✗ FALHA: {fictitious} - {result.stderr.strip()}")
            except subprocess.TimeoutExpired:
                self.log(f"⏱ TIMEOUT: {fictitious}")
            except Exception as e:
                self.log(f"⚠ ERRO: {fictitious} - {e}")
            finally:
                try:
                    os.remove(temp_path)
                except:
                    pass

            self.after(0, lambda v=idx: self.progress.configure(value=v))
            percent_depois = int(idx / total * 100)
            self.after(0, lambda p=percent_depois: self.percent_label.config(text=f"{p}%"))

        self.after(0, lambda: self.current_file_label.config(text="Envio finalizado" if success == total else "Envio com falhas"))
        self.after(0, self.cancel_btn.config, {'state': tk.DISABLED})
        self.after(0, self.send_btn.config, {'state': tk.NORMAL})
        self.sending = False

        if not self.cancel_send:
            if success == total:
                self.log(f"✅ Todos os {total} arquivos enviados com sucesso!")
                self.after(0, lambda: messagebox.showinfo("Sucesso", f"{total} arquivo(s) enviado(s)"))
                self.after(0, self.clear_all)
                self.after(0, self.reset_progress)
            else:
                self.log(f"⚠ Enviados {success} de {total} (com falhas).")
                self.after(0, lambda: messagebox.showerror("Erro parcial", f"{success}/{total} enviados. Veja o log."))
        else:
            self.log("Envio cancelado pelo usuário.")
            self.after(0, self.reset_progress)

    def _on_destroy(self, event=None):
        try:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        except:
            pass
import os
import io
import shutil
import subprocess
import tempfile
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime

import fitz  # PyMuPDF
from PIL import Image, ImageOps
import openpyxl
class RefazerTab(ttk.Frame):
    # ---------- CONFIGURAÇÕES (agora atributos da classe) ----------
    CAMINHO_CAPAS_POD_AGRUPADAS = r"\\192.168.200.253\producao\____REPOSITORIO GERAL POD\_CAPAS DE POD AGRUPADAS"
    CAMINHO_UICLAP_LOTES = r"\\192.168.200.253\producao\____REPOSITORIO GERAL POD\UICLAP\LOTES"
    CAMINHO_RELATORIOS_PERDA = r"\\192.168.200.253\producao\____REPOSITORIO GERAL POD\_CAPAS DE POD AGRUPADAS\Z_RELATÓRIOS\RELATÓRIOS DE PERDA.xlsx"
    GS_PATH = r"C:\Ghostscript\gswin64c.exe"

    LARGURA_CM = 45
    ALTURA_CM = 32
    CM_TO_PT = 28.3465
    SANGRIA_CM = 0.3

    LARGURA_PT = LARGURA_CM * CM_TO_PT
    ALTURA_PT = ALTURA_CM * CM_TO_PT
    SANGRIA_PT = SANGRIA_CM * CM_TO_PT

    TAM_MARCA = 5 * CM_TO_PT / 10   # 5 mm
    GAP_MARCA = 2 * CM_TO_PT / 10   # 2 mm

    SETORES_PERDA = [ "PRE-IMPRESSÃO", "PROBLEMA DE FONTE","IMPRESSÃO", "REFILE INICIAL MIOLO","REFILE INICIAL CAPA","SEPARAÇÃO", "ENCADERNAÇÃO", "LAMINAÇÃO","REFILE ORELHA",
                     "REFILE FINAL", "SHIRINK","EXPEDIÇÃO"]
    MAQUINAS_REF = ["RICOH", "OCÉ 1", "OCÉ 3", "BALTORO","ATEXCO"]
    OPERADORES = ["FRAN", "MONALISA", "PALOMA", "PEDRO", "ISMENIO", "CAMILA", "EDUARDO", "DIOGO", "KENNEDY"]

    def __init__(self, parent, app=None):
        super().__init__(parent)
        self.parent = parent
        self.app = app  # pode ser usado para acessar recursos do app principal

        # Variáveis de estado
        self.lista_completa = []
        self.pasta_var = tk.StringVar()
        self.subpasta_var = tk.StringVar()
        self.operador_var = tk.StringVar()
        self.filtro_var = tk.StringVar()
        self.idx_capa = -1
        self.idx_miolo = -1
        self.idx_caminho = -1

        self._criar_interface()
        self.carregar_clientes()

    # ---------- FUNÇÕES AUXILIARES (métodos da classe) ----------
    @staticmethod
    def get_border_pixmap(page, rect, zoom=2.0):
        pix = page.get_pixmap(clip=rect, dpi=int(72 * zoom))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    @staticmethod
    def espelhar_borda(img, flip_horizontal=False, flip_vertical=False):
        if flip_horizontal:
            img = ImageOps.mirror(img)
        if flip_vertical:
            img = ImageOps.flip(img)
        return img

    @staticmethod
    def inserir_imagem_na_pagina(page, img, rect, qualidade=95):
        with io.BytesIO() as output:
            img.save(output, format='PNG', quality=qualidade, optimize=True)
            data = output.getvalue()
        page.insert_image(rect, stream=data, keep_proportion=False)

    @staticmethod
    def desenhar_marca_central_vertical(page, x, y0, y1, tam=None, gap=None):
        if tam is None:
            tam = RefazerTab.TAM_MARCA
        if gap is None:
            gap = RefazerTab.GAP_MARCA
        s = page.new_shape()
        s.draw_line((x, y0 - tam), (x, y0 - gap))
        s.draw_line((x, y1 + tam), (x, y1 + gap))
        s.finish(width=0.5)
        s.commit()

    @staticmethod
    def aplicar_sangria_espelhada_externa(page, nova, trim_original, destino, sangria_pt, bordas, rotate=0):
        zoom = 2.0
        mapa_bordas = {
            0:   {'top': 'top', 'bottom': 'bottom', 'left': 'left', 'right': 'right'},
            90:  {'top': 'left', 'bottom': 'right', 'left': 'bottom', 'right': 'top'},
            180: {'top': 'bottom', 'bottom': 'top', 'left': 'right', 'right': 'left'},
            270: {'top': 'right', 'bottom': 'left', 'left': 'top', 'right': 'bottom'}
        }
        mapa = mapa_bordas.get(rotate % 360, mapa_bordas[0])

        for borda_dest in bordas:
            borda_orig = mapa[borda_dest]
            if borda_orig == 'top':
                rect = fitz.Rect(trim_original.x0, trim_original.y0,
                                 trim_original.x1, trim_original.y0 + sangria_pt)
                flip_h, flip_v = False, True
            elif borda_orig == 'bottom':
                rect = fitz.Rect(trim_original.x0, trim_original.y1 - sangria_pt,
                                 trim_original.x1, trim_original.y1)
                flip_h, flip_v = False, True
            elif borda_orig == 'left':
                rect = fitz.Rect(trim_original.x0, trim_original.y0,
                                 trim_original.x0 + sangria_pt, trim_original.y1)
                flip_h, flip_v = True, False
            elif borda_orig == 'right':
                rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y0,
                                 trim_original.x1, trim_original.y1)
                flip_h, flip_v = True, False
            else:
                continue

            img = RefazerTab.get_border_pixmap(page, rect, zoom)
            if rotate == 90:
                img = img.transpose(Image.ROTATE_90)
            elif rotate == 180:
                img = img.transpose(Image.ROTATE_180)
            elif rotate == 270:
                img = img.transpose(Image.ROTATE_270)

            img = RefazerTab.espelhar_borda(img, flip_horizontal=flip_h, flip_vertical=flip_v)
            if borda_dest in ('top', 'bottom'):
                img = img.resize((int(destino.width), int(sangria_pt)), Image.Resampling.LANCZOS)
            else:
                img = img.resize((int(sangria_pt), int(destino.height)), Image.Resampling.LANCZOS)

            if borda_dest == 'top':
                dest_rect = fitz.Rect(destino.x0, destino.y0 - sangria_pt, destino.x1, destino.y0)
            elif borda_dest == 'bottom':
                dest_rect = fitz.Rect(destino.x0, destino.y1, destino.x1, destino.y1 + sangria_pt)
            elif borda_dest == 'left':
                dest_rect = fitz.Rect(destino.x0 - sangria_pt, destino.y0, destino.x0, destino.y1)
            elif borda_dest == 'right':
                dest_rect = fitz.Rect(destino.x1, destino.y0, destino.x1 + sangria_pt, destino.y1)
            else:
                continue

            RefazerTab.inserir_imagem_na_pagina(nova, img, dest_rect)

        cantos = [('top','left'), ('top','right'), ('bottom','left'), ('bottom','right')]
        for b1, b2 in cantos:
            if b1 in bordas and b2 in bordas:
                borda_orig1 = mapa[b1]
                borda_orig2 = mapa[b2]
                if (borda_orig1 == 'top' and borda_orig2 == 'left') or (borda_orig1 == 'left' and borda_orig2 == 'top'):
                    rect = fitz.Rect(trim_original.x0, trim_original.y0,
                                     trim_original.x0 + sangria_pt, trim_original.y0 + sangria_pt)
                elif (borda_orig1 == 'top' and borda_orig2 == 'right') or (borda_orig1 == 'right' and borda_orig2 == 'top'):
                    rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y0,
                                     trim_original.x1, trim_original.y0 + sangria_pt)
                elif (borda_orig1 == 'bottom' and borda_orig2 == 'left') or (borda_orig1 == 'left' and borda_orig2 == 'bottom'):
                    rect = fitz.Rect(trim_original.x0, trim_original.y1 - sangria_pt,
                                     trim_original.x0 + sangria_pt, trim_original.y1)
                elif (borda_orig1 == 'bottom' and borda_orig2 == 'right') or (borda_orig1 == 'right' and borda_orig2 == 'bottom'):
                    rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y1 - sangria_pt,
                                     trim_original.x1, trim_original.y1)
                else:
                    continue

                img = RefazerTab.get_border_pixmap(page, rect, zoom)
                if rotate == 90:
                    img = img.transpose(Image.ROTATE_90)
                elif rotate == 180:
                    img = img.transpose(Image.ROTATE_180)
                elif rotate == 270:
                    img = img.transpose(Image.ROTATE_270)
                img = RefazerTab.espelhar_borda(img, flip_horizontal=True, flip_vertical=True)
                img = img.resize((int(sangria_pt), int(sangria_pt)), Image.Resampling.LANCZOS)

                if b1 == 'top' and b2 == 'left':
                    dest_rect = fitz.Rect(destino.x0 - sangria_pt, destino.y0 - sangria_pt, destino.x0, destino.y0)
                elif b1 == 'top' and b2 == 'right':
                    dest_rect = fitz.Rect(destino.x1, destino.y0 - sangria_pt, destino.x1 + sangria_pt, destino.y0)
                elif b1 == 'bottom' and b2 == 'left':
                    dest_rect = fitz.Rect(destino.x0 - sangria_pt, destino.y1, destino.x0, destino.y1 + sangria_pt)
                elif b1 == 'bottom' and b2 == 'right':
                    dest_rect = fitz.Rect(destino.x1, destino.y1, destino.x1 + sangria_pt, destino.y1 + sangria_pt)
                else:
                    continue
                RefazerTab.inserir_imagem_na_pagina(nova, img, dest_rect)

    @staticmethod
    def pagina_tem_conteudo(page):
        if page.get_contents():
            return True
        if page.get_text("text").strip():
            return True
        if page.get_images():
            return True
        return False

    @staticmethod
    def obter_area_corte_segura(page, w_ref, h_ref):
        if page.trimbox and page.trimbox.width > 0 and page.trimbox.height > 0:
            return page.trimbox
        rect = page.rect
        if rect.width > 0 and rect.height > 0:
            return rect
        return fitz.Rect(0, 0, w_ref, h_ref)

    @staticmethod
    def calcular_y_com_sangria(altura_pt, h, sangria_pt):
        if h + 2 * sangria_pt > altura_pt:
            y_center = (altura_pt - h) / 2
            return y_center, False
        else:
            y_center = (altura_pt - h) / 2
            y_min = sangria_pt
            y_max = altura_pt - h - sangria_pt
            y = max(y_min, min(y_center, y_max))
            return y, True

    @staticmethod
    def ajustar_para_media_box(rect, largura, altura):
        x0 = max(0, min(rect.x0, largura - 1))
        y0 = max(0, min(rect.y0, altura - 1))
        x1 = max(x0 + 1, min(rect.x1, largura))
        y1 = max(y0 + 1, min(rect.y1, altura))
        return fitz.Rect(x0, y0, x1, y1)

    @staticmethod
    def desenhar_marcas(pagina, destino, lado):
        s = pagina.new_shape()
        x0, y0, x1, y1 = destino.x0, destino.y0, destino.x1, destino.y1
        tam = RefazerTab.TAM_MARCA
        gap = RefazerTab.GAP_MARCA

        # MARCAS VERTICAIS
        s.draw_line((x0, y0 - tam), (x0, y0 - gap))
        s.draw_line((x1, y0 - tam), (x1, y0 - gap))
        s.draw_line((x0, y1 + tam), (x0, y1 + gap))
        s.draw_line((x1, y1 + tam), (x1, y1 + gap))

        # MARCAS HORIZONTAIS
        if lado == 'esquerda':
            s.draw_line((x0 - tam, y0), (x0 - gap, y0))
            s.draw_line((x0 - tam, y1), (x0 - gap, y1))
        elif lado == 'direita':
            s.draw_line((x1 + tam, y0), (x1 + gap, y0))
            s.draw_line((x1 + tam, y1), (x1 + gap, y1))

        s.finish(width=0.5, color=(0,0,0))
        s.commit()

    def montar_pdf(self, entrada, saida, tipo):
        doc = fitz.open(entrada)
        out = fitz.open()
        w_ref = h_ref = 595
        if len(doc) > 0:
            w_ref, h_ref = doc[0].rect.width, doc[0].rect.height

        largura_pt = self.LARGURA_PT
        altura_pt = self.ALTURA_PT
        sangria_pt = self.SANGRIA_PT
        espaco_pt = 0
        paginas_vazias = []

        if tipo == 'brochura':
            total_original = len(doc)
            if total_original % 4 != 0:
                faltam = 4 - (total_original % 4)
                for _ in range(faltam):
                    doc.new_page(width=w_ref, height=h_ref)
            total = len(doc)
            metade = total // 2
            num_faces = total // 2

            for i in range(num_faces):
                if i % 2 == 0:
                    idx_esq = metade + i
                    idx_dir = i
                else:
                    idx_esq = i
                    idx_dir = metade + i
                idx_esq = idx_esq if idx_esq < total else None
                idx_dir = idx_dir if idx_dir < total else None

                page_esq = doc[idx_esq] if idx_esq is not None else None
                page_dir = doc[idx_dir] if idx_dir is not None else None
                rot_esq = (idx_esq is not None and idx_esq >= metade)
                rot_dir = (idx_dir is not None and idx_dir >= metade)

                trim_esq = self.obter_area_corte_segura(page_esq, w_ref, h_ref) if page_esq else None
                w_esq = trim_esq.width if trim_esq else w_ref
                h_esq = trim_esq.height if trim_esq else h_ref
                trim_dir = self.obter_area_corte_segura(page_dir, w_ref, h_ref) if page_dir else None
                w_dir = trim_dir.width if trim_dir else w_ref
                h_dir = trim_dir.height if trim_dir else h_ref

                nova = out.new_page(width=largura_pt, height=altura_pt)

                if sangria_pt > 0:
                    y_esq, cabe_esq = self.calcular_y_com_sangria(altura_pt, h_esq, sangria_pt)
                    y_dir, cabe_dir = self.calcular_y_com_sangria(altura_pt, h_dir, sangria_pt)
                    sangria_aplicar_esq = cabe_esq and page_esq is not None
                    sangria_aplicar_dir = cabe_dir and page_dir is not None
                else:
                    y_esq = (altura_pt - h_esq) / 2
                    y_dir = (altura_pt - h_dir) / 2
                    sangria_aplicar_esq = sangria_aplicar_dir = False

                largura_total = w_esq + w_dir + espaco_pt + 2 * sangria_pt
                x_inicio = (largura_pt - largura_total) / 2
                x1 = x_inicio + sangria_pt
                destino1 = fitz.Rect(x1, y_esq, x1 + w_esq, y_esq + h_esq)
                destino1 = self.ajustar_para_media_box(destino1, largura_pt, altura_pt)
                x2 = x1 + w_esq + espaco_pt
                destino2 = fitz.Rect(x2, y_dir, x2 + w_dir, y_dir + h_dir)
                destino2 = self.ajustar_para_media_box(destino2, largura_pt, altura_pt)

                if page_esq and self.pagina_tem_conteudo(page_esq):
                    try:
                        nova.show_pdf_page(destino1, doc, idx_esq, clip=trim_esq, rotate=180 if rot_esq else 0)
                    except Exception:
                        paginas_vazias.append(str(idx_esq+1))
                elif page_esq is not None:
                    paginas_vazias.append(str(idx_esq+1))

                if page_dir and self.pagina_tem_conteudo(page_dir):
                    try:
                        nova.show_pdf_page(destino2, doc, idx_dir, clip=trim_dir, rotate=180 if rot_dir else 0)
                    except Exception:
                        paginas_vazias.append(str(idx_dir+1))
                elif page_dir is not None:
                    paginas_vazias.append(str(idx_dir+1))

                if sangria_pt > 0:
                    if page_esq and self.pagina_tem_conteudo(page_esq) and sangria_aplicar_esq:
                        self.aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1, sangria_pt,
                                                               bordas=['top','bottom','left'],
                                                               rotate=180 if rot_esq else 0)
                    if page_dir and self.pagina_tem_conteudo(page_dir) and sangria_aplicar_dir:
                        self.aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2, sangria_pt,
                                                               bordas=['top','bottom','right'],
                                                               rotate=180 if rot_dir else 0)
                    bleed_total = fitz.Rect(
                        destino1.x0 - sangria_pt,
                        min(destino1.y0, destino2.y0) - sangria_pt,
                        destino2.x1 + sangria_pt,
                        max(destino1.y1, destino2.y1) + sangria_pt
                    )
                    bleed_total = self.ajustar_para_media_box(bleed_total, largura_pt, altura_pt)
                    if bleed_total.width > 0 and bleed_total.height > 0:
                        nova.set_bleedbox(bleed_total)
                nova.set_trimbox(destino1)
                if espaco_pt == 0:
                    self.desenhar_marcas(nova, destino1, lado='esquerda')
                    self.desenhar_marcas(nova, destino2, lado='direita')
                    x_centro = destino1.x1
                    self.desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1)

        elif tipo == 'grampo':
            total = len(doc)
            num_folhas = (total + 1) // 2
            for i in range(num_folhas):
                if i % 2 == 0:
                    idx_esq = total - i if total - i >= 1 else None
                    idx_dir = i + 1 if i + 1 <= total else None
                else:
                    idx_esq = i + 1 if i + 1 <= total else None
                    idx_dir = total - i if total - i >= 1 else None
                page_esq = doc[idx_esq - 1] if idx_esq is not None else None
                page_dir = doc[idx_dir - 1] if idx_dir is not None else None

                trim_esq = self.obter_area_corte_segura(page_esq, w_ref, h_ref) if page_esq else None
                w_esq = trim_esq.width if trim_esq else w_ref
                h_esq = trim_esq.height if trim_esq else h_ref
                trim_dir = self.obter_area_corte_segura(page_dir, w_ref, h_ref) if page_dir else None
                w_dir = trim_dir.width if trim_dir else w_ref
                h_dir = trim_dir.height if trim_dir else h_ref

                nova = out.new_page(width=largura_pt, height=altura_pt)

                if sangria_pt > 0:
                    y_esq, cabe_esq = self.calcular_y_com_sangria(altura_pt, h_esq, sangria_pt)
                    y_dir, cabe_dir = self.calcular_y_com_sangria(altura_pt, h_dir, sangria_pt)
                    sangria_aplicar_esq = cabe_esq and page_esq is not None
                    sangria_aplicar_dir = cabe_dir and page_dir is not None
                else:
                    y_esq = (altura_pt - h_esq) / 2
                    y_dir = (altura_pt - h_dir) / 2
                    sangria_aplicar_esq = sangria_aplicar_dir = False

                largura_total = w_esq + w_dir + espaco_pt + 2 * sangria_pt
                x_inicio = (largura_pt - largura_total) / 2
                x1 = x_inicio + sangria_pt
                destino1 = fitz.Rect(x1, y_esq, x1 + w_esq, y_esq + h_esq)
                destino1 = self.ajustar_para_media_box(destino1, largura_pt, altura_pt)
                x2 = x1 + w_esq + espaco_pt
                destino2 = fitz.Rect(x2, y_dir, x2 + w_dir, y_dir + h_dir)
                destino2 = self.ajustar_para_media_box(destino2, largura_pt, altura_pt)

                if page_esq and self.pagina_tem_conteudo(page_esq):
                    try:
                        nova.show_pdf_page(destino1, doc, idx_esq - 1, clip=trim_esq)
                    except Exception:
                        paginas_vazias.append(str(idx_esq))
                elif page_esq is not None:
                    paginas_vazias.append(str(idx_esq))

                if page_dir and self.pagina_tem_conteudo(page_dir):
                    try:
                        nova.show_pdf_page(destino2, doc, idx_dir - 1, clip=trim_dir)
                    except Exception:
                        paginas_vazias.append(str(idx_dir))
                elif page_dir is not None:
                    paginas_vazias.append(str(idx_dir))

                if sangria_pt > 0:
                    if page_esq and self.pagina_tem_conteudo(page_esq) and sangria_aplicar_esq:
                        self.aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1, sangria_pt,
                                                               bordas=['top','bottom','left'])
                    if page_dir and self.pagina_tem_conteudo(page_dir) and sangria_aplicar_dir:
                        self.aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2, sangria_pt,
                                                               bordas=['top','bottom','right'])
                    bleed_total = fitz.Rect(
                        destino1.x0 - sangria_pt,
                        min(destino1.y0, destino2.y0) - sangria_pt,
                        destino2.x1 + sangria_pt,
                        max(destino1.y1, destino2.y1) + sangria_pt
                    )
                    bleed_total = self.ajustar_para_media_box(bleed_total, largura_pt, altura_pt)
                    if bleed_total.width > 0 and bleed_total.height > 0:
                        nova.set_bleedbox(bleed_total)
                nova.set_trimbox(destino1)
                if espaco_pt == 0:
                    self.desenhar_marcas(nova, destino1, lado='esquerda')
                    self.desenhar_marcas(nova, destino2, lado='direita')
                    x_centro = destino1.x1
                    self.desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1)

        elif tipo == '1up':
            for idx, page in enumerate(doc):
                trim_original = self.obter_area_corte_segura(page, w_ref, h_ref)
                w = trim_original.width if trim_original else w_ref
                h = trim_original.height if trim_original else h_ref

                nova = out.new_page(width=largura_pt, height=altura_pt)

                x = (largura_pt - w) / 2
                if sangria_pt > 0:
                    y, cabe = self.calcular_y_com_sangria(altura_pt, h, sangria_pt)
                    sangria_aplicar = cabe
                else:
                    y = (altura_pt - h) / 2
                    sangria_aplicar = False

                destino = fitz.Rect(x, y, x + w, y + h)
                destino = self.ajustar_para_media_box(destino, largura_pt, altura_pt)

                if not self.pagina_tem_conteudo(page):
                    paginas_vazias.append(str(page.number + 1))
                else:
                    try:
                        nova.show_pdf_page(destino, doc, page.number, clip=trim_original)
                    except Exception:
                        paginas_vazias.append(str(page.number + 1))

                nova.set_trimbox(destino)

                if sangria_pt > 0 and sangria_aplicar and self.pagina_tem_conteudo(page):
                    bleed = fitz.Rect(
                        destino.x0 - sangria_pt,
                        destino.y0 - sangria_pt,
                        destino.x1 + sangria_pt,
                        destino.y1 + sangria_pt
                    )
                    bleed = self.ajustar_para_media_box(bleed, largura_pt, altura_pt)
                    if bleed.width > 0 and bleed.height > 0:
                        nova.set_bleedbox(bleed)
                    self.aplicar_sangria_espelhada_externa(
                        page, nova, trim_original, destino, sangria_pt,
                        bordas=['top', 'bottom', 'left', 'right']
                    )

                if espaco_pt == 0:
                    self.desenhar_marcas(nova, destino, lado='centro')  # adapte se necessário
        


        out.save(saida, garbage=4, deflate=True, clean=True)
        doc.close()
        out.close()
        if paginas_vazias:
            print(f"Páginas vazias (não incluídas): {', '.join(paginas_vazias)}")

    def adicionar_marca_dagua(self, caminho_pdf, texto, distancia_mm=3, caminho_saida=None):
        if not caminho_pdf.lower().endswith('.pdf'):
            return
        try:
            doc = fitz.open(caminho_pdf)
        except Exception as e:
            print(f"Erro ao abrir PDF {caminho_pdf}: {e}")
            return

        margem_pt = distancia_mm * 2.83464567
        fonte_tamanho = 11
        cor = (1, 0, 0)
        for pagina in doc:
            w = pagina.rect.width
            x0 = 50
            x1 = w - 50
            y0 = margem_pt
            y1 = y0 + 40
            rect = fitz.Rect(x0, y0, x1, y1)
            pagina.insert_textbox(rect, texto, fontsize=fonte_tamanho, color=cor, fontname="helv", align=1)

        destino = caminho_saida if caminho_saida else caminho_pdf
        dir_destino = os.path.dirname(destino)
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(dir=dir_destino, delete=False, suffix='.pdf') as tmp:
                tmp_path = tmp.name
            doc.save(tmp_path)
            doc.close()
            os.replace(tmp_path, destino)
        except Exception as e:
            print(f"Erro ao salvar PDF {destino}: {e}")
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)
            doc.close()

    def comprimir_pdf(self, entrada, saida):
        if not os.path.exists(self.GS_PATH):
            raise FileNotFoundError(f"Ghostscript não encontrado em: {self.GS_PATH}")
        comando = [
            self.GS_PATH,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=/ebook",
            "-dDownsampleColorImages=true",
            "-dDownsampleGrayImages=true",
            "-dDownsampleMonoImages=true",
            "-dColorImageResolution=150",
            "-dGrayImageResolution=150",
            "-dMonoImageResolution=150",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={saida}",
            entrada
        ]
        try:
            subprocess.run(comando, check=True)
        except subprocess.CalledProcessError as e:
            print(f"Erro ao comprimir: {e}")
            shutil.copy(entrada, saida)

    def converter_em_contornos(self, entrada, saida):
        if not os.path.exists(self.GS_PATH):
            raise FileNotFoundError(f"Ghostscript não encontrado em: {self.GS_PATH}")
        comando = [
            self.GS_PATH,
            "-dNOPAUSE",
            "-dBATCH",
            "-dQUIET",
            "-sDEVICE=pdfwrite",
            "-dNoOutputFonts",
            f"-sOutputFile={saida}",
            entrada
        ]
        try:
            subprocess.run(comando, check=True)
        except subprocess.CalledProcessError as e:
            print(f"Erro ao converter em contornos: {e}")
            if entrada != saida:
                shutil.copy(entrada, saida)

    def listar_pastas(self, diretorio, reverse=False):
        try:
            pastas = [p for p in os.listdir(diretorio) if os.path.isdir(os.path.join(diretorio, p))]
            pastas.sort(reverse=reverse)
            return pastas
        except Exception as e:
            messagebox.showerror("Erro", str(e), parent=self)
            return []

    @staticmethod
    def obter_indice_coluna(ws, nome_coluna):
        cabecalho = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not cabecalho:
            return -1
        for i, valor in enumerate(cabecalho):
            if valor and str(valor).strip().lower() == nome_coluna.lower():
                return i
        return -1

    def obter_lista_arquivos_da_planilha(self, caminho_excel, cliente):
        wb = openpyxl.load_workbook(caminho_excel, data_only=True)
        ws = wb.active
        idx_capa = self.obter_indice_coluna(ws, "ARQUIVO CAPA")
        idx_miolo = self.obter_indice_coluna(ws, "ARQUIVO MIOLO")
        idx_caminho = self.obter_indice_coluna(ws, "CAMINHO DO ARQUIVO")

        if cliente == "uiclap":
            idx_acabamento = self.obter_indice_coluna(ws, "Acabamento")
            idx_formato_miolo = self.obter_indice_coluna(ws, "FORMATO MIOLO")
        else:
            idx_acabamento = self.obter_indice_coluna(ws, "Encadernação")
            idx_formato_miolo = self.obter_indice_coluna(ws, "FORMATO MIOLO")

        lista = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if idx_capa != -1 and row[idx_capa]:
                nome = str(row[idx_capa]).strip()
                caminho_fisico = None
                if idx_caminho != -1 and row[idx_caminho]:
                    caminho_fisico = os.path.join(str(row[idx_caminho]).strip(), nome)
                acabamento = str(row[idx_acabamento]).strip() if idx_acabamento != -1 and row[idx_acabamento] else None
                formato_miolo = str(row[idx_formato_miolo]).strip() if idx_formato_miolo != -1 and row[idx_formato_miolo] else None
                lista.append({"nome": nome, "tipo": "CAPA", "linha": row, "caminho_fisico": caminho_fisico,
                              "acabamento": acabamento, "formato_miolo": formato_miolo})
            if idx_miolo != -1 and row[idx_miolo]:
                nome = str(row[idx_miolo]).strip()
                caminho_fisico = None
                if idx_caminho != -1 and row[idx_caminho]:
                    caminho_fisico = os.path.join(str(row[idx_caminho]).strip(), nome)
                acabamento = str(row[idx_acabamento]).strip() if idx_acabamento != -1 and row[idx_acabamento] else None
                formato_miolo = str(row[idx_formato_miolo]).strip() if idx_formato_miolo != -1 and row[idx_formato_miolo] else None
                lista.append({"nome": nome, "tipo": "MIOLO", "linha": row, "caminho_fisico": caminho_fisico,
                              "acabamento": acabamento, "formato_miolo": formato_miolo})
        wb.close()
        return lista

    def copiar_para_refazer(self, origem, cliente, lote, tipo):
        if not os.path.exists(origem):
            return None
        if cliente == "uiclap":
            destino_dir = os.path.join(self.CAMINHO_UICLAP_LOTES, lote, "REFAZER", tipo)
        else:
            destino_dir = os.path.join(self.CAMINHO_CAPAS_POD_AGRUPADAS, cliente, lote, "REFAZER", tipo)
        os.makedirs(destino_dir, exist_ok=True)
        destino = os.path.join(destino_dir, os.path.basename(origem))
        try:
            shutil.copy2(origem, destino)
            return destino
        except Exception as e:
            print(f"Erro ao copiar {origem}: {e}")
            return None

    # ---------- TOOLTIP (classe interna) ----------
    class ToolTip:
        def __init__(self, widget):
            self.widget = widget
            self.tipwindow = None
            self.id = None
            self.x = self.y = 0
            self.text = ''
            widget.bind('<Enter>', self.enter)
            widget.bind('<Leave>', self.leave)
            widget.bind('<Motion>', self.motion)

        def enter(self, event=None):
            self.schedule()

        def leave(self, event=None):
            self.unschedule()
            self.hidetip()

        def motion(self, event=None):
            self.x, self.y = event.x, event.y

        def schedule(self):
            self.unschedule()
            self.id = self.widget.after(500, self.showtip)

        def unschedule(self):
            id_ = self.id
            self.id = None
            if id_:
                self.widget.after_cancel(id_)

        def showtip(self, event=None):
            item = self.widget.identify_row(self.y)
            col = self.widget.identify_column(self.x)
            if not item or col != '#1':
                return
            nome = self.widget.item(item, 'values')[0]
            self.text = nome
            x, y, _, _ = self.widget.bbox(item, col)
            x += self.widget.winfo_rootx() + 25
            y += self.widget.winfo_rooty() + 25
            self.tipwindow = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            label = tk.Label(tw, text=self.text, justify='left', background="#ffffe0", relief='solid', borderwidth=1,
                             font=("Arial", 9))
            label.pack()

        def hidetip(self):
            tw = self.tipwindow
            self.tipwindow = None
            if tw:
                tw.destroy()

    # ---------- CONSTRUÇÃO DA INTERFACE ----------
    def _criar_interface(self):
        # Configurar grid do frame principal
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)  # PanedWindow vai expandir

        # ================= TOPO - FILTROS =================
        top = ttk.LabelFrame(main_frame, text="Filtros", bootstyle=PRIMARY)
        top.grid(row=0, column=0, sticky="ew", padx=8, pady=5)
        top.grid_columnconfigure(1, weight=1)
        top.grid_columnconfigure(3, weight=2)

        ttk.Label(top, text="Cliente").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.cb_cliente = ttk.Combobox(top, textvariable=self.pasta_var)
        self.cb_cliente.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.cb_cliente.bind("<<ComboboxSelected>>", self.on_cliente_change)

        ttk.Label(top, text="Lote").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.cb_lote = ttk.Combobox(top, textvariable=self.subpasta_var)
        self.cb_lote.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        self.cb_lote.bind("<<ComboboxSelected>>", self.on_lote_change)

        ttk.Label(top, text="Operador").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.cb_operador = ttk.Combobox(top, values=self.OPERADORES, textvariable=self.operador_var)
        self.cb_operador.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.cb_operador.current(0)

        ttk.Label(top, text="Filtro").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        filtro_entry = ttk.Entry(top, textvariable=self.filtro_var)
        filtro_entry.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        filtro_entry.bind("<Return>", lambda e: self.buscar_arquivos())

        # Botões principais (sem progressbar aqui)
        botoes_top = ttk.Frame(top)
        botoes_top.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 5))
        botoes_top.grid_columnconfigure(0, weight=1)
        botoes_top.grid_columnconfigure(1, weight=1)
        botoes_top.grid_columnconfigure(2, weight=1)

        self.btn_buscar = ttk.Button(botoes_top, text="Buscar", command=self.buscar_arquivos,
                                    bootstyle=PRIMARY)
        self.btn_buscar.grid(row=0, column=0, padx=5, sticky="ew")

        self.btn_preparar = ttk.Button(botoes_top, text="Preparar Processamento",
                                    command=self.preparar_processamento, bootstyle=INFO)
        self.btn_preparar.grid(row=0, column=1, padx=5, sticky="ew")

        self.btn_limpar = ttk.Button(botoes_top, text="Limpar", command=self.limpar,
                                    bootstyle=DANGER)
        self.btn_limpar.grid(row=0, column=2, padx=5, sticky="ew")

        paned = tk.PanedWindow(main_frame, orient="vertical", sashrelief="raised", sashwidth=8, bg="#000000")
        paned.grid(row=1, column=0, sticky="nsew", pady=5,padx=10)

        # --- Lista de arquivos ---
        frame_lista = ttk.LabelFrame(paned, text="Arquivos", bootstyle=PRIMARY)
        paned.add(frame_lista, stretch="always")

        tree_container = ttk.Frame(frame_lista)
        tree_container.pack(fill="both", expand=True, padx=5, pady=5)
        tree_container.rowconfigure(0, weight=1)
        tree_container.columnconfigure(0, weight=1)

        self.tree_selecao = ttk.Treeview(tree_container, columns=("Nome",), show="headings", selectmode="extended")
        self.tree_selecao.heading("Nome", text="Nome")
        self.tree_selecao.column("Nome", width=120)
        vsb1 = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree_selecao.yview)
        hsb1 = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree_selecao.xview)
        self.tree_selecao.configure(yscrollcommand=vsb1.set, xscrollcommand=hsb1.set)
        self.tree_selecao.grid(row=0, column=0, sticky="nsew")
        vsb1.grid(row=0, column=1, sticky="ns")
        hsb1.grid(row=1, column=0, sticky="ew")
        self.ToolTip(self.tree_selecao)

        # --- Edição inteligente ---
        self.frame_edicao = ttk.LabelFrame(paned, text="Edição Inteligente", bootstyle=PRIMARY)
        paned.add(self.frame_edicao, stretch="always")

        self.update_idletasks()
        paned.sash_place(0, 0, 330)

        self.frame_edicao.grid_rowconfigure(1, weight=1)
        self.frame_edicao.grid_columnconfigure(0, weight=1)

        toolbar = ttk.Frame(self.frame_edicao)
        toolbar.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 0))
        ttk.Label(toolbar, text="💡 Clique ÚNICO nas colunas Quantidade / Setor / Máquina para editar",
                foreground="#1e5eff", font=("Arial", 10, "bold"), anchor="w").pack(fill="x", padx=5, pady=5)

        edit_container = ttk.Frame(self.frame_edicao)
        edit_container.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        edit_container.rowconfigure(0, weight=1)
        edit_container.columnconfigure(0, weight=1)

        self.tree_edicao = ttk.Treeview(edit_container,
                                        columns=("Nome", "Código", "Quantidade", "Setor", "Máquina"),
                                        show="headings")
        self.tree_edicao.heading("Nome", text="Arquivo")
        self.tree_edicao.heading("Código", text="Código")
        self.tree_edicao.heading("Quantidade", text="Quantidade")
        self.tree_edicao.heading("Setor", text="Setor")
        self.tree_edicao.heading("Máquina", text="Máquina")
        self.tree_edicao.column("Nome", width=80, minwidth=250)
        self.tree_edicao.column("Código", width=20, anchor="center")
        self.tree_edicao.column("Quantidade", width=20, anchor="center")
        self.tree_edicao.column("Setor", width=80, anchor="center")
        self.tree_edicao.column("Máquina", width=60, anchor="center")

        vsb2 = ttk.Scrollbar(edit_container, orient="vertical", command=self.tree_edicao.yview)
        hsb2 = ttk.Scrollbar(edit_container, orient="horizontal", command=self.tree_edicao.xview)
        self.tree_edicao.configure(yscrollcommand=vsb2.set, xscrollcommand=hsb2.set)
        self.tree_edicao.grid(row=0, column=0, sticky="nsew")
        vsb2.grid(row=0, column=1, sticky="ns")
        hsb2.grid(row=1, column=0, sticky="ew")

        # --- Painel inferior de edição (sem botão de processar) ---
        painel_inferior = ttk.Frame(self.frame_edicao)
        painel_inferior.grid(row=2, column=0, sticky="ew", padx=5, pady=(0, 5))

        controles = ttk.Frame(painel_inferior)
        controles.pack(side="left", fill="x", expand=True)

        ttk.Label(controles, text="Setor:").pack(side="left", padx=(0, 5))
        self.cb_setor = ttk.Combobox(controles, values=self.SETORES_PERDA, width=18)
        self.cb_setor.pack(side="left", padx=5)
        self.cb_setor.current(0)

        ttk.Label(controles, text="Máquina:").pack(side="left", padx=(10, 5))
        self.cb_maquina = ttk.Combobox(controles, values=self.MAQUINAS_REF, width=15)
        self.cb_maquina.pack(side="left", padx=5)
        self.cb_maquina.current(0)

        self.btn_aplicar_sel = ttk.Button(controles, text="Aplicar nos Selecionados",
                                        bootstyle=SECONDARY, command=self.aplicar_nos_selecionados)
        self.btn_aplicar_sel.pack(side="left", padx=10)

        self.btn_aplicar_todos = ttk.Button(controles, text="Aplicar em TODOS",
                                            bootstyle=PRIMARY, command=self.aplicar_em_todos)
        self.btn_aplicar_todos.pack(side="left", padx=5)

        # NOTA: O botão "PROCESSAR" que estava aqui foi removido e será criado no status_frame abaixo.

        self.tree_edicao.bind("<Button-1>", self.edicao_inline)

        # ================= ANDAMENTO (PROGRESSBAR) =================
        progress_frame = ttk.LabelFrame(main_frame, text="⏳ Andamento", bootstyle=SUCCESS)
        progress_frame.grid(row=2, column=0, sticky="ew", pady=5, padx=10)
        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            bootstyle="primary-striped"
        )
        self.progress.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 5))

        self.progress_label = ttk.Label(progress_frame, text="0/0")
        self.progress_label.grid(row=0, column=1, padx=(0, 10), pady=5)

        # ================= STATUS E BOTÃO PROCESSAR =================
        # Cria a variável de status se não existir
        if not hasattr(self, 'status_var'):
            self.status_var = tk.StringVar(value="✅ Pronto para processar")

        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=3, column=0, sticky="ew", pady=5, padx=10)
        status_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10, "bold"),
            bootstyle="success"
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.btn_processar = ttk.Button(
            status_frame,
            text="▶ PROCESSAR",
            command=self.executar,   # comando original do botão removido
            bootstyle=SUCCESS,        # pode trocar para SUCCESS se preferir
            width=22
        )
        self.btn_processar.grid(row=0, column=1, sticky="e")

    # ---------- HANDLERS E LÓGICA ----------
    def _limpar_tudo_exceto_cliente_lote(self):
        for i in self.tree_selecao.get_children():
            self.tree_selecao.delete(i)
        for i in self.tree_edicao.get_children():
            self.tree_edicao.delete(i)
        self.lista_completa = []

    def on_cliente_change(self, event=None):
        self._limpar_tudo_exceto_cliente_lote()
        self.atualizar_subpastas()

    def on_lote_change(self, event=None):
        lote = self.subpasta_var.get()
        if not lote:
            return
        self._limpar_tudo_exceto_cliente_lote()
        self.buscar_arquivos()

    def atualizar_subpastas(self, event=None):
        cliente = self.pasta_var.get()
        if cliente == "uiclap":
            base = self.CAMINHO_UICLAP_LOTES
        else:
            base = os.path.join(self.CAMINHO_CAPAS_POD_AGRUPADAS, cliente)
        lotes = self.listar_pastas(base, reverse=True)
        self.cb_lote["values"] = lotes
        if lotes:
            self.cb_lote.current(0)
        else:
            self.cb_lote.set("")

    def carregar_clientes(self):
        clientes = self.listar_pastas(self.CAMINHO_CAPAS_POD_AGRUPADAS)
        clientes.append("uiclap")
        self.cb_cliente["values"] = clientes
        if clientes:
            self.cb_cliente.current(0)
            self.atualizar_subpastas()

    def buscar_arquivos(self):
        cliente = self.pasta_var.get()
        lote = self.subpasta_var.get()
        if not cliente or not lote:
            return
        if cliente == "uiclap":
            pasta = os.path.join(self.CAMINHO_UICLAP_LOTES, lote)
        else:
            pasta = os.path.join(self.CAMINHO_CAPAS_POD_AGRUPADAS, cliente, lote)
        arquivos_excel = [f for f in os.listdir(pasta) if f.lower().endswith(".xlsx")]
        planilha = None
        for arq in arquivos_excel:
            caminho = os.path.join(pasta, arq)
            try:
                wb = openpyxl.load_workbook(caminho, data_only=True)
                ws = wb.active
                if self.obter_indice_coluna(ws, "ARQUIVO CAPA") != -1 or self.obter_indice_coluna(ws, "ARQUIVO MIOLO") != -1:
                    planilha = caminho
                    wb.close()
                    break
                wb.close()
            except:
                pass
        if not planilha:
            messagebox.showerror("Erro", "Planilha não encontrada.", parent=self)
            return

        wb = openpyxl.load_workbook(planilha, data_only=True)
        ws = wb.active
        self.idx_capa = self.obter_indice_coluna(ws, "ARQUIVO CAPA")
        self.idx_miolo = self.obter_indice_coluna(ws, "ARQUIVO MIOLO")
        self.idx_caminho = self.obter_indice_coluna(ws, "CAMINHO DO ARQUIVO")
        wb.close()

        self.lista_completa = self.obter_lista_arquivos_da_planilha(planilha, cliente)
        filtro = self.filtro_var.get().strip().lower()
        if filtro:
            filtros = [x.strip() for x in filtro.split(",")]
            self.lista_completa = [item for item in self.lista_completa if any(f in item["nome"].lower() for f in filtros)]

        for i in self.tree_selecao.get_children():
            self.tree_selecao.delete(i)
        for item in self.lista_completa:
            self.tree_selecao.insert("", "end", values=(item["nome"],))
        self.tree_selecao.column("Nome", width=min(600, self.tree_selecao.winfo_width()))

    def preparar_processamento(self):
        selecionados = self.tree_selecao.selection()
        if not selecionados:
            messagebox.showinfo("Aviso", "Selecione arquivos.", parent=self)
            return
        for item in self.tree_edicao.get_children():
            self.tree_edicao.delete(item)
        cliente = self.pasta_var.get()
        coluna_codigo = 8 if cliente == "uiclap" else 3
        for iid in selecionados:
            nome = self.tree_selecao.item(iid, "values")[0]
            item = next((x for x in self.lista_completa if x["nome"] == nome), None)
            if not item:
                continue
            linha = item["linha"]
            codigo = "N/A"
            if coluna_codigo < len(linha):
                codigo = linha[coluna_codigo]
            self.tree_edicao.insert("", "end", values=(nome, codigo, "1", self.SETORES_PERDA[0], self.MAQUINAS_REF[0]))

    def limpar(self):
        self._limpar_tudo_exceto_cliente_lote()
        self.filtro_var.set("")
        self.operador_var.set(self.OPERADORES[0])

    def aplicar_nos_selecionados(self):
        itens = self.tree_edicao.selection()
        if not itens:
            messagebox.showinfo("Aviso", "Selecione linhas na tabela de edição.", parent=self)
            return
        setor = self.cb_setor.get()
        maquina = self.cb_maquina.get()
        for item in itens:
            valores = list(self.tree_edicao.item(item, "values"))
            valores[3] = setor
            valores[4] = maquina
            self.tree_edicao.item(item, values=valores)

    def aplicar_em_todos(self):
        setor = self.cb_setor.get()
        maquina = self.cb_maquina.get()
        for item in self.tree_edicao.get_children():
            valores = list(self.tree_edicao.item(item, "values"))
            valores[3] = setor
            valores[4] = maquina
            self.tree_edicao.item(item, values=valores)

    def edicao_inline(self, event):
        region = self.tree_edicao.identify_region(event.x, event.y)
        if region != "cell":
            return
        item = self.tree_edicao.identify_row(event.y)
        coluna = self.tree_edicao.identify_column(event.x)
        if not item:
            return
        col = int(coluna.replace("#", ""))
        if col not in [3, 4, 5]:
            return
        valores = list(self.tree_edicao.item(item, "values"))
        bbox = self.tree_edicao.bbox(item, coluna)
        if not bbox:
            return
        x, y, w, h = bbox
        for child in self.tree_edicao.winfo_children():
            if isinstance(child, (tk.Entry, ttk.Combobox)):
                child.destroy()
        if col == 3:
            entry = tk.Entry(self.tree_edicao, justify="center", font=("Arial", 10, "bold"))
            entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, valores[2])
            entry.focus()
            def salvar(event=None):
                novo = entry.get().strip()
                if not novo:
                    novo = "1"
                valores[2] = novo
                self.tree_edicao.item(item, values=valores)
                entry.destroy()
            entry.bind("<Return>", salvar)
            entry.bind("<FocusOut>", salvar)
        elif col == 4:
            cb = ttk.Combobox(self.tree_edicao, values=self.SETORES_PERDA, state="readonly")
            cb.place(x=x, y=y, width=w, height=h)
            cb.set(valores[3])
            cb.focus()
            def salvar(event=None):
                valores[3] = cb.get()
                self.tree_edicao.item(item, values=valores)
                cb.destroy()
            cb.bind("<<ComboboxSelected>>", salvar)
            cb.bind("<FocusOut>", salvar)
        elif col == 5:
            cb = ttk.Combobox(self.tree_edicao, values=self.MAQUINAS_REF, state="readonly")
            cb.place(x=x, y=y, width=w, height=h)
            cb.set(valores[4])
            cb.focus()
            def salvar(event=None):
                valores[4] = cb.get()
                self.tree_edicao.item(item, values=valores)
                cb.destroy()
            cb.bind("<<ComboboxSelected>>", salvar)
            cb.bind("<FocusOut>", salvar)

    def executar(self):
        cliente = self.pasta_var.get()
        lote = self.subpasta_var.get()
        operador = self.operador_var.get()
        if not cliente or not lote:
            return
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        if cliente == "uiclap":
            pasta_lote = os.path.join(self.CAMINHO_UICLAP_LOTES, lote)
        else:
            pasta_lote = os.path.join(self.CAMINHO_CAPAS_POD_AGRUPADAS, cliente, lote)

        try:
            wb = openpyxl.load_workbook(self.CAMINHO_RELATORIOS_PERDA)
            aba = "UICLAP" if cliente == "uiclap" else "PODS"
            ws = wb[aba]
        except Exception as e:
            messagebox.showerror("Erro", str(e), parent=self)
            return

        copias_capas = []
        miolos_info = []
        processados = 0

        total_itens = len(self.tree_edicao.get_children())
        self.progress["maximum"] = total_itens
        self.progress["value"] = 0
        self.progress_label.config(text=f"Copiando: 0/{total_itens}")
        self.update_idletasks()

        for idx, iid in enumerate(self.tree_edicao.get_children(), start=1):
            valores = self.tree_edicao.item(iid, "values")
            nome = valores[0]
            qtd = valores[2]
            setor = valores[3]
            maquina = valores[4]

            item_original = next((x for x in self.lista_completa if x["nome"] == nome), None)
            if not item_original:
                continue

            tipo_atual = item_original["tipo"]
            linha = item_original["linha"]
            caminho_fisico = item_original["caminho_fisico"]
            if not caminho_fisico or not os.path.exists(caminho_fisico):
                caminho_fisico = os.path.join(pasta_lote, nome)

            destino = self.copiar_para_refazer(caminho_fisico, cliente, lote, tipo_atual)

            if tipo_atual == "CAPA" and destino and destino.lower().endswith('.pdf'):
                copias_capas.append((destino, os.path.splitext(nome)[0], setor))
            elif tipo_atual == "MIOLO" and destino and destino.lower().endswith('.pdf'):
                miolos_info.append((
                    destino,
                    os.path.splitext(nome)[0],
                    item_original.get("acabamento"),
                    item_original.get("formato_miolo"),
                    setor
                ))

            # Copia o outro tipo (miolo/capa) se existir
            tipo_outro = "MIOLO" if tipo_atual == "CAPA" else "CAPA"
            idx_nome_outro = self.idx_miolo if tipo_atual == "CAPA" else self.idx_capa
            if idx_nome_outro != -1 and idx_nome_outro < len(linha):
                nome_outro = str(linha[idx_nome_outro]).strip() if linha[idx_nome_outro] else None
                if nome_outro:
                    caminho_outro = None
                    if self.idx_caminho != -1 and self.idx_caminho < len(linha) and linha[self.idx_caminho]:
                        caminho_outro = os.path.join(str(linha[self.idx_caminho]).strip(), nome_outro)
                    else:
                        caminho_outro = os.path.join(pasta_lote, nome_outro)
                    self.copiar_para_refazer(caminho_outro, cliente, lote, tipo_outro)

            nova_linha = linha + (qtd, setor, maquina, operador, data_hora, tipo_atual)
            ws.append(nova_linha)
            processados += 1

            self.progress["value"] = idx
            self.progress_label.config(text=f"Copiando: {idx}/{total_itens}")
            self.update_idletasks()

        wb.save(self.CAMINHO_RELATORIOS_PERDA)
        wb.close()

        # Etapa 2: marca d'água nas capas
        total_capas = len(copias_capas)
        if total_capas > 0:
            self.progress["maximum"] = total_capas
            self.progress["value"] = 0
            self.progress_label.config(text=f"Capas: 0/{total_capas}")
            self.update_idletasks()
            for idx, (destino, nome_base, setor_capa) in enumerate(copias_capas, start=1):
                try:
                    self.adicionar_marca_dagua(destino, nome_base)
                    if setor_capa == "PROBLEMA DE FONTE":
                        temp_outlines = destino.replace('.pdf', '_TEMP_OUTLINES.pdf')
                        self.converter_em_contornos(destino, temp_outlines)
                        os.replace(temp_outlines, destino)
                except Exception as e:
                    print(f"Erro na marca d'água/contornos de {nome_base}: {e}")
                self.progress["value"] = idx
                self.progress_label.config(text=f"Capas: {idx}/{total_capas}")
                self.update_idletasks()

        # Etapa 3: processar miolos
        total_miolos = len(miolos_info)
        if total_miolos > 0:
            self.progress["maximum"] = total_miolos
            self.progress["value"] = 0
            self.progress_label.config(text=f"Miolos: 0/{total_miolos}")
            self.update_idletasks()

            for idx, (copia_dest, nome_base, acab, fmt_miolo, setor_item) in enumerate(miolos_info, start=1):
                self.progress_label.config(text=f"Miolo {idx}/{total_miolos}: {nome_base}")
                self.update_idletasks()
                try:
                    tipo_montagem = None
                    acab_upper = acab.upper() if acab else ""
                    if "BROCHURA" in acab_upper:
                        tipo_montagem = "brochura"

                    elif "GRAMPO" in acab_upper:
                        tipo_montagem = "grampo"

                    elif "GRAMPEADO" in acab_upper:
                        tipo_montagem = "grampo"
                    
                    elif "WIREO" in acab_upper:
                        tipo_montagem = "brochura"

                    elif "CADERNO" in acab_upper:
                        tipo_montagem = "brochura"    

                    if tipo_montagem:
                        if setor_item == "PRE-IMPRESSÃO":
                            nome_final = copia_dest.replace('.pdf', '_MONTADO_REDUZIDO.pdf')
                            reduzido_temp = copia_dest.replace('.pdf', '_TEMP_RED.pdf')
                            self.comprimir_pdf(copia_dest, reduzido_temp)
                            self.montar_pdf(reduzido_temp, nome_final, tipo_montagem)
                            if os.path.exists(reduzido_temp): os.remove(reduzido_temp)
                        elif setor_item == "PROBLEMA DE FONTE":
                            nome_final = copia_dest.replace('.pdf', '_MONTADO_CURVAS.pdf')
                            montado_temp = copia_dest.replace('.pdf', '_TEMP_MONT.pdf')
                            self.montar_pdf(copia_dest, montado_temp, tipo_montagem)
                            self.converter_em_contornos(montado_temp, nome_final)
                            if os.path.exists(montado_temp): os.remove(montado_temp)
                        else:
                            nome_final = copia_dest.replace('.pdf', '_MONTADO.pdf')
                            self.montar_pdf(copia_dest, nome_final, tipo_montagem)

                        if os.path.exists(copia_dest):
                            os.remove(copia_dest)
                    else:
                        if setor_item == "PROBLEMA DE FONTE":
                            nome_final = copia_dest.replace('.pdf', '_CURVAS.pdf')
                            self.converter_em_contornos(copia_dest, nome_final)
                            if os.path.exists(copia_dest): os.remove(copia_dest)
                        print(f"Miolo {nome_base}: sem tipo de montagem definido (acabamento: {acab}), mantido como cópia simples.")
                except Exception as e:
                    print(f"Erro ao processar miolo {nome_base}: {e}")

                self.progress["value"] = idx
                self.update_idletasks()

        # Limpeza final da interface
        for item in self.tree_edicao.get_children():
            self.tree_edicao.delete(item)
        for item in self.tree_selecao.get_children():
            self.tree_selecao.delete(item)
        self.lista_completa = []
        self.progress["value"] = 0
        self.progress_label.config(text="")
        messagebox.showinfo("Concluído", f"{processados} arquivos processados.", parent=self)

import os
import csv
import sqlite3
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

def gprint_material_historico(pilha, arquivo, modo):
    """O TXT de origem tem prioridade sobre nomes de arquivos e modos."""
    origem = str(pilha or '').upper()
    if 'UICLAP_MIOLO_' in origem:
        return 'Miolo'
    if 'UICLAP_CAPA_' in origem:
        return 'Capa'
    if modo in ('capa_marcacao', 'mesclagem', 'envio'):
        return 'Capa'
    arquivo = str(arquivo or '').upper()
    if '-MIOLO-' in arquivo or 'MIOLO_' in arquivo:
        return 'Miolo'
    if '-CAPA-' in arquivo or 'CAPA_' in arquivo:
        return 'Capa'
    return 'Outro'


def gprint_resumir_lote(caminho):
    """Exibe a pasta do lote, sem perder o caminho completo no banco."""
    componentes = str(caminho or '').replace('\\', '/').split('/')
    for parte in componentes:
        if re.match(r'^\d{3,}[-_]', parte):
            return parte
    return componentes[-1] if componentes else ''


def gprint_resumir_pilha(nome):
    texto = str(nome or '')
    texto = re.sub(r'^UICLAP_(?:CAPA|MIOLO)_', '', texto, flags=re.I)
    texto = re.sub(r'_\d+ ITENS?_\d+ TOTAL$', '', texto, flags=re.I)
    return texto


def gprint_nome_equipamento(maquina):
    return str(maquina or '').split(' - ')[0]


def gprint_descricao_modo(modo):
    nomes = {
        'capa_marcacao': 'Marcação de capa',
        'mesclagem': 'Mesclagem de capas',
        'envio': 'Envio do PDF mesclado',
        'corta_empilha_rotacionado': 'Corta/empilha rotacionado',
        'Corta_empilha': 'Corta e empilha',
        'grampo': 'Grampo',
        'dois_na_folha': 'Dois na folha',
        'dois_na_folha_rotacionado': 'Dois na folha rotacionado',
        'um_na_folha_atexco': 'Uma na folha ATEXCO',
        'um_na_folha': 'Uma na folha',
    }
    return nomes.get(modo, str(modo or ''))


def gprint_situacao_historico(status, material, maquina, copia):
    """Texto didatico de exibicao; nao muda o status salvo no SQLite."""
    destino = gprint_nome_equipamento(maquina)
    rotulos = {
        'Montagem OK': 'Capa preparada' if material == 'Capa' else 'PDF de miolo preparado',
        'Mesclagem OK': 'Capas mescladas em PDF final',
        'Enviado': f'Enviado à {destino}',
        'Falha no envio': f'Falha no envio à {destino}',
        'Montagem Erro': 'Erro na montagem',
        'Não enviado (manual)': 'PDF salvo, sem envio LPR',
        'Salvo': 'PDF salvo localmente',
        'Falha ao salvar': 'Falha ao salvar PDF',
    }
    return rotulos.get(str(status or ''), str(status or ''))


def gprint_resumo_historico_agrupado(grupos):
    """Resumo de operações obtido por SQL agrupado, sem LIMIT da tabela visual.

    grupos: (tipo, status, modo, total_eventos, soma_quantidade).
    - Capa preparada: soma quantidades, pois registros antigos eram por copia
      (qtd=1) e os novos sao por arquivo com a quantidade total.
    - Capa mesclada/enviada: soma quantidades dos eventos correspondentes.
    - Miolo: conta eventos, porque um envio LPR pode conter mais de um livro.
    """
    r = {
        'registros': 0,
        'capa_preparadas': 0,
        'capa_mescladas': 0, 'capa_pdfs_mesclados': 0,
        'capa_pdfs_enviados': 0, 'capa_capas_enviadas': 0,
        'capa_falhas_lpr': 0,
        'miolo_pdfs_montados': 0, 'miolo_envios_lpr': 0,
        'miolo_falhas_lpr': 0, 'miolo_pdfs_salvos': 0,
        'outros': 0,
    }
    for material, status, modo, eventos, quantidade in grupos:
        eventos = int(eventos or 0)
        quantidade = int(quantidade or 0)
        status, modo = str(status or ''), str(modo or '')
        r['registros'] += eventos
        if material == 'Capa':
            if status == 'Montagem OK':
                r['capa_preparadas'] += quantidade
            elif status == 'Mesclagem OK':
                r['capa_mescladas'] += quantidade
                r['capa_pdfs_mesclados'] += eventos
            elif status == 'Enviado':
                r['capa_pdfs_enviados'] += eventos
                if modo == 'envio':
                    r['capa_capas_enviadas'] += quantidade
            elif status == 'Falha no envio':
                r['capa_falhas_lpr'] += eventos
        elif material == 'Miolo':
            if status == 'Montagem OK':
                r['miolo_pdfs_montados'] += eventos
            elif status == 'Enviado':
                r['miolo_envios_lpr'] += eventos
            elif status == 'Falha no envio':
                r['miolo_falhas_lpr'] += eventos
            elif status == 'Salvo':
                r['miolo_pdfs_salvos'] += eventos
        else:
            r['outros'] += eventos
    return r


def gprint_tipo_historico_sql():
    """Classificacao no SQL, compativel com a exibicao de eventos antigos."""
    return ("CASE "
            "WHEN instr(upper(coalesce(pilha,'')), 'UICLAP_MIOLO_') > 0 THEN 'Miolo' "
            "WHEN instr(upper(coalesce(pilha,'')), 'UICLAP_CAPA_') > 0 THEN 'Capa' "
            "WHEN modo IN ('capa_marcacao','mesclagem','envio') THEN 'Capa' "
            "WHEN instr(upper(coalesce(arquivo,'')), '-MIOLO-') > 0 "
            "OR instr(upper(coalesce(arquivo,'')), 'MIOLO_') > 0 THEN 'Miolo' "
            "WHEN instr(upper(coalesce(arquivo,'')), '-CAPA-') > 0 "
            "OR instr(upper(coalesce(arquivo,'')), 'CAPA_') > 0 THEN 'Capa' "
            "ELSE 'Outro' END")


class HistoricoTab(ttk.Frame):
    """Busca explicitamente acionada pelo operador; nao consulta DB na abertura."""
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.dados = []
        self.dados_brutos = {}
        self.ordem_atual = None
        self._busca_em_andamento = False
        self._usuarios_carregados = False
        self._resultado_busca = queue.Queue()
        self.setup_ui()
        # Sem self.after(..., buscar): nao acessar rede apenas por abrir a aba.

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        filtro_frame = ttk.LabelFrame(main_frame, text="🔍 Filtros de Histórico", bootstyle="primary")
        filtro_frame.pack(fill=tk.X, pady=(0, 10))

        row1 = ttk.Frame(filtro_frame)
        row1.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(row1, text="De:").pack(side=tk.LEFT, padx=(0, 2))
        self.data_inicio = ttk.Entry(row1, width=12)
        self.data_inicio.pack(side=tk.LEFT, padx=2)
        self.data_inicio.insert(0, (datetime.now() - timedelta(days=30)).strftime("%d/%m/%Y"))
        ttk.Label(row1, text="Até:").pack(side=tk.LEFT, padx=(10, 2))
        self.data_fim = ttk.Entry(row1, width=12)
        self.data_fim.pack(side=tk.LEFT, padx=2)
        self.data_fim.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.usar_data = tk.BooleanVar(value=True)
        ttk.Checkbutton(row1, text="Filtrar por data", variable=self.usar_data).pack(side=tk.LEFT, padx=20)

        row2 = ttk.Frame(filtro_frame)
        row2.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(row2, text="Lote:").pack(side=tk.LEFT)
        self.lote_entry = ttk.Entry(row2, width=26)
        self.lote_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="Pilha:").pack(side=tk.LEFT, padx=(12, 0))
        self.pilha_entry = ttk.Entry(row2, width=25)
        self.pilha_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="Arquivo:").pack(side=tk.LEFT, padx=(12, 0))
        self.arquivo_entry = ttk.Entry(row2, width=25)
        self.arquivo_entry.pack(side=tk.LEFT, padx=5)

        row3 = ttk.Frame(filtro_frame)
        row3.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(row3, text="Tipo:").pack(side=tk.LEFT)
        self.tipo_combo = ttk.Combobox(row3, values=["Todos", "Capa", "Miolo"], width=10, state="readonly")
        self.tipo_combo.pack(side=tk.LEFT, padx=5)
        self.tipo_combo.current(0)
        ttk.Label(row3, text="Operação:").pack(side=tk.LEFT, padx=(12, 0))
        self.operacao_combo = ttk.Combobox(
            row3, values=["Todas", "Montagem", "Mesclagem", "Envio LPR", "Falha LPR", "Salvo sem envio"],
            width=18, state="readonly")
        self.operacao_combo.pack(side=tk.LEFT, padx=5)
        self.operacao_combo.current(0)
        ttk.Label(row3, text="Máquina:").pack(side=tk.LEFT, padx=(12, 0))
        maquinas = ["Todas"]
        self.maquina_combo = ttk.Combobox(row3, values=maquinas, width=27, state="readonly")
        self.maquina_combo.pack(side=tk.LEFT, padx=5)
        self.maquina_combo.current(0)

        row4 = ttk.Frame(filtro_frame)
        row4.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(row4, text="Usuário:").pack(side=tk.LEFT)
        self.usuario_combo = ttk.Combobox(row4, values=["Todos"], width=17, state="readonly")
        self.usuario_combo.pack(side=tk.LEFT, padx=5)
        self.usuario_combo.current(0)
        ttk.Label(row4, text="Status:").pack(side=tk.LEFT, padx=(12, 0))
        self.status_combo = ttk.Combobox(
            row4, values=["", "Montagem OK", "Mesclagem OK", "Enviado", "Falha no envio",
                          "Não enviado (manual)", "Montagem Erro", "Salvo", "Falha ao salvar"],
            width=22, state="readonly")
        self.status_combo.pack(side=tk.LEFT, padx=5)
        self.status_combo.current(0)
        ttk.Label(row4, text="Modo:").pack(side=tk.LEFT, padx=(12, 0))
        modos = ["", "mesclagem", "capa_marcacao", "envio", "dois_na_folha",
                 "dois_na_folha_rotacionado", "grampo", "Corta_empilha",
                 "corta_empilha_rotacionado", "um_na_folha", "um_na_folha_atexco"]
        self.modo_combo = ttk.Combobox(row4, values=modos, width=27, state="readonly")
        self.modo_combo.pack(side=tk.LEFT, padx=5)
        self.modo_combo.current(0)

        btn_frame = ttk.Frame(filtro_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        self.btn_buscar = ttk.Button(btn_frame, text="🔎 Buscar", command=self.buscar, bootstyle="primary")
        self.btn_buscar.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🧹 Limpar Filtros", command=self.limpar_filtros,
                   bootstyle="secondary").pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_frame, text="Mostrar:").pack(side=tk.LEFT, padx=(18, 3))
        self.limite_combo = ttk.Combobox(btn_frame, values=["100", "500", "1000", "3000"],
                                        width=6, state="readonly")
        self.limite_combo.pack(side=tk.LEFT, padx=5)
        self.limite_combo.set("500")
        self.lbl_estado = ttk.Label(btn_frame, text="Defina os filtros e clique Buscar. O banco não é carregado ao abrir.")
        self.lbl_estado.pack(side=tk.LEFT, padx=12)

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        colunas = ("Data/Hora", "Tipo", "Lote", "Pilha", "Arquivo", "Máquina", "Formato", "Operação",
                   "Qtd", "Cópia", "Situação", "Erro", "Usuário")
        self.tree = ttk.Treeview(tree_frame, columns=colunas, show="headings", height=14)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        larguras = [145, 65, 175, 190, 260, 115, 115, 165, 50, 60, 195, 140, 100]
        for col, larg in zip(colunas, larguras):
            self.tree.heading(col, text=col, command=lambda c=col: self.ordenar_por(c))
            self.tree.column(col, width=larg, minwidth=50)
        self.tree.bind('<Double-1>', self.mostrar_detalhes)

        resumo_frame = ttk.LabelFrame(main_frame, text="📊 Resumo do lote / operação pesquisada", bootstyle="secondary")
        resumo_frame.pack(fill=tk.X, pady=(5, 0))
        self.lbl_total_registros = ttk.Label(resumo_frame, text="Nenhuma consulta realizada",
                                              font=("Segoe UI", 9, "bold"))
        self.lbl_total_registros.pack(anchor="w", padx=10, pady=(4, 0))
        self.lbl_capa = ttk.Label(resumo_frame, text="CAPA | Preparadas: — | Mescladas: — | PDFs enviados: — | Capas nos PDFs enviados: —",
                                   font=("Segoe UI", 9, "bold"))
        self.lbl_capa.configure(wraplength=1120, justify=tk.LEFT)
        self.lbl_capa.pack(anchor="w", padx=10, pady=2)
        self.lbl_miolo = ttk.Label(resumo_frame, text="MIOLO | PDFs montados: — | Envios LPR: — | Falhas: —",
                                    font=("Segoe UI", 9, "bold"))
        self.lbl_miolo.configure(wraplength=1120, justify=tk.LEFT)
        self.lbl_miolo.pack(anchor="w", padx=10, pady=(0, 4))
        self.lbl_nota = ttk.Label(resumo_frame,
                                  text="'Enviado' = LPR aceitou o trabalho; não confirma impressão física. Totais não misturam montagem e envio.",
                                  font=("Segoe UI", 8))
        self.lbl_nota.pack(anchor="w", padx=10, pady=(0, 4))

    def _montar_filtro(self):
        """Cria cláusula WHERE/parametros a partir dos controles, sem acessar o BD."""
        filtros, params = [], []
        if self.usar_data.get():
            try:
                di = datetime.strptime(self.data_inicio.get().strip(), "%d/%m/%Y")
                df = datetime.strptime(self.data_fim.get().strip(), "%d/%m/%Y")
                if df < di:
                    raise ValueError("Data final anterior à inicial")
            except ValueError as exc:
                raise ValueError("Informe datas válidas (dd/mm/aaaa), com início anterior ao fim.") from exc
            filtros += ["data_hora >= ?", "data_hora < ?"]
            params += [di.strftime("%Y-%m-%d"), (df + timedelta(days=1)).strftime("%Y-%m-%d")]
        for campo, widget in (("lote", self.lote_entry), ("pilha", self.pilha_entry),
                              ("arquivo", self.arquivo_entry)):
            valor = widget.get().strip()
            if valor:
                # Pesquisa literal: % e _ digitados nao sao wildcards SQL.
                valor = valor.replace('\\', '\\\\').replace('%', '\\%').replace('_', '\\_')
                filtros.append(f"{campo} LIKE ? ESCAPE '\\'")
                params.append(f"%{valor}%")
        tipo = self.tipo_combo.get()
        if tipo in ("Capa", "Miolo"):
            filtros.append(gprint_tipo_historico_sql() + " = ?")
            params.append(tipo)
        etapa = self.operacao_combo.get()
        etapas = {"Montagem": ("Montagem OK", "Montagem Erro"), "Mesclagem": ("Mesclagem OK",),
                  "Envio LPR": ("Enviado",), "Falha LPR": ("Falha no envio",),
                  "Salvo sem envio": ("Salvo", "Não enviado (manual)", "Falha ao salvar")}
        if etapa in etapas:
            choices = etapas[etapa]
            filtros.append("status IN (" + ",".join("?" for _ in choices) + ")")
            params.extend(choices)
        maquina = self.maquina_combo.get()
        if maquina and maquina != "Todas":
            filtros.append("maquina = ?")
            params.append(maquina)
        usuario = self.usuario_combo.get()
        if usuario and usuario != "Todos":
            filtros.append("usuario = ?")
            params.append(usuario)
        status = self.status_combo.get().strip()
        if status:
            filtros.append("status = ?")
            params.append(status)
        modo = self.modo_combo.get().strip()
        if modo:
            filtros.append("modo = ?")
            params.append(modo)
        return (' WHERE ' + ' AND '.join(filtros)) if filtros else '', params

    def buscar(self):
        if self._busca_em_andamento:
            return
        try:
            where, params = self._montar_filtro()
            limite = int(self.limite_combo.get())
        except ValueError as exc:
            messagebox.showerror("Filtros inválidos", str(exc), parent=self)
            return
        self._busca_em_andamento = True
        self.btn_buscar.configure(state='disabled')
        self.lbl_estado.configure(text="Consultando o histórico…")
        # Queries SQLite e checagem de rede fora da thread de Tkinter.
        threading.Thread(target=self._consultar_em_background,
                         args=(where, tuple(params), limite), daemon=True).start()
        self.after(100, self._verificar_resultado)

    def _consultar_em_background(self, where, params, limite):
        conn = None
        try:
            caminho = self.app.log_db_path
            origem = 'central'
            try:
                gprint_checar_banco_central(caminho)
            except (GPrintBancoIndisponivel, OSError):
                caminho = self.app.log_db_local_path
                origem = 'local'
                if not os.path.isfile(caminho):
                    raise FileNotFoundError("O histórico central está indisponível e não existe banco local.")
            # As consultas abaixo sao apenas SELECT e nao alteram o esquema SQL.
            # O caminho nativo funciona tanto com unidade mapeada quanto com UNC.
            if not os.path.isfile(caminho):
                raise FileNotFoundError('O banco de histórico não foi encontrado.')
            conn = sqlite3.connect(caminho, timeout=15)
            cursor = conn.cursor()
            campos = ("data_hora, lote, pilha, arquivo, maquina, ip, formato, modo, "
                      "quantidade, status, erro, usuario, copia")
            cursor.execute("SELECT " + campos + " FROM logs" + where +
                           " ORDER BY data_hora DESC, id DESC LIMIT ?", params + (limite,))
            rows = cursor.fetchall()
            tipo_sql = gprint_tipo_historico_sql()
            cursor.execute("SELECT " + tipo_sql + " AS material, status, modo, "
                           "COUNT(*) AS eventos, COALESCE(SUM(CASE WHEN quantidade > 0 "
                           "THEN quantidade ELSE 0 END),0) AS unidades "
                           "FROM logs" + where + " GROUP BY material, status, modo", params)
            grupos = cursor.fetchall()
            resumo = gprint_resumo_historico_agrupado(grupos)
            usuarios = None
            if not self._usuarios_carregados:
                cursor.execute("SELECT DISTINCT usuario FROM logs WHERE usuario IS NOT NULL "
                               "AND usuario != '' ORDER BY usuario")
                usuarios = [r[0] for r in cursor.fetchall()]
            self._resultado_busca.put((True, rows, resumo, usuarios, origem))
        except Exception as exc:
            self._resultado_busca.put((False, exc))
        finally:
            if conn is not None:
                conn.close()

    def _verificar_resultado(self):
        try:
            resultado = self._resultado_busca.get_nowait()
        except queue.Empty:
            if self.winfo_exists():
                self.after(100, self._verificar_resultado)
            return
        self._busca_em_andamento = False
        self.btn_buscar.configure(state='normal')
        if not resultado[0]:
            erro = resultado[1]
            self.lbl_estado.configure(text="Não foi possível consultar o histórico.")
            if isinstance(erro, FileNotFoundError):
                messagebox.showwarning("GPrint - Histórico indisponível",
                                       "Não foi possível acessar o banco de dados da produção.\n\n"
                                       "Verifique se a pasta 'produção' está online.\n"
                                       "Se o problema continuar, INFORME O GUILHERME.", parent=self)
            else:
                messagebox.showerror("GPrint - Histórico", f"Falha ao consultar histórico: {erro}", parent=self)
            return
        _, rows, resumo, usuarios, origem = resultado
        if origem == 'local':
            messagebox.showwarning("GPrint - Histórico local",
                                   "O histórico central da produção não está acessível.\n"
                                   "Os dados exibidos são somente do banco local desta estação.\n\n"
                                   "Verifique o diretório 'produção' e, se necessário, INFORME O GUILHERME.",
                                   parent=self)
        if usuarios is not None:
            anterior = self.usuario_combo.get() or 'Todos'
            self.usuario_combo.configure(values=['Todos'] + usuarios)
            self.usuario_combo.set(anterior if anterior in ['Todos'] + usuarios else 'Todos')
            self._usuarios_carregados = True
        self.tree.delete(*self.tree.get_children())
        self.dados = []
        self.dados_brutos = {}
        for row in rows:
            material = gprint_material_historico(row[2], row[3], row[7])
            visivel = (
                row[0], material, gprint_resumir_lote(row[1]),
                gprint_resumir_pilha(row[2]), row[3],
                gprint_nome_equipamento(row[4]), row[6], gprint_descricao_modo(row[7]),
                row[8], row[12] or '-',
                gprint_situacao_historico(row[9], material, row[4], row[12]),
                row[10], row[11]
            )
            self.dados.append(visivel)
            iid = self.tree.insert('', tk.END, values=visivel)
            self.dados_brutos[iid] = row
        self.lbl_total_registros.configure(text=f"Exibindo {len(rows)} de {resumo['registros']} registros filtrados")
        self.lbl_capa.configure(text=(
            f"CAPA | Preparadas: {resumo['capa_preparadas']} | "
            f"Mescladas: {resumo['capa_mescladas']} em {resumo['capa_pdfs_mesclados']} PDFs | "
            f"PDFs enviados: {resumo['capa_pdfs_enviados']} | "
            f"Capas nesses envios: {resumo['capa_capas_enviadas']} | "
            f"Falhas LPR: {resumo['capa_falhas_lpr']}"))
        self.lbl_miolo.configure(text=(
            f"MIOLO | PDFs montados: {resumo['miolo_pdfs_montados']} | "
            f"Envios LPR: {resumo['miolo_envios_lpr']} | "
            f"Falhas: {resumo['miolo_falhas_lpr']} | "
            f"Cópias salvas sem LPR: {resumo['miolo_pdfs_salvos']}"))
        self.lbl_estado.configure(text="Consulta concluída." +
                                 (" Histórico LOCAL." if origem == 'local' else " Histórico CENTRAL."))

    def limpar_filtros(self):
        self.data_inicio.delete(0, tk.END)
        self.data_inicio.insert(0, (datetime.now() - timedelta(days=30)).strftime("%d/%m/%Y"))
        self.data_fim.delete(0, tk.END)
        self.data_fim.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.usar_data.set(True)
        for entry in (self.lote_entry, self.pilha_entry, self.arquivo_entry):
            entry.delete(0, tk.END)
        for combo, valor in ((self.tipo_combo, 'Todos'), (self.operacao_combo, 'Todas'),
                             (self.maquina_combo, 'Todas'), (self.usuario_combo, 'Todos'),
                             (self.status_combo, ''), (self.modo_combo, '')):
            combo.set(valor)
        self.limite_combo.set('500')
        self.tree.delete(*self.tree.get_children())
        self.dados, self.dados_brutos = [], {}
        self.lbl_total_registros.configure(text='Nenhuma consulta realizada')
        self.lbl_capa.configure(text='CAPA | Preparadas: — | Mescladas: — | PDFs enviados: — | Capas nos PDFs enviados: —')
        self.lbl_miolo.configure(text='MIOLO | PDFs montados: — | Envios LPR: — | Falhas: —')
        self.lbl_estado.configure(text='Defina os filtros e clique Buscar.')

    def ordenar_por(self, coluna):
        if not self.dados:
            return
        idx_map = {"Data/Hora": 0, "Tipo": 1, "Lote": 2, "Pilha": 3, "Arquivo": 4,
                   "Máquina": 5, "Formato": 6, "Operação": 7, "Qtd": 8,
                   "Cópia": 9, "Situação": 10, "Erro": 11, "Usuário": 12}
        reverse = self.ordem_atual == coluna
        self.ordem_atual = coluna
        itens = list(self.tree.get_children())
        itens.sort(key=lambda iid: self._chave_ordenacao(self.tree.set(iid, coluna), coluna),
                   reverse=reverse)
        for indice, iid in enumerate(itens):
            self.tree.move(iid, '', indice)

    @staticmethod
    def _chave_ordenacao(valor, coluna):
        if coluna == 'Qtd':
            try:
                return int(valor)
            except (ValueError, TypeError):
                return 0
        return str(valor).casefold()

    def mostrar_detalhes(self, _event=None):
        selecao = self.tree.selection()
        if not selecao:
            return
        row = self.dados_brutos.get(selecao[0])
        if row is None:
            return
        titulos = ('Data/hora', 'Diretório do lote', 'Arquivo TXT / Pilha',
                   'Arquivo PDF', 'Máquina', 'IP', 'Fila / Formato', 'Modo',
                   'Quantidade', 'Status', 'Erro', 'Usuário', 'Cópia')
        mensagem = '\n'.join(f'{titulo}: {valor}' for titulo, valor in zip(titulos, row))
        material = gprint_material_historico(row[2], row[3], row[7])
        if material == 'Miolo':
            significado = ('Um PDF montado' if row[9] == 'Montagem OK'
                           else 'Uma tentativa de envio LPR' if row[9] in ('Enviado', 'Falha no envio')
                           else 'Registro desta etapa')
        elif material == 'Capa':
            significado = ('Capas preparadas' if row[9] == 'Montagem OK' else
                           'Capas no PDF final' if row[9] in ('Mesclagem OK', 'Enviado') else
                           'Quantidade registrada pela rotina de capa')
        else:
            significado = 'Quantidade registrada pela rotina original'
        mensagem += f'\n\nSignificado da quantidade: {significado}'
        if row[9] == 'Enviado':
            mensagem += ('\n\nEnviado significa sucesso da rotina LPR, '
                         'não confirma que a impressora terminou fisicamente.')
        messagebox.showinfo('GPrint - Detalhes do registro', mensagem, parent=self)

if __name__ == "__main__":
    
    logging.basicConfig(
        filename='processamento_capas.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    if verificar_data_limite():
        root = TkinterDnD.Tk()
        style = Style(theme=CONFIG.get('tema', 'darkly'))
        app = Application(root)
        root.mainloop()