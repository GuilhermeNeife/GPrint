from builtins import super
import os
import ntpath
from pathlib import Path
import sys
import multiprocessing

if __name__ == "__main__":
    multiprocessing.freeze_support()

from concurrent.futures import ProcessPoolExecutor, TimeoutError as FutureTimeout
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
from datetime import datetime, timedelta, date
from typing import List
import fitz  
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
from contextlib import closing, suppress
from tkinter import simpledialog

# ============================================================================
# GPRINT - CONFIGURACAO CENTRAL DA PRODUCAO
# Ajuste SOMENTE este diretorio para localizar ambos os bancos.
# Nao crie a pasta da rede automaticamente se o servidor estiver offline.
# ============================================================================
GPRINT_DIRETORIO_DADOS = r"\\192.168.200.253\producao\____REPOSITORIO GERAL POD\UICLAP\LOGS"
GPRINT_BANCO_HISTORICO = os.path.join(GPRINT_DIRETORIO_DADOS, "impressoes.db")
GPRINT_BANCO_EQUIPAMENTOS = os.path.join(GPRINT_DIRETORIO_DADOS, "equipamentos.db")
# ============================================================================
# GPRINT - BASE DO REFAZER (perdas e resultados), SEM ALTERAR MONTAGEM OU LPR
# Banco local por estacao para durabilidade + sincronizacao em lote para rede.
# ============================================================================
GPRINT_BANCO_REFAZER = os.path.join(GPRINT_DIRETORIO_DADOS, 'refazer.db')
GPRINT_REFAZER_LOCAL = os.path.join(
    os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'),
    'GPrint', 'refazer_pendente.db'
)

GPRINT_REFAZER_CAMPOS = (
    'id', 'sessao_id', 'criado_em', 'atualizado_em', 'cliente', 'lote',
    'arquivo_original', 'codigo', 'tipo', 'quantidade', 'setor', 'operador',
    'maquina', 'fila_lpr', 'pdf_resultado', 'montagem_status',
    'envio_status', 'erro', 'estacao'
)


def gprint_criar_schema_refazer(conexao, local=False):
    """Schema aditivo: nao modifica equipamentos, filas ou historico de impressao."""
    conexao.execute('''CREATE TABLE IF NOT EXISTS refazer_registros (
        id TEXT PRIMARY KEY,
        sessao_id TEXT NOT NULL,
        criado_em TEXT NOT NULL,
        atualizado_em TEXT NOT NULL,
        cliente TEXT NOT NULL,
        lote TEXT NOT NULL,
        arquivo_original TEXT NOT NULL,
        codigo TEXT,
        tipo TEXT NOT NULL CHECK(tipo IN ('CAPA','MIOLO')),
        quantidade INTEGER NOT NULL DEFAULT 1 CHECK(quantidade >= 0),
        setor TEXT,
        operador TEXT,
        maquina TEXT,
        fila_lpr TEXT,
        pdf_resultado TEXT,
        montagem_status TEXT NOT NULL DEFAULT 'PENDENTE',
        envio_status TEXT NOT NULL DEFAULT 'PENDENTE',
        erro TEXT NOT NULL DEFAULT '',
        estacao TEXT NOT NULL DEFAULT ''
    )''')
    conexao.execute('''CREATE INDEX IF NOT EXISTS idx_refazer_data
        ON refazer_registros(criado_em)''')
    conexao.execute('''CREATE INDEX IF NOT EXISTS idx_refazer_lote
        ON refazer_registros(lote, tipo)''')
    conexao.execute('''CREATE INDEX IF NOT EXISTS idx_refazer_setor
        ON refazer_registros(setor, operador)''')
    if local:
        conexao.execute('''CREATE TABLE IF NOT EXISTS refazer_sincronizacao (
            registro_id TEXT PRIMARY KEY REFERENCES refazer_registros(id) ON DELETE CASCADE,
            pendente INTEGER NOT NULL DEFAULT 1
        )''')
        conexao.execute('''CREATE INDEX IF NOT EXISTS idx_refazer_pendente
            ON refazer_sincronizacao(pendente)''')


class GPrintRefazerRepositorio:
    """Grava primeiro no computador; envia em lote ao servidor, sem repetir LPR."""
    def __init__(self, central=GPRINT_BANCO_REFAZER, local=GPRINT_REFAZER_LOCAL):
        self.central = central
        self.local = local
        self._mutex = threading.RLock()
        self._preparado = False

    def _conectar_local(self):
        os.makedirs(os.path.dirname(self.local), exist_ok=True)
        con = sqlite3.connect(self.local, timeout=15)
        con.row_factory = sqlite3.Row
        con.execute('PRAGMA busy_timeout=15000')
        con.execute('PRAGMA foreign_keys=ON')
        return con

    def preparar_local(self):
        with self._mutex:
            if not self._preparado:
                with self._conectar_local() as con:
                    gprint_criar_schema_refazer(con, local=True)
                self._preparado = True

    def registrar(self, **dados):
        """INSERT/UPDATE duravel; nao acessa a rede durante a montagem."""
        self.preparar_local()
        dados.setdefault('id', str(uuid.uuid4()))
        agora = datetime.now().isoformat(timespec='microseconds')
        colunas = GPRINT_REFAZER_CAMPOS
        placeholders = ', '.join('?' for _ in colunas)
        updates = ', '.join(f'{col}=excluded.{col}' for col in colunas if col != 'id')
        sql = (f"INSERT INTO refazer_registros ({', '.join(colunas)}) "
               f"VALUES ({placeholders}) ON CONFLICT(id) DO UPDATE SET {updates}")
        with self._mutex:
            with self._conectar_local() as con:
                antigo = con.execute('SELECT * FROM refazer_registros WHERE id=?',
                                     (dados['id'],)).fetchone()
                if antigo is not None:
                    completo = dict(antigo)
                else:
                    completo = {
                        'sessao_id': str(uuid.uuid4()), 'criado_em': agora,
                        'cliente': '', 'lote': '', 'arquivo_original': '',
                        'tipo': 'MIOLO', 'quantidade': 1,
                        'montagem_status': 'PENDENTE', 'envio_status': 'PENDENTE',
                        'erro': '', 'estacao': socket.gethostname()
                    }
                completo.update(dados)
                completo['atualizado_em'] = agora
                con.execute(sql, tuple(completo.get(col) for col in colunas))
                con.execute('''INSERT INTO refazer_sincronizacao(registro_id, pendente)
                    VALUES(?, 1) ON CONFLICT(registro_id) DO UPDATE SET pendente=1''',
                    (dados['id'],))
        return dados['id']

    def sincronizar(self, tamanho_lote=200):
        """Idempotente por UUID; DB central nunca recebe dois registros iguais."""
        self.preparar_local()
        with self._mutex:
            if not os.path.isdir(os.path.dirname(self.central)):
                return 0, 'Pasta de produção indisponível'
            total = 0
            try:
                with self._conectar_local() as local:
                    while True:
                        registros = local.execute('''SELECT r.* FROM refazer_registros r
                            JOIN refazer_sincronizacao s ON s.registro_id=r.id
                            WHERE s.pendente=1 ORDER BY r.criado_em LIMIT ?''',
                            (tamanho_lote,)).fetchall()
                        if not registros:
                            return total, ''
                        colunas = GPRINT_REFAZER_CAMPOS
                        placeholders = ', '.join('?' for _ in colunas)
                        updates = ', '.join(
                            f'{c}=excluded.{c}' for c in colunas if c != 'id')
                        sql = (f"INSERT INTO refazer_registros ({', '.join(colunas)}) "
                               f"VALUES ({placeholders}) ON CONFLICT(id) DO UPDATE SET "
                               f"{updates} WHERE excluded.atualizado_em >= refazer_registros.atualizado_em")
                        with closing(sqlite3.connect(self.central, timeout=20)) as remoto:
                            remoto.execute('PRAGMA busy_timeout=20000')
                            # Usar rollback journal; nao ativar WAL em arquivo compartilhado.
                            gprint_criar_schema_refazer(remoto, local=False)
                            with remoto:
                                remoto.executemany(sql, [tuple(r[col] for col in colunas)
                                                        for r in registros])
                        # Marcacao separada: reenvio posterior e seguro se o processo cair aqui.
                        with local:
                            local.executemany('''UPDATE refazer_sincronizacao
                                SET pendente=0 WHERE registro_id=?''',
                                [(r['id'],) for r in registros])
                        total += len(registros)
            except (OSError, sqlite3.Error) as exc:
                logging.warning('Refazer: banco central indisponivel para sincronizacao: %s', exc)
                return total, 'Não foi possível atualizar o banco central; registros preservados localmente.'

    def pendentes(self):
        self.preparar_local()
        with self._conectar_local() as con:
            return con.execute('SELECT COUNT(*) FROM refazer_sincronizacao WHERE pendente=1').fetchone()[0]

    def resumo_periodo(self, dias=30):
        """Soma perdas por setor, sem carregar todos os registros na tela."""
        if not os.path.isfile(self.central):
            raise FileNotFoundError('Banco central de retrabalhos indisponivel.')
        inicio = (datetime.now() - timedelta(days=dias)).isoformat()
        with sqlite3.connect(self.central, timeout=10) as con:
            return con.execute('''SELECT setor, COUNT(*), SUM(quantidade),
                SUM(CASE WHEN envio_status='ENVIADO' THEN 1 ELSE 0 END),
                SUM(CASE WHEN envio_status='FALHOU' THEN 1 ELSE 0 END)
                FROM refazer_registros WHERE criado_em >= ?
                GROUP BY setor ORDER BY SUM(quantidade) DESC''', (inicio,)).fetchall()


    def opcoes_historicas_perdas(self):
        if not os.path.isfile(self.central):
            return [], []
        with closing(sqlite3.connect(self.central, timeout=10)) as conexao:
            if not conexao.execute(
                "SELECT 1 FROM sqlite_schema WHERE name='refazer_registros' AND type='table'"
            ).fetchone():
                return [], []
            return tuple(
                [r[0] for r in conexao.execute(
                    f"SELECT DISTINCT {coluna} FROM refazer_registros WHERE TRIM({coluna}) <> ''")]
                for coluna in ('setor', 'operador')
            )

    def resumo_entre_datas(self, data_inicio, data_fim, setor='', operador='', tipo='Todos'):
        """Resumo central por periodo inclusivo, consultado APENAS ao clicar em Buscar.

        Recebe objetos date. Limites sao timestamp ISO em intervalos [inicio, fim+1 dia).
        Mantem os registros antigos e o banco atual sem migrações.
        """
        if not isinstance(data_inicio, date) or not isinstance(data_fim, date):
            raise ValueError('Informe datas validas.')
        if data_inicio > data_fim:
            raise ValueError('A data inicial deve ser anterior ou igual a data final.')
        if not os.path.isfile(self.central):
            raise FileNotFoundError('Banco central de retrabalhos indisponivel.')
        inicio = data_inicio.isoformat() + 'T00:00:00'
        fim_exclusivo = (data_fim + timedelta(days=1)).isoformat() + 'T00:00:00'
        filtros = ['criado_em >= ?', 'criado_em < ?']
        parametros = [inicio, fim_exclusivo]
        if setor.strip():
            filtros.append('TRIM(setor) = ? COLLATE NOCASE')
            parametros.append(setor.strip())
        if operador.strip():
            filtros.append('TRIM(operador) = ? COLLATE NOCASE')
            parametros.append(operador.strip())
        if tipo in ('CAPA', 'MIOLO'):
            filtros.append('tipo = ?')
            parametros.append(tipo)
        where = ' AND '.join(filtros)
        consulta = f"""SELECT setor, COUNT(*) AS arquivos,
            COALESCE(SUM(quantidade), 0) AS unidades,
            COALESCE(SUM(CASE WHEN envio_status='ENVIADO' THEN 1 ELSE 0 END), 0) AS enviados,
            COALESCE(SUM(CASE WHEN envio_status='FALHOU' THEN 1 ELSE 0 END), 0) AS falhas,
            COALESCE(SUM(CASE WHEN montagem_status IN ('MONTADO', 'PREPARADO') THEN 1 ELSE 0 END), 0) AS montados
            FROM refazer_registros WHERE {where}
            GROUP BY setor ORDER BY unidades DESC, setor COLLATE NOCASE"""
        with sqlite3.connect(self.central, timeout=10) as conexao:
            return conexao.execute(consulta, parametros).fetchall()


# Senha de acesso a Configuracoes > Equipamentos.
# ALTERE antes de compilar o EXE. Nunca use a senha real em um GitHub publico.
# Protege a interface do cadastro; nao substitui permissoes da pasta de rede.
GPRINT_SENHA_ADMIN_EQUIPAMENTOS = "ADMINISTRADOR123"


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


def gprint_criar_schema_impressoes(conexao):
    conexao.execute(_GPRINT_CRIAR_LOGS)
    conexao.execute("""
        CREATE TABLE IF NOT EXISTS gprint_ids_recebidos (
            evento_id TEXT PRIMARY KEY,
            data_hora TEXT NOT NULL
        )
    """)


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
            gprint_criar_schema_impressoes(conexao)
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
MAX_ARQUIVOS_TXT_MIOLO = 4
MAX_MONTAGENS_SIMULTANEAS_MIOLO = 61  # Limite do ProcessPoolExecutor no Windows.

TIMEOUT_LPR_MIOLO_SEGUNDOS = 300

DATA_LIMITE = "29/07/2027"
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
    "sangria_padrao_miolo": 3,
    "montagens_simultaneas_miolo": 1,
}
def obter_montagens_simultaneas_miolo():
    try:
        valor = int(str(CONFIG.get("montagens_simultaneas_miolo", 1)))
        return valor if 1 <= valor <= MAX_MONTAGENS_SIMULTANEAS_MIOLO else 1
    except ValueError:
        return 1


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
GPRINT_MODOS_PASTAS_REFAZER = {
    'clientes': 'Pastas de clientes',
    'planilha': 'Lotes com planilha',
    'pdf': 'Lotes com PDFs',
    'uiclap_legado': 'Planilha UICLAP antiga',
}
GPRINT_PASTAS_REFAZER_INICIAIS = (
    (0, 'POD', r"\\192.168.200.253\producao\____REPOSITORIO GERAL POD\_CAPAS DE POD AGRUPADAS", 'clientes'),
    (1, 'uiclap', r"\\Sp-dek-srvap-07\api_docuprint", 'pdf'),
    (2, 'uiclap antigo', r"\\192.168.200.253\producao\____REPOSITORIO GERAL POD\UICLAP\LOTES", 'uiclap_legado'),
    (3, '', '', 'clientes'),
    (4, '', '', 'clientes'),
    (5, '', '', 'clientes'),
)


GPRINT_TRATAMENTOS_SETORES = {
    'padrao': 'Padrão',
    'comprimir_miolo': 'Comprimir miolo',
    'curvas': 'Converter em curvas',
}
GPRINT_SETORES_REFAZER_INICIAIS = (
    ('PRE-IMPRESSÃO', 'comprimir_miolo'),
    ('PROBLEMA DE FONTE', 'curvas'),
    ('IMPRESSÃO', 'padrao'),
    ('REFILE INICIAL MIOLO', 'padrao'),
    ('REFILE INICIAL CAPA', 'padrao'),
    ('SEPARAÇÃO', 'padrao'),
    ('ENCADERNAÇÃO', 'padrao'),
    ('LAMINAÇÃO', 'padrao'),
    ('REFILE ORELHA', 'padrao'),
    ('REFILE FINAL', 'padrao'),
    ('SHIRINK', 'padrao'),
    ('EXPEDIÇÃO', 'padrao'),
)


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

    def preparar_banco(self, pastas_iniciais=True):
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
                conn.execute('''CREATE TABLE IF NOT EXISTS operadores (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    ativo INTEGER NOT NULL DEFAULT 1,
                    criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )''')

                conn.execute("""CREATE TABLE IF NOT EXISTS pastas_clientes_refazer (
                    slot INTEGER PRIMARY KEY CHECK(slot BETWEEN 0 AND 5),
                    nome TEXT NOT NULL,
                    caminho TEXT NOT NULL,
                    modo TEXT NOT NULL CHECK(modo IN ('clientes', 'planilha', 'pdf', 'uiclap_legado'))
                )""")
                conn.executemany(
                    'INSERT OR IGNORE INTO pastas_clientes_refazer(slot, nome, caminho, modo) VALUES (?, ?, ?, ?)',
                    GPRINT_PASTAS_REFAZER_INICIAIS if pastas_iniciais
                    else [(slot, '', '', 'clientes') for slot in range(6)])

                setor_novo = conn.execute(
                    "SELECT 1 FROM sqlite_schema WHERE name='setores_refazer'").fetchone() is None
                conn.execute("""CREATE TABLE IF NOT EXISTS setores_refazer (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    tratamento TEXT NOT NULL DEFAULT 'padrao'
                        CHECK(tratamento IN ('padrao', 'comprimir_miolo', 'curvas')),
                    ativo INTEGER NOT NULL DEFAULT 1 CHECK(ativo IN (0, 1))
                )""")
                if setor_novo and pastas_iniciais:
                    conn.executemany('INSERT INTO setores_refazer(nome, tratamento) VALUES (?, ?)',
                                     GPRINT_SETORES_REFAZER_INICIAIS)

    def listar_setores(self, incluir_inativos=False):
        sql = 'SELECT id, nome, tratamento, ativo FROM setores_refazer'
        if not incluir_inativos:
            sql += ' WHERE ativo=1'
        with self._conectar() as conn:
            return [dict(r) for r in conn.execute(sql + ' ORDER BY id')]

    def salvar_setor(self, nome, tratamento, ativo, setor_id=None):
        nome = nome.strip().upper()
        if not nome or len(nome) > 80:
            raise ValueError('Informe um nome de setor com até 80 caracteres.')
        if tratamento not in GPRINT_TRATAMENTOS_SETORES:
            raise ValueError('Selecione um tratamento de PDF válido.')
        campos = nome, tratamento, int(bool(ativo))
        with self._conectar() as conn, conn:
            if setor_id is None:
                return conn.execute(
                    'INSERT INTO setores_refazer(nome, tratamento, ativo) VALUES (?, ?, ?)',
                    campos).lastrowid
            atualizado = conn.execute(
                'UPDATE setores_refazer SET nome=?, tratamento=?, ativo=? WHERE id=?',
                (*campos, setor_id))
            if not atualizado.rowcount:
                raise ValueError('O setor não está mais no cadastro.')
            return setor_id

    def listar_pastas_refazer(self):
        with self._conectar() as conn:
            return [dict(linha) for linha in conn.execute(
                'SELECT slot, nome, caminho, modo FROM pastas_clientes_refazer ORDER BY slot')]

    def salvar_pastas_refazer(self, pastas, versao_anterior):
        """Salva os seis espaços como um conjunto, sem sobrescrever edição de outra estação."""
        if [p['slot'] for p in pastas] != list(range(6)):
            raise ValueError('Informe a pasta principal e os cinco espaços adicionais.')
        caminhos = set()
        nomes = set()
        valores = []
        for pasta in pastas:
            slot, nome, caminho, modo = (pasta[k] for k in ('slot', 'nome', 'caminho', 'modo'))
            nome, caminho = nome.strip(), caminho.strip()
            if not caminho:
                if slot == 0:
                    raise ValueError('Informe o caminho da pasta principal.')
                valores.append(('', '', 'clientes', slot))
                continue
            caminho = ntpath.normpath(caminho)
            if not ntpath.splitdrive(caminho)[0] or not ntpath.isabs(caminho):
                raise ValueError(f'{nome or "Pasta adicional"}: informe um caminho completo de rede ou disco.')
            if modo not in GPRINT_MODOS_PASTAS_REFAZER or (slot == 0 and modo != 'clientes'):
                raise ValueError('Selecione uma estrutura válida para a pasta.')
            if not nome:
                raise ValueError(f'Informe o nome da pasta no espaço {slot}.')
            chave = ntpath.normcase(caminho)
            if chave in caminhos or nome.casefold() in nomes:
                raise ValueError('Cada pasta deve ter um nome e um caminho diferentes.')
            caminhos.add(chave)
            nomes.add(nome.casefold())
            valores.append((nome, caminho, modo, slot))
        with self._conectar() as conn, conn:
            conn.execute('BEGIN IMMEDIATE')
            atual = [dict(linha) for linha in conn.execute(
                'SELECT slot, nome, caminho, modo FROM pastas_clientes_refazer ORDER BY slot')]
            if atual != versao_anterior:
                raise ValueError('Outra estação alterou o cadastro. Recarregue antes de salvar.')
            conn.executemany('UPDATE pastas_clientes_refazer SET nome=?, caminho=?, modo=? WHERE slot=?',
                             valores)


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


    def listar_operadores(self, incluir_inativos=False):
        sql = 'SELECT id, nome, ativo FROM operadores'
        if not incluir_inativos:
            sql += ' WHERE ativo=1'
        sql += ' ORDER BY nome COLLATE NOCASE'
        with self._conectar() as conn:
            return [dict(r) for r in conn.execute(sql)]

    def salvar_operador(self, nome, operador_id=None):
        nome = str(nome).strip().upper()
        if not nome or len(nome) > 80:
            raise ValueError('Informe nome de operador com ate 80 caracteres.')
        with self._conectar() as conn:
            with conn:
                if operador_id:
                    conn.execute('UPDATE operadores SET nome=? WHERE id=?', (nome, operador_id))
                    return operador_id
                cur = conn.execute('INSERT INTO operadores(nome) VALUES(?)', (nome,))
                return cur.lastrowid

    def alterar_ativo_operador(self, operador_id, ativo):
        with self._conectar() as conn:
            with conn:
                conn.execute('UPDATE operadores SET ativo=? WHERE id=?',
                             (int(bool(ativo)), operador_id))


    def excluir_operador(self, operador_id):
        """Exclui cadastro; nao apaga nomes gravados no banco de perdas."""
        with self._conectar() as conn:
            with conn:
                return conn.execute('DELETE FROM operadores WHERE id=?', (operador_id,)).rowcount


def gprint_nome_maquina(equipamento):
    return f"{equipamento['nome']} - {equipamento['ip']}"


def gprint_centralizar_janela(janela, parent=None, largura=None, altura=None):
    """Centraliza um Toplevel em relacao a janela principal, respeitando a tela.

    Nao altera dialogos nativos do Windows nem os tooltips.
    """
    janela.update_idletasks()
    tela_largura = janela.winfo_screenwidth()
    tela_altura = janela.winfo_screenheight()
    largura = min(int(largura or max(janela.winfo_width(), janela.winfo_reqwidth())), tela_largura)
    altura = min(int(altura or max(janela.winfo_height(), janela.winfo_reqheight())), tela_altura)
    if parent is not None and parent.winfo_exists():
        referencia = parent.winfo_toplevel()
        x = referencia.winfo_rootx() + (referencia.winfo_width() - largura) // 2
        y = referencia.winfo_rooty() + (referencia.winfo_height() - altura) // 2
    else:
        x = (tela_largura - largura) // 2
        y = (tela_altura - altura) // 2
    x = max(0, min(x, tela_largura - largura))
    y = max(0, min(y, tela_altura - altura))
    janela.geometry(f'{largura}x{altura}+{x}+{y}')
    janela.deiconify()


def gprint_autorizar_relatorio(parent):
    """Solicita a MESMA senha definida para administrar os equipamentos."""
    senha_configurada = GPRINT_SENHA_ADMIN_EQUIPAMENTOS
    if not senha_configurada.strip() or senha_configurada == 'ALTERE_ESTA_SENHA_ANTES_DE_USAR':
        messagebox.showwarning('GPrint - Acesso restrito',
            'Configure a senha administrativa antes de consultar os relatórios.',
            parent=parent)
        return False
    senha = simpledialog.askstring('GPrint - Relatorio de perdas',
        'Informe a senha administrativa para consultar as perdas:',
        show='*', parent=parent)
    if senha is None:
        return False
    if not hmac.compare_digest(senha, senha_configurada):
        messagebox.showerror('GPrint - Acesso negado',
            'Senha incorreta. O resumo de perdas permanece protegido.',
            parent=parent)
        return False
    return True


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
        self.withdraw()
        self._montar()
        self.atualizar()
        gprint_centralizar_janela(self, cadastro, 980, 650)
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

class GPrintCadastroOperadoresTab(ttk.Frame):
    """Cadastro e edicao dos operadores do Refazer na subaba restrita de Configuracoes.

    Os nomes continuam em equipamentos.db. A exclusao nunca apaga o historico
    de retrabalhos no refazer.db, que guarda o nome usado em cada registro.
    """
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.repo = app.repositorio_equipamentos
        self.nome_var = tk.StringVar()
        self.estado_var = tk.StringVar(value='Selecione um operador para editar, ou clique em Novo.')
        self.operador_id = None
        self.dados = {}
        self._montar_tela()

    def _montar_tela(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        formulario = ttk.LabelFrame(self, text='Operadores do Refazer — cadastro central', padding=12)
        formulario.grid(row=0, column=0, sticky='ew', padx=10, pady=8)
        formulario.columnconfigure(1, weight=1)
        ttk.Label(formulario, text='Nome:').grid(row=0, column=0, sticky='w', padx=(0, 8))
        self.entry_nome = ttk.Entry(formulario, textvariable=self.nome_var)
        self.entry_nome.grid(row=0, column=1, sticky='ew', padx=5)
        self.entry_nome.bind('<Return>', lambda _event: self.salvar())
        botoes = ttk.Frame(formulario)
        botoes.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(12, 4))
        ttk.Button(botoes, text='Novo operador', command=self.novo).pack(side=tk.LEFT, padx=3)
        ttk.Button(botoes, text='Editar selecionado', command=self.editar,
                   bootstyle=INFO).pack(side=tk.LEFT, padx=3)
        ttk.Button(botoes, text='Salvar cadastro / edicao', command=self.salvar,
                   bootstyle=SUCCESS).pack(side=tk.LEFT, padx=3)
        ttk.Label(formulario, textvariable=self.estado_var, bootstyle=SECONDARY).grid(
            row=2, column=0, columnspan=2, sticky='w', pady=(6, 0))

        painel = ttk.LabelFrame(self, text='Operadores cadastrados', padding=8)
        painel.grid(row=1, column=0, sticky='nsew', padx=10, pady=5)
        painel.columnconfigure(0, weight=1)
        painel.rowconfigure(0, weight=1)
        self.lista = ttk.Treeview(painel, columns=('nome', 'situacao'),
                                  show='headings', selectmode='browse', height=12)
        self.lista.heading('nome', text='Operador')
        self.lista.heading('situacao', text='Situacao')
        self.lista.column('nome', width=400, anchor='w')
        self.lista.column('situacao', width=125, anchor='center')
        self.lista.grid(row=0, column=0, sticky='nsew')
        barra = ttk.Scrollbar(painel, orient=tk.VERTICAL, command=self.lista.yview)
        barra.grid(row=0, column=1, sticky='ns')
        self.lista.configure(yscrollcommand=barra.set)
        self.lista.bind('<<TreeviewSelect>>', self.selecionar)
        self.lista.bind('<Double-1>', self._editar_por_duplo_clique)

        acoes = ttk.Frame(self, padding=(10, 4))
        acoes.grid(row=2, column=0, sticky='ew')
        ttk.Button(acoes, text='Ativar / Desativar', command=self.alternar,
                   bootstyle=WARNING).pack(side=tk.LEFT, padx=3)
        ttk.Button(acoes, text='Excluir operador', command=self.excluir,
                   bootstyle=DANGER).pack(side=tk.LEFT, padx=3)
        ttk.Button(acoes, text='Atualizar lista', command=self.atualizar).pack(side=tk.RIGHT, padx=3)
        ttk.Label(self,
                  text='Alteracoes afetam os operadores disponiveis no Refazer. '
                       'Registros antigos de perdas permanecem preservados.',
                  bootstyle=SECONDARY).grid(row=3, column=0, sticky='w', padx=15, pady=8)

    def _erro(self, exc):
        messagebox.showerror('GPrint - Operadores',
                             f'Nao foi possivel concluir a operacao.\n{exc}', parent=self)

    def novo(self):
        self.operador_id = None
        self.nome_var.set('')
        selecionados = self.lista.selection()
        if selecionados:
            self.lista.selection_remove(*selecionados)
        self.estado_var.set('Novo operador: digite o nome e clique em Salvar.')
        self.entry_nome.focus_set()

    def atualizar(self, selecionar_id=None):
        """Recarrega somente sob demanda, sem consultar o banco na abertura geral."""
        try:
            anterior = str(selecionar_id or self.operador_id or '')
            linhas = self.repo.listar_operadores(incluir_inativos=True)
            self.dados = {str(r['id']): r for r in linhas}
            antigos = self.lista.get_children()
            if antigos:
                self.lista.delete(*antigos)
            for r in linhas:
                self.lista.insert('', 'end', iid=str(r['id']),
                                  values=(r['nome'], 'Ativo' if r['ativo'] else 'Inativo'))
            if anterior and anterior in self.dados:
                self.lista.selection_set(anterior)
                self.lista.see(anterior)
                self.selecionar()
            else:
                self.operador_id = None
                self.nome_var.set('')
                self.estado_var.set('Selecione um operador para editar, ou clique em Novo.')
        except Exception as exc:
            self._erro(exc)

    def selecionar(self, _event=None):
        selecionado = self.lista.selection()
        if not selecionado:
            return
        linha = self.dados.get(selecionado[0])
        if linha:
            self.operador_id = linha['id']
            self.nome_var.set(linha['nome'])
            self.estado_var.set(f"Operador selecionado: {linha['nome']}. Pode editar e salvar.")

    def _editar_por_duplo_clique(self, event):
        item = self.lista.identify_row(event.y)
        if item and item in self.dados:
            self.lista.selection_set(item)
            self.selecionar()
            self.editar()

    def editar(self):
        if self.operador_id is None or str(self.operador_id) not in self.dados:
            messagebox.showwarning('GPrint - Operadores',
                                   'Selecione um operador na tabela para editar.', parent=self)
            return
        self.entry_nome.focus_set()
        self.entry_nome.selection_range(0, tk.END)

    def salvar(self):
        nome = self.nome_var.get().strip()
        if not nome:
            messagebox.showwarning('GPrint - Operadores',
                                   'Informe o nome do operador.', parent=self)
            return
        try:
            operador_id = self.repo.salvar_operador(nome, self.operador_id)
            self.atualizar(selecionar_id=operador_id)
            self.app.aba_refazer.atualizar_operadores_refazer()
            self.estado_var.set('Cadastro salvo com sucesso. Alteracoes disponiveis no Refazer.')
        except Exception as exc:
            self._erro(exc)

    def alternar(self):
        linha = self.dados.get(str(self.operador_id))
        if linha is None:
            messagebox.showwarning('GPrint - Operadores',
                                   'Selecione um operador para ativar ou desativar.', parent=self)
            return
        novo_ativo = not bool(linha['ativo'])
        acao = 'Ativar' if novo_ativo else 'Desativar'
        if not messagebox.askyesno('GPrint - Operadores',
                                   f"{acao} o operador '{linha['nome']}'?", parent=self):
            return
        try:
            self.repo.alterar_ativo_operador(linha['id'], novo_ativo)
            self.atualizar(selecionar_id=linha['id'])
            self.app.aba_refazer.atualizar_operadores_refazer()
        except Exception as exc:
            self._erro(exc)

    def excluir(self):
        linha = self.dados.get(str(self.operador_id))
        if linha is None:
            messagebox.showwarning('GPrint - Operadores',
                                   'Selecione um operador para excluir.', parent=self)
            return
        if not messagebox.askyesno('GPrint - Excluir operador',
                f"Excluir o cadastro de '{linha['nome']}'?\n\n"
                'Ele deixara de aparecer para novos retrabalhos.\n'
                'O historico de perdas ja registrado sera preservado.', parent=self):
            return
        try:
            if not self.repo.excluir_operador(linha['id']):
                raise ValueError('Operador nao encontrado no cadastro central.')
            self.novo()
            self.atualizar()
            self.app.aba_refazer.atualizar_operadores_refazer()
        except Exception as exc:
            self._erro(exc)


class GPrintCadastroSetoresTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=14)
        self.app = app
        self.repo = app.repositorio_equipamentos
        self.dados = {}
        self.setor_id = None
        self.nome = tk.StringVar()
        self.tratamento = tk.StringVar(value=GPRINT_TRATAMENTOS_SETORES['padrao'])
        self.ativo = tk.BooleanVar(value=True)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(3, weight=1)
        ttk.Label(self, text='Nome do setor').grid(row=0, column=0, sticky='w', padx=(0, 8))
        entrada = ttk.Entry(self, textvariable=self.nome)
        entrada.grid(row=0, column=1, sticky='ew')
        ttk.Label(self, text='Tratamento do PDF').grid(row=1, column=0, sticky='w', pady=6)
        ttk.Combobox(self, textvariable=self.tratamento, state='readonly',
                     values=list(GPRINT_TRATAMENTOS_SETORES.values())).grid(row=1, column=1, sticky='ew')
        acoes = ttk.Frame(self)
        acoes.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(3, 10))
        ttk.Button(acoes, text='Novo setor', command=self._novo).pack(side='left')
        ttk.Button(acoes, text='Salvar setor', command=self._salvar, bootstyle=SUCCESS).pack(side='left', padx=8)
        ttk.Checkbutton(acoes, text='Ativo para novos retrabalhos', variable=self.ativo).pack(side='left', padx=6)
        host = ttk.Frame(self)
        host.grid(row=3, column=0, columnspan=2, sticky='nsew')
        host.columnconfigure(0, weight=1)
        host.rowconfigure(0, weight=1)
        self.lista = ttk.Treeview(host, columns=('nome', 'tratamento', 'estado'),
                                  show='headings', selectmode='browse')
        for coluna, titulo, largura in (
            ('nome', 'Setor', 200), ('tratamento', 'Tratamento do PDF', 230), ('estado', 'Situação', 90)
        ):
            self.lista.heading(coluna, text=titulo)
            self.lista.column(coluna, width=largura)
        self.lista.grid(row=0, column=0, sticky='nsew')
        rolagem = ttk.Scrollbar(host, command=self.lista.yview)
        rolagem.grid(row=0, column=1, sticky='ns')
        self.lista.configure(yscrollcommand=rolagem.set)
        self.lista.bind('<<TreeviewSelect>>', self._selecionar)

    def _novo(self):
        self.setor_id = None
        selecionados = self.lista.selection()
        if selecionados:
            self.lista.selection_remove(*selecionados)
        self.nome.set('')
        self.tratamento.set(GPRINT_TRATAMENTOS_SETORES['padrao'])
        self.ativo.set(True)

    def _selecionar(self, _event=None):
        selecao = self.lista.selection()
        if selecao:
            setor = self.dados[selecao[0]]
            self.setor_id = setor['id']
            self.nome.set(setor['nome'])
            self.tratamento.set(GPRINT_TRATAMENTOS_SETORES[setor['tratamento']])
            self.ativo.set(bool(setor['ativo']))

    def atualizar(self, setor_id=None):
        try:
            setores = self.repo.listar_setores(incluir_inativos=True)
        except GPrintCadastroErro as erro:
            messagebox.showerror('GPrint - Setores', str(erro), parent=self)
            return
        self._novo()
        self.lista.delete(*self.lista.get_children())
        self.dados = {str(r['id']): r for r in setores}
        for iid, setor in self.dados.items():
            self.lista.insert('', 'end', iid=iid, values=(
                setor['nome'], GPRINT_TRATAMENTOS_SETORES[setor['tratamento']],
                'Ativo' if setor['ativo'] else 'Inativo'))
        if str(setor_id) in self.dados:
            self.lista.selection_set(str(setor_id))
            self._selecionar()

    def _salvar(self):
        tratamento = next(codigo for codigo, rotulo in GPRINT_TRATAMENTOS_SETORES.items()
                          if rotulo == self.tratamento.get())
        try:
            setor_id = self.repo.salvar_setor(self.nome.get(), tratamento, self.ativo.get(), self.setor_id)
        except (GPrintCadastroErro, ValueError) as erro:
            messagebox.showerror('GPrint - Setores', str(erro), parent=self)
            return
        self.atualizar(setor_id)
        self.app.aba_refazer.atualizar_setores_refazer()


class GPrintCadastroPastasRefazerTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=14)
        self.app = app
        self.repo = app.repositorio_equipamentos
        self.versao = None
        self.campos = []
        self.columnconfigure(2, weight=1)
        ttk.Label(
            self, text='A pasta principal contém as pastas dos clientes. Cadastre até cinco pastas adicionais.',
            wraplength=680,
        ).grid(row=0, column=0, columnspan=5, sticky='w', pady=(0, 14))
        for coluna, texto in enumerate(('Pasta', 'Nome', 'Caminho', 'Estrutura')):
            ttk.Label(self, text=texto).grid(row=1, column=coluna, sticky='w', padx=(0, 8))
        for slot in range(6):
            linha = slot + 2
            nome, caminho, modo = tk.StringVar(), tk.StringVar(), tk.StringVar()
            self.campos.append((nome, caminho, modo))
            label = 'Principal' if slot == 0 else f'Adicional {slot}'
            ttk.Label(self, text=label).grid(row=linha, column=0, sticky='w', padx=(0, 8), pady=4)
            ttk.Entry(self, textvariable=nome, width=16).grid(row=linha, column=1, sticky='ew', padx=(0, 8))
            ttk.Entry(self, textvariable=caminho, width=22).grid(row=linha, column=2, sticky='ew', padx=(0, 8))
            tipos = GPRINT_MODOS_PASTAS_REFAZER
            ttk.Combobox(self, textvariable=modo, values=list(tipos.values()), width=20,
                         state='readonly' if slot else 'disabled').grid(
                             row=linha, column=3, sticky='ew', padx=(0, 8))
            ttk.Button(self, text='Procurar…', command=lambda var=caminho: self._procurar(var)).grid(
                row=linha, column=4, sticky='w')
        acoes = ttk.Frame(self)
        acoes.grid(row=8, column=0, columnspan=5, sticky='ew', pady=(15, 0))
        self.salvar = ttk.Button(acoes, text='Salvar pastas', command=self._salvar, bootstyle=SUCCESS)
        self.salvar.pack(side='left', padx=(0, 8))
        ttk.Button(acoes, text='Recarregar', command=self.atualizar).pack(side='left')
        ttk.Label(self, text='Para remover uma pasta adicional, limpe o campo Caminho e salve.',
                  wraplength=680).grid(row=9, column=0, columnspan=5, sticky='w', pady=(8, 0))

    def _procurar(self, caminho):
        escolhido = filedialog.askdirectory(parent=self, title='Selecionar pasta de clientes ou lotes')
        if escolhido:
            caminho.set(escolhido)

    def atualizar(self):
        try:
            dados = self.repo.listar_pastas_refazer()
        except GPrintCadastroErro as erro:
            self.versao = None
            self.salvar.configure(state='disabled')
            messagebox.showerror('GPrint - Pastas de clientes', str(erro), parent=self)
            return
        self.versao = dados
        self.salvar.configure(state='normal')
        for pasta, (nome, caminho, modo) in zip(dados, self.campos):
            nome.set(pasta['nome'])
            caminho.set(pasta['caminho'])
            modo.set(GPRINT_MODOS_PASTAS_REFAZER[pasta['modo']])

    def _salvar(self):
        codigos = {rotulo: codigo for codigo, rotulo in GPRINT_MODOS_PASTAS_REFAZER.items()}
        pastas = [
            dict(slot=slot, nome=nome.get(), caminho=caminho.get(), modo=codigos.get(modo.get()))
            for slot, (nome, caminho, modo) in enumerate(self.campos)
        ]
        try:
            self.repo.salvar_pastas_refazer(pastas, self.versao)
        except (ValueError, GPrintCadastroErro) as erro:
            messagebox.showerror('GPrint - Pastas de clientes', str(erro), parent=self)
            return
        self.atualizar()
        self.app.aba_refazer.carregar_clientes(recarregar=True)
        messagebox.showinfo('GPrint - Pastas de clientes', 'Cadastro atualizado.', parent=self)


class GPrintBancosAdministrativos:
    """Backups externos e restauração dos três bancos centrais."""
    def __init__(self, equipamentos):
        self.equipamentos = equipamentos
        self.caminhos = {
            'equipamentos.db': Path(equipamentos.db_path),
            'impressoes.db': Path(GPRINT_BANCO_HISTORICO),
            'refazer.db': Path(GPRINT_BANCO_REFAZER),
        }

    @staticmethod
    def _uri_leitura(caminho):
        caminho = caminho.resolve()
        uri = caminho.as_uri()
        if caminho.drive.startswith('\\\\'):
            # SQLite espera um caminho UNC com autoridade vazia na URI.
            uri = 'file:////' + uri.removeprefix('file://')
        return uri + '?mode=ro'

    def gerar_pacote(self, pasta_destino, vazio=False):
        pasta_destino = Path(pasta_destino)
        tipo = 'GPrint_bancos_vazios' if vazio else 'GPrint_backup'
        destino = pasta_destino / f'{tipo}_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:6]}'
        with tempfile.TemporaryDirectory(prefix='gprint_', dir=pasta_destino) as temporario:
            pasta = Path(temporario) / 'bancos'
            pasta.mkdir()
            if vazio:
                self._preparar_bancos({nome: pasta / nome for nome in self.caminhos}, vazio=True)
            else:
                for nome, central in self.caminhos.items():
                    if not central.is_file():
                        raise ValueError(f'Não foi possível localizar {nome} na pasta central LOGS.')
                    with closing(sqlite3.connect(self._uri_leitura(central), uri=True, timeout=15)) as origem:
                        with closing(sqlite3.connect(pasta / nome)) as copia:
                            origem.backup(copia)
            pasta.rename(destino)
        return destino

    @staticmethod
    def _preparar_bancos(caminhos, vazio=False):
        GPrintEquipamentosRepository(caminhos['equipamentos.db']).preparar_banco(pastas_iniciais=not vazio)
        for nome, criar_schema in (
            ('impressoes.db', gprint_criar_schema_impressoes), ('refazer.db', gprint_criar_schema_refazer)
        ):
            with closing(sqlite3.connect(caminhos[nome], timeout=15)) as conn, conn:
                criar_schema(conn)

    def importar_pacote(self, pasta_origem):
        fontes = {nome: Path(pasta_origem) / nome for nome in self.caminhos}
        for nome, caminho in fontes.items():
            if not caminho.is_file():
                raise ValueError(f'Arquivo obrigatório ausente na pasta selecionada: {nome}.')
            if any(central.exists() and caminho.samefile(central) for central in self.caminhos.values()):
                raise ValueError('Selecione uma pasta de backup diferente dos bancos centrais em uso.')

        estruturas = (
            ('main', 'equipamentos', ('id', 'nome', 'ip', 'tipo', 'ativo', 'criado_em')),
            ('main', 'filas', ('id', 'equipamento_id', 'nome', 'descricao', 'ignorar_folha_obra', 'ativo')),
            ('main', 'operadores', ('id', 'nome', 'ativo', 'criado_em')),
            ('main', 'setores_refazer', ('id', 'nome', 'tratamento', 'ativo')),
            ('main', 'pastas_clientes_refazer', ('slot', 'nome', 'caminho', 'modo')),
            ('impressoes', 'logs', (
                'id', 'data_hora', 'lote', 'pilha', 'arquivo', 'maquina', 'ip', 'formato',
                'modo', 'quantidade', 'status', 'erro', 'copia', 'usuario', 'montagem_ok', 'envio_ok')),
            ('impressoes', 'gprint_ids_recebidos', ('evento_id', 'data_hora')),
            ('refazer', 'refazer_registros', GPRINT_REFAZER_CAMPOS),
        )
        aliases = dict(zip(self.caminhos, ('main', 'impressoes', 'refazer')))
        self._preparar_bancos(self.caminhos)
        with closing(sqlite3.connect(self.caminhos['equipamentos.db'], uri=True, timeout=15)) as conn, conn:
            conn.execute('PRAGMA foreign_keys = ON')
            conn.execute('PRAGMA trusted_schema = OFF')
            for nome, alias in aliases.items():
                if alias != 'main':
                    conn.execute(f'ATTACH DATABASE ? AS {alias}', (str(self.caminhos[nome]),))
                conn.execute(f'ATTACH DATABASE ? AS origem_{alias}', (self._uri_leitura(fontes[nome]),))
                modo = conn.execute(f'PRAGMA {alias}.journal_mode').fetchone()[0]
                if modo not in ('delete', 'truncate', 'persist'):
                    raise ValueError(f'{nome}: modo de journal {modo} incompatível com restauração conjunta.')
                conn.execute(f'PRAGMA {alias}.synchronous = FULL')
            conn.execute('BEGIN IMMEDIATE')
            operacoes = []
            for alias, tabela, colunas in estruturas:
                origem = f'origem_{alias}'
                objeto = conn.execute(
                    f'SELECT type FROM {origem}.sqlite_schema WHERE name=?', (tabela,)).fetchone()
                if objeto is None and tabela in ('pastas_clientes_refazer', 'setores_refazer'):
                    continue  # Preserva cadastros que ainda não existiam no backup.
                if objeto is None and tabela == 'gprint_ids_recebidos':
                    operacoes.append((alias, tabela, None))
                    continue
                disponiveis = {r[1] for r in conn.execute(f'PRAGMA {origem}.table_info({tabela})')}
                if objeto != ('table',) or not set(colunas) <= disponiveis:
                    raise ValueError(f'O backup não tem uma estrutura compatível para {tabela}.')
                if tabela == 'pastas_clientes_refazer':
                    slots = conn.execute(f'SELECT slot FROM {origem}.{tabela} ORDER BY slot').fetchall()
                    if slots != [(slot,) for slot in range(6)]:
                        raise ValueError('O backup de pastas precisa ter os seis espaços de cadastro.')
                operacoes.append((alias, tabela, ', '.join(colunas)))

            for alias, tabela, _ in reversed(operacoes):
                conn.execute(f'DELETE FROM {alias}.{tabela}')
            for alias, tabela, colunas in operacoes:
                if colunas:
                    conn.execute(f'INSERT INTO {alias}.{tabela} ({colunas}) '
                                 f'SELECT {colunas} FROM origem_{alias}.{tabela}')


class GPrintBancoAdministrativoTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding=18)
        self.app = app
        self.bancos = GPrintBancosAdministrativos(app.repositorio_equipamentos)
        self.columnconfigure(0, weight=1)
        ttk.Label(self, text='Bancos centrais da pasta LOGS', font=('Segoe UI', 12, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky='w')
        ttk.Label(self, text='equipamentos.db  |  impressoes.db  |  refazer.db').grid(
            row=1, column=0, columnspan=2, sticky='w', pady=(5, 14))
        opcoes = (
            ('Exportar bancos', 'Cria uma pasta de backup com os três arquivos .db.', self._exportar),
            ('Importar bancos', 'Restaura os três arquivos de uma pasta de backup.', self._importar),
            ('Gerar bancos vazios', 'Cria os três arquivos com suas tabelas, prontos para uso.',
             self._gerar_vazios),
        )
        for linha, (titulo, descricao, acao) in enumerate(opcoes, 2):
            ttk.Label(self, text=descricao, wraplength=490).grid(
                row=linha, column=0, sticky='w', pady=7, padx=(0, 12))
            ttk.Button(self, text=titulo, command=acao).grid(row=linha, column=1, sticky='ew', pady=7)

    def _exportar(self):
        self._salvar_pacote('Exportar bancos', vazio=False)

    def _gerar_vazios(self):
        self._salvar_pacote('Gerar bancos vazios', vazio=True)

    def _salvar_pacote(self, titulo, vazio):
        destino = filedialog.askdirectory(parent=self, title=f'{titulo} — escolha a pasta de destino')
        if not destino:
            return
        try:
            pasta = self.bancos.gerar_pacote(destino, vazio=vazio)
        except (ValueError, GPrintCadastroErro, OSError, sqlite3.Error) as erro:
            messagebox.showerror(titulo, str(erro), parent=self)
        else:
            messagebox.showinfo(titulo, f'Arquivos salvos em:\n{pasta}', parent=self)

    def _importar(self):
        origem = filedialog.askdirectory(
            parent=self, title='Selecione a pasta com equipamentos.db, impressoes.db e refazer.db')
        if not origem or not messagebox.askyesno(
            'GPrint - Confirmar restauração',
            'Substituir os cadastros, o histórico de impressões e os registros de retrabalho '
            'pelos dados da pasta selecionada?\n\n'
            'A alteração vale para todas as estações que usam os bancos centrais.',
            parent=self,
        ):
            return
        try:
            self.bancos.importar_pacote(origem)
        except (ValueError, GPrintCadastroErro, OSError, sqlite3.Error) as erro:
            messagebox.showerror('GPrint - Importar bancos', str(erro), parent=self)
            return
        self.app.aba_config.atualizar_cadastros_apos_importacao()
        messagebox.showinfo('GPrint - Importar bancos', 'Os três bancos foram restaurados.', parent=self)


class ConfigTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self._admin_liberado = False
        self._cadastro_widget = None
        self._widgets_cadastros = []
        self._criar_configuracoes_internas()
        self.setup_ui()

    def _criar_configuracoes_internas(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.abas_config = ttk.Notebook(self)
        self.abas_config.grid(row=0, column=0, sticky='nsew')
        self.tab_preferencias = ttk.Frame(self.abas_config)
        self.tab_administracao = ttk.Frame(self.abas_config)
        self.abas_config.add(self.tab_preferencias, text='Preferências pessoais')
        self.abas_config.add(self.tab_administracao, text='Administração central (restrito)')

        self.tab_administracao.columnconfigure(0, weight=1)
        self.tab_administracao.rowconfigure(0, weight=1)
        self._bloqueio_admin = ttk.Frame(self.tab_administracao, padding=22)
        self._bloqueio_admin.grid(row=0, column=0, sticky='nsew')
        ttk.Label(self._bloqueio_admin, text='Administração central',
                  font=('Segoe UI', 15, 'bold')).pack(anchor='w')
        ttk.Label(self._bloqueio_admin, text='Equipamentos, filas, operadores, setores e pastas de clientes.',
                  wraplength=600).pack(anchor='w', pady=(8, 12))
        ttk.Button(self._bloqueio_admin, text='Desbloquear cadastros',
                   command=self.solicitar_acesso_administrativo, bootstyle=PRIMARY).pack(anchor='w')
        self.abas_config.bind('<<NotebookTabChanged>>', self._ao_mudar_configuracao)

    def _ao_mudar_configuracao(self, _event=None):
        if self.abas_config.select() != str(self.tab_administracao):
            self.bloquear_cadastro()

    def mostrar_equipamentos(self):
        self.abas_config.select(self.tab_administracao)
        self.solicitar_acesso_administrativo()

    def solicitar_acesso_administrativo(self):
        if self._admin_liberado or self.abas_config.select() != str(self.tab_administracao):
            return
        senha_configurada = GPRINT_SENHA_ADMIN_EQUIPAMENTOS
        if not senha_configurada.strip():
            messagebox.showwarning('GPrint - Cadastro restrito',
                                   'Configure a senha administrativa.', parent=self)
            return
        senha = simpledialog.askstring('GPrint - Acesso administrativo',
                                       'Informe a senha da administração central:', show='*', parent=self)
        if senha is None:
            return
        if not hmac.compare_digest(senha.encode('utf-8'), senha_configurada.encode('utf-8')):
            messagebox.showerror('GPrint - Acesso negado', 'Senha incorreta.', parent=self)
            return
        self._admin_liberado = True
        self._bloqueio_admin.grid_remove()
        painel = self._cadastro_widget = ttk.Frame(self.tab_administracao, padding=10)
        painel.grid(row=0, column=0, sticky='nsew')
        painel.columnconfigure(0, weight=1)
        painel.rowconfigure(1, weight=1)
        barra = ttk.Frame(painel)
        barra.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        ttk.Label(barra, text='Alterações compartilhadas entre as estações').pack(side='left')
        ttk.Button(barra, text='Bloquear cadastros', command=self.bloquear_cadastro,
                   bootstyle=WARNING).pack(side='right')
        abas = ttk.Notebook(painel)
        abas.grid(row=1, column=0, sticky='nsew')
        secoes = {}
        for titulo, classe in (
            ('Equipamentos e filas', GPrintCadastroEquipamentosTab),
            ('Operadores do Refazer', GPrintCadastroOperadoresTab),
            ('Setores do Refazer', GPrintCadastroSetoresTab),
            ('Pastas de clientes', GPrintCadastroPastasRefazerTab),
            ('Banco de dados', GPrintBancoAdministrativoTab),
        ):
            aba = ttk.Frame(abas)
            abas.add(aba, text=titulo)
            secoes[str(aba)] = (aba, classe)

        def abrir_secao(evento=None):
            secao = secoes.get(abas.select())
            if not secao:
                return
            aba, classe = secao
            if not aba.winfo_children():
                widget = classe(aba, self.app)
                widget.pack(fill='both', expand=True)
                if not isinstance(widget, GPrintBancoAdministrativoTab):
                    self._widgets_cadastros.append(widget)
                    widget.atualizar()

        abas.bind('<<NotebookTabChanged>>', abrir_secao)
        abrir_secao()

    def atualizar_cadastros_apos_importacao(self):
        for widget in self._widgets_cadastros:
            widget.atualizar()
        self.app.atualizar_listas_equipamentos()
        self.app.aba_refazer.atualizar_destinos_refazer()
        self.app.aba_refazer.atualizar_operadores_refazer()
        self.app.aba_refazer.atualizar_setores_refazer()
        self.app.aba_refazer.carregar_clientes(recarregar=True)

    def bloquear_cadastro(self):
        self._admin_liberado = False
        self._widgets_cadastros.clear()
        if self._cadastro_widget is not None:
            self._cadastro_widget.destroy()
            self._cadastro_widget = None
        self._bloqueio_admin.grid()

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

        ttk.Label(miolo_frame, text="Montagens simultâneas:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w")
        self.montagens_miolo_var = tk.StringVar()
        ttk.Label(miolo_frame, textvariable=self.montagens_miolo_var).grid(
            row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(miolo_frame, text="Alterar com senha", command=self.alterar_montagens_miolo).grid(
            row=1, column=2, padx=5, pady=5, sticky="w")

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
        
        self.montagens_miolo_var.set(str(obter_montagens_simultaneas_miolo()))

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
    def alterar_montagens_miolo(self):
        senha_configurada = GPRINT_SENHA_ADMIN_EQUIPAMENTOS
        if not senha_configurada.strip():
            messagebox.showwarning("GPrint - Miolo", "Configure a senha administrativa.", parent=self)
            return
        senha = simpledialog.askstring(
            "GPrint - Acesso administrativo", "Informe a senha para alterar as montagens simultâneas.",
            show="*", parent=self)
        if senha is None:
            return
        if not hmac.compare_digest(senha.encode("utf-8"), senha_configurada.encode("utf-8")):
            messagebox.showerror("GPrint - Acesso negado", "Senha incorreta.", parent=self)
            return
        quantidade = simpledialog.askinteger(
            "Montagens simultâneas", "Quantos PDFs de miolo podem ser montados ao mesmo tempo?",
            initialvalue=obter_montagens_simultaneas_miolo(), minvalue=1,
            maxvalue=MAX_MONTAGENS_SIMULTANEAS_MIOLO, parent=self)
        if quantidade is None:
            return
        novas_preferencias = dict(CONFIG, montagens_simultaneas_miolo=quantidade)
        if salvar_configuracoes(novas_preferencias):
            CONFIG.update(novas_preferencias)
            self.montagens_miolo_var.set(str(quantidade))
        else:
            messagebox.showerror("GPrint - Miolo", "Não foi possível salvar a preferência.", parent=self)

    def restaurar_padroes(self):
        global CONFIG
        
        if messagebox.askyesno(
            "Confirmar",
            "Deseja restaurar as preferências gerais para os valores padrão?"
        ):
            montagens = obter_montagens_simultaneas_miolo()
            CONFIG = dict(DEFAULT_CONFIG, montagens_simultaneas_miolo=montagens)
            self.carregar_configuracoes()
            salvar_configuracoes(CONFIG)
            messagebox.showinfo("Sucesso", "Configurações padrão restauradas!")           
def gprint_enviar_pdf_lpr(log, caminho_arquivo, ip, formato, job_name, copia_numero=1, total_copias=1):
    arquivo_temp = None
    try:
        tamanho_mib = os.path.getsize(caminho_arquivo) / (1024 ** 2)
        log(
            f"Enviando cópia {copia_numero}/{total_copias} | {tamanho_mib:.1f} MiB | "
            f"IP: {ip} | Fila: {formato} | Job: {job_name}", "ENVIO"
        )
        arquivo_temp = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4()}.pdf")
        shutil.copy2(caminho_arquivo, arquivo_temp)
        time.sleep(1)
        resultado = subprocess.run(
            ["lpr", "-S", ip, "-P", formato, "-J", job_name, arquivo_temp],
            capture_output=True, text=True, timeout=TIMEOUT_LPR_MIOLO_SEGUNDOS,
            errors="replace",
        )
        if resultado.returncode == 0:
            log(f"Cópia {copia_numero} aceita pelo LPR", "SUCESSO")
            return True, "", False
        detalhe = (resultado.stderr or resultado.stdout or "").strip()
        erro = f"LPR retornou código {resultado.returncode}"
        if detalhe:
            erro += f": {detalhe}"
    except subprocess.TimeoutExpired:
        erro = f"Timeout LPR de {TIMEOUT_LPR_MIOLO_SEGUNDOS}s; aceite pela fila não confirmado"
        log(erro, "ERRO")
        return False, erro, True
    except Exception as exc:
        erro = str(exc)
    finally:
        if arquivo_temp:
            try:
                os.remove(arquivo_temp)
            except FileNotFoundError:
                pass
            except OSError as exc:
                logging.warning("Não foi possível remover temporário LPR: %s", exc)
    log(erro, "ERRO")
    return False, erro, False

def gprint_registrar_pendencias_lpr(emitir_log, caminho, etapa, itens):
    try:
        with open(caminho, 'a', encoding='utf-8-sig') as arquivo:
            arquivo.write(f"\n{etapa} | {datetime.now():%d/%m/%Y %H:%M:%S}\n")
            for item in itens:
                log = item['log']
                arquivo.write(
                    f"\nSituação: {item['situacao']} | Envio {item['copia']}/{item['total_envios']}\n"
                    f"PDF montado: {item['pdf']}\n"
                    f"Origem: {item['processo']['TXT_ORIGEM']} | Linha {item['processo']['LINHA']}\n"
                    f"Impressora: {log['maquina']} | IP: {log['ip']} | Fila: {log['formato']}\n"
                    f"Job: {item['job']}\n"
                    f"Resultado: {log['erro'] or log['status']}\n"
                )
                if 'linhas' in item:
                    arquivo.write(f"Total de capas no PDF: {log['quantidade']}\n")
                    for linha in item['linhas']:
                        arquivo.write(f"  Linha {linha['LINHA']} | {linha['ARQUIVO']} | QTD {linha['QTD']}\n")
    except OSError as erro:
        emitir_log(f"Falha ao gravar o relatório de LPR em {caminho}: {erro}", "ERRO")
        return False
    emitir_log(f"Relatório de envio LPR: {caminho}")
    return True

def gprint_exportar_txt_pendentes(itens, subpasta, sessao, titulo, parent, emitir_log):
    por_txt = {}
    for processo, quantidade in itens:
        linha = processo['LINHA_ORIGINAL']
        if quantidade != processo['QTD']:
            campos = linha.split('|')
            campos[2] = str(quantidade)
            linha = '|'.join(campos)
        por_txt.setdefault(processo['TXT_ORIGEM'], []).append(linha)
    salvos, falhas = [], []
    for origem, linhas in por_txt.items():
        pasta = os.path.join(os.path.dirname(origem), subpasta, sessao)
        destino = os.path.join(pasta, os.path.basename(origem))
        temporario = None
        try:
            os.makedirs(pasta, exist_ok=True)
            with tempfile.NamedTemporaryFile(
                mode='w', dir=pasta, delete=False, encoding='utf-8-sig', newline=''
            ) as arquivo:
                temporario = arquivo.name
                arquivo.writelines(linhas)
            os.replace(temporario, destino)
        except OSError as erro:
            falhas.append(f"{origem}: {erro}")
            if temporario:
                try:
                    os.remove(temporario)
                except OSError:
                    logging.warning("Não foi possível remover o TXT temporário: %s", temporario)
        else:
            salvos.append(destino)
            emitir_log(f"TXT de miolo exportado: {destino}")
    if salvos:
        messagebox.showinfo(titulo, "\n".join(salvos), parent=parent)
    if falhas:
        messagebox.showerror("GPrint - Falha na exportação", "\n".join(falhas), parent=parent)

def gprint_montar_item_miolo(processo, pasta_temporaria, sangria_mm, comprimir):
    """Monta um PDF em pasta exclusiva, sem acessar Tkinter ou a impressora."""
    p = processo
    pasta_item = os.path.join(pasta_temporaria, f"TXT{p['TXT_NUMERO']:02d}_L{p['LINHA']}")
    os.mkdir(pasta_item)
    origem = os.path.normpath(os.path.join(p["CAMINHO"], p["ARQUIVO"]))
    local = os.path.join(pasta_item, f"TXT{p['TXT_NUMERO']:02d}_L{p['LINHA']}_{os.path.basename(p['ARQUIVO'])}")
    shutil.copy2(origem, local)
    saida = os.path.join(pasta_item, "montado.pdf")
    w_pt, h_pt = p["LARGURA_MM"] * MM, p["ALTURA_MM"] * MM
    modo = p["MODO"]

    if modo == "dois_na_folha":
        dois_na_folha(local, saida, w_pt, h_pt, espacamento_mm=p["ESPACAMENTO"], sangria_mm=sangria_mm)
    elif modo == "grampo":
        grampo(local, saida, w_pt, h_pt, sangria_mm=sangria_mm)
    elif modo == "Corta_empilha":
        Corta_empilha(local, saida, w_pt, h_pt, sangria_mm=sangria_mm, comprimir_entrada=comprimir)
    elif modo == "corta_empilha_rotacionado":
        corta_empilha_rotacionado(local, saida, w_pt, h_pt, sangria_mm=sangria_mm, comprimir_entrada=comprimir)
    elif modo == "dois_na_folha_rotacionado":
        dois_na_folha_rotacionado(local, saida, w_pt, h_pt, espacamento_mm=p["ESPACAMENTO"], sangria_mm=sangria_mm)
    elif modo == "um_na_folha_atexco":
        um_na_folha_atexco(local, saida, w_pt, h_pt, sangria_mm=sangria_mm, comprimir_entrada=comprimir)
    else:
        um_na_folha(local, saida, w_pt, h_pt, sangria_mm=sangria_mm)

    if not os.path.isfile(saida) or os.path.getsize(saida) == 0:
        raise RuntimeError("A montagem não gerou um PDF válido.")
    return saida


def gprint_criar_seletor_impressao(painel, impressora_var, envio_automatico_var):
    """Exibe na mesma linha a máquina (nome e IP) e o envio automático."""
    painel.columnconfigure(1, weight=1)
    ttk.Label(painel, text="Impressora:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    seletor = ttk.Combobox(painel, textvariable=impressora_var, values=[], state="readonly")
    seletor.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    ttk.Checkbutton(
        painel, text="Enviar para impressora automaticamente",
        variable=envio_automatico_var, bootstyle="success-round-toggle",
    ).grid(row=0, column=2, padx=5, pady=5, sticky="w")
    return seletor


class MioloTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.processos = []
        self.arquivos_txt = []  # arquivos selecionados para a proxima execucao
        self.processando = False
        self.thread = None
        self.cancelamento = threading.Event()
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

        self.impressao_auto_var = tk.BooleanVar(value=True)
        self.combo_impressora = gprint_criar_seletor_impressao(
            config_frame, self.impressora_var, self.impressao_auto_var)

        opcoes_miolo = ttk.Frame(config_frame)
        opcoes_miolo.grid(row=1, column=0, columnspan=3, sticky="w")

        self.comprimir_pdf_var = tk.BooleanVar(value=False)
        self.checkbox_comprimir = ttk.Checkbutton(
            opcoes_miolo,
            text="Comprimir PDF antes do envio",
            variable=self.comprimir_pdf_var,
            bootstyle="info-square-toggle"
        )
        self.checkbox_comprimir.pack(side="left", padx=(5, 16), pady=5)

        self.enviar_uma_copia_var = tk.BooleanVar(value=False)
        self.chk_enviar_uma = ttk.Checkbutton(
            opcoes_miolo,
            text="Enviar apenas uma cópia",
            variable=self.enviar_uma_copia_var,
            bootstyle="info-square-toggle"
        )
        self.chk_enviar_uma.pack(side="left", padx=(0, 5), pady=5)

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
        self.btn_cancelar = ttk.Button(
            status_frame, text="Cancelar", command=self.cancelar_processamento,
            bootstyle="danger-outline", state="disabled"
        )
        self.btn_cancelar.grid(row=0, column=2, padx=(8, 0), sticky="e")

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
        if not self.processos:
            messagebox.showwarning("Aviso", "Carregue um arquivo TXT primeiro.")
            return
        if self.processando:
            return

        # Captura os valores de Tkinter antes de iniciar o trabalho em segundo plano.
        self.equipamento_execucao = equipamento.copy()
        opcoes = {
            'sangria': obter_sangria_padrao_miolo(),
            'paralelo': obter_montagens_simultaneas_miolo(),
            'comprimir': self.comprimir_pdf_var.get(),
            'automatico': self.impressao_auto_var.get(),
            'uma_copia': self.enviar_uma_copia_var.get(),
            'impressora': self.impressora_var.get(),
            'ip': self.equipamento_execucao['ip'],
        }
        self.cancelamento.clear()
        self.processando = True
        self.btn_processar.config(state='disabled', text='⏳ Processando...')
        self.btn_select.config(state='disabled')
        self.btn_cancelar.config(state='normal')
        self.status_var.set(f"Processando miolo | até {opcoes['paralelo']} montagem(ns) simultânea(s)")
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.processos)
        self.progress_label.config(text=f"0/{len(self.processos)}")
        self.indicadores_miolo.iniciar()

        self.thread = threading.Thread(target=self._executar, args=(opcoes,), daemon=True)
        self.thread.start()

    def cancelar_processamento(self):
        if self.processando and not self.cancelamento.is_set():
            self.cancelamento.set()
            self.btn_cancelar.config(state='disabled')
            self.status_var.set("Cancelando… aguardando as operações em andamento")
            self.log("Cancelamento solicitado. Aguardando as operações em andamento.")

    def enviar_por_lpr(self, caminho_arquivo, ip, formato, job_name, copia_numero, total_copias):
        return gprint_enviar_pdf_lpr(
            self.log, caminho_arquivo, ip, formato, job_name, copia_numero, total_copias)

    @staticmethod
    def _gerar_resumo_final_miolo(dados):
        """O log e a janela final mostram os mesmos totais desta execução."""
        d = dados
        situacao = ('Cancelado' if d['cancelado'] else
                    'Concluído com falhas' if d['falhas'] else 'Concluído')
        horas, resto = divmod(d['duracao'], 3600)
        minutos, segundos = divmod(resto, 60)
        nao_iniciados = max(0, d['total'] - d['iniciados'])
        tentativas = d['envios'] + d['erros_envio']
        linhas = [
            '=' * 56,
            f"RESUMO FINAL DO MIOLO — {situacao.upper()}",
            '=' * 56,
            f"Início: {d['inicio']:%d/%m/%Y %H:%M:%S}",
            f"Término: {d['fim']:%d/%m/%Y %H:%M:%S}",
            f"Tempo total: {horas:02d}:{minutos:02d}:{segundos:02d}",
            f"Operador Windows: {d['usuario']}",
            f"Equipamento: {d['maquina']} | IP: {d['ip']}",
            f"Modo: {'Envio automático' if d['automatico'] else 'Salvar cópias, sem envio'}",
            f"Apenas uma cópia: {'Sim' if d['uma_copia'] else 'Não'} | Sangria: {d['sangria']:g} mm",
            '-' * 56,
            'ARQUIVOS E MONTAGENS',
            f"TXT utilizados: {d['txt']}",
            f"Arquivos na lista: {d['total']} | Iniciados: {d['iniciados']}",
            f"Arquivos não iniciados: {nao_iniciados}",
            f"PDFs montados com sucesso: {d['montados']}",
            f"Arquivos que chegaram ao fim do fluxo: {d['concluidos']}",
            f"Arquivos com erro no processamento: {d['erros']}",
            'Montagens dos arquivos iniciados: ' + (', '.join(
                f'{modo}: {qtd}' for modo, qtd in sorted(d['modos'].items())) or 'Nenhuma'),
            'Filas dos arquivos iniciados: ' + (', '.join(d['filas']) or 'Nenhuma'),
            '-' * 56,
            'ENVIO E SALVAMENTO',
            f"Tentativas de envio LPR: {tentativas}",
            f"Envios aceitos pelo LPR: {d['envios']} | Falhas de envio: {d['erros_envio']}",
            f"Cópias salvas: {d['salvos']} | Falhas ao salvar: {d['erros_salvar']}",
            'Aceite pelo LPR não confirma a impressão física.',
            '-' * 56,
            f"Pasta dos PDFs montados: {d['pasta']}",
        ]
        if d['pastas_salvas']:
            linhas.append('Pastas das cópias salvas:')
            linhas.extend(f'  {pasta}' for pasta in d['pastas_salvas'])
        if d['pendencias_lpr']:
            reenvios = sum(item['situacao'] == 'reenviar' for item in d['pendencias_lpr'])
            linhas.append(
                f"Envios para reenvio: {reenvios} | "
                f"Envios para conferência: {len(d['pendencias_lpr']) - reenvios}"
            )
        if d['cancelado']:
            linhas.append(f"Linhas restantes para retomar: {d['pendentes']}")
        if d['falhas']:
            linhas.extend(['-' * 56, f"FALHAS REGISTRADAS: {len(d['falhas'])}"])
            linhas.extend(f'  {falha}' for falha in d['falhas'][:10])
            if len(d['falhas']) > 10:
                linhas.append(f"  Mais {len(d['falhas']) - 10} falha(s); consulte o log detalhado acima.")
        linhas.append('=' * 56)
        return '\n'.join(linhas), situacao

    def _executar(self, opcoes):
        try:
            dados, pendentes = self._executar_lote_miolo(opcoes)
        except Exception as erro:
            logging.exception("Falha na execução de miolo")
            self.after(0, self._finalizar_miolo, None, [], str(erro))
        else:
            self.after(0, self._finalizar_miolo, dados, pendentes)

    def _executar_lote_miolo(self, opcoes):
        inicio_execucao = datetime.now()
        inicio_cronometro = time.monotonic()
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
        sangria_mm = opcoes['sangria']
        paralelo = min(opcoes['paralelo'], len(self.processos))
        self.log(f"Sangria: {sangria_mm:g} mm | Montagens simultâneas: {paralelo}")

        success_count = 0
        error_count = 0
        envio_count = 0
        envio_erro_count = 0
        salvos_count = 0
        salvos_erro_count = 0
        montados_count = 0
        iniciados_count = 0
        falhas_resumo = []
        pastas_salvas = set()
        txt_processados = set()
        filas_processadas = set()
        modos_processados = {}
        envio_automatico = opcoes['automatico']
        somente_uma_copia = opcoes['uma_copia']
        impressora_selecionada = opcoes['impressora']
        ip_extraido = opcoes['ip']

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

        varios_txt = len(self.arquivos_txt) > 1
        retomada = 0
        pendencias_lpr = []
        sessao_temp = tempfile.mkdtemp(prefix="sessao_miolo_", dir=PASTA_TRABALHO)
        executor = None
        futuros = {}
        proximo = 0
        try:
            if paralelo > 1 and not self.cancelamento.is_set():
                executor = ProcessPoolExecutor(
                    max_workers=paralelo, mp_context=multiprocessing.get_context("spawn"))
            for i, p in enumerate(self.processos):
                if self.cancelamento.is_set():
                    break

                iniciados_count += 1
                etapa_atual = 'Preparação'
                try:
                    origem_txt = p.get("TXT_ORIGEM", getattr(self, "arquivo_txt", ""))
                    lote = os.path.dirname(origem_txt) if origem_txt else "desconhecido"
                    pilha = os.path.splitext(os.path.basename(origem_txt))[0] if origem_txt else "desconhecido"
                    numero_txt = p.get("TXT_NUMERO", 1)
                    txt_processados.add(origem_txt or 'TXT não informado')
                    filas_processadas.add(p['FORMATO'])
                    modos_processados[p['MODO']] = modos_processados.get(p['MODO'], 0) + 1
                    self.after(0, self._atualizar_progresso_miolo, i, len(self.processos))
                
                    self.log(f"\n--- TXT {numero_txt}: {os.path.basename(origem_txt)} | L{p['LINHA']}: {p['ARQUIVO']} ---", "INFO")
                    self.log(f"📄 Formato: {p['FORMATO']}", "INFO")
                    self.log(f"🔧 Modo: {p['MODO']}", "INFO")
                    self.log(f"📊 QTD para envio: {p['QTD']}, QTD Total: {p['QTD_total']}", "INFO")
                
                    etapa_atual = 'Montagem'
                    if executor is None:
                        montado = gprint_montar_item_miolo(p, sessao_temp, sangria_mm, opcoes['comprimir'])
                    else:
                        # Limita também os PDFs prontos aguardando sua vez no envio.
                        while (proximo < len(self.processos)
                               and len(futuros) < paralelo
                               and not self.cancelamento.is_set()):
                            futuros[proximo] = executor.submit(
                                gprint_montar_item_miolo, self.processos[proximo],
                                sessao_temp, sangria_mm, opcoes['comprimir'])
                            proximo += 1
                        if self.cancelamento.is_set():
                            break
                        futuro = futuros.pop(i)
                        while not self.cancelamento.is_set():
                            try:
                                montado = futuro.result(timeout=0.1)
                                break
                            except FutureTimeout:
                                if futuro.done():
                                    montado = futuro.result()
                                    break

                    if self.cancelamento.is_set():
                        break
                    retomada = i + 1
                    # Isola as saídas de TXT diferentes, mesmo com NOMEFIC iguais.
                    pasta_saida_txt = os.path.join(pasta_lote, f"TXT_{numero_txt:02d}") if varios_txt else pasta_lote
                    os.makedirs(pasta_saida_txt, exist_ok=True)
                    pasta_salvos = os.path.join(pasta_saida_txt, "COPIAS_SALVAS")
                    output_path = os.path.join(pasta_saida_txt, f"{p['NOMEFIC']}.pdf")
                    if os.path.exists(output_path):
                        # Mantém o PDF de cada linha disponível para eventual reenvio.
                        saida_linha = os.path.join(pasta_saida_txt, f"L{p['LINHA']}_{uuid.uuid4().hex[:6]}")
                        os.makedirs(saida_linha)
                        output_path = os.path.join(saida_linha, os.path.basename(output_path))
                        pasta_salvos = os.path.join(saida_linha, "COPIAS_SALVAS")
                    shutil.move(montado, output_path)
                    shutil.rmtree(os.path.dirname(montado), ignore_errors=True)
                    montados_count += 1
                    etapa_atual = 'Finalização/envio'
                    self.log(f"✅ PDF gerado: {os.path.basename(output_path)}", "SUCESSO")
                
                    # Quantidade de envios conforme o modo de montagem
                    job_unico_por_montagem = (
                        p["MODO"] == "dois_na_folha"
                        or (p["MODO"] == "um_na_folha_atexco" and p["QTD"] >= 10)
                    )
                    if job_unico_por_montagem:
                        qtd_envios = 1
                        self.log(f"🔢 Modo 1 envio/: {p['QTD']} exemplares -> 1 envio", "INFO")
                    else:
                        qtd_envios = p["QTD"]
                        self.log(f"🔢 Modo {p['MODO']}: {p['QTD']} exemplares -> {qtd_envios} envios", "INFO")

                    # Forçar apenas uma cópia caso a opção esteja marcada
                    if somente_uma_copia:
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
                
                    copias_a_conferir = 0
                    # Loop de envios (ou salvamentos)
                    for envio_idx in range(qtd_envios):
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

                        if envio_automatico:
                            self.log("🖨️ Impressão automática ativada", "INFO")
                            sucesso, erro, timeout_lpr = self.enviar_por_lpr(
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
                                status = "Timeout no envio" if timeout_lpr else "Falha no envio"
                                if timeout_lpr:
                                    # Inclui a cópia de aceite incerto e as que ficaram sem tentativa.
                                    copias_a_conferir = qtd_envios - envio_idx
                                falhas_resumo.append(
                                    f"TXT {numero_txt} | {p['ARQUIVO']} | envio {copia_num}/{qtd_envios}: {erro}")
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
                                pastas_salvas.add(pasta_salvos)
                                sucesso = True
                                status = "Salvo"
                                erro = ""
                            except Exception as e:
                                self.log(f"❌ Erro ao salvar: {e}", "ERRO")
                                sucesso = False
                                status = "Falha ao salvar"
                                erro = str(e)
                                salvos_erro_count += 1
                                falhas_resumo.append(
                                    f"TXT {numero_txt} | {p['ARQUIVO']} | cópia {copia_num}: {erro}")
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
                        if envio_automatico and not sucesso:
                            pendencias_lpr.append({
                                'processo': p, 'pdf': output_path, 'job': job_name,
                                'copia': copia_num, 'total_envios': qtd_envios,
                                'job_unico': job_unico_por_montagem, 'log': dados_log,
                                'situacao': 'conferir' if timeout_lpr else 'reenviar',
                            })
                            if timeout_lpr:
                                for numero in range(copia_num + 1, qtd_envios + 1):
                                    pendencias_lpr.append(dict(
                                        pendencias_lpr[-1], copia=numero, situacao='aguardando',
                                        log=dict(dados_log, copia=f"{numero}/{qtd_envios}",
                                                 erro="Envio não iniciado após timeout da cópia anterior.")
                                    ))


                        if copias_a_conferir:
                            self.log(
                                f"{copias_a_conferir - 1} cópia(s) restante(s) sem tentativa. "
                                "Confira a cópia com timeout na fila antes de reenviar.", "ERRO"
                            )
                            break
                        # Mantém a pausa entre envios mesmo ao terminar uma linha cancelada.
                        if envio_automatico:
                            time.sleep(1.5)
                        else:
                            self.cancelamento.wait(1.5)

                    success_count += 1
               
                except Exception as e:
                    if self.cancelamento.is_set() and retomada <= i:
                        break
                    retomada = i + 1
                    self.log(f"❌ Erro: {str(e)}", "ERRO")
                    error_count += 1
                    falhas_resumo.append(
                        f"TXT {p.get('TXT_NUMERO', 1)} | {p.get('ARQUIVO', '?')} | {etapa_atual}: {e}")
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
            if executor is not None:
                executor.shutdown(wait=True, cancel_futures=True)
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

        pendentes = self.processos[retomada:] if self.cancelamento.is_set() else []
        cancelado = bool(pendentes)
        fim_execucao = datetime.now()
        duracao = max(0, int(time.monotonic() - inicio_cronometro))
        resumo_dados = {
            'cancelado': cancelado, 'inicio': inicio_execucao, 'fim': fim_execucao,
            'duracao': duracao, 'total': len(self.processos), 'iniciados': iniciados_count,
            'montados': montados_count, 'concluidos': success_count, 'erros': error_count,
            'envios': envio_count, 'erros_envio': envio_erro_count,
            'salvos': salvos_count, 'erros_salvar': salvos_erro_count,
            'txt': len(txt_processados), 'filas': sorted(filas_processadas),
            'modos': modos_processados, 'maquina': impressora_selecionada,
            'ip': ip_extraido, 'usuario': usuario, 'sangria': sangria_mm,
            'automatico': envio_automatico, 'uma_copia': somente_uma_copia,
            'pasta': pasta_lote, 'pastas_salvas': sorted(pastas_salvas),
            'falhas': falhas_resumo, 'pendentes': len(pendentes),
            'pendencias_lpr': pendencias_lpr,
        }
        return resumo_dados, pendentes

    def _atualizar_progresso_miolo(self, feitos, total):
        self.progress['value'] = feitos
        self.progress_label.config(text=f"{feitos}/{total}")

    def _finalizar_miolo(self, dados, pendentes, erro=None):
        self.indicadores_miolo.finalizar()
        self.btn_processar.config(state='normal', text='▶ Processar e Enviar')
        self.btn_select.config(state='normal')
        self.btn_cancelar.config(state='disabled')
        self.processando = False
        if erro is not None:
            self.status_var.set("Falha no processamento do miolo")
            messagebox.showerror("GPrint - Miolo", erro, parent=self)
            return

        resumo, situacao = self._gerar_resumo_final_miolo(dados)
        for linha in resumo.splitlines():
            self.log(linha, 'ERRO' if dados['falhas'] else 'INFO')
        self._atualizar_progresso_miolo(dados['total'] - len(pendentes), dados['total'])
        self.status_var.set(
            f"{situacao}: {dados['montados']}/{dados['total']} montados | "
            f"LPR: {dados['envios']} OK / {dados['erros_envio']} falhas | Salvos: {dados['salvos']}"
        )
        sessao = os.path.basename(dados['pasta'])
        if pendentes:
            # O botão Processar também retoma apenas as linhas restantes nesta sessão.
            self.processos = pendentes
            self._atualizar_lista_miolo()
            if messagebox.askyesno(
                "GPrint - Miolo cancelado",
                f"Restam {len(pendentes)} linha(s) na fila. Deseja exportar os TXTs para a pasta cancelados?",
                parent=self,
            ):
                self._exportar_miolo_cancelado(pendentes, sessao)

        exibir = messagebox.showwarning if dados['cancelado'] or dados['falhas'] else messagebox.showinfo
        exibir(f'GPrint - Miolo: {situacao}', resumo, parent=self)

        itens_lpr = dados['pendencias_lpr']
        if itens_lpr:
            relatorio = os.path.join(dados['pasta'], 'erros_envio_lpr.txt')
            gravado = gprint_registrar_pendencias_lpr(self.log, relatorio, "Envio original", itens_lpr)
            reenviaveis = sum(item['situacao'] == 'reenviar' for item in itens_lpr)
            mensagem = (
                f"{reenviaveis} envio(s) com erro podem ser repetidos usando os PDFs já montados.\n"
                f"{len(itens_lpr) - reenviaveis} envio(s) precisam de conferência por causa de timeout.\n\n"
                + (f"Detalhes: {relatorio}" if gravado else "Falha ao salvar os detalhes. Consulte o log.")
            )
            if reenviaveis and messagebox.askyesno(
                "GPrint - Reenviar miolos?", mensagem + "\n\nDeseja reenviar agora os envios com erro?",
                parent=self,
            ):
                self.cancelamento.clear()
                self.processando = True
                self.btn_processar.config(state='disabled')
                self.btn_select.config(state='disabled')
                self.btn_cancelar.config(state='normal')
                self.status_var.set("Reenviando PDFs de miolo já montados…")
                self.thread = threading.Thread(
                    target=self._reenviar_miolo_lpr, args=(itens_lpr, relatorio, sessao), daemon=True)
                self.thread.start()
            else:
                self._exportar_pendencias_lpr(itens_lpr, sessao)

    def _reenviar_miolo_lpr(self, itens, relatorio, sessao):
        bloqueados = set()
        aceitos = 0
        try:
            for item in itens:
                if self.cancelamento.is_set():
                    break
                if item['situacao'] != 'reenviar':
                    continue
                log = item['log']
                if item['pdf'] in bloqueados:
                    item['situacao'] = 'aguardando'
                    item['log'] = dict(log, erro="Reenvio adiado após timeout de outra cópia deste PDF.")
                    gprint_registrar_pendencias_lpr(self.log, relatorio, "Reenvio adiado", [item])
                    continue
                sucesso, erro, timeout = self.enviar_por_lpr(
                    item['pdf'], log['ip'], log['formato'], item['job'], item['copia'], item['total_envios']
                )
                item['situacao'] = 'aceito' if sucesso else 'conferir' if timeout else 'reenviar'
                status = 'Enviado' if sucesso else 'Timeout no reenvio' if timeout else 'Falha no reenvio'
                item['log'] = dict(log, status=status, erro=erro, envio_ok=int(sucesso))
                gprint_registrar_pendencias_lpr(self.log, relatorio, "Tentativa de reenvio", [item])
                self.registrar_log_sqlite(item['log'])
                aceitos += int(sucesso)
                if timeout:
                    bloqueados.add(item['pdf'])
                self.cancelamento.wait(1.5)
        except Exception:
            logging.exception("Falha ao finalizar o reenvio de miolo")
        finally:
            gprint_sincronizar_pendentes(
                self.log_db_path, root=self.app.root, emitir_log=self.log)
            self.after(0, self._finalizar_reenvio_miolo, itens, sessao, aceitos)

    def _finalizar_reenvio_miolo(self, itens, sessao, aceitos):
        self.processando = False
        self.btn_processar.config(state='normal')
        self.btn_select.config(state='normal')
        self.btn_cancelar.config(state='disabled')
        restantes = sum(item['situacao'] != 'aceito' for item in itens)
        self.status_var.set(f"Reenvio LPR: {aceitos} aceito(s) | {restantes} pendente(s)")
        self._exportar_pendencias_lpr(itens, sessao)
        messagebox.showinfo("GPrint - Reenvio concluído", self.status_var.get(), parent=self)

    def _exportar_pendencias_lpr(self, itens, sessao):
        por_destino = {'reenvio_lpr': {}, 'conferir_lpr': {}}
        for item in itens:
            if item['situacao'] == 'aceito':
                continue
            subpasta = 'reenvio_lpr' if item['situacao'] == 'reenviar' else 'conferir_lpr'
            p = item['processo']
            chave = (p['TXT_ORIGEM'], p['LINHA'])
            contagens = por_destino[subpasta]
            quantidade = p['QTD'] if item['job_unico'] else contagens.get(chave, (p, 0))[1] + 1
            contagens[chave] = (p, quantidade)
        titulos = {
            'reenvio_lpr': "TXTs com cópias que ainda precisam de reenvio",
            'conferir_lpr': "Confira as cópias com timeout na fila antes de usar estes TXTs",
        }
        for subpasta, contagens in por_destino.items():
            if contagens:
                self._exportar_miolo_txt(contagens.values(), subpasta, sessao, titulos[subpasta])

    def _exportar_miolo_cancelado(self, pendentes, sessao):
        self._exportar_miolo_txt(
            [(p, p['QTD']) for p in pendentes], 'cancelados', sessao, "TXTs pendentes exportados"
        )

    def _exportar_miolo_txt(self, itens, subpasta, sessao, titulo):
        gprint_exportar_txt_pendentes(itens, subpasta, sessao, titulo, self, self.log)

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
                        with open(caminho_txt, 'r', encoding=enc, newline='') as f:
                            lines = f.readlines()
                        self.log(f"TXT {numero_txt}: {os.path.basename(caminho_txt)} (encoding: {enc})", "INFO")
                        break
                    except UnicodeDecodeError:
                        continue
                if lines is None:
                    raise ValueError(f"Nao foi possivel identificar a codificacao do TXT: {caminho_txt}")

                for idx, linha_original in enumerate(lines):
                    linha = linha_original.strip()
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
                    processo["LINHA_ORIGINAL"] = linha_original
                    novos_processos.append(processo)


            # Confirmacao atomica da nova fila de trabalhos: nao deixa selecao parcial.
            self.processos = novos_processos
            self.arquivos_txt = arquivos
            self.arquivo_txt = arquivos[0]  # compatibilidade com fluxo de TXT unico
            self.nome_base_txt = os.path.splitext(os.path.basename(arquivos[0]))[0]

            self._atualizar_lista_miolo()
            quantidade = len(arquivos)
            self.status_var.set(f"✅ {len(self.processos)} processos de {quantidade} TXT carregados")
            self.log(f"✅ {quantidade} TXT carregado(s), total {len(self.processos)} processos", "SUCESSO")
            for numero, caminho in enumerate(arquivos, start=1):
                self.log(f"   TXT {numero}: {caminho}", "INFO")
            return True
        except Exception as e:
            messagebox.showerror("Erro ao carregar TXT", f"Nao foi possivel carregar todos os TXT:\n{e}")
            self.log(f"ERRO ao carregar TXT: {e}", "ERRO")
            return False

    def _atualizar_lista_miolo(self):
        self.tree.delete(*self.tree.get_children())
        for processo in self.processos:
            partes = processo['FORMATO'].split('_')
            self.tree.insert('', 'end', values=(
                processo['TXT_NUMERO'],
                f"L{processo['LINHA']}",
                os.path.basename(processo['ARQUIVO']),
                partes[min(1, len(partes) - 1)].upper(),
                processo['CODIGOACAB'],
                processo['MODO'],
                processo['QTD'],
                processo['QTD_total'],
            ))
        arquivos = list(dict.fromkeys(p['TXT_ORIGEM'] for p in self.processos))
        self.file_label.config(text=(
            os.path.basename(arquivos[0]) if len(arquivos) == 1
            else f"{len(arquivos)} TXT selecionados | {len(self.processos)} processos"
        ))

def inserir_recorte_pdf_espelhado(origem, saida, recorte, destino, rotate=0,
                                  flip_horizontal=False, flip_vertical=False):
    """Copia conteúdo PDF vetorial e reflete a matriz, sem rasterizar as cores."""
    existentes = {
        item[0] for item in saida.get_xobjects()
        if len(item) >= 3 and item[2] == 0
    }
    saida.show_pdf_page(
        destino,
        origem.parent,
        origem.number,
        clip=recorte,
        rotate=rotate % 360,
        overlay=True,
    )

    novos = [
        item for item in saida.get_xobjects()
        if len(item) >= 4 and item[2] == 0 and item[0] not in existentes
    ]
    if not novos:
        raise RuntimeError("Não foi possível localizar o recorte vetorial da sangria.")

    xref = novos[-1][0]
    tipo, matriz_txt = saida.parent.xref_get_key(xref, "Matrix")
    numeros = [float(n) for n in re.findall(r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?", matriz_txt)]
    if tipo != "array" or len(numeros) != 6:
        raise RuntimeError("A matriz do recorte vetorial está em formato inesperado.")

    a, b, c, d, e, f = numeros
    if flip_horizontal:
        # Reflexão no eixo vertical central do retângulo de destino.
        a, c, e = -a, -c, destino.x0 + destino.x1 - e
    if flip_vertical:
        # Reflexão no eixo horizontal central; converte Y da página para PDF.
        b, d, f = -b, -d, 2 * saida.rect.height - destino.y0 - destino.y1 - f

    matriz = "[{} {} {} {} {} {}]".format(*(f"{v:.10g}" for v in (a, b, c, d, e, f)))
    saida.parent.xref_set_key(xref, "Matrix", matriz)



def aplicar_sangria_espelhada_externa(page, nova, trim_original, destino, sangria_pt, bordas, rotate=0):
    """
    Cria a sangria espelhada copiando os vetores e imagens originais do PDF.
    Assim, não renderiza os recortes em RGB/CMYK nem recalcula suas separações.
    """
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
        elif borda_orig == 'bottom':
            rect = fitz.Rect(trim_original.x0, trim_original.y1 - sangria_pt,
                             trim_original.x1, trim_original.y1)
        elif borda_orig == 'left':
            rect = fitz.Rect(trim_original.x0, trim_original.y0,
                             trim_original.x0 + sangria_pt, trim_original.y1)
        elif borda_orig == 'right':
            rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y0,
                             trim_original.x1, trim_original.y1)
        else:
            continue

        if borda_dest == 'top':
            dest_rect = fitz.Rect(destino.x0, destino.y0 - sangria_pt, destino.x1, destino.y0)
            flip_h, flip_v = False, True
        elif borda_dest == 'bottom':
            dest_rect = fitz.Rect(destino.x0, destino.y1, destino.x1, destino.y1 + sangria_pt)
            flip_h, flip_v = False, True
        elif borda_dest == 'left':
            dest_rect = fitz.Rect(destino.x0 - sangria_pt, destino.y0, destino.x0, destino.y1)
            flip_h, flip_v = True, False
        elif borda_dest == 'right':
            dest_rect = fitz.Rect(destino.x1, destino.y0, destino.x1 + sangria_pt, destino.y1)
            flip_h, flip_v = True, False
        else:
            continue

        inserir_recorte_pdf_espelhado(
            page, nova, rect, dest_rect, rotate=rotate,
            flip_horizontal=flip_h, flip_vertical=flip_v,
        )

    cantos = [
        ('top', 'left'),
        ('top', 'right'),
        ('bottom', 'left'),
        ('bottom', 'right')
    ]
    for b1, b2 in cantos:
        if b1 in bordas and b2 in bordas:
            borda_orig1 = mapa[b1]
            borda_orig2 = mapa[b2]
            lados_originais = {borda_orig1, borda_orig2}
            x0 = trim_original.x0 if 'left' in lados_originais else trim_original.x1 - sangria_pt
            x1 = x0 + sangria_pt
            y0 = trim_original.y0 if 'top' in lados_originais else trim_original.y1 - sangria_pt
            y1 = y0 + sangria_pt
            rect = fitz.Rect(x0, y0, x1, y1)

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

            inserir_recorte_pdf_espelhado(
                page, nova, rect, dest_rect, rotate=rotate,
                flip_horizontal=True, flip_vertical=True,
            )


def desenhar_marcas(page, trim, tam, gap=2*MM, lado=None):
    """Desenha marcas de corte vetoriais ao redor do retângulo de corte."""
    x0, y0, x1, y1 = trim
    shape = page.new_shape()
    if lado in (None, 'esquerda'):
        shape.draw_line((x0, y0 - tam), (x0, y0 - gap))
        shape.draw_line((x0 - tam, y0), (x0 - gap, y0))
        shape.draw_line((x0, y1 + tam), (x0, y1 + gap))
        shape.draw_line((x0 - tam, y1), (x0 - gap, y1))
    if lado in (None, 'direita'):
        shape.draw_line((x1, y0 - tam), (x1, y0 - gap))
        shape.draw_line((x1 + tam, y0), (x1 + gap, y0))
        shape.draw_line((x1, y1 + tam), (x1, y1 + gap))
        shape.draw_line((x1 + tam, y1), (x1 + gap, y1))
    shape.finish(width=0.5)
    shape.commit()


def desenhar_marca_central_vertical(page, x, y0, y1, tam=5*MM, gap=2*MM):
    """Desenha as marcas de dobra no centro, acima e abaixo da página."""
    shape = page.new_shape()
    shape.draw_line((x, y0 - tam), (x, y0 - gap))
    shape.draw_line((x, y1 + tam), (x, y1 + gap))
    shape.finish(width=0.5)
    shape.commit()


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
def gprint_mesclar_capas(itens, destino):
    """Repete cada documento inteiro para preservar a sequência frente/verso."""
    with fitz.open() as agrupado:
        for arquivo, quantidade in itens:
            with fitz.open(arquivo) as capa:
                for copia in range(quantidade):
                    agrupado.insert_pdf(capa, final=copia == quantidade - 1)
        agrupado.save(destino)


def gprint_reservar_destino_capa(pasta_destino, identific, fila, quantidade):
    """Reserva um nome para o PDF final sem sobrescrever arquivos existentes.

    Usa IDENTIFIC_CAPAS_QUANTIDADE_FILA.pdf e acrescenta _1, _2... em caso de colisão.
    A reserva O_EXCL evita que duas estacoes escolham o mesmo destino.
    """
    nome_base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(identific)).strip(' .') or 'sem_nome'
    nome_fila = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', fila).strip(' .')

    # Guarda margem para extensao e sufixo, considerando caminhos do Windows.
    limite = min(190, max(32, 240 - len(os.path.abspath(pasta_destino))))
    sufixo_capas = f'_CAPAS_{quantidade}'
    limite_fila = max(0, limite - len(sufixo_capas) - 8)
    sufixo_fila = f'_{nome_fila[:limite_fila]}' if nome_fila and limite_fila else ''
    total_lote = re.search(
        r'(?<![A-Za-z0-9])TOTAL[ _-]+[0-9]+(?![A-Za-z0-9])',
        nome_base, re.IGNORECASE)

    for numero in range(0, 10001):
        sufixo = sufixo_capas + sufixo_fila
        if numero:
            sufixo += f'_{numero}'
        espaco = limite - len(sufixo)
        identificacao = nome_base[:espaco]
        if total_lote and total_lote.end() > espaco and len(total_lote[0]) <= espaco:
            prefixo = nome_base[:min(total_lote.start(), espaco - len(total_lote[0]) - 1)].rstrip('_ .-')
            identificacao = f'{prefixo}_{total_lote[0]}' if prefixo else total_lote[0]
        nome = f"{identificacao}{sufixo}.pdf"
        caminho = os.path.join(pasta_destino, nome)
        try:
            descritor = os.open(caminho, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o666)
        except FileExistsError:
            continue
        else:
            os.close(descritor)
            return caminho
    raise RuntimeError('Nao foi possivel reservar um nome livre para a capa mesclada')


def gprint_aplicar_icone(root):
    pasta_recursos = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    caminhos = (
        os.path.join(pasta_recursos, 'GPrint.ico'),
        r"Q:\___LISTA DE PRIORIDADES\GUILHERME\GPrint.ico",
    )
    for caminho in caminhos:
        try:
            if os.path.isfile(caminho):
                root.iconbitmap(default=caminho)
                return
        except (OSError, tk.TclError) as erro:
            logging.warning("Falha ao carregar o ícone do GPrint: %s", erro)
    logging.warning("GPrint.ico indisponível para a janela do GPrint.")


class Application:
    def __init__(self, root):
        self.root = root
        self.arquivos_selecionados = []
        self.imprimir_automaticamente = tk.BooleanVar(value=True)
        self.tema_atual = tk.StringVar(value=CONFIG.get('tema', 'darkly'))
        self.style = Style(theme=self.tema_atual.get())
        self.repositorio_equipamentos = GPrintEquipamentosRepository()
        self.equipamentos_cache = []
        self.logo_path = r"Q:\___LISTA DE PRIORIDADES\GUILHERME\GPrint.png"
        
        # Controle de progresso
        self.processamento_em_andamento = False
        self.cancelamento_capa = threading.Event()
        self.tempo_inicio = 0
        self.log_queue = queue.Queue()
        
        self.setup_ui()
        self.root.after(600, self.iniciar_cadastro_equipamentos)
        self.setup_drag_and_drop()
        self.atualizar_log_queue()

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
        self.root.minsize(800, 600)
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
        self.aba_refazer.atualizar_setores_refazer()
        if not equipamentos:
            self.notebook.select(self.aba_config)
            messagebox.showinfo('Configuracao inicial do GPrint',
                'Nenhum equipamento ativo foi encontrado.\n\n'
                'Acesse Configurações > Administração central > Equipamentos e filas, entre com a senha '
                'administrativa e cadastre as impressoras e suas filas.\n'
                'O cadastro sera compartilhado pelas estacoes.', parent=self.root)
            self.aba_config.mostrar_equipamentos()

    def _ao_mudar_aba_configuracoes(self, _event=None):
        aba_atual = self.notebook.select()
        if aba_atual != str(self.aba_config):
            # Sair de Configuracoes revoga o acesso; novo acesso pede senha.
            self.aba_config.bloquear_cadastro()
        if aba_atual == str(self.aba_refazer):
            self.aba_refazer.atualizar_destinos_refazer()
            self.aba_refazer.atualizar_operadores_refazer()
            self.aba_refazer.atualizar_setores_refazer()
            self.aba_refazer.carregar_clientes()
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
        config_frame = ttk.LabelFrame(main_frame, text="⚙️ Configurações de Impressão", bootstyle=PRIMARY)
        config_frame.grid(row=0, column=0, sticky="ew", pady=5)
        self.equipamento_capa_var = tk.StringVar()
        self.equipamentos_capa = {}
        self.combo_equipamento_capa = gprint_criar_seletor_impressao(
            config_frame, self.equipamento_capa_var, self.imprimir_automaticamente)

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
        self.btn_cancelar_capa = ttk.Button(
            status_frame, text="Cancelar", command=self.cancelar_capa,
            bootstyle="danger-outline", state=DISABLED)
        self.btn_cancelar_capa.grid(row=0, column=2, padx=(8, 0))

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
        if self.processamento_em_andamento:
            return
        txt_files = list(dict.fromkeys(
            os.path.abspath(f) for f in file_paths if f.lower().endswith('.txt')))
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
        if self.processamento_em_andamento or not verificar_data_limite():
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
        
        self.cancelamento_capa.clear()
        self.btn_cancelar_capa.config(state=NORMAL)
        arquivos = tuple(self.arquivos_selecionados)
        automatico = bool(self.imprimir_automaticamente.get())
        thread = threading.Thread(target=self.processar_arquivos, args=(arquivos, automatico), daemon=True)
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
    
    def cancelar_capa(self):
        if self.processamento_em_andamento:
            self.cancelamento_capa.set()
            self.btn_cancelar_capa.config(state=DISABLED)
            self.status_label.config(text="Cancelando… aguardando a etapa em andamento")

    def finalizar_processamento(self, sucesso=True, cancelado=False):
        self.processamento_em_andamento = False
        self.indicadores_capa.finalizar()
        self.btn_processar.config(state=NORMAL)
        self.btn_select.config(state=NORMAL)
        self.btn_cancelar_capa.config(state=DISABLED)
        texto = "Processamento cancelado" if cancelado else "Concluído" if sucesso else "Concluído com falhas"
        self.status_label.config(text=texto, bootstyle="success" if sucesso and not cancelado else "warning")
        if sucesso and not cancelado:
            self.progress_bar['value'] = 100
            self.progress_percent_label.config(text="100%")

    def _log_lpr_capa(self, mensagem, nivel="INFO"):
        self.adicionar_log(mensagem)
        (logging.error if nivel == "ERRO" else logging.info)("Capa: %s", mensagem)

    def processar_arquivos(self, arquivos, automatico):
        sessao = f"CAPA_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}"
        restantes = {}
        falhas_preparo = []
        falhas_txt = []
        falhas_envio = []
        montados = 0
        enviados = 0
        try:
            for txt in arquivos:
                try:
                    for linha in ler_arquivo_txt(txt, propagar_erro=True):
                        restantes[(txt, linha['LINHA'])] = linha
                except (OSError, UnicodeError) as erro:
                    falhas_txt.append(f"{txt}: {erro}")

            with tempfile.TemporaryDirectory(prefix="capa_", dir=PASTA_TRABALHO) as pasta:
                grupos = {}
                total = len(restantes)
                usuario = os.environ.get('USERNAME', 'desconhecido')
                for indice, (chave, linha) in enumerate(restantes.copy().items(), 1):
                    if self.cancelamento_capa.is_set():
                        break
                    try:
                        origem = os.path.join(linha['CAMINHO'], linha['ARQUIVO'])
                        if not origem.lower().endswith('.pdf') or linha['QTD'] < 1:
                            raise ValueError("A linha precisa indicar um PDF e pelo menos uma capa.")
                        with fitz.open(origem) as pdf:
                            frente_verso = pdf.page_count > 1
                        fila = linha['FILACOLOR']
                        formato = fila[:-3] if fila.upper().endswith('_FV') else fila
                        if frente_verso:
                            fila = f"{formato}_FV"
                        pasta_item = tempfile.mkdtemp(prefix="item_", dir=pasta)
                        entrada = os.path.join(pasta_item, f"{linha['NOMEFIC']}_temp.pdf")
                        shutil.copy2(origem, entrada)
                        marcado = adicionar_marca_dagua(
                            entrada, linha['LAMINAÇÃO'], linha['CODIGOACAB'], linha['IDENTIFIC'],
                            formato, propagar_erro=True)
                        if not marcado:
                            raise RuntimeError("A marca d’água não gerou um PDF.")
                        grupo = grupos.setdefault((chave[0], linha['IDENTIFIC'], fila), [])
                        grupo.append((chave, linha, marcado))
                        self.registrar_log_sqlite(self._dados_capa(
                            chave[0], linha['ARQUIVO'], fila, linha['QTD'], usuario,
                            'capa_marcacao', 'Montagem OK', montagem_ok=1))
                    except Exception as erro:
                        falhas_preparo.append((restantes.pop(chave), str(erro)))
                    self.root.after(0, self.atualizar_progresso_ui, 80 * indice / total,
                                    f"Preparando capas: {indice}/{total}")

                for indice, ((txt, identific, fila), itens) in enumerate(grupos.items(), 1):
                    if self.cancelamento_capa.is_set():
                        break
                    quantidade = sum(linha['QTD'] for _, linha, _ in itens)
                    try:
                        agrupado = os.path.join(pasta, f"mescla_{indice}.pdf")
                        gprint_mesclar_capas(((pdf, linha['QTD']) for _, linha, pdf in itens), agrupado)
                        destino = gprint_reservar_destino_capa(os.path.dirname(txt), identific, fila, quantidade)
                        try:
                            shutil.move(agrupado, destino)
                        except Exception:
                            with suppress(OSError):
                                os.remove(destino)
                            raise
                    except Exception as erro:
                        falhas_preparo.extend((linha, str(erro)) for _, linha, _ in itens)
                        for chave, _, _ in itens:
                            restantes.pop(chave)
                        continue

                    for chave, _, _ in itens:
                        restantes.pop(chave)
                    montados += 1
                    self.adicionar_log(f"PDF gerado: {destino} | Fila: {fila} | Capas: {quantidade}")
                    self.registrar_log_sqlite(self._dados_capa(
                        txt, os.path.basename(destino), fila, quantidade, usuario, 'mesclagem', 'Mesclagem OK'))
                    titulo = f'{identific}_CAPAS_{quantidade}' if identific else f'CAPAS_{quantidade}'
                    if fila.upper().endswith('_FV'):
                        titulo += '_FV'
                    job = f'{titulo}.pdf'
                    dados = self._dados_capa(txt, os.path.basename(destino), fila, quantidade, usuario,
                                             'envio', 'Não enviado (manual)', copia='1/1')
                    item = {'pdf': destino, 'job': job, 'log': dados, 'copia': 1, 'total_envios': 1,
                            'processo': itens[0][1], 'linhas': [linha for _, linha, _ in itens]}
                    if automatico:
                        if self.cancelamento_capa.is_set():
                            item['situacao'] = 'reenviar'
                            dados.update(status='Envio cancelado', erro='Envio não iniciado após cancelamento.')
                        else:
                            enviados += self._enviar_capa_pendente(item)
                        if item['situacao'] != 'aceito':
                            falhas_envio.append(item)
                        self.cancelamento_capa.wait(1)
                    else:
                        self.registrar_log_sqlite(dados)
                    self.root.after(0, self.atualizar_progresso_ui, 80 + 20 * indice / len(grupos),
                                    f"Mesclando capas: {indice}/{len(grupos)}")
        except Exception as erro:
            logging.exception("Falha no processamento de capas")
            falhas_txt.append(str(erro))
        finally:
            gprint_sincronizar_pendentes(self.log_db_path, root=self.root, emitir_log=self.adicionar_log)
            self.root.after(0, self._finalizar_lote_capa, sessao, montados, enviados,
                            list(restantes.values()), falhas_preparo, falhas_txt, falhas_envio)

    def _dados_capa(self, txt, arquivo, fila, quantidade, usuario, modo, status, **campos):
        return dict({
            'lote': os.path.dirname(txt), 'pilha': os.path.splitext(os.path.basename(txt))[0],
            'arquivo': arquivo, 'formato': fila, 'quantidade': quantidade, 'usuario': usuario,
            'maquina': self.maquina_capa_execucao, 'ip': self.ip_capa_execucao,
            'modo': modo, 'status': status, 'montagem_ok': 0, 'envio_ok': 0, 'erro': '', 'copia': '',
        }, **campos)

    def _enviar_capa_pendente(self, item):
        dados = item['log']
        aceito, erro, timeout = gprint_enviar_pdf_lpr(
            self._log_lpr_capa, item['pdf'], dados['ip'], dados['formato'], item['job'])
        item['situacao'] = 'aceito' if aceito else 'conferir' if timeout else 'reenviar'
        dados.update(status='Enviado' if aceito else 'Timeout no envio' if timeout else 'Falha no envio',
                     erro=erro, envio_ok=int(aceito))
        self.registrar_log_sqlite(dados)
        return int(aceito)

    def _finalizar_lote_capa(self, sessao, montados, enviados, restantes, falhas_preparo, falhas_txt, falhas_envio):
        cancelado = self.cancelamento_capa.is_set()
        if restantes and (not cancelado or messagebox.askyesno(
            "GPrint - Capa cancelada",
            f"Restam {len(restantes)} linha(s) sem PDF final. Exportar os TXTs em cancelados?",
            parent=self.root,
        )):
            self._exportar_capa_txt(restantes, 'cancelados', sessao)

        # Evita reabrir todo o lote pelo botão Processar após um cancelamento parcial.
        if cancelado:
            self.arquivos_selecionados = []
            self.tree_arquivos.delete(*self.tree_arquivos.get_children())
            self.file_label.config(text="Selecione os TXTs para continuar")

        if falhas_preparo:
            for linha, erro in falhas_preparo:
                self._log_lpr_capa(f"{linha['TXT_ORIGEM']} | linha {linha['LINHA']}: {erro}", "ERRO")
            self._exportar_capa_txt(
                [linha for linha, _ in falhas_preparo], 'erros_preparo_capa', sessao, falhas_preparo)
        if falhas_txt:
            messagebox.showwarning("GPrint - Capa", "\n".join(falhas_txt), parent=self.root)

        self._relatorio_envio_capa(falhas_envio, sessao, "Envio original")
        reenviaveis = sum(item['situacao'] == 'reenviar' for item in falhas_envio)
        if reenviaveis and not cancelado and messagebox.askyesno(
            "GPrint - Reenviar capas?",
            f"Reenviar agora {reenviaveis} PDF(s) agrupado(s) já montado(s), na fila original?",
            parent=self.root,
        ):
            self.status_label.config(text="Reenviando PDFs de capa…")
            threading.Thread(
                target=self._reenviar_capas,
                args=(falhas_envio, sessao, montados, enviados, bool(falhas_preparo or falhas_txt or restantes)),
                daemon=True).start()
            return
        self._encerrar_capas(sessao, montados, enviados, falhas_envio, cancelado,
                             bool(falhas_preparo or falhas_txt or restantes))

    def _reenviar_capas(self, itens, sessao, montados, enviados, falha_preparo):
        try:
            for item in itens:
                if self.cancelamento_capa.is_set():
                    break
                if item['situacao'] == 'reenviar':
                    enviados += self._enviar_capa_pendente(item)
                    self._relatorio_envio_capa([item], sessao, "Tentativa de reenvio")
                    self.cancelamento_capa.wait(1)
        except Exception:
            falha_preparo = True
            logging.exception("Falha ao finalizar o reenvio de capas")
        finally:
            gprint_sincronizar_pendentes(self.log_db_path, root=self.root, emitir_log=self.adicionar_log)
            self.root.after(0, self._encerrar_capas, sessao, montados, enviados, itens,
                            self.cancelamento_capa.is_set(), falha_preparo)

    def _relatorio_envio_capa(self, itens, sessao, etapa):
        por_pasta = {}
        for item in itens:
            por_pasta.setdefault(os.path.dirname(item['pdf']), []).append(item)
        for pasta, registros in por_pasta.items():
            relatorio = os.path.join(pasta, f'{sessao}_erros_envio_lpr.txt')
            gprint_registrar_pendencias_lpr(self._log_lpr_capa, relatorio, etapa, registros)

    def _encerrar_capas(self, sessao, montados, enviados, itens, cancelado, falha_preparo):
        reenvio = [linha for item in itens if item['situacao'] == 'reenviar' for linha in item['linhas']]
        conferir = [linha for item in itens if item['situacao'] == 'conferir' for linha in item['linhas']]
        for linhas, pasta in ((reenvio, 'reenvio_lpr'), (conferir, 'conferir_lpr')):
            if linhas:
                self._exportar_capa_txt(linhas, pasta, sessao)
        tem_falhas = bool(falha_preparo or reenvio or conferir)
        self.finalizar_processamento(not tem_falhas, cancelado)
        aviso = f"PDFs agrupados: {montados} | Envios aceitos pelo LPR: {enviados}."
        if conferir:
            aviso += "\nConfira os envios com timeout na impressora antes de reenviar."
        if reenvio:
            aviso += "\nOs PDFs agrupados continuam disponíveis para envio manual. Consulte o relatório no log."
        exibir = messagebox.showwarning if tem_falhas or cancelado else messagebox.showinfo
        exibir("GPrint - Capa", aviso, parent=self.root)

    def _exportar_capa_txt(self, linhas, subpasta, sessao, falhas=()):
        itens = [(linha, linha['QTD']) for linha in sorted(linhas, key=lambda p: (p['TXT_ORIGEM'], p['LINHA']))]
        if falhas:
            for origem in dict.fromkeys(linha['TXT_ORIGEM'] for linha, _ in falhas):
                pasta = os.path.join(os.path.dirname(origem), subpasta, sessao)
                nome = os.path.basename(origem) + '.erros.txt'
                try:
                    os.makedirs(pasta, exist_ok=True)
                    with open(os.path.join(pasta, nome), 'w', encoding='utf-8-sig') as arquivo:
                        arquivo.writelines(f"Linha {linha['LINHA']} | {linha['ARQUIVO']}: {erro}\n"
                                          for linha, erro in falhas if linha['TXT_ORIGEM'] == origem)
                except OSError as erro:
                    self._log_lpr_capa(f"Falha ao registrar erros de preparo: {erro}", "ERRO")
        gprint_exportar_txt_pendentes(itens, subpasta, sessao, f"GPrint - {subpasta}", self.root,
                                     self.adicionar_log)

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

        # Centraliza em relacao ao GPrint (sem alterar botoes/menus).
        gprint_centralizar_janela(about_window, self.root, 450, 350)

        main = tk.Frame(about_window, bg="#0f172a", padx=30, pady=20)
        main.pack(fill="both", expand=True)

        tk.Label(main, text="GPrint", font=("Segoe UI", 18, "bold"),
                bg="#0f172a", fg="#f8fafc").pack(pady=(10, 0))

        versao_badge = tk.Label(main, text="  v14.0  ", font=("Segoe UI", 10, "bold"),
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
def ler_arquivo_txt(caminho_txt, propagar_erro=False):
    dados = []
    try:
        # utf-8-sig remove BOM automaticamente
        with open(caminho_txt, "r", encoding="utf-8-sig", newline="") as arquivo:
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
                            "TXT_ORIGEM": caminho_txt,
                            "LINHA_ORIGINAL": linha,
                            "LINHA": idx,
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
        if propagar_erro:
            raise
        messagebox.showerror("Erro", f"Não foi possível ler o arquivo {os.path.basename(caminho_txt)}")
        return []
def referencia_capa(pagina):
    """Obtém a apara, a área de arte ou os limites dos elementos impressos."""
    area = pagina.rect
    origem = pagina.cropbox_position
    para_area_visivel = fitz.Matrix(1, 0, 0, 1, -origem.x, -origem.y)
    for chave, caixa in (("TrimBox", pagina.trimbox), ("ArtBox", pagina.artbox)):
        if pagina.parent.xref_get_key(pagina.xref, chave)[0] != "null":
            referencia = caixa * para_area_visivel
            if not referencia.is_empty and area.contains(referencia):
                return referencia

    conteudo = fitz.Rect()
    for tipo, coordenadas in pagina.get_bboxlog():
        if tipo != "ignore-text":
            elemento = fitz.Rect(coordenadas) & area
            if not elemento.is_empty:
                conteudo |= elemento
    return area if conteudo.is_empty else conteudo


def adicionar_marca_dagua(pdf_path, laminacao, codigoacab, identific, filacolor, *, propagar_erro=False):
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

        mm = 72 / 25.4
        novo_doc = fitz.open()

        for pagina in doc:
            largura_pt = pagina.rect.width
            altura_pt = pagina.rect.height
            # O fluxo de capa pronta é o original de 32 × 60, incluindo
            # o recorte de 20 mm por lado e 40 mm no topo para a fila 28 × 56.
            capa_pronta = (
                filacolor in ("Capas_32X60", "Capas_28X56")
                and abs(largura_pt / mm - 600) <= 5
                and abs(altura_pt / mm - 320) <= 5
            )
            if capa_pronta:
                nova_pagina = novo_doc.new_page(width=largura_pt, height=altura_pt)
                if filacolor == "Capas_28X56":
                    nova_pagina.set_cropbox(fitz.Rect(
                        20 * mm, 40 * mm, largura_pt - 20 * mm, altura_pt
                    ))
            else:
                largura_dest_cm, altura_dest_cm = formatos_cm.get(filacolor, (60, 32))
                nova_pagina = novo_doc.new_page(
                    width=largura_dest_cm * 10 * mm, height=altura_dest_cm * 10 * mm
                )

            visivel = nova_pagina.rect
            rotacao = pagina.rotation
            para_visual = pagina.rotation_matrix
            # show_pdf_page recebe o recorte sem /Rotate. Repõe a orientação na montagem.
            pagina.set_rotation(0)
            area_original = pagina.rect * para_visual
            x_original = (visivel.width - area_original.width) / 2
            y_original = visivel.y1 - area_original.y1
            x, y = x_original, y_original

            if not capa_pronta:
                referencia = referencia_capa(pagina) * para_visual
                x = (visivel.width - referencia.x0 - referencia.x1) / 2
                y_maximo = visivel.y1 - referencia.y1
                y = y_maximo - 10 * mm
                y_minimo = visivel.y0 - referencia.y0
                if y < y_minimo - 0.01 * mm:
                    if y_minimo - y_maximo > 0.01 * mm:
                        raise ValueError(
                            f"Página {pagina.number + 1}: a área da capa tem "
                            f"{referencia.height / mm:.2f} mm de altura e excede "
                            f"os {visivel.height / mm:.2f} mm da folha."
                        )
                    y = min(max(y_original, y_minimo), y_maximo)
                    logging.info(
                        "Capa %s, página %d: recuo inferior de %.2f mm para manter a área da capa inteira.",
                        nome_arquivo, pagina.number + 1, (y_maximo - y) / mm
                    )

            deslocamento = fitz.Matrix(1, 0, 0, 1, x, y)
            destino = (area_original * deslocamento) & visivel
            if not destino.is_empty and pagina.get_contents():
                origem = destino * ~deslocamento * ~para_visual
                nova_pagina.show_pdf_page(
                    destino, doc, pagina.number, clip=origem, rotate=-rotacao
                )

            x_marca = visivel.x0 + 50 + CONFIG.get('deslocamento_x', 0)
            largura_marca = visivel.width - 100
            y0_marca = visivel.y0 + recuo_topo_pt
            posicao_marca = fitz.Rect(
                x_marca, y0_marca, x_marca + largura_marca, y0_marca + altura_caixa_texto
            )
            nova_pagina.insert_textbox(
                posicao_marca,
                f"{nome_arquivo} - {identific}",
                fontsize=CONFIG['tamanho_fonte'],
                color=cor_marca,
                fontname=CONFIG['fonte'],
                align=1
            )

        # Salva o documento final
        novo_caminho = os.path.splitext(pdf_path)[0] + "_marcado.pdf"
        novo_doc.save(novo_caminho)
        novo_doc.close()
        doc.close()
        return novo_caminho

    except Exception as e:
        logging.error(f"Erro ao adicionar marca d'água em {pdf_path}: {e}")
        if propagar_erro:
            raise
        messagebox.showerror("Erro", f"Falha ao montar {os.path.basename(pdf_path)}\n\n{e}")
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

        ttk.Label(print_frame, text="Fila de impressão:", width=18).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.format_cb = ttk.Combobox(print_frame, state="readonly", width=36, bootstyle=(PRIMARY))
        self.format_cb.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(print_frame, text="Identificação (opcional):", width=22).grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.id_entry = ttk.Entry(print_frame, width=36, bootstyle=(PRIMARY))
        self.id_entry.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
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

    def on_machine_change(self, event=None):
        maquina=self.equipamentos_disponiveis.get(self.machine_var.get()) if hasattr(self,'equipamentos_disponiveis') else None
        if not maquina:
            self.ip_label.configure(text='IP: ---')
            self.format_cb.configure(values=[])
            self.format_cb.set('')
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
def gprint_validar_destino_refazer(equipamentos, filas_por_equipamento, nome_maquina, nome_fila, tipo):
    """Valida o destino de um retrabalho sem alterar o nome da fila LPR.

    Retorna um cadastro (nome, IP, tipo) ou gera ValueError.
    Equipamentos e filas vêm do cadastro central; só ativos são aceitos.
    """
    maquina = equipamentos.get(str(nome_maquina).strip())
    if not maquina:
        raise ValueError('Selecione uma máquina cadastrada e ativa.')
    if maquina['tipo'] not in (str(tipo).upper(), 'AMBOS'):
        raise ValueError(f"A máquina {maquina['nome']} não está habilitada para {tipo}.")
    if not str(nome_fila).strip():
        raise ValueError(f"Escolha uma fila para a máquina {maquina['nome']}.")
    filas = filas_por_equipamento.get(maquina['id'], {})
    if str(nome_fila).strip() not in filas:
        raise ValueError(f"A fila escolhida não está ativa em {maquina['nome']}.")
    return {'maquina': maquina.copy(), 'fila': str(nome_fila).strip()}


class GPrintJanelaResumoPerdas(tk.Toplevel):
    """Resumo restrito. Apenas SQL agregado; nada e carregado ao abrir."""
    PERIODOS = ('Ultimos 30 dias', 'Mes anterior', 'Ultimos 3 meses (90 dias)',
                'Mes atual', 'Personalizado')

    def __init__(self, parent, repositorio, cadastro):
        super().__init__(parent)
        self.withdraw()
        self.repositorio = repositorio
        self.cadastro = cadastro
        self.title('GPrint - Relatorio de perdas (acesso restrito)')
        self.geometry('940x590')
        self.minsize(760, 500)
        self.transient(parent.winfo_toplevel())
        self.periodo = tk.StringVar(value='Ultimos 30 dias')
        self.inicio = tk.StringVar()
        self.fim = tk.StringVar()
        self.setor = tk.StringVar()
        self.operador = tk.StringVar()
        self.tipo = tk.StringVar(value='Todos')
        self.status = tk.StringVar(value='Selecione o periodo e clique em Buscar.')
        self.total = tk.StringVar(value='Nenhuma consulta realizada.')
        self._buscando = False
        self._montar()
        self._carregar_filtros()
        self._atualizar_periodo()
        gprint_centralizar_janela(self, parent, 940, 590)
        self.grab_set()
        self.focus_set()

    @staticmethod
    def intervalo_periodo(nome, hoje=None):
        """Periodos prontos; datas inclusivas com inicio/fim do calendario."""
        hoje = hoje or date.today()
        if nome == 'Ultimos 30 dias':
            return hoje - timedelta(days=29), hoje
        if nome == 'Mes anterior':
            primeiro_atual = hoje.replace(day=1)
            ultimo_anterior = primeiro_atual - timedelta(days=1)
            return ultimo_anterior.replace(day=1), ultimo_anterior
        if nome == 'Ultimos 3 meses (90 dias)':
            return hoje - timedelta(days=89), hoje
        if nome == 'Mes atual':
            return hoje.replace(day=1), hoje
        raise ValueError('Selecione um periodo valido.')

    def _montar(self):
        base = ttk.Frame(self, padding=14)
        base.pack(fill='both', expand=True)
        ttk.Label(base, text='Resumo de perdas e retrabalhos',
                  font=('Segoe UI', 16, 'bold')).pack(anchor='w', pady=(0, 8))
        ttk.Label(base, text='Dados do refazer.db central. Somente consulta; nenhum registro e alterado.',
                  bootstyle=SECONDARY).pack(anchor='w', pady=(0, 10))
        filtros = ttk.LabelFrame(base, text='Filtros', padding=10)
        filtros.pack(fill='x')
        for col in range(6):
            filtros.columnconfigure(col, weight=1)
        ttk.Label(filtros, text='Periodo').grid(row=0, column=0, sticky='w')
        self.combo_periodo = ttk.Combobox(filtros, state='readonly',
                                          values=self.PERIODOS, textvariable=self.periodo)
        self.combo_periodo.grid(row=1, column=0, columnspan=2, sticky='ew', padx=(0, 10))
        self.combo_periodo.bind('<<ComboboxSelected>>', lambda _event: self._atualizar_periodo())
        ttk.Label(filtros, text='Inicio (DD/MM/AAAA)').grid(row=0, column=2, sticky='w')
        self.entrada_inicio = ttk.Entry(filtros, textvariable=self.inicio)
        self.entrada_inicio.grid(row=1, column=2, sticky='ew', padx=(0, 10))
        ttk.Label(filtros, text='Fim (DD/MM/AAAA)').grid(row=0, column=3, sticky='w')
        self.entrada_fim = ttk.Entry(filtros, textvariable=self.fim)
        self.entrada_fim.grid(row=1, column=3, sticky='ew', padx=(0, 10))
        ttk.Button(filtros, text='Buscar', command=self.buscar,
                   bootstyle=PRIMARY).grid(row=1, column=4, columnspan=2, sticky='ew')
        ttk.Label(filtros, text='Tipo').grid(row=2, column=0, sticky='w', pady=(9, 0))
        ttk.Combobox(filtros, state='readonly', textvariable=self.tipo,
                     values=('Todos', 'CAPA', 'MIOLO')).grid(row=3, column=0, columnspan=2,
                                                             sticky='ew', padx=(0, 10))
        ttk.Label(filtros, text='Setor (opcional)').grid(row=2, column=2, sticky='w', pady=(9, 0))
        self.combo_setor = ttk.Combobox(filtros, textvariable=self.setor, values=('',), state='readonly')
        self.combo_setor.grid(row=3, column=2, columnspan=2, sticky='ew', padx=(0, 10))
        ttk.Label(filtros, text='Operador (opcional)').grid(row=2, column=4, sticky='w', pady=(9, 0))
        self.combo_operador = ttk.Combobox(filtros, textvariable=self.operador, values=('',), state='readonly')
        self.combo_operador.grid(row=3, column=4, columnspan=2, sticky='ew')
        ttk.Label(base, textvariable=self.total, font=('Segoe UI', 11, 'bold')).pack(
            anchor='w', pady=(14, 6))
        painel = ttk.Frame(base)
        painel.pack(fill='both', expand=True)
        painel.rowconfigure(0, weight=1)
        painel.columnconfigure(0, weight=1)
        cols = ('setor', 'arquivos', 'unidades', 'montados', 'enviados', 'falhas')
        self.tabela = ttk.Treeview(painel, columns=cols, show='headings', selectmode='none')
        for col, label, width in [
            ('setor','Setor',210),('arquivos','Arquivos',100),
            ('unidades','Unidades',100),('montados','Preparados',100),
            ('enviados','Envios LPR',110),('falhas','Falhas LPR',110)]:
            self.tabela.heading(col, text=label)
            self.tabela.column(col, width=width, anchor='w' if col=='setor' else 'center',
                               stretch=True)
        self.tabela.grid(row=0, column=0, sticky='nsew')
        rolagem = ttk.Scrollbar(painel, orient='vertical', command=self.tabela.yview)
        rolagem.grid(row=0, column=1, sticky='ns')
        self.tabela.configure(yscrollcommand=rolagem.set)
        rodape = ttk.Frame(base)
        rodape.pack(fill='x', pady=(9, 0))
        ttk.Label(rodape, textvariable=self.status, bootstyle=SECONDARY).pack(side='left')
        ttk.Button(rodape, text='Fechar', command=self.destroy).pack(side='right')

    def _carregar_filtros(self):
        def consultar():
            setores, operadores = [], []
            try:
                setores, operadores = self.repositorio.opcoes_historicas_perdas()
            except (OSError, sqlite3.Error) as exc:
                logging.warning('Refazer: opções do histórico indisponíveis: %s', exc)
            for buscar, nomes in (
                (self.cadastro.listar_setores, setores), (self.cadastro.listar_operadores, operadores)
            ):
                try:
                    nomes.extend(r['nome'] for r in buscar(incluir_inativos=True))
                except GPrintCadastroErro as exc:
                    logging.warning('Refazer: opções do cadastro indisponíveis: %s', exc)
            try:
                self.after(0, self._aplicar_filtros, setores, operadores)
            except (RuntimeError, tk.TclError):
                pass
        threading.Thread(target=consultar, daemon=True).start()

    def _aplicar_filtros(self, setores, operadores):
        if self.winfo_exists():
            for combo, nomes in ((self.combo_setor, setores), (self.combo_operador, operadores)):
                opcoes = {nome.strip().casefold(): nome.strip() for nome in nomes if nome.strip()}
                combo.configure(values=('', *sorted(opcoes.values(), key=str.casefold)))

    def _atualizar_periodo(self):
        personalizado = self.periodo.get() == 'Personalizado'
        if not personalizado:
            inicio, fim = self.intervalo_periodo(self.periodo.get())
            self.inicio.set(inicio.strftime('%d/%m/%Y'))
            self.fim.set(fim.strftime('%d/%m/%Y'))
        self.entrada_inicio.configure(state='normal' if personalizado else 'readonly')
        self.entrada_fim.configure(state='normal' if personalizado else 'readonly')

    def buscar(self):
        if self._buscando:
            return
        try:
            inicio = datetime.strptime(self.inicio.get().strip(), '%d/%m/%Y').date()
            fim = datetime.strptime(self.fim.get().strip(), '%d/%m/%Y').date()
            if inicio > fim:
                raise ValueError('A data inicial deve ser anterior ou igual a data final.')
            if fim > date.today():
                raise ValueError('O fim do periodo nao pode estar no futuro.')
        except ValueError as exc:
            messagebox.showwarning('GPrint - Periodo invalido',
                'Informe datas validas em DD/MM/AAAA.\n' + str(exc), parent=self)
            return
        self._buscando = True
        self.status.set('Consultando o banco central...')
        periodo_texto = f'{inicio:%d/%m/%Y} a {fim:%d/%m/%Y}'
        setor, operador, tipo = self.setor.get(), self.operador.get(), self.tipo.get()

        def tarefa():
            try:
                linhas = self.repositorio.resumo_entre_datas(inicio, fim, setor, operador, tipo)
                erro = None
            except Exception as exc:
                linhas, erro = [], exc
            try:
                self.after(0, lambda: self._concluir_busca(linhas, erro, periodo_texto))
            except (tk.TclError, RuntimeError):
                pass
        threading.Thread(target=tarefa, daemon=True).start()

    def _concluir_busca(self, linhas, erro, periodo):
        if not self.winfo_exists():
            return
        self._buscando = False
        if erro is not None:
            self.status.set('Consulta nao concluida.')
            messagebox.showwarning('GPrint - Relatorio de perdas',
                'Nao foi possivel consultar o banco de retrabalhos. '
                'Verifique a pasta producao e, se necessario, informe o Guilherme.',
                parent=self)
            logging.warning('Refazer: falha na consulta de perdas: %s', erro)
            return
        antigos = self.tabela.get_children()
        if antigos:
            self.tabela.delete(*antigos)
        for setor, arquivos, unidades, enviados, falhas, montados in linhas:
            self.tabela.insert('', 'end', values=(setor or 'Nao informado',
                arquivos, unidades, montados, enviados, falhas))
        arquivos = sum(int(r[1] or 0) for r in linhas)
        unidades = sum(int(r[2] or 0) for r in linhas)
        montados = sum(int(r[5] or 0) for r in linhas)
        enviados = sum(int(r[3] or 0) for r in linhas)
        falhas = sum(int(r[4] or 0) for r in linhas)
        self.total.set(f'{periodo}  |  {arquivos} arquivos  |  {unidades} unidades de retrabalho')
        self.status.set(f'PDFs preparados: {montados}   |   Envios LPR aceitos: {enviados}   |   Falhas: {falhas}')


def gprint_clientes_das_pastas(pastas):
    candidatos = []
    avisos = []
    for pasta in pastas:
        caminho = pasta['caminho']
        if not caminho:
            continue
        modo = pasta['modo']
        try:
            if modo == 'clientes':
                with os.scandir(caminho) as entradas:
                    clientes = sorted(
                        ((e.name, e.path) for e in entradas if e.is_dir()),
                        key=lambda item: item[0].casefold())
                candidatos.extend((nome, local, 'planilha', pasta, False) for nome, local in clientes)
            else:
                # Verifica o acesso sem carregar todos os lotes antecipadamente.
                with os.scandir(caminho):
                    pass
                candidatos.append((pasta['nome'], caminho, modo, pasta, True))
        except OSError as erro:
            avisos.append(f"{pasta['nome']} ({caminho}): {erro}")

    diretos = {nome.casefold() for nome, _, _, _, direto in candidatos if direto}
    destinos = {}
    rotulos = set()
    for nome, caminho, modo, pasta, direto in candidatos:
        rotulo = nome
        if rotulo.casefold() in rotulos or (not direto and rotulo.casefold() in diretos):
            rotulo = f"{nome} — {pasta['nome']}"
        while rotulo.casefold() in rotulos or (not direto and rotulo.casefold() in diretos):
            rotulo += f" [{pasta['slot'] + 1}]"
        rotulos.add(rotulo.casefold())
        destinos[rotulo] = (caminho, modo)
    return destinos, avisos


class RefazerTab(ttk.Frame):
    # ---------- CONFIGURAÇÕES (agora atributos da classe) ----------
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

    MAQUINAS_REF = ["RICOH", "OCÉ 1", "OCÉ 3", "BALTORO","ATEXCO"]
    OPERADORES = []  # Carregados do equipamentos.db, nao fixos no codigo.

    def __init__(self, parent, app=None):
        super().__init__(parent)
        self.parent = parent
        self.app = app  # pode ser usado para acessar recursos do app principal

        # Variáveis de estado
        self.lista_completa = []
        self.pasta_var = tk.StringVar()
        self.subpasta_var = tk.StringVar()
        self.busca_lote_var = tk.StringVar()
        self.lotes_disponiveis = []
        self.operador_var = tk.StringVar()
        self.filtro_var = tk.StringVar()
        self.idx_capa = -1
        self.idx_miolo = -1
        self.idx_caminho = -1
        # O refazer antigo (preparar/salvar) permanece como modo padrão.
        # Envio marcado por padrao, mas operador pode desativar para apenas montar.
        self.enviar_refazer_var = tk.BooleanVar(value=True)
        self.repositorio_refazer = GPrintRefazerRepositorio()
        self._consulta_operadores_em_andamento = False
        self._equipamentos_refazer = {}
        self._filas_refazer = {}
        self._fila_refazer = {}
        self._origem_lista = None
        self._setores_refazer = {}
        self._consulta_setores_refazer = False
        self._resumo_fila = tk.StringVar(value='Fila vazia')
        self._refazer_em_execucao = False

        self._criar_interface()
        self._clientes_por_pasta = {}
        self._carregando_clientes = False
        self._recarregar_clientes = False

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

    @staticmethod
    def _modo_miolo_refazer(tipo, quantidade):
        """A quantidade é a editada no Refazer, não a original da planilha."""
        if tipo == 'brochura':
            return 'dois_na_folha' if int(quantidade) > 10 else 'Corta_empilha'
        if tipo == 'grampo':
            return 'grampo'
        if tipo == '1up':
            return 'um_na_folha'
        raise ValueError(f'Tipo de montagem do Refazer desconhecido: {tipo}')

    def montar_pdf(self, entrada, saida, tipo, quantidade=1, lote='', nome_original=None,
                   sangria_mm=None):
        """Reutiliza montagem, marcas e identificação da aba Miolo.

        A cópia de entrada mantém a identificação original mesmo após redução.
        A quantidade escolhe o layout; o envio continua sendo um trabalho por PDF.
        """
        modo = self._modo_miolo_refazer(tipo, quantidade)
        montadores = {
            'Corta_empilha': Corta_empilha,
            'dois_na_folha': dois_na_folha,
            'grampo': grampo,
            'um_na_folha': um_na_folha,
        }
        if sangria_mm is None:
            sangria_mm = obter_sangria_padrao_miolo()
        nome = os.path.splitext(os.path.basename(nome_original or entrada))[0]
        lote_curto = self._identificacao_lote(lote)
        identificacao = f'LOTE_{lote_curto}_{nome}' if lote_curto else nome
        identificacao = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', identificacao).strip(' .')
        # Limita somente a cópia temporária, preservando o lote no início.
        identificacao = identificacao[:160].rstrip(' .') or 'GPrint'
        with tempfile.TemporaryDirectory(prefix='gprint_ref_miolo_') as pasta_temp:
            entrada_identificada = os.path.join(pasta_temp, identificacao + '.pdf')
            shutil.copy2(entrada, entrada_identificada)
            # A redução, quando solicitada, já ocorreu antes da montagem.
            # Corta_empilha tem compressão ligada por padrão: desliga aqui
            # para respeitar o setor escolhido e evitar uma segunda redução.
            opcoes = {'comprimir_entrada': False} if modo == 'Corta_empilha' else {}
            resultado = montadores[modo](
                entrada_identificada, saida, self.LARGURA_PT, self.ALTURA_PT,
                sangria_mm=sangria_mm, add_marcas=True, **opcoes)
            if resultado is False or not os.path.isfile(saida) or os.path.getsize(saida) == 0:
                raise RuntimeError('A montagem do miolo não gerou um PDF válido.')
        return modo

    def adicionar_marca_dagua(self, caminho_pdf, texto, distancia_mm=3, caminho_saida=None):
        if not caminho_pdf.lower().endswith('.pdf'):
            return
        doc = fitz.open(caminho_pdf)

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
            raise

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

    def _cliente_uiclap(self, cliente):
        return self._clientes_por_pasta.get(cliente, ('', ''))[1] in ('pdf', 'uiclap_legado')

    def _pasta_cliente_refazer(self, cliente):
        return self._clientes_por_pasta[cliente][0]

    def obter_lista_arquivos_da_pasta(self, pasta):
        """UICLAP novo: PDFs na raiz do lote, sem depender de planilha."""
        lista = []
        with os.scandir(pasta) as entradas:
            nomes = sorted((e.name for e in entradas
                            if e.is_file() and e.name.lower().endswith('.pdf')),
                           key=str.casefold)
        for nome in nomes:
            tipo_match = re.search(r'(?:^|[\s_()\-])(capa|miolo)(?=$|[\s_()\-])',
                                   os.path.splitext(nome)[0], re.IGNORECASE)
            if not tipo_match:
                continue
            tipo = tipo_match.group(1).upper()
            prefixo = nome[:tipo_match.start(1)].strip(' -_').casefold()
            restante = nome[tipo_match.end(1):].lstrip(' -_')
            codigo_match = re.match(r'(\d+)(?:v\d+)?(?=$|[\s_\-])', restante, re.IGNORECASE)
            codigo = codigo_match.group(1) if codigo_match else 'N/A'
            formato = re.search(r'(\d+(?:[.,]\d+)?)\s*[xX]\s*(\d+(?:[.,]\d+)?)', restante)
            # Sem Excel, brochura é o padrão do retrabalho. Grampo explícito
            # no nome mantém sua montagem própria.
            acabamento = 'Grampo' if re.search(r'\b(?:grampo|grampeado)\b', nome, re.IGNORECASE) else 'Brochura'
            lista.append({
                'nome': nome, 'tipo': tipo, 'linha': (), 'codigo': codigo,
                'caminho_fisico': os.path.join(pasta, nome),
                'acabamento': acabamento,
                'formato_miolo': formato.group(0) if formato else None,
                '_chave_par': (prefixo, codigo),
            })
        # Pareia somente quando há um único arquivo do outro tipo no lote.
        # A lista é completa antes de aplicar o filtro digitado pelo operador.
        pares = {}
        for item in lista:
            if item['codigo'] != 'N/A':
                pares.setdefault((item['_chave_par'], item['tipo']), []).append(item)
        for item in lista:
            outro_tipo = 'MIOLO' if item['tipo'] == 'CAPA' else 'CAPA'
            candidatos = pares.get((item['_chave_par'], outro_tipo), [])
            if len(candidatos) == 1:
                item['arquivo_relacionado'] = candidatos[0]['caminho_fisico']
            item.pop('_chave_par', None)
        return lista

    def obter_lista_arquivos_da_planilha(self, caminho_excel, cliente):
        wb = openpyxl.load_workbook(caminho_excel, data_only=True)
        ws = wb.active
        idx_capa = self.obter_indice_coluna(ws, "ARQUIVO CAPA")
        idx_miolo = self.obter_indice_coluna(ws, "ARQUIVO MIOLO")
        idx_caminho = self.obter_indice_coluna(ws, "CAMINHO DO ARQUIVO")

        if self._cliente_uiclap(cliente):
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

    def copiar_para_refazer(self, origem, pasta_lote, tipo, subpasta=None):
        if not os.path.exists(origem):
            return None
        destino_dir = os.path.join(pasta_lote, "REFAZER", tipo)
        if subpasta:
            destino_dir = os.path.join(destino_dir, subpasta)
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
        self.cb_cliente = ttk.Combobox(top, textvariable=self.pasta_var, state="readonly")
        self.cb_cliente.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.cb_cliente.bind("<<ComboboxSelected>>", self.on_cliente_change)
        ttk.Button(top, text="Atualizar clientes", command=self.carregar_clientes).grid(
            row=0, column=4, padx=(0, 5), pady=5)


        ttk.Label(top, text="Lote").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        campo_lote = ttk.Frame(top)
        campo_lote.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        campo_lote.columnconfigure(2, weight=1)
        ttk.Label(campo_lote, text="Buscar nº").grid(row=0, column=0, padx=(0, 4))
        busca_lote = ttk.Entry(campo_lote, textvariable=self.busca_lote_var, width=10)
        busca_lote.grid(row=0, column=1, padx=(0, 6))
        busca_lote.bind("<Return>", self._confirmar_busca_lote)
        self.cb_lote = ttk.Combobox(campo_lote, textvariable=self.subpasta_var, state="readonly")
        self.cb_lote.grid(row=0, column=2, sticky="ew")
        self.cb_lote.bind("<<ComboboxSelected>>", self.on_lote_change)
        self.busca_lote_var.trace_add("write", self._filtrar_lotes)

        ttk.Label(top, text="Operador").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.cb_operador = ttk.Combobox(top, values=[], textvariable=self.operador_var, state='readonly')
        self.cb_operador.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        # Operadores sao recebidos do cadastro central ao entrar na aba.

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

        self.btn_preparar = ttk.Button(botoes_top, text="Adicionar à fila",
                                    command=self.preparar_processamento, bootstyle=INFO)
        self.btn_preparar.grid(row=0, column=1, padx=5, sticky="ew")

        self.btn_limpar = ttk.Button(botoes_top, text="Limpar filtros", command=self.limpar,
                                    bootstyle=DANGER)
        self.btn_limpar.grid(row=0, column=2, padx=5, sticky="ew")
        ttk.Button(botoes_top, text='Resumo de perdas (restrito)',
                   command=self.mostrar_resumo_perdas, bootstyle=SECONDARY
                   ).grid(row=1, column=0, columnspan=3, sticky='ew', padx=5, pady=(6, 0))

        paned = tk.PanedWindow(main_frame, orient="vertical", sashrelief="raised", sashwidth=8, bg="#000000")
        paned.grid(row=1, column=0, sticky="nsew", pady=5,padx=10)

        # --- Lista de arquivos ---
        frame_lista = ttk.LabelFrame(paned, text="Arquivos", bootstyle=PRIMARY)
        paned.add(frame_lista, stretch="always", minsize=110)

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
        self.frame_edicao = ttk.LabelFrame(paned, text="Edição Inteligente · Fila de retrabalho", bootstyle=PRIMARY)
        paned.add(self.frame_edicao, stretch="always", minsize=195)

        self.update_idletasks()
        paned.sash_place(0, 0, 190)

        self.frame_edicao.grid_rowconfigure(1, weight=1)
        self.frame_edicao.grid_columnconfigure(0, weight=1)

        toolbar = ttk.Frame(self.frame_edicao)
        toolbar.grid(row=0, column=0, sticky="ew", padx=5, pady=(5, 0))
        ttk.Label(toolbar, textvariable=self._resumo_fila).pack(side='left', padx=5)
        ttk.Button(toolbar, text='Remover selecionados', command=self.remover_da_fila,
                   bootstyle='danger-outline').pack(side='right', padx=5)
        ttk.Button(toolbar, text='Limpar fila', command=self.limpar_fila,
                   bootstyle='secondary-outline').pack(side='right')

        edit_container = ttk.Frame(self.frame_edicao)
        edit_container.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        edit_container.rowconfigure(0, weight=1)
        edit_container.columnconfigure(0, weight=1)

        colunas = ('Nome', 'Código', 'Quantidade', 'Setor', 'Máquina', 'Fila LPR', 'Cliente', 'Lote', 'Tipo')
        self.tree_edicao = ttk.Treeview(
            edit_container, columns=colunas, show='headings', selectmode='extended',
            displaycolumns=('Cliente', 'Lote', 'Tipo', 'Nome', 'Quantidade', 'Setor', 'Máquina', 'Fila LPR', 'Código'))
        for coluna, titulo, largura, ancora in (
            ('Cliente', 'Cliente', 125, 'w'), ('Lote', 'Lote', 145, 'w'), ('Tipo', 'Tipo', 55, 'center'),
            ('Nome', 'Arquivo', 235, 'w'), ('Quantidade', 'Qtd.', 55, 'center'), ('Setor', 'Setor', 150, 'w'),
            ('Máquina', 'Máquina', 115, 'w'), ('Fila LPR', 'Fila LPR', 180, 'w'), ('Código', 'Código', 75, 'center'),
        ):
            self.tree_edicao.heading(coluna, text=titulo)
            self.tree_edicao.column(coluna, width=largura, minwidth=45, anchor=ancora, stretch=coluna == 'Nome')

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
        self.cb_setor = ttk.Combobox(controles, width=17, state="readonly")
        self.cb_setor.pack(side="left", padx=5)

        ttk.Label(controles, text="Máquina:").pack(side="left", padx=(10, 5))
        self.cb_maquina = ttk.Combobox(controles, values=self.MAQUINAS_REF, width=18, state='readonly')
        self.cb_maquina.pack(side="left", padx=5)
        self.cb_maquina.current(0)
        self.cb_maquina.bind('<<ComboboxSelected>>', self._alterar_maquina_refazer)

        ttk.Label(controles, text='Fila:').pack(side='left', padx=(7, 2))
        self.cb_fila_refazer = ttk.Combobox(controles, width=18, values=[], state='readonly')
        self.cb_fila_refazer.pack(side='left', padx=3)

        self.btn_aplicar_sel = ttk.Button(controles, text="Aplicar seleção",
                                        bootstyle=SECONDARY, command=self.aplicar_nos_selecionados)
        self.btn_aplicar_sel.pack(side="left", padx=10)

        self.btn_aplicar_todos = ttk.Button(controles, text="Aplicar à fila",
                                            bootstyle=PRIMARY, command=self.aplicar_em_todos)
        self.btn_aplicar_todos.pack(side="left", padx=5)

        # NOTA: O botão "PROCESSAR" que estava aqui foi removido e será criado no status_frame abaixo.

        self.tree_edicao.bind("<Button-1>", self.edicao_inline)
        self.tree_edicao.bind('<Delete>', self.remover_da_fila)
        # Ao selecionar um trabalho, mostrar no rodapé sua máquina e fila.
        self.tree_edicao.bind('<<TreeviewSelect>>', self._sincronizar_selecao_refazer)

        # Só manda para a impressora quando o operador marcar esta opção.
        opcoes_envio = ttk.Frame(main_frame)
        opcoes_envio.grid(row=4, column=0, sticky='ew', padx=12, pady=(0, 5))
        ttk.Checkbutton(opcoes_envio, text='Enviar os PDFs preparados às impressoras após o refazer',
                        variable=self.enviar_refazer_var, bootstyle='info-square-toggle').pack(side='left')
        ttk.Label(opcoes_envio, text='(capas repetidas no PDF conforme a quantidade)',
                  bootstyle=SECONDARY).pack(side='left', padx=12)

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

    def atualizar_operadores_refazer(self):
        """Busca operadores ativos sem bloquear a interface durante leitura da rede."""
        if self._consulta_operadores_em_andamento:
            return
        self._consulta_operadores_em_andamento = True
        def consultar():
            try:
                operadores = self.app.repositorio_equipamentos.listar_operadores()
                nomes = [o['nome'] for o in operadores]
            except Exception as exc:
                logging.warning('Refazer: cadastro de operadores indisponivel: %s', exc)
                nomes = []
            def aplicar():
                self._consulta_operadores_em_andamento = False
                if not self.winfo_exists():
                    return
                atual = self.operador_var.get()
                self.cb_operador.configure(values=nomes)
                if atual not in nomes:
                    self.operador_var.set(nomes[0] if len(nomes) == 1 else '')
            try:
                self.after(0, aplicar)
            except (tk.TclError, RuntimeError):
                pass
        threading.Thread(target=consultar, daemon=True).start()

    def atualizar_setores_refazer(self):
        if self._consulta_setores_refazer:
            return
        self._consulta_setores_refazer = True

        def consultar():
            try:
                setores = self.app.repositorio_equipamentos.listar_setores()
            except GPrintCadastroErro as exc:
                logging.warning('Refazer: cadastro de setores indisponível: %s', exc)
                setores = None
            try:
                self.after(0, self._aplicar_setores_refazer, setores)
            except (tk.TclError, RuntimeError):
                pass
        threading.Thread(target=consultar, daemon=True).start()

    def _aplicar_setores_refazer(self, setores):
        self._consulta_setores_refazer = False
        if setores is None or not self.winfo_exists():
            return
        self._setores_refazer = {r['nome']: r['tratamento'] for r in setores}
        nomes = tuple(self._setores_refazer)
        self.cb_setor.configure(values=nomes)
        if self.cb_setor.get() not in self._setores_refazer:
            self.cb_setor.set(next(iter(nomes), ''))

    def atualizar_destinos_refazer(self):
        """Lê cadastro em segundo plano ao entrar na aba; não bloqueia a abertura."""
        if getattr(self, '_consulta_refazer_em_andamento', False):
            return
        self._consulta_refazer_em_andamento = True
        def consultar():
            try:
                repo = self.app.repositorio_equipamentos
                equipamentos = repo.listar_equipamentos()
                filas = {e['id']: {f['nome']: f for f in repo.listar_filas(e['id'])}
                         for e in equipamentos}
                self.after(0, lambda: self._destinos_refazer_carregados(equipamentos, filas))
            except Exception as exc:
                logging.warning('Refazer: cadastro de equipamentos indisponível: %s', exc)
                self.after(0, lambda: setattr(self, '_consulta_refazer_em_andamento', False))
        threading.Thread(target=consultar, daemon=True).start()

    def _destinos_refazer_carregados(self, equipamentos, filas):
        self._consulta_refazer_em_andamento = False
        self._equipamentos_refazer = {e['nome']: e for e in equipamentos}
        self._filas_refazer = filas
        nomes = list(self._equipamentos_refazer)
        self.cb_maquina.configure(values=nomes)
        if self.cb_maquina.get() not in nomes:
            self.cb_maquina.set(nomes[0] if len(nomes) == 1 else '')
        self._alterar_maquina_refazer()

    def _filas_para_refazer(self, maquina, tipo=None):
        e = self._equipamentos_refazer.get(maquina)
        if not e or (tipo and e['tipo'] not in (tipo, 'AMBOS')):
            return []
        return list(self._filas_refazer.get(e['id'], {}))

    def _alterar_maquina_refazer(self, _event=None):
        # Usa exatamente as filas ATIVAS cadastradas para essa máquina.
        # Não acrescenta/remove prefixos; não modifica argumentos LPR.
        filas = self._filas_para_refazer(self.cb_maquina.get())
        self.cb_fila_refazer.configure(values=filas)
        if self.cb_fila_refazer.get() not in filas:
            self.cb_fila_refazer.set('')

    def _sincronizar_selecao_refazer(self, _event=None):
        """Mostra os dados da linha selecionada nos controles inferiores."""
        escolhidos = self.tree_edicao.selection()
        if len(escolhidos) != 1:
            return
        dados = self.tree_edicao.item(escolhidos[0], 'values')
        if len(dados) < 6:
            return
        self.cb_setor.set(dados[3])
        self.cb_maquina.set(dados[4])
        self._alterar_maquina_refazer()
        filas = self._filas_para_refazer(dados[4], self._tipo_linha_refazer(escolhidos[0]))
        self.cb_fila_refazer.set(dados[5] if dados[5] in filas else '')

    def _tipo_linha_refazer(self, iid):
        return self._fila_refazer.get(iid, {}).get('tipo', '')

    def _validar_destinos_refazer(self):
        """Valida TODOS os destinos antes de modificar arquivos ou relatório."""
        if not self.enviar_refazer_var.get():
            return {}
        if not self._equipamentos_refazer:
            raise ValueError('O cadastro de equipamentos está indisponível. Abra a aba novamente.')
        plano = {}
        for iid in self.tree_edicao.get_children():
            linha = self.tree_edicao.item(iid, 'values')
            tipo = self._tipo_linha_refazer(iid)
            if not tipo:
                raise ValueError(f'Não foi possível identificar Capa/Miolo de {linha[0]}.')
            try:
                plano[iid] = gprint_validar_destino_refazer(
                    self._equipamentos_refazer, self._filas_refazer,
                    linha[4], linha[5], tipo)
            except ValueError as exc:
                raise ValueError(f"{linha[0]}: {exc}") from exc
        return plano

    @staticmethod
    def _identificacao_lote(lote):
        texto = str(lote or '').strip()
        numero = re.match(r'^([0-9]+)(?=$|[\s_\-\[])', texto)
        return numero.group(1) if numero else texto

    @staticmethod
    def _nome_trabalho_refazer(nome_pdf, lote=''):
        """Identificação curta e ASCII para compatibilidade com o LPR do Windows.

        Não altera o nome do PDF final e não altera o nome da fila.
        """
        import hashlib
        import unicodedata
        base = os.path.splitext(os.path.basename(nome_pdf))[0]
        ascii_nome = unicodedata.normalize('NFKD', base).encode('ascii', 'ignore').decode('ascii')
        ascii_nome = re.sub(r'[^A-Za-z0-9_.-]+', '_', ascii_nome).strip('_.-')
        lote_curto = RefazerTab._identificacao_lote(lote)
        lote_ascii = unicodedata.normalize('NFKD', lote_curto).encode('ascii', 'ignore').decode('ascii')
        lote_ascii = re.sub(r'[^A-Za-z0-9_.-]+', '_', lote_ascii).strip('_.-')
        identidade = f'{lote}|{base}' if lote else base
        sufixo = hashlib.sha1(identidade.encode('utf-8')).hexdigest()[:6]
        prefixo = f'LOTE_{lote_ascii[:32]}_' if lote_ascii else ''
        return prefixo + (ascii_nome[:32].rstrip('_.-') or 'GPrint') + '_' + sufixo

    def _enviar_refazer_lpr(self, arquivo_pdf, destino):
        """Envio exclusivo do Refazer com copia TEMPORARIA de nome curto.

        A montagem e o PDF final NAO sao modificados. A fila/IP continuam exatos.
        As outras rotinas LPR do programa permanecem inalteradas.
        """
        maquina = destino['maquina']
        fila = destino['fila']
        nome_trabalho = self._nome_trabalho_refazer(arquivo_pdf, destino.get('lote', ''))
        if not os.path.isfile(arquivo_pdf):
            return False, 'PDF final não encontrado.'
        try:
            if os.path.getsize(arquivo_pdf) == 0:
                return False, 'PDF final está vazio.'
        except OSError as exc:
            return False, f'Não foi possível acessar o PDF final: {exc}'

        # Diferentemente da montagem, somente o arquivo TEMPORÁRIO recebe nome curto.
        # Em especial, evita caminhos enormes/com acentos que o lpr.exe não consegue ler.
        temp_dir = tempfile.mkdtemp(prefix='gpr_')
        try:
            nome_temp = 'job_' + uuid.uuid4().hex[:10] + '.pdf'
            pdf_temp = os.path.join(temp_dir, nome_temp)
            shutil.copy2(arquivo_pdf, pdf_temp)
            if not os.path.isfile(pdf_temp) or os.path.getsize(pdf_temp) == 0:
                return False, 'Não foi possível preparar o arquivo temporário para envio.'
            # Confirma que a cópia temporária pode ser lida antes de chamar o LPR.
            with open(pdf_temp, 'rb') as copia:
                if not copia.read(5).startswith(b'%PDF-'):
                    return False, 'Arquivo temporário não é um PDF válido.'
            comando = ['lpr', '-S', maquina['ip'], '-P', fila, '-J', nome_trabalho, pdf_temp]
            resultado = subprocess.run(comando, capture_output=True, text=True, timeout=120)
            if resultado.returncode != 0:
                return False, (resultado.stderr or resultado.stdout or
                               f'Código de saída {resultado.returncode}').strip()
            return True, ''
        except subprocess.TimeoutExpired:
            return False, 'Tempo de resposta esgotado'
        except Exception as exc:
            return False, str(exc)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _enviar_resultados_refazer(self, trabalhos):
        """Mantem a mesma chamada LPR; devolve resumo para UMA mensagem final."""
        total = len(trabalhos)
        self.progress['maximum'] = max(1, total)
        self.progress['value'] = 0
        enviados = 0
        falhas = []
        for idx, (arquivo, destino, registro_id) in enumerate(trabalhos, 1):
            self.progress_label.configure(text=f'Enviando: {idx}/{total}')
            self.update_idletasks()
            if not os.path.isfile(arquivo):
                ok, detalhe = False, 'PDF final nao encontrado para envio'
            else:
                # Rotina LPR existente preservada; nenhuma nova chamada por quantidade.
                ok, detalhe = self._enviar_refazer_lpr(arquivo, destino)
            if ok:
                enviados += 1
                logging.info('Refazer: trabalho aceito via LPR (%s / %s / %s)',
                             os.path.basename(arquivo), destino['maquina']['nome'], destino['fila'])
            else:
                falha = (f"{os.path.basename(arquivo)} ({destino['maquina']['nome']} / "
                         f"{destino['fila']}): {detalhe}")
                falhas.append(falha)
                logging.error('Refazer - falha LPR: %s', falha)
            try:
                self.repositorio_refazer.registrar(
                    id=registro_id, envio_status='ENVIADO' if ok else 'FALHOU',
                    erro='' if ok else detalhe)
            except Exception:
                logging.exception('Refazer: erro ao registrar retorno do LPR localmente')
                falhas.append('Falha ao registrar o resultado no historico local.')
            self.progress['value'] = idx
        return enviados, falhas


    # ---------- HANDLERS E LÓGICA ----------
    def _limpar_lista_origem(self):
        self.tree_selecao.delete(*self.tree_selecao.get_children())
        self.lista_completa = []
        self._origem_lista = None

    def _atualizar_resumo_fila(self):
        lotes = {item['pasta_lote'] for item in self._fila_refazer.values()}
        clientes = {item['cliente'] for item in self._fila_refazer.values()}
        self._resumo_fila.set(
            f'{len(self._fila_refazer)} item(ns)  ·  {len(lotes)} lote(s)  ·  {len(clientes)} cliente(s)'
            if self._fila_refazer else 'Fila vazia'
        )

    def remover_da_fila(self, _event=None):
        if self._refazer_em_execucao:
            return
        self._finalizar_editor_refazer(cancelar=True)
        selecionados = self.tree_edicao.selection()
        for iid in selecionados:
            self._fila_refazer.pop(iid, None)
        self.tree_edicao.delete(*selecionados)
        self._atualizar_resumo_fila()

    def limpar_fila(self):
        if self._refazer_em_execucao or not self._fila_refazer:
            return
        if messagebox.askyesno('GPrint - Limpar fila', 'Remover todos os itens da fila de retrabalho?', parent=self):
            self._esvaziar_fila()

    def _esvaziar_fila(self):
        self._finalizar_editor_refazer(cancelar=True)
        self.tree_edicao.delete(*self.tree_edicao.get_children())
        self._fila_refazer.clear()
        self._atualizar_resumo_fila()

    def on_cliente_change(self, event=None):
        self._limpar_lista_origem()
        self.atualizar_subpastas()

    def on_lote_change(self, event=None):
        lote = self.subpasta_var.get()
        if not lote:
            return
        self._limpar_lista_origem()
        self.buscar_arquivos()

    def atualizar_subpastas(self, event=None):
        cliente = self.pasta_var.get()
        base = self._clientes_por_pasta.get(cliente)
        self.lotes_disponiveis = self.listar_pastas(base[0], reverse=True) if base else []
        self.busca_lote_var.set("")
        if self.lotes_disponiveis:
            self.cb_lote.current(0)
        else:
            self.cb_lote.set("")

    def _filtrar_lotes(self, *_):
        termo = self.busca_lote_var.get().strip().casefold()
        if not termo:
            lotes = self.lotes_disponiveis
        elif termo.isdecimal():
            lotes = [lote for lote in self.lotes_disponiveis
                     if self._identificacao_lote(lote).startswith(termo)]
            lotes.sort(key=lambda lote: self._identificacao_lote(lote) != termo)
        else:
            lotes = [lote for lote in self.lotes_disponiveis if termo in lote.casefold()]
        self.cb_lote["values"] = lotes

    def _confirmar_busca_lote(self, event=None):
        termo = self.busca_lote_var.get().strip()
        lotes = self.cb_lote["values"]
        exatos = [lote for lote in lotes if self._identificacao_lote(lote) == termo]
        if len(exatos) == 1 or len(lotes) == 1:
            self.cb_lote.set(exatos[0] if len(exatos) == 1 else lotes[0])
            self.on_lote_change()
        elif lotes:
            self.cb_lote.focus_set()
            self.cb_lote.event_generate("<Down>")
        else:
            messagebox.showinfo('GPrint - Refazer', 'Nenhum lote corresponde à busca.', parent=self)
        return "break"

    def carregar_clientes(self, recarregar=False):
        if self._carregando_clientes:
            self._recarregar_clientes |= recarregar
            return
        if self._refazer_em_execucao:
            return
        self._carregando_clientes = True

        def carregar():
            repositorio = self.app.repositorio_equipamentos if self.app else GPrintEquipamentosRepository()
            try:
                # Também funciona se o operador abrir Refazer durante a inicialização.
                repositorio.preparar_banco()
                clientes, avisos = gprint_clientes_das_pastas(repositorio.listar_pastas_refazer())
            except GPrintCadastroErro as erro:
                clientes, avisos = None, [str(erro)]
            self.after(0, self._aplicar_clientes, clientes, avisos)

        threading.Thread(target=carregar, daemon=True).start()

    def _aplicar_clientes(self, clientes, avisos):
        self._carregando_clientes = False
        if self._recarregar_clientes:
            self._recarregar_clientes = False
            self.carregar_clientes()
            return
        if self._refazer_em_execucao:
            return
        if clientes is not None:
            escolhido = self.pasta_var.get()
            anterior = self._clientes_por_pasta.get(escolhido)
            self._clientes_por_pasta = clientes
            self.cb_cliente.configure(values=list(clientes))
            if escolhido not in clientes:
                self.pasta_var.set(next(iter(clientes), ''))
            if self.pasta_var.get() != escolhido or clientes.get(escolhido) != anterior:
                self.on_cliente_change()
        if avisos:
            self.status_var.set(f"{len(avisos)} pasta(s) ou cadastro indisponível. Use Atualizar clientes.")
            for aviso in avisos:
                logging.warning('Refazer: %s', aviso)
        else:
            self.status_var.set(f'{len(clientes)} cliente(s) disponíveis')

    def buscar_arquivos(self):
        self._limpar_lista_origem()
        cliente = self.pasta_var.get()
        lote = self.subpasta_var.get()
        if not cliente or not lote:
            return
        pasta = os.path.join(self._pasta_cliente_refazer(cliente), lote)
        if self._clientes_por_pasta[cliente][1] == 'pdf':
            # Não reutiliza índices de uma planilha carregada anteriormente.
            self.idx_capa = self.idx_miolo = self.idx_caminho = -1
            try:
                self.lista_completa = self.obter_lista_arquivos_da_pasta(pasta)
            except OSError as exc:
                self.lista_completa = []
                for iid in self.tree_selecao.get_children():
                    self.tree_selecao.delete(iid)
                messagebox.showerror('GPrint - Refazer',
                                     f'Não foi possível acessar a pasta do lote:\n{exc}', parent=self)
                return
        else:
            try:
                arquivos_excel = [f for f in os.listdir(pasta) if f.lower().endswith(".xlsx")]
            except OSError as erro:
                messagebox.showerror('GPrint - Refazer', str(erro), parent=self)
                return
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
        for indice, item in enumerate(self.lista_completa):
            self.tree_selecao.insert("", "end", iid=str(indice), values=(item["nome"],))
        self._origem_lista = (cliente, lote, pasta)
        self.tree_selecao.column("Nome", width=min(600, self.tree_selecao.winfo_width()))

    def preparar_processamento(self):
        if self._refazer_em_execucao:
            return
        selecionados = self.tree_selecao.selection()
        if not selecionados or not self._origem_lista:
            messagebox.showinfo("GPrint - Refazer", "Selecione arquivos para adicionar à fila.", parent=self)
            return
        self._finalizar_editor_refazer()
        cliente, lote, pasta_lote = self._origem_lista
        coluna_codigo = 8 if self._cliente_uiclap(cliente) else 3
        chaves = {p['chave'] for p in self._fila_refazer.values()}
        adicionados = 0
        for iid in selecionados:
            item = self.lista_completa[int(iid)]
            linha = item['linha']
            nome = item['nome']
            tipo = item['tipo']
            origem = item.get('caminho_fisico') or os.path.join(pasta_lote, nome)
            chave = (ntpath.normcase(pasta_lote), tipo, ntpath.normcase(origem))
            if chave in chaves:
                continue
            chaves.add(chave)
            codigo = linha[coluna_codigo] if coluna_codigo < len(linha) else item.get('codigo', 'N/A')
            relacionadas = item.get('arquivo_relacionado')
            indice_outro = self.idx_miolo if tipo == 'CAPA' else self.idx_capa
            if not relacionadas and 0 <= indice_outro < len(linha) and linha[indice_outro]:
                pasta = (str(linha[self.idx_caminho]).strip()
                         if 0 <= self.idx_caminho < len(linha) and linha[self.idx_caminho] else pasta_lote)
                relacionadas = os.path.join(pasta, str(linha[indice_outro]).strip())
            opcoes = [e['nome'] for e in self._equipamentos_refazer.values() if e['tipo'] in (tipo, 'AMBOS')]
            maquina = opcoes[0] if len(opcoes) == 1 else (
                '' if self.enviar_refazer_var.get() else self.MAQUINAS_REF[0])
            setor = next(iter(self._setores_refazer), '')
            linha_id = self.tree_edicao.insert('', 'end', values=(
                nome, codigo, '1', setor, maquina, '', cliente, lote, tipo))
            self._fila_refazer[linha_id] = dict(
                item, chave=chave, cliente=cliente, lote=lote, pasta_lote=pasta_lote,
                arquivo_relacionado=relacionadas)
            adicionados += 1
        self._atualizar_resumo_fila()
        repetidos = len(selecionados) - adicionados
        texto = f'{adicionados} item(ns) adicionado(s) à fila.'
        if repetidos:
            texto += f' {repetidos} já estava(m) na fila.'
        self.status_var.set(texto)

    def limpar(self):
        self._limpar_lista_origem()
        self.filtro_var.set("")
        self.operador_var.set('')

    def _aplicar_edicao_em_linhas(self, iids):
        """Atualiza setor/máquina/fila sem misturar filas de equipamentos distintos."""
        self._finalizar_editor_refazer()
        if not iids:
            messagebox.showinfo('GPrint - Refazer', 'Selecione pelo menos uma linha.', parent=self)
            return
        setor = self.cb_setor.get()
        if setor not in self._setores_refazer:
            messagebox.showwarning('GPrint - Refazer', 'Selecione um setor ativo.', parent=self)
            return
        maquina = self.cb_maquina.get()
        fila = self.cb_fila_refazer.get()
        enviar = self.enviar_refazer_var.get()

        # Validação atômica: não aplica somente parte de um conjunto misto.
        if enviar:
            for item in iids:
                tipo = self._tipo_linha_refazer(item)
                try:
                    gprint_validar_destino_refazer(
                        self._equipamentos_refazer, self._filas_refazer,
                        maquina, fila, tipo)
                except ValueError as exc:
                    messagebox.showwarning(
                        'GPrint - Refazer: equipamento e fila',
                        f'{exc}\n\nConfira a máquina e a fila antes de aplicar.', parent=self)
                    return

        for item in iids:
            valores = list(self.tree_edicao.item(item, 'values'))
            valores[3] = setor
            valores[4] = maquina
            # Se o envio estiver desativado, mantém o relatório editável sem
            # exigir cadastro de fila. Não transporta fila de outra máquina.
            filas = self._filas_para_refazer(maquina, self._tipo_linha_refazer(item))
            valores[5] = fila if fila in filas else ''
            self.tree_edicao.item(item, values=valores)

    def aplicar_nos_selecionados(self):
        self._aplicar_edicao_em_linhas(list(self.tree_edicao.selection()))

    def aplicar_em_todos(self):
        self._aplicar_edicao_em_linhas(list(self.tree_edicao.get_children()))

    def _finalizar_editor_refazer(self, cancelar=False, editor_esperado=None):
        """Finaliza a edicao no Treeview com seguranca, inclusive apos trocar de lote.

        As alteracoes de linha acontecem na thread principal do Tkinter.
        Eventos de FocusOut atrasados nao devem recriar nem alterar linhas apagadas.
        """
        estado = getattr(self, '_editor_refazer_ativo', None)
        if not estado:
            return
        widget, iid, indice = estado
        if editor_esperado is not None and widget is not editor_esperado:
            return  # Evento atrasado de outro editor.

        # Desativa a edicao ANTES de destruir o widget (destroy pode gerar FocusOut).
        self._editor_refazer_ativo = None
        try:
            if not cancelar and self.tree_edicao.winfo_exists() and self.tree_edicao.exists(iid):
                # Captura somente a coluna editada. Nao escreve valores antigos de
                # outras colunas, que podem ter mudado com os controles inferiores.
                novo = widget.get()
                atual = list(self.tree_edicao.item(iid, 'values'))
                if len(atual) >= 6:
                    if indice == 2:
                        novo = novo.strip() or '1'
                    if indice == 4 and novo != atual[4]:
                        atual[5] = ''  # Fila da maquina anterior nao e reaproveitada.
                    atual[indice] = novo
                    self.tree_edicao.item(iid, values=atual)
        except tk.TclError:
            # A janela ou a linha pode ter sido removida durante a troca de lote.
            logging.debug('Refazer: editor antigo descartado; linha nao existe mais.')
        finally:
            try:
                if widget.winfo_exists():
                    widget.destroy()
            except tk.TclError:
                pass

        if not cancelar and indice in (4, 5):
            try:
                if self.tree_edicao.winfo_exists():
                    self.after_idle(self._sincronizar_selecao_refazer)
            except tk.TclError:
                pass

    def edicao_inline(self, event):
        if event.state & 0x0005:  # Shift ou Ctrl para selecionar várias linhas.
            return
        """Edicao inteligente com callbacks de salvamento de uso unico."""
        if self.tree_edicao.identify_region(event.x, event.y) != 'cell':
            return
        iid = self.tree_edicao.identify_row(event.y)
        coluna = self.tree_edicao.identify_column(event.x)
        if self._refazer_em_execucao or not iid or not self.tree_edicao.exists(iid):
            return
        try:
            posicao = int(coluna.lstrip('#')) - 1
            if posicao < 0:
                return
            coluna_id = self.tree_edicao['displaycolumns'][posicao]
            indice = self.tree_edicao['columns'].index(coluna_id)
        except (ValueError, IndexError):
            return
        if indice not in (2, 3, 4, 5):
            return

        # Ao abrir outro editor, consolida o anterior ANTES de criar o novo.
        self._finalizar_editor_refazer()
        if not self.tree_edicao.exists(iid):
            return
        valores = list(self.tree_edicao.item(iid, 'values'))
        if len(valores) < 6:
            return
        bbox = self.tree_edicao.bbox(iid, coluna)
        if not bbox:
            return
        x, y, largura, altura = bbox

        if indice == 2:
            widget = tk.Entry(self.tree_edicao, justify='center', font=('Arial', 10, 'bold'))
        else:
            if indice == 3:
                opcoes = list(self._setores_refazer)
            elif indice == 4:
                tipo = self._tipo_linha_refazer(iid)
                opcoes = [e['nome'] for e in self._equipamentos_refazer.values()
                          if e['tipo'] in (tipo, 'AMBOS')]
                if not opcoes and not self.enviar_refazer_var.get():
                    opcoes = list(self.MAQUINAS_REF)
            else:
                opcoes = self._filas_para_refazer(valores[4], self._tipo_linha_refazer(iid))
                if not opcoes:
                    messagebox.showinfo('GPrint - Refazer',
                        'Selecione primeiro uma maquina compativel com filas cadastradas.', parent=self)
                    return
            widget = ttk.Combobox(self.tree_edicao, values=opcoes, state='readonly')

        widget.place(x=x, y=y, width=largura, height=altura)
        if indice == 2:
            widget.insert(0, valores[indice])
        else:
            widget.set(valores[indice])
        self._editor_refazer_ativo = (widget, iid, indice)
        widget.focus_set()

        # O editor pode emitir <<ComboboxSelected>> e FocusOut na mesma edicao.
        # Apenas o primeiro evento salva. Os demais sao ignorados pela identidade.
        finalizar = lambda _event=None, w=widget: self._finalizar_editor_refazer(editor_esperado=w)
        if indice == 2:
            widget.bind('<Return>', finalizar)
        else:
            widget.bind('<<ComboboxSelected>>', finalizar)
        widget.bind('<FocusOut>', finalizar)
        widget.bind('<Escape>', lambda _event=None, w=widget:
                    self._finalizar_editor_refazer(cancelar=True, editor_esperado=w))

    def executar(self):
        """Bloqueia execução duplicada; sempre restaura o botão, mesmo após falha."""
        if self._refazer_em_execucao:
            return
        self._refazer_em_execucao = True
        self.btn_processar.configure(state='disabled')
        try:
            self._executar_refazer_impl()
        except Exception as exc:
            logging.exception('Falha ao preparar o Refazer')
            messagebox.showerror('GPrint - Refazer',
                                 'Não foi possível concluir o processamento.\n' + str(exc),
                                 parent=self)
        finally:
            # Zera inclusive se houver cancelamento ou erro de banco.
            self.progress['value'] = 0
            self.progress_label.configure(text='0/0')
            self._refazer_em_execucao = False
            self.btn_processar.configure(state='normal')

    def mostrar_resumo_perdas(self):
        """Relatorio restrito: consulta e renderizacao somente apos senha correta."""
        if not gprint_autorizar_relatorio(self):
            return
        GPrintJanelaResumoPerdas(self, self.repositorio_refazer, self.app.repositorio_equipamentos)

    def _executar_refazer_impl(self):
        self._finalizar_editor_refazer()
        iids = self.tree_edicao.get_children()
        if not iids:
            messagebox.showinfo('GPrint - Refazer', 'Adicione arquivos à fila antes de processar.', parent=self)
            return
        try:
            for iid in iids:
                valores = self.tree_edicao.item(iid, "values")
                minimo = 1 if self._tipo_linha_refazer(iid) == 'CAPA' else 0
                if not str(valores[2]).isdecimal() or int(valores[2]) < minimo:
                    raise ValueError(f"{valores[0]}: informe uma quantidade de pelo menos {minimo}.")
                if valores[3] not in self._setores_refazer:
                    raise ValueError(f"{valores[0]}: selecione um setor ativo.")
            plano_destinos = self._validar_destinos_refazer()
        except ValueError as exc:
            messagebox.showwarning('GPrint - Refazer: destino de impressão', str(exc), parent=self)
            return
        if plano_destinos:
            if not messagebox.askyesno('GPrint - Confirmar envio do Refazer',
                    f'Serão preparados e enviados {len(plano_destinos)} arquivo(s), '
                    'um trabalho LPR por PDF final.\n\n'
                    'Confirme se máquina e fila estão corretas.\n'
                    'Cada PDF de capa terá a quantidade informada na sua linha.', parent=self):
                return
        operador = self.operador_var.get()
        sangria_miolo = obter_sangria_padrao_miolo()

        if not operador:
            messagebox.showwarning('GPrint - Refazer',
                'Selecione um operador cadastrado em Configurações > Administração central > Operadores do Refazer.',
                parent=self)
            return
        # Inicializa a fila DURAVEL antes de copiar ou montar qualquer PDF.
        try:
            self.repositorio_refazer.preparar_local()
        except (OSError, sqlite3.Error) as exc:
            messagebox.showerror('GPrint - Refazer',
                f'Nao foi possivel preparar o banco LOCAL de retrabalhos.\n{exc}\n'
                'Nenhum processamento foi iniciado.', parent=self)
            return
        sessao_id = str(uuid.uuid4())

        copias_capas = []
        miolos_info = []
        arquivos_para_envio = []
        processados = 0

        total_itens = len(iids)
        destinos_ocupados = set()
        relacionados = []
        self.progress["maximum"] = total_itens
        self.progress["value"] = 0
        self.progress_label.config(text=f"Copiando: 0/{total_itens}")
        self.update_idletasks()

        for idx, iid in enumerate(iids, start=1):
            valores = self.tree_edicao.item(iid, "values")
            nome = valores[0]
            qtd = valores[2]
            setor = valores[3]
            maquina = valores[4]

            item_original = self._fila_refazer[iid]
            cliente, lote, pasta_lote = (item_original[k] for k in ('cliente', 'lote', 'pasta_lote'))
            tipo_atual = item_original['tipo']
            destino_impressao = plano_destinos.get(iid)
            if destino_impressao:
                destino_impressao['lote'] = lote
            tratamento = self._setores_refazer[setor]
            registro_id = self.repositorio_refazer.registrar(
                sessao_id=sessao_id, cliente=cliente, lote=lote,
                arquivo_original=nome, codigo=str(valores[1]), tipo=tipo_atual,
                quantidade=max(0, int(qtd)), setor=setor, operador=operador,
                maquina=maquina, fila_lpr=str(valores[5]),
                montagem_status='PENDENTE',
                envio_status='PENDENTE' if plano_destinos.get(iid) else 'DESATIVADO')
            caminho_fisico = item_original['caminho_fisico']
            if not caminho_fisico or not os.path.exists(caminho_fisico):
                caminho_fisico = os.path.join(pasta_lote, nome)
            chave_destino = (ntpath.normcase(pasta_lote), tipo_atual,
                             ntpath.normcase(os.path.basename(caminho_fisico)))
            subpasta = f'{sessao_id[:8]}_{idx:03d}' if chave_destino in destinos_ocupados else None
            destinos_ocupados.add(chave_destino)
            destino = self.copiar_para_refazer(caminho_fisico, pasta_lote, tipo_atual, subpasta)
            self.repositorio_refazer.registrar(
                id=registro_id, montagem_status='COPIADO' if destino else 'FALHOU',
                pdf_resultado=destino or '',
                erro='' if destino else 'Arquivo original nao localizado ou copia falhou')

            if tipo_atual == "CAPA" and destino and destino.lower().endswith('.pdf'):
                copias_capas.append((destino, os.path.splitext(nome)[0], tratamento, destino_impressao, registro_id, int(qtd), lote))
            elif tipo_atual == "MIOLO" and destino and destino.lower().endswith('.pdf'):
                miolos_info.append((
                    destino,
                    os.path.splitext(nome)[0],
                    item_original.get("acabamento"),
                    item_original.get("formato_miolo"),
                    tratamento,
                    destino_impressao,
                    registro_id,
                    int(qtd),
                    lote
                ))

            tipo_outro = 'MIOLO' if tipo_atual == 'CAPA' else 'CAPA'
            relacionado = item_original.get('arquivo_relacionado')
            if relacionado:
                relacionados.append((relacionado, pasta_lote, tipo_outro))

            self.progress["value"] = idx
            self.progress_label.config(text=f"Copiando: {idx}/{total_itens}")
            self.update_idletasks()

        for origem, pasta_lote, tipo in relacionados:
            chave_destino = (ntpath.normcase(pasta_lote), tipo, ntpath.normcase(os.path.basename(origem)))
            if chave_destino not in destinos_ocupados:
                self.copiar_para_refazer(origem, pasta_lote, tipo)
                destinos_ocupados.add(chave_destino)

        # Etapa 2: marca d'água nas capas
        total_capas = len(copias_capas)
        if total_capas > 0:
            self.progress["maximum"] = total_capas
            self.progress["value"] = 0
            self.progress_label.config(text=f"Capas: 0/{total_capas}")
            self.update_idletasks()
            for idx, (destino, nome_base, tratamento, destino_impressao, registro_id, quantidade, lote) in enumerate(copias_capas, start=1):
                try:
                    self.adicionar_marca_dagua(destino, f'LOTE {self._identificacao_lote(lote)} - {nome_base}')
                    if tratamento == "curvas":
                        temp_outlines = os.path.splitext(destino)[0] + '_TEMP_OUTLINES.pdf'
                        self.converter_em_contornos(destino, temp_outlines)
                        os.replace(temp_outlines, destino)
                    if quantidade > 1:
                        with tempfile.TemporaryDirectory(prefix="capas_", dir=os.path.dirname(destino)) as pasta:
                            agrupado = os.path.join(pasta, "capas.pdf")
                            gprint_mesclar_capas([(destino, quantidade)], agrupado)
                            os.replace(agrupado, destino)
                    self.repositorio_refazer.registrar(
                        id=registro_id, montagem_status='PREPARADO', pdf_resultado=destino)
                    processados += 1
                    if destino_impressao and os.path.isfile(destino):
                        arquivos_para_envio.append((destino, destino_impressao, registro_id))
                except Exception as e:
                    self.repositorio_refazer.registrar(
                        id=registro_id, montagem_status='FALHOU', erro=str(e))
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

            for idx, (copia_dest, nome_base, acab, fmt_miolo, tratamento, destino_impressao, registro_id, quantidade, lote) in enumerate(miolos_info, start=1):
                self.progress_label.config(text=f"Miolo {idx}/{total_miolos}: {nome_base}")
                self.update_idletasks()
                try:
                    arquivo_final_refazer = copia_dest
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
                        if tratamento == "comprimir_miolo":
                            nome_final = os.path.splitext(copia_dest)[0] + '_MONTADO_REDUZIDO.pdf'
                            reduzido_temp = os.path.splitext(copia_dest)[0] + '_TEMP_RED.pdf'
                            self.comprimir_pdf(copia_dest, reduzido_temp)
                            self.montar_pdf(reduzido_temp, nome_final, tipo_montagem,
                                            quantidade=quantidade, lote=lote, nome_original=nome_base + '.pdf',
                                            sangria_mm=sangria_miolo)
                            if os.path.exists(reduzido_temp): os.remove(reduzido_temp)
                        elif tratamento == "curvas":
                            nome_final = os.path.splitext(copia_dest)[0] + '_MONTADO_CURVAS.pdf'
                            montado_temp = os.path.splitext(copia_dest)[0] + '_TEMP_MONT.pdf'
                            self.montar_pdf(copia_dest, montado_temp, tipo_montagem,
                                            quantidade=quantidade, lote=lote, nome_original=nome_base + '.pdf',
                                            sangria_mm=sangria_miolo)
                            self.converter_em_contornos(montado_temp, nome_final)
                            if os.path.exists(montado_temp): os.remove(montado_temp)
                        else:
                            nome_final = os.path.splitext(copia_dest)[0] + '_MONTADO.pdf'
                            self.montar_pdf(copia_dest, nome_final, tipo_montagem,
                                            quantidade=quantidade, lote=lote, nome_original=nome_base + '.pdf',
                                            sangria_mm=sangria_miolo)

                        if os.path.exists(copia_dest):
                            os.remove(copia_dest)
                        arquivo_final_refazer = nome_final
                    else:
                        if tratamento == "curvas":
                            nome_final = os.path.splitext(copia_dest)[0] + '_CURVAS.pdf'
                            self.converter_em_contornos(copia_dest, nome_final)
                            if os.path.exists(copia_dest): os.remove(copia_dest)
                            arquivo_final_refazer = nome_final
                        print(f"Miolo {nome_base}: sem tipo de montagem definido (acabamento: {acab}), mantido como cópia simples.")
                    self.repositorio_refazer.registrar(
                        id=registro_id, montagem_status='MONTADO',
                        pdf_resultado=arquivo_final_refazer)
                    processados += 1
                    if destino_impressao and os.path.isfile(arquivo_final_refazer):
                        arquivos_para_envio.append((arquivo_final_refazer, destino_impressao, registro_id))
                except Exception as e:
                    self.repositorio_refazer.registrar(
                        id=registro_id, montagem_status='FALHOU', erro=str(e))
                    print(f"Erro ao processar miolo {nome_base}: {e}")

                self.progress["value"] = idx
                self.update_idletasks()

        self._esvaziar_fila()
        enviados, falhas = (0, [])
        if plano_destinos:
            enviados, falhas = self._enviar_resultados_refazer(arquivos_para_envio)
            if len(arquivos_para_envio) < len(plano_destinos):
                falhas.append(f'{len(plano_destinos)-len(arquivos_para_envio)} arquivo(s) sem PDF pronto para envio.')
        # Registros novos e pendentes de execucoes anteriores sao sincronizados.
        sincronizados, erro_sincronizacao = self.repositorio_refazer.sincronizar()
        pendentes = self.repositorio_refazer.pendentes()
        self.progress['value'] = 0
        self.progress_label.configure(text='0/0')
        self.status_var.set('Pronto para o próximo Refazer')
        if plano_destinos:
            resumo = (f'{processados} arquivo(s) preparados.\n'
                      f'{enviados}/{len(arquivos_para_envio)} envio(s) aceitos pelo LPR.\n')
            if falhas:
                resumo += f'{len(falhas)} falha(s) no envio/registro.\n' + '\n'.join(falhas[:5])
        else:
            resumo = f'{processados} arquivo(s) preparados.\nEnvio às impressoras desativado.'
        if pendentes:
            resumo += (f'\n\nATENÇÃO: {pendentes} registro(s) permanecem no computador.'
                       '\nVerifique a rede da produção e informe o Guilherme.')
        else:
            resumo += f'\n\nHistórico de retrabalho sincronizado ({sincronizados} registro(s)).'
        (messagebox.showwarning if falhas or pendentes else messagebox.showinfo)(
            'GPrint - Resultado do Refazer', resumo, parent=self)

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

    if sys.platform == 'win32':
        try:
            resultado = ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                ctypes.c_wchar_p('Docuprint.GPrint')
            )
            if resultado != 0:
                logging.warning("Falha ao identificar GPrint na barra de tarefas: %s", resultado)
        except (AttributeError, OSError) as erro:
            logging.warning("Falha ao identificar GPrint na barra de tarefas: %s", erro)

    if verificar_data_limite():
        root = TkinterDnD.Tk()
        gprint_aplicar_icone(root)
        style = Style(theme=CONFIG.get('tema', 'darkly'))
        app = Application(root)
        root.mainloop()
