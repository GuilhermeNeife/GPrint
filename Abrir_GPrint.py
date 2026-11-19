"""Lançador do GPrint para Windows. Publique o EXE ao lado das versões do GPrint."""

import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import re
import queue
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk


PADRAO = re.compile(r"GPrint_(\d{8})_(\d{6})_(\d{3})\.exe", re.IGNORECASE)
PASTA = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
ESTADO = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "GPrint" / "versao_notificada.txt"


def listar_versoes(pasta):
    versoes = [p for p in pasta.glob("GPrint_*.exe") if PADRAO.fullmatch(p.name) and p.is_file()]
    return sorted(versoes, key=lambda p: PADRAO.fullmatch(p.name).groups(), reverse=True)


def proteger_executavel(arquivo):
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    abrir = kernel.CreateFileW
    abrir.argtypes = (wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                      wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE)
    abrir.restype = wintypes.HANDLE
    fechar = kernel.CloseHandle
    fechar.argtypes = (wintypes.HANDLE,)
    fechar.restype = wintypes.BOOL
    # Acesso somente para leitura, compartilhando leitura e sem compartilhar exclusão.
    handle = abrir(str(arquivo), 0x80000000, 1, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value:
        raise ctypes.WinError(ctypes.get_last_error())
    return handle, fechar


def limpar_antigas(versoes):
    ocupadas, indisponiveis = [], []
    for antiga in versoes:
        try:
            antiga.unlink(missing_ok=True)
        except OSError as erro:
            if getattr(erro, "winerror", None) in (32, 33, 1224):
                ocupadas.append(antiga.name)
            else:
                indisponiveis.append(antiga.name)
    return ocupadas, indisponiveis


def memorizar_versao(versao):
    ESTADO.parent.mkdir(parents=True, exist_ok=True)
    temporario = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=ESTADO.parent,
                                         delete=False) as arquivo:
            temporario = Path(arquivo.name)
            arquivo.write(versao)
        temporario.replace(ESTADO)
    finally:
        if temporario:
            temporario.unlink(missing_ok=True)


def janelas_gprint():
    """Identifica janelas principais já visíveis no desktop desta estação."""
    user32 = ctypes.WinDLL("user32")
    callback = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows.argtypes = (callback, wintypes.LPARAM)
    user32.IsWindowVisible.argtypes = (wintypes.HWND,)
    user32.GetWindowTextW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
    encontradas = set()

    @callback
    def visitar(handle, _dados):
        if user32.IsWindowVisible(handle):
            titulo = ctypes.create_unicode_buffer(32)
            user32.GetWindowTextW(handle, titulo, len(titulo))
            if titulo.value == "GPrint":
                encontradas.add(handle)
        return True

    user32.EnumWindows(visitar, 0)
    return encontradas


def main():
    janela = tk.Tk()
    janela.withdraw()
    janela.title("Iniciando GPrint")
    janela.resizable(False, False)
    janela.configure(bg="#111a2c")
    quadro = tk.Frame(janela, bg="#111a2c", padx=32, pady=26)
    quadro.pack(fill="both", expand=True)
    recursos = Path(getattr(sys, "_MEIPASS", PASTA))
    for pasta in dict.fromkeys((recursos, PASTA)):
        try:
            janela.iconbitmap(default=str(pasta / "GPrint.ico"))
            break
        except tk.TclError:
            continue
    logo = None
    for pasta in dict.fromkeys((recursos, PASTA)):
        try:
            imagem = tk.PhotoImage(file=str(pasta / "GPrint.png"))
            escala = max(1, (max(imagem.width(), imagem.height()) + 159) // 160)
            logo = imagem.subsample(escala, escala)
            break
        except tk.TclError:
            continue
    if logo:
        tk.Label(quadro, image=logo, bg="#111a2c").pack(pady=(0, 12))
    tk.Label(quadro, text="GPrint", bg="#111a2c", fg="#f5f7fb",
             font=("Segoe UI", 24, "bold")).pack()
    destaque = tk.StringVar(value="Verificando versões")
    status = tk.StringVar(value="Consultando a pasta de versões na rede…")
    aviso = tk.StringVar()
    tk.Label(quadro, textvariable=destaque, bg="#111a2c", fg="#8ec5ff",
             font=("Segoe UI", 13, "bold")).pack(pady=(10, 0))
    tk.Label(quadro, textvariable=status, bg="#111a2c", fg="#e1e8f4",
             font=("Segoe UI", 10), wraplength=405, justify="center").pack(pady=8)
    barra = ttk.Progressbar(quadro, mode="indeterminate", length=390)
    barra.pack(pady=(5, 12))
    tk.Label(quadro, textvariable=aviso, bg="#111a2c", fg="#ffc879",
             font=("Segoe UI", 9), wraplength=405, justify="center").pack()

    janela.update_idletasks()
    largura, altura = 500, max(340, janela.winfo_reqheight() + 48)
    x = max(0, (janela.winfo_screenwidth() - largura) // 2)
    y = max(0, (janela.winfo_screenheight() - altura) // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")
    janela.attributes("-topmost", True)
    janela.deiconify()
    barra.start(12)
    eventos = queue.SimpleQueue()

    def abrir():
        protecao = None
        processo = None
        try:
            if sys.platform != "win32":
                raise OSError("Execute o lançador no Windows.")
            for tentativa in range(3):
                versoes = listar_versoes(PASTA)
                if not versoes:
                    raise OSError(f"Nenhuma versão publicada foi encontrada em {PASTA}.")
                atual = versoes[0]
                try:
                    protecao = proteger_executavel(atual)
                    break
                except OSError as erro:
                    if tentativa == 2 or getattr(erro, "winerror", None) not in (2, 3, 32):
                        raise

            try:
                anterior = ESTADO.read_text(encoding="utf-8").strip()
            except (OSError, UnicodeError):
                anterior = ""
            nova = anterior != atual.name
            eventos.put(("versao", atual.name, nova))
            ocupadas, indisponiveis = limpar_antigas(versoes[1:])
            if ocupadas:
                eventos.put(("aviso", "Versões anteriores estão bloqueadas. Feche e reabra o GPrint após terminar o lote."))
            elif indisponiveis:
                eventos.put(("aviso", "Algumas versões anteriores ficaram na rede. A exclusão depende das permissões."))

            anteriores = janelas_gprint()
            ambiente = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")
            processo = subprocess.Popen([str(atual)], cwd=PASTA, env=ambiente)
            # O processo inicial do PyInstaller ainda precisa extrair e carregar o aplicativo.
            inicio = time.monotonic()
            tempo_minimo = 4 if nova or ocupadas or indisponiveis else 0.7
            janela_visivel = False
            while processo.poll() is None:
                decorrido = time.monotonic() - inicio
                janela_visivel = janela_visivel or bool(janelas_gprint() - anteriores)
                if (janela_visivel and decorrido >= tempo_minimo) or decorrido >= 40:
                    break
                time.sleep(0.2)

            if processo.poll() is not None and not janela_visivel:
                raise OSError(f"O GPrint encerrou durante a abertura (código {processo.returncode}).")
            if nova and janela_visivel:
                try:
                    memorizar_versao(atual.name)
                except OSError:
                    pass
            eventos.put(("ocultar",))
            processo.wait()
        except OSError as erro:
            eventos.put(("erro", str(erro)))
        finally:
            if processo is not None and processo.poll() is None:
                processo.wait()
            if protecao:
                handle, fechar = protecao
                fechar(handle)
            eventos.put(("fim",))

    def atualizar():
        try:
            while True:
                evento, *dados = eventos.get_nowait()
                if evento == "versao":
                    nome, nova = dados
                    destaque.set("Nova versão disponível" if nova else "Versão atual")
                    status.set(f"{nome}\nVerificando versões antigas e iniciando o GPrint…")
                elif evento == "aviso":
                    aviso.set(dados[0])
                elif evento == "ocultar":
                    barra.stop()
                    janela.withdraw()
                elif evento == "erro":
                    barra.stop()
                    janela.deiconify()
                    messagebox.showerror("GPrint · Falha ao abrir", dados[0], parent=janela)
                elif evento == "fim":
                    janela.destroy()
                    return
        except queue.Empty:
            janela.after(100, atualizar)

    # O lançador protege o EXE enquanto ele roda, inclusive após ocultar a tela.
    janela.protocol("WM_DELETE_WINDOW", janela.withdraw)
    janela.after(100, atualizar)
    threading.Thread(target=abrir, daemon=False).start()
    janela.mainloop()


if __name__ == "__main__":
    main()
