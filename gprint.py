import os
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, END
import fitz  # PyMuPDF
from PyPDF2 import PdfMerger
import subprocess
import logging
from datetime import datetime
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style
from typing import List
from tkinterdnd2 import DND_FILES, TkinterDnD
import json
from tkinter import colorchooser
import sys
import os
import tempfile
from datetime import datetime
from tkinter import ttk
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from pikepdf import Pdf, Rectangle
from reportlab.lib.pagesizes import A3, A4
from reportlab.pdfgen import canvas
import io
import math

# --- CONSTANTES ---
DATA_LIMITE = "31/07/2027"
PASTA_BASE = r"N:\____REPOSITORIO GERAL POD\UICLAP\LOTES"
PASTA_TRABALHO = os.path.join(os.environ.get('TEMP', os.path.expanduser("~")), "UICLAP_TEMP")
CONFIG_FILE = os.path.expanduser("~/uiclap_config.json")

# Lista de temas do ttkbootstrap
TEMAS_DISPONIVEIS = [
    "cosmo", "flatly", "journal", "litera", "lumen", "minty", "pulse",
    "sandstone", "united", "yeti", "morph", "darkly", "superhero",
    "solar", "cyborg", "vapor"]

DEFAULT_CONFIG = {
    'cor_fosco': (0, 0, 1),
    'cor_brilho': (1, 0.7, 0.1),
    'fonte': 'helv',
    'tamanho_fonte': 11,
    'deslocamento_padrao': -30,
    'deslocamento_capa_28x56': 92,
    'deslocamento_x': 70 * 2.83465,
    'tema': 'darkly',
    'formatos_pdf': [
        {'nome': 'Padrão 1', 'tamanho_entrada': (150, 210), 'tamanho_saida': (220, 320)},
        {'nome': 'Padrão 2', 'tamanho_entrada': (160, 230), 'tamanho_saida': (240, 320)},
        {'nome': 'Padrão 3', 'tamanho_entrada': (170, 240), 'tamanho_saida': (252, 362)},
        {'nome': 'Grande 1', 'tamanho_entrada': (220, 310), 'tamanho_saida': (320, 450)},
        {'nome': 'Grande 2', 'tamanho_entrada': (235, 310), 'tamanho_saida': (320, 480)}
    ]
}
# --- FUNÇÕES DE UTILIDADE ---
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
        with open(caminho_txt, "r", encoding="utf-8") as arquivo:
            for idx, linha in enumerate(arquivo, 1):
                campos = linha.strip().split("|")
                if len(campos) >= 14:
                    try:
                        dados.append({
                            "CAMINHO": campos[0].strip(),
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

def adicionar_marca_dagua(pdf_path, laminacao, codigoacab, identific,filacolor):
    try:
        doc = fitz.open(pdf_path)
        nome_arquivo = os.path.splitext(os.path.basename(pdf_path))[0]

        cor_marca = CONFIG['cor_fosco'] if laminacao.upper() == "FOSCO" else CONFIG['cor_brilho']

        if codigoacab == "3":
            deslocamento_y = CONFIG['deslocamento_padrao']+8
        elif codigoacab == "2":
            deslocamento_y = CONFIG['deslocamento_padrao'] + 12
        else:
            deslocamento_y = CONFIG['deslocamento_capa_28x56']

        for pagina in doc:
            largura, altura = pagina.rect.width, pagina.rect.height

            posicao_topo = fitz.Rect(
                50 + CONFIG['deslocamento_x'], 
                30 + deslocamento_y,  
                largura - 50 + CONFIG['deslocamento_x'], 
                70 + deslocamento_y
            )

            texto_marca = f"{nome_arquivo} - {identific}"
            pagina.insert_textbox(
                posicao_topo, texto_marca,
                fontsize=CONFIG['tamanho_fonte'],
                color=cor_marca,
                fontname=CONFIG['fonte'],
                align=1
            )
            if filacolor == "Capas_28X56":                
                # CORTAR 4CM DA PARTE DE CIMA DO PDF
                corte_cm = 4
                corte_pontos = corte_cm * 28.3465  # Converter cm para pontos (1 cm = 28.3465 pontos)
                
                # Definir nova área da página (cortar a parte superior)
                novo_retangulo = fitz.Rect(
                    0,                      # x0 (esquerda)
                    corte_pontos,           # y0 (topo - cortando 4cm)
                    largura,                # x1 (direita)
                    altura                  # y1 (base)
                )
                
                # Aplicar o corte definindo o cropbox
                pagina.set_cropbox(novo_retangulo)

        novo_caminho = pdf_path.replace(".pdf", "_marcado.pdf")
        doc.save(novo_caminho)
        doc.close()
        return novo_caminho

    except Exception as e:
        logging.error(f"Erro ao adicionar marca d'água em {pdf_path}: {e}")
        messagebox.showerror("Erro", f"Falha ao marcar o arquivo {os.path.basename(pdf_path)}")
        return None

# --- CLASSES DA INTERFACE ---
class ProgressDialog(ttk.Toplevel):
    def __init__(self, parent, title="Processando"):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x150")
        self.resizable(False, False)
        
        self.center_window()

        self.progress = ttk.Progressbar(
            self, 
            orient=HORIZONTAL, 
            length=460, 
            mode=DETERMINATE,
            bootstyle=PRIMARY
        )
        self.progress.pack(pady=(20, 10), padx=20)

        self.label = ttk.Label(
            self, 
            text="0% concluído", 
            anchor=W, 
            width=60,
            bootstyle=SECONDARY
        )
        self.label.pack(padx=20)

        self.timer_label = ttk.Label(
            self,
            text="Tempo decorrido: 0 segundos",
            bootstyle=SECONDARY
        )
        self.timer_label.pack(pady=(0, 10))

        self.grab_set()
        self.focus_force()

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        self.geometry(f"+{x}+{y}")

    def update_progress(self, value, text=None):
        self.progress['value'] = value
        if text:
            self.label.config(text=text)
        else:
            self.label.config(text=f"{int(value)}% concluído")
        self.update_idletasks()

    def update_timer(self, seconds):
        self.timer_label.config(text=f"Tempo decorrido: {seconds} segundos")

class ConfigTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Configurações de cores
        colors_frame = ttk.Labelframe(
            scrollable_frame, 
            text="Configurações de Cores", 
            padding=10,
            bootstyle=PRIMARY
        )
        colors_frame.pack(fill=X, pady=5)
        
        ttk.Label(
            colors_frame, 
            text="Cor Fosco:", 
            width=20
        ).grid(row=0, column=0, padx=5, pady=5, sticky=W)
        
        self.cor_fosco_btn = ttk.Button(
            colors_frame,
            text="Selecionar",
            command=lambda: self.selecionar_cor('cor_fosco'),
            bootstyle=INFO
        )
        self.cor_fosco_btn.grid(row=0, column=1, padx=5, pady=5)
        
        self.cor_fosco_preview = tk.Frame(
            colors_frame, 
            width=50, 
            height=25,
            relief=SOLID,
            bg="#0000FF"
        )
        self.cor_fosco_preview.grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(
            colors_frame, 
            text="Cor Brilho:", 
            width=20
        ).grid(row=1, column=0, padx=5, pady=5, sticky=W)
        
        self.cor_brilho_btn = ttk.Button(
            colors_frame,
            text="Selecionar",
            command=lambda: self.selecionar_cor('cor_brilho'),
            bootstyle=INFO
        )
        self.cor_brilho_btn.grid(row=1, column=1, padx=5, pady=5)
        
        self.cor_brilho_preview = tk.Frame(
            colors_frame, 
            width=50, 
            height=25,
            relief=SOLID,
            bg="#FFB300"
        )
        self.cor_brilho_preview.grid(row=1, column=2, padx=5, pady=5)
        
        # Configurações de fonte
        font_frame = ttk.Labelframe(
            scrollable_frame, 
            text="Configurações de Fonte", 
            padding=10,
            bootstyle=PRIMARY
        )
        font_frame.pack(fill=X, pady=5)
        
        ttk.Label(
            font_frame, 
            text="Nome da Fonte:", 
            width=20
        ).grid(row=0, column=0, padx=5, pady=5, sticky=W)
        
        self.fonte_entry = ttk.Entry(font_frame)
        self.fonte_entry.grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        
        ttk.Label(
            font_frame, 
            text="Tamanho da Fonte:", 
            width=20
        ).grid(row=1, column=0, padx=5, pady=5, sticky=W)
        
        self.tamanho_fonte_spin = ttk.Spinbox(
            font_frame,
            from_=6,
            to=24,
            increment=1
        )
        self.tamanho_fonte_spin.grid(row=1, column=1, padx=5, pady=5, sticky=W)
        
        # Configurações de deslocamento
        offset_frame = ttk.Labelframe(
            scrollable_frame, 
            text="Configurações de Deslocamento", 
            padding=10,
            bootstyle=PRIMARY
        )
        offset_frame.pack(fill=X, pady=5)
        
        ttk.Label(
            offset_frame, 
            text="Deslocamento Padrão (Y):", 
            width=25
        ).grid(row=0, column=0, padx=5, pady=5, sticky=W)
        
        self.deslocamento_padrao_spin = ttk.Spinbox(
            offset_frame,
            from_=-100,
            to=100,
            increment=1
        )
        self.deslocamento_padrao_spin.grid(row=0, column=1, padx=5, pady=5, sticky=W)
        
        ttk.Label(
            offset_frame, 
            text="Deslocamento 28x56 (Y):", 
            width=25
        ).grid(row=1, column=0, padx=5, pady=5, sticky=W)
        
        self.deslocamento_capa_spin = ttk.Spinbox(
            offset_frame,
            from_=-100,
            to=100,
            increment=1
        )
        self.deslocamento_capa_spin.grid(row=1, column=1, padx=5, pady=5, sticky=W)
        
        ttk.Label(
            offset_frame, 
            text="Deslocamento X:", 
            width=25
        ).grid(row=2, column=0, padx=5, pady=5, sticky=W)
        
        self.deslocamento_x_spin = ttk.Spinbox(
            offset_frame,
            from_=0,
            to=500,
            increment=1
        )
        self.deslocamento_x_spin.grid(row=2, column=1, padx=5, pady=5, sticky=W)
        
        # Configurações de Formatos PDF
        formatos_frame = ttk.Labelframe(
            scrollable_frame,
            text="Formatos de PDF Automáticos",
            padding=10,
            bootstyle=PRIMARY
        )
        formatos_frame.pack(fill=X, pady=5)

        # Treeview para exibir os formatos
        columns = ('nome', 'entrada', 'saida')
        self.formatos_tree = ttk.Treeview(
            formatos_frame,
            columns=columns,
            show='headings',
            height=5
        )
        
        self.formatos_tree.heading('nome', text='Nome')
        self.formatos_tree.heading('entrada', text='Entrada (mm)')
        self.formatos_tree.heading('saida', text='Saída (mm)')
        
        self.formatos_tree.column('nome', width=120)
        self.formatos_tree.column('entrada', width=100)
        self.formatos_tree.column('saida', width=100)
        
        self.formatos_tree.pack(fill=X, pady=5)

        # Frame para controles de edição
        edit_frame = ttk.Frame(formatos_frame)
        edit_frame.pack(fill=X, pady=5)

        # Campos para edição
        ttk.Label(edit_frame, text="Nome:").grid(row=0, column=0, padx=5, sticky=W)
        self.fmt_nome = ttk.Entry(edit_frame, width=15)
        self.fmt_nome.grid(row=0, column=1, padx=5, sticky=W)

        ttk.Label(edit_frame, text="Entrada (LxA mm):").grid(row=0, column=2, padx=5, sticky=W)
        self.fmt_entrada_larg = ttk.Spinbox(edit_frame, from_=50, to=500, width=5)
        self.fmt_entrada_larg.grid(row=0, column=3, padx=2, sticky=W)
        ttk.Label(edit_frame, text="x").grid(row=0, column=4)
        self.fmt_entrada_alt = ttk.Spinbox(edit_frame, from_=50, to=500, width=5)
        self.fmt_entrada_alt.grid(row=0, column=5, padx=2, sticky=W)

        ttk.Label(edit_frame, text="Saída (LxA mm):").grid(row=0, column=6, padx=5, sticky=W)
        self.fmt_saida_larg = ttk.Spinbox(edit_frame, from_=50, to=600, width=5)
        self.fmt_saida_larg.grid(row=0, column=7, padx=2, sticky=W)
        ttk.Label(edit_frame, text="x").grid(row=0, column=8)
        self.fmt_saida_alt = ttk.Spinbox(edit_frame, from_=50, to=600, width=5)
        self.fmt_saida_alt.grid(row=0, column=9, padx=2, sticky=W)

        # Botões de ação
        btn_frame = ttk.Frame(edit_frame)
        btn_frame.grid(row=1, column=0, columnspan=10, pady=5)

        ttk.Button(
            btn_frame,
            text="Adicionar",
            command=self.adicionar_formato,
            bootstyle=INFO,
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            btn_frame,
            text="Editar",
            command=self.editar_formato,
            bootstyle=WARNING,
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            btn_frame,
            text="Remover",
            command=self.remover_formato,
            bootstyle=DANGER,
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            btn_frame,
            text="Resetar",
            command=self.resetar_formatos,
            bootstyle=SECONDARY,
            width=10
        ).pack(side=LEFT, padx=2)

        # Botões de ação principais
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.pack(fill=X, pady=10)
        
        ttk.Button(
            btn_frame,
            text="Salvar Configurações",
            command=self.salvar_configuracoes,
            bootstyle=SUCCESS,
            width=20
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="Restaurar Padrões",
            command=self.restaurar_padroes,
            bootstyle=WARNING,
            width=20
        ).pack(side=LEFT, padx=5)
        
        self.carregar_configuracoes()
    
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
        
        # Formatos PDF
        self.formatos_tree.delete(*self.formatos_tree.get_children())
        for formato in CONFIG['formatos_pdf']:
            entrada = f"{formato['tamanho_entrada'][0]}x{formato['tamanho_entrada'][1]}"
            saida = f"{formato['tamanho_saida'][0]}x{formato['tamanho_saida'][1]}"
            self.formatos_tree.insert('', END, values=(formato['nome'], entrada, saida))
    
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
    
    def adicionar_formato(self):
        try:
            nome = self.fmt_nome.get()
            entrada_larg = int(self.fmt_entrada_larg.get())
            entrada_alt = int(self.fmt_entrada_alt.get())
            saida_larg = int(self.fmt_saida_larg.get())
            saida_alt = int(self.fmt_saida_alt.get())

            novo_formato = {
                'nome': nome,
                'tamanho_entrada': (entrada_larg, entrada_alt),
                'tamanho_saida': (saida_larg, saida_alt)
            }

            CONFIG['formatos_pdf'].append(novo_formato)
            self.carregar_configuracoes()
            
            # Limpar campos
            self.fmt_nome.delete(0, END)
            self.fmt_entrada_larg.delete(0, END)
            self.fmt_entrada_alt.delete(0, END)
            self.fmt_saida_larg.delete(0, END)
            self.fmt_saida_alt.delete(0, END)

        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira valores numéricos válidos")

    def editar_formato(self):
        selected = self.formatos_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um formato para editar")
            return

        try:
            # Obter dados atuais
            item = self.formatos_tree.item(selected[0])
            nome = item['values'][0]
            
            # Encontrar o formato na lista
            idx = next(i for i, f in enumerate(CONFIG['formatos_pdf']) if f['nome'] == nome)
            formato = CONFIG['formatos_pdf'][idx]

            # Atualizar com novos valores
            formato['nome'] = self.fmt_nome.get() or formato['nome']
            formato['tamanho_entrada'] = (
                int(self.fmt_entrada_larg.get() or formato['tamanho_entrada'][0]),
                int(self.fmt_entrada_alt.get() or formato['tamanho_entrada'][1])
            )
            formato['tamanho_saida'] = (
                int(self.fmt_saida_larg.get() or formato['tamanho_saida'][0]),
                int(self.fmt_saida_alt.get() or formato['tamanho_saida'][1])
            )

            self.carregar_configuracoes()

        except (ValueError, StopIteration):
            messagebox.showerror("Erro", "Não foi possível editar o formato")

    def remover_formato(self):
        selected = self.formatos_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um formato para remover")
            return

        item = self.formatos_tree.item(selected[0])
        nome = item['values'][0]
        
        CONFIG['formatos_pdf'] = [f for f in CONFIG['formatos_pdf'] if f['nome'] != nome]
        self.carregar_configuracoes()

    def resetar_formatos(self):
        if messagebox.askyesno("Confirmar", "Resetar todos os formatos para os padrões?"):
            CONFIG['formatos_pdf'] = DEFAULT_CONFIG['formatos_pdf'].copy()
            self.carregar_configuracoes()
    
    def salvar_configuracoes(self):
        global CONFIG
        
        try:
            CONFIG['fonte'] = self.fonte_entry.get()
            CONFIG['tamanho_fonte'] = int(self.tamanho_fonte_spin.get())
            
            CONFIG['deslocamento_padrao'] = int(self.deslocamento_padrao_spin.get())
            CONFIG['deslocamento_capa_28x56'] = int(self.deslocamento_capa_spin.get())
            CONFIG['deslocamento_x'] = float(self.deslocamento_x_spin.get())
            
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
        self.impressora_var = tk.StringVar()
        self.arquivo_selecionado = None
        self.setup_ui()

    def setup_ui(self):
        self.status_var = tk.StringVar(value="Status: aguardando ação...")  # ← Adicionado aqui
        main_frame = ttk.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Frame de configurações
        config_frame = ttk.LabelFrame(main_frame, text="Configurações", bootstyle="INFO")
        config_frame.pack(fill='x', pady=5)

        ttk.Label(config_frame, text="Impressora:").grid(row=0, column=0, padx=5, pady=5, sticky='e')

        self.combo_impressora = ttk.Combobox(
            config_frame,
            textvariable=self.impressora_var,
            values=[
                "OCÉ 1 - 192.168.200.80",
                "OCÉ 2 - 192.168.200.81",
                "OCÉ 3 - 192.168.200.85"
            ],
            state="readonly",
            bootstyle="INFO"
        )
        self.combo_impressora.current(0)
        self.combo_impressora.grid(row=0, column=1, padx=5, pady=5, sticky='we')

        # Frame de arquivo
        file_frame = ttk.LabelFrame(main_frame, text="Arquivo TXT", bootstyle="INFO")
        file_frame.pack(fill='x', pady=5)

        self.btn_selecionar = ttk.Button(
            file_frame,
            text="Selecionar TXT",
            command=self.selecionar_txt,
            bootstyle="INFO"
        )
        self.btn_selecionar.pack(pady=5)

        self.file_label = ttk.Label(
            file_frame,
            text="Clique para selecionar um arquivo .txt",
            relief="groove",
            anchor="center",
            padding=10
        )
        self.file_label.pack(fill='x', pady=5)

        # Visualização do conteúdo
        ttk.Label(file_frame, text="Visualização do conteúdo:").pack(anchor='w', pady=(10, 0))
        self.txt_view = tk.Text(
            file_frame,
            height=10,
            wrap='word',
            state='disabled'
        )
        scrollbar = ttk.Scrollbar(file_frame, command=self.txt_view.yview)
        self.txt_view.configure(yscrollcommand=scrollbar.set)

        self.txt_view.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Botão de processamento
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)

        self.btn_gerar = ttk.Button(
            btn_frame,
            text="Processar TXT e Gerar PDFs",
            command=self.gerar_pdf,
            bootstyle="INFO",
            
        )
        self.btn_gerar.pack(pady=5, ipadx=20, ipady=5)
    

    # Métodos que você deve implementar conforme a lógica do seu projeto
    def selecionar_txt(self):
        """Seleciona um arquivo TXT usando diálogo de arquivo"""
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos TXT", "*.txt")])
        if file_path:
            self.arquivo_txt = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.ler_arquivo_txt_miolo()

    def ler_arquivo_txt_miolo(self):
        """Lê e processa o arquivo TXT selecionado"""
        try:
            self.processos = []  # Limpa lista de processos
            with open(self.arquivo_txt, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                self.txt_view.config(state='normal')
                self.txt_view.delete(1.0, tk.END)
                
                for idx, linha in enumerate(lines):
                    linha = linha.strip()
                    if not linha:
                        continue
                    
                    campos = linha.split('|')
                    
                    if len(campos) < 13:
                        self.txt_view.insert(tk.END, f"Linha {idx+1}: Formato inválido - ignorada\n")
                        continue
                    
                    processo = {
                        "LINHA": idx+1,
                        "CAMINHO": campos[0].strip(),
                        "ARQUIVO": campos[1].strip(),
                        "QTD": int(campos[2].strip()) if campos[2].strip().isdigit() else 1,
                        "CODIGOACAB": campos[9].strip(),
                        "NOMEFIC": campos[11].strip(),
                        "FORMATO": campos[12].strip()
                    }
                    
                    # Extrair dimensões
                    formato = processo["FORMATO"].split()[0]
                    try:
                        largura, altura = map(int, formato.lower().split('x'))
                        processo.update({
                            "LARGURA_MM": largura * 10,
                            "ALTURA_MM": altura * 10,
                            "ESPACAMENTO": 3 if f"{largura}x{altura}" == "22x32" else 0
                        })
                    except:
                        processo.update({
                            "LARGURA_MM": 480,  # Default A4
                            "ALTURA_MM": 320,
                            "ESPACAMENTO": 0
                        })
                        self.txt_view.insert(tk.END, f"Linha {idx+1}: Formato inválido - usando padrão A4\n")
                    
                    self.processos.append(processo)
                    self.txt_view.insert(
                        tk.END, 
                        f"Linha {idx+1}: {processo['ARQUIVO']} | Qtd: {processo['QTD']} | " +
                        f"Acab: {processo['CODIGOACAB']} | Formato: {processo['FORMATO']}\n"
                    )

            self.status_var.set(f"✅ Arquivo TXT carregado com {len(self.processos)} processos válidos")
            self.txt_view.config(state='disabled')

        except Exception as e:
            self.status_var.set(f"❌ Erro ao ler arquivo TXT: {str(e)}")
            
    def gerar_pdf(self):
        """Processa os arquivos e gera os PDFs conforme configuração"""
        if not self.processos:
            self.status_var.set("❌ Nenhum processo válido para gerar PDF!")
            return

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        temp_dir = os.path.join(desktop_path, "PDFs_Processados")
        os.makedirs(temp_dir, exist_ok=True)

        success_count = 0
        error_count = 0
        
        self.txt_view.config(state='normal')
        self.txt_view.insert(tk.END, "\nIniciando processamento...\n")
        self.txt_view.config(state='disabled')
        self.winfo_toplevel().update()
 # Atualiza a interface

        for processo in self.processos:
            try:
                caminho_pdf = os.path.join(processo["CAMINHO"], processo["ARQUIVO"])
                
                if not os.path.exists(caminho_pdf):
                    self.atualizar_log(f"Linha {processo['LINHA']}: Arquivo não encontrado - {caminho_pdf}")
                    error_count += 1
                    continue

                # Determinar modo
                codigo_acab = processo["CODIGOACAB"]
                qtd = processo["QTD"]

                if codigo_acab == "1":
                    modo = "corta_e_empilha" if qtd == 1 else "2_na_folha"
                elif codigo_acab == "2":
                    modo = "grampeado"
                elif codigo_acab == "3":
                    modo = "1_na_folha"
                else:
                    modo = "corta_e_empilha"

                # Nome do arquivo de saída
                output_name = f"{processo['NOMEFIC']}_MONTADO.pdf"
                output_path = os.path.join(temp_dir, output_name)

                # Converter dimensões para pontos
                output_width = processo["LARGURA_MM"] * 2.83494
                output_height = processo["ALTURA_MM"] * 2.83494
                espacamento = processo["ESPACAMENTO"]

                # Processar conforme o modo
                if modo == "2_na_folha":
                    create_2up_pdf(
                        caminho_pdf, output_path,
                        output_width=output_width,
                        output_height=output_height,
                        espacamento_mm=espacamento
                    )
                elif modo == "grampeado":
                    create_booklet_pdf(
                        caminho_pdf, output_path,
                        output_width=output_width,
                        output_height=output_height
                    )
                elif modo == "corta_e_empilha":
                    create_stack_pdf(
                        caminho_pdf, output_path,
                        output_width=output_width,
                        output_height=output_height
                    )
                elif modo == "1_na_folha":
                    create_1up_pdf(
                        caminho_pdf, output_path,
                        output_width=output_width,
                        output_height=output_height,
                        centralizado=True,
                        add_cropmarks=True,
                        redimensionar=False
                    )

                success_count += 1
                self.atualizar_log(f"Linha {processo['LINHA']}: PDF gerado - {output_name}")

                # Enviar para impressora
                try:
                    
                    impressora_selecionada = self.impressora_var.get()
                    ip = impressora_selecionada.split("-")[1].strip()
                    formato = processo["FORMATO"]
                    nome_fic = processo["NOMEFIC"]

                    
                    comando = f'lpr -S {ip} -P "{formato}" "{output_path}"'

                    print(f"Comando: {comando}")
                    print(f"ip: {ip}")
                    retorno = os.system(comando)
                    
                    if retorno == 0:
                        self.atualizar_log(f"Linha {processo['LINHA']}: Enviado para {impressora_selecionada}")
                    else:
                        self.atualizar_log(f"Linha {processo['LINHA']}: ❌ Falha no envio para impressora (código {retorno})")
                except Exception as e:
                    self.atualizar_log(f"Linha {processo['LINHA']}: ❌ Erro ao enviar para impressora: {str(e)}")

            except Exception as e:
                error_count += 1
                self.atualizar_log(f"Linha {processo['LINHA']}: ❌ Erro - {str(e)}")

        self.status_var.set(f"✅ Processamento completo! {success_count} sucesso(s), {error_count} erro(s)")
        messagebox.showinfo("Processamento Concluído", 
                          f"Processamento finalizado com {success_count} sucesso(s) e {error_count} erro(s).\n"
                          f"Arquivos salvos em: {temp_dir}")
    
    def atualizar_log(self, mensagem):
        """Atualiza o log de mensagens"""
        self.txt_view.config(state='normal')
        self.txt_view.insert(tk.END, mensagem + "\n")
        self.txt_view.see(tk.END)
        self.txt_view.config(state='disabled')
        self.winfo_toplevel().update()  # Atualiza a interface

def create_1up_pdf(input_path, output_path, output_width, output_height, 
                  centralizado=True, pos_horizontal="Centro", pos_vertical="Centro",
                  margin_mm=10, add_cropmarks=True, redimensionar=False):
    """Cria PDF com posicionamento avançado para impressão frente e verso"""
    margin = margin_mm * 2.83494  # Converter margem para pontos
    
    with Pdf.open(input_path) as pdf:
        output_pdf = Pdf.new()
        
        for idx, page in enumerate(pdf.pages):
            # Obter dimensões da página original
            trim_box = page.trimbox or page.cropbox or page.mediabox
            page_width = float(trim_box[2]) - float(trim_box[0])
            page_height = float(trim_box[3]) - float(trim_box[1])
            
            # Calcular escala se necessário
            scale = 1.0
            if redimensionar:
                scale_w = (output_width - 2*margin) / page_width
                scale_h = (output_height - 2*margin) / page_height
                scale = min(scale_w, scale_h)
                page_width *= scale
                page_height *= scale
            
            # Cálculo da posição
            if centralizado:
                x = (output_width - page_width) / 2
                y = (output_height - page_height) / 2
            else:
                # Posicionamento horizontal
                if pos_horizontal == "Esquerda":
                    x = margin
                elif pos_horizontal == "Direita":
                    x = output_width - page_width - margin
                else:  # Centro
                    x = (output_width - page_width) / 2
                
                # Posicionamento vertical
                if pos_vertical == "Topo":
                    y = output_height - page_height - margin
                elif pos_vertical == "Base":
                    y = margin
                else:  # Centro
                    y = (output_height - page_height) / 2
            
            # Alternância para impressão frente/verso (quando Esquerda/Direita)
            if pos_horizontal in ["Esquerda", "Direita"] and idx % 2 == 1:
                x = output_width - x - page_width  # Inverte a posição para páginas ímpares
                
            # Criar nova página
            new_page = output_pdf.add_blank_page(page_size=(output_width, output_height))
            new_page.add_overlay(page, Rectangle(x, y, x + page_width, y + page_height))
            
            # Adicionar marcas de corte se necessário
            if add_cropmarks:
                crop_pdf = generate_crop_marks(
                    page_width=output_width,
                    page_height=output_height,
                    rects=[(x, y, page_width, page_height)],
                    espacamento_mm=0,  # Não usado no modo 1-up
                    deslocamento_mm=0,  # Não usado no modo 1-up
                    grossura_mm=0.5,    # Espessura padrão
                    mark_length_mm=5,   # Comprimento padrão
                    offset_marks_mm=2,  # Deslocamento padrão
                    is_1up=True         # Modo 1 página por folha
                )
                new_page.add_overlay(crop_pdf.pages[0])
        
        output_pdf.save(output_path)

def create_2up_pdf(input_path, output_path, output_width, output_height, espacamento_mm=5):
    """
    Cria PDF com 2 páginas idênticas por folha (1|1, 2|2) com verificação rigorosa de formatos
    Parâmetros:
    - input_path: caminho do PDF de entrada
    - output_path: caminho do PDF de saída
    - output_width: largura desejada em pontos
    - output_height: altura desejada em pontos
    - espacamento_mm: espaço entre as páginas em milímetros (padrão: 5mm)
    """
    espacamento = espacamento_mm * 2.83494  # Converter mm para pontos

    with Pdf.open(input_path) as pdf:
        num_pages = len(pdf.pages)
        first_page = pdf.pages[0]
        trim_box = first_page.trimbox or first_page.cropbox or first_page.mediabox
        page_width = float(trim_box[2]) - float(trim_box[0])
        page_height = float(trim_box[3]) - float(trim_box[1])

        # Converter para milímetros para mensagens mais amigáveis
        width_mm = page_width / 2.83494
        height_mm = page_height / 2.83494
        required_width_mm = (width_mm * 2) + espacamento_mm
        required_height_mm = height_mm
        
        print(f"\n=== Informações da Página ===")
        print(f"Tamanho original: {width_mm:.2f}x{height_mm:.2f} mm")
        print(f"Espaçamento entre páginas: {espacamento_mm} mm")
        print(f"Necessário mínimo: {required_width_mm:.2f}x{required_height_mm:.2f} mm")

        # Forçar orientação paisagem se necessário
        if output_width < output_height:
            output_width, output_height = output_height, output_width

        output_width_mm = output_width / 2.83494
        output_height_mm = output_height / 2.83494
        print(f"Formato selecionado: {output_width_mm:.2f}x{output_height_mm:.2f} mm")

        # Verificação rigorosa do espaço necessário
        total_width_pt = (page_width * 2) + espacamento
        if total_width_pt > output_width or page_height > output_height:
            error_msg = [
                "\n❌ ERRO: As páginas não cabem no formato selecionado",
                f"- Tamanho necessário: {required_width_mm:.2f}x{required_height_mm:.2f} mm",
                f"- Formato selecionado: {output_width_mm:.2f}x{output_height_mm:.2f} mm",
                "\nSoluções possíveis:",
                f"1. Reduza o espaçamento (atual: {espacamento_mm} mm)",
                "2. Use um formato de saída maior",
                f"3. Reduza o tamanho das páginas para no máximo {(output_width-espacamento)/2/2.83494:.2f}x{output_height/2.83494:.2f} mm"
            ]
            raise ValueError("\n".join(error_msg))
        else:
            print("✅ Verificação de tamanho: OK (páginas cabem no formato)")

        # Calcular posições centralizadas
        center_x_offset = (output_width - total_width_pt) / 2
        center_y_offset = (output_height - page_height) / 2

        # Gerar marcas de corte
        crop_pdf = generate_crop_marks(
            output_width, output_height,
            [
                (center_x_offset, center_y_offset, page_width, page_height),
                (center_x_offset + page_width + espacamento, center_y_offset, page_width, page_height)
            ],
            is_1up=False
        )
        print("✅ Marcas de corte geradas")

        output_pdf = Pdf.new()

        print(f"\n=== Processando Páginas ===")
        print(f"Total de páginas: {num_pages}")
        print(f"Folhas geradas: {num_pages} (2 cópias por página)")

        for i in range(num_pages):  # 1 folha por página original
            new_page = output_pdf.add_blank_page(page_size=(output_width, output_height))

            # Página esquerda
            new_page.add_overlay(
                pdf.pages[i],
                Rectangle(
                    center_x_offset,
                    center_y_offset,
                    center_x_offset + page_width,
                    center_y_offset + page_height
                )
            )

            # Página direita (mesma página)
            new_page.add_overlay(
                pdf.pages[i],
                Rectangle(
                    center_x_offset + page_width + espacamento,
                    center_y_offset,
                    center_x_offset + (page_width * 2) + espacamento,
                    center_y_offset + page_height
                )
            )

            # Adicionar marcas de corte
            new_page.add_overlay(crop_pdf.pages[0])

        output_pdf.save(output_path)
        print(f"\n✅ PDF 2-up criado com sucesso em: {output_path}")
        print(f"Formato final: {output_width_mm:.2f}x{output_height_mm:.2f} mm")
        return True

def create_booklet_pdf(input_path, output_path, output_width, output_height, add_cropmarks=True, redimensionar=False):
    """
    Cria PDF no formato grampeado (booklet) com verificação rigorosa de formatos
    Parâmetros:
    - input_path: caminho do PDF de entrada
    - output_path: caminho do PDF de saída
    - output_width: largura desejada em pontos (1/72 de polegada)
    - output_height: altura desejada em pontos
    - add_cropmarks: se True, adiciona marcas de corte
    - redimensionar: se False, verifica rigorosamente o tamanho sem redimensionar
    """
    with Pdf.open(input_path) as pdf:
        num_pages = len(pdf.pages)
        first_page = pdf.pages[0]
        trim_box = first_page.trimbox or first_page.cropbox or first_page.mediabox
        page_width = float(trim_box[2]) - float(trim_box[0])
        page_height = float(trim_box[3]) - float(trim_box[1])

        # Converter para milímetros para mensagens mais amigáveis
        width_mm = page_width / 2.83494
        height_mm = page_height / 2.83494
        required_width_mm = width_mm * 2  # 2 páginas lado a lado
        required_height_mm = height_mm
        
        # Forçar orientação paisagem
        if output_width < output_height:
            output_width, output_height = output_height, output_width

        # Verificação rigorosa do formato
        output_width_mm = output_width / 2.83494
        output_height_mm = output_height / 2.83494
        
        print(f"\n=== Verificação de Formato ===")
        print(f"Tamanho da página original: {width_mm:.2f}x{height_mm:.2f} mm")
        print(f"Necessário para booklet: {required_width_mm:.2f}x{required_height_mm:.2f} mm")
        print(f"Formato selecionado: {output_width_mm:.2f}x{output_height_mm:.2f} mm")

        if not redimensionar:
            # Verificação rigorosa sem redimensionamento
            required_width_pt = page_width * 2
            if (required_width_pt > output_width or page_height > output_height):
                error_msg = [
                    "\n❌ ERRO: As páginas não cabem no formato selecionado sem redimensionar",
                    f"- Tamanho necessário: {required_width_mm:.2f}x{required_height_mm:.2f} mm",
                    f"- Formato selecionado: {output_width_mm:.2f}x{output_height_mm:.2f} mm",
                    "\nSoluções possíveis:",
                    "1. Use um formato de saída maior",
                    "2. Ative a opção 'redimensionar' (se apropriado)",
                    "3. Reduza o tamanho das páginas originais"
                ]
                raise ValueError("\n".join(error_msg))
            else:
                print("✅ Verificação de tamanho: OK (cabe sem redimensionar)")
        else:
            print("ℹ️ Modo redimensionar ativado - verificações de tamanho ignoradas")

        # Calcular posições centralizadas
        total_width = page_width * 2
        center_x_offset = (output_width - total_width) / 2
        center_y_offset = (output_height - page_height) / 2

        # Gerar marcas de corte se necessário
        crop_pdf = None
        if add_cropmarks:
            crop_pdf = generate_crop_marks(
                output_width, output_height,
                [
                    (center_x_offset, center_y_offset, page_width, page_height),
                    (center_x_offset + page_width, center_y_offset, page_width, page_height)
                ],
                is_1up=False
            )
            print("✅ Marcas de corte geradas")

        output_pdf = Pdf.new()

        # Calcular número total de páginas (múltiplo de 4)
        total_sheets = math.ceil(num_pages / 4) * 4
        print(f"\n=== Configuração do Booklet ===")
        print(f"Páginas originais: {num_pages}")
        print(f"Páginas totais (arredondado para múltiplo de 4): {total_sheets}")
        print(f"Folhas necessárias: {total_sheets // 2} (frente e verso)")

        left = 0  # Página inicial
        right = total_sheets - 1  # Última página

        while left < right:
            # Página da frente (ex: 32 | 1)
            new_page = output_pdf.add_blank_page(page_size=(output_width, output_height))

            if right < num_pages:
                new_page.add_overlay(
                    pdf.pages[right],
                    Rectangle(
                        center_x_offset,
                        center_y_offset,
                        center_x_offset + page_width,
                        center_y_offset + page_height
                    )
                )

            if left < num_pages:
                new_page.add_overlay(
                    pdf.pages[left],
                    Rectangle(
                        center_x_offset + page_width,
                        center_y_offset,
                        center_x_offset + (page_width * 2),
                        center_y_offset + page_height
                    )
                )

            right -= 1
            left += 1

            # Página do verso (ex: 2 | 31)
            new_page_back = output_pdf.add_blank_page(page_size=(output_width, output_height))

            if left < num_pages:
                new_page_back.add_overlay(
                    pdf.pages[left],
                    Rectangle(
                        center_x_offset,
                        center_y_offset,
                        center_x_offset + page_width,
                        center_y_offset + page_height
                    )
                )

            if right < num_pages:
                new_page_back.add_overlay(
                    pdf.pages[right],
                    Rectangle(
                        center_x_offset + page_width,
                        center_y_offset,
                        center_x_offset + (page_width * 2),
                        center_y_offset + page_height
                    )
                )

            right -= 1
            left += 1

            # Adicionar marcas de corte se necessário
            if add_cropmarks:
                new_page.add_overlay(crop_pdf.pages[0])
                new_page_back.add_overlay(crop_pdf.pages[0])

        output_pdf.save(output_path)
        print(f"\n✅ Booklet criado com sucesso em: {output_path}")
        print(f"Formato final: {output_width_mm:.2f}x{output_height_mm:.2f} mm")
        return True

def create_stack_pdf(input_path, output_path, output_width=None, output_height=None):
    """
    Cria PDF no formato corta e empilha com fallback automático para formatos configurados
    - Primeiro tenta usar o formato especificado
    - Se não couber, busca no DEFAULT_CONFIG um formato adequado
    - Se ainda não encontrar, calcula o formato mínimo necessário
    """
    with Pdf.open(input_path) as pdf:
        num_pages = len(pdf.pages)
        first_page = pdf.pages[0]
        trim_box = first_page.trimbox or first_page.cropbox or first_page.mediabox
        page_width = float(trim_box[2]) - float(trim_box[0])
        page_height = float(trim_box[3]) - float(trim_box[1])

        # Converter para milímetros para comparação
        width_mm = page_width / 2.83494
        height_mm = page_height / 2.83494
        required_width_mm = width_mm * 2  # Stack precisa de 2 páginas lado a lado
        required_height_mm = height_mm
        
        print(f"\n=== Informações da Página ===")
        print(f"Tamanho original: {width_mm:.2f}x{height_mm:.2f} mm")
        print(f"Necessário mínimo: {required_width_mm:.2f}x{required_height_mm:.2f} mm")

        # Função para encontrar o melhor formato configurado
        def encontrar_formato_configurado():
            for formato in sorted(DEFAULT_CONFIG['formatos_pdf'], 
                                key=lambda x: x['tamanho_saida'][0]):  # Ordena por largura
                saida_w, saida_h = formato['tamanho_saida']
                if (saida_w >= required_width_mm and 
                    saida_h >= required_height_mm):
                    return formato
            return None

        # Verificar formato especificado (se fornecido)
        formato_selecionado = None
        if output_width is not None and output_height is not None:
            print(f"\n=== Tentando Formato Especificado ===")
            print(f"Formato fornecido: {output_width/2.83494:.2f}x{output_height/2.83494:.2f} mm")
            
            # Forçar orientação paisagem
            if output_width < output_height:
                output_width, output_height = output_height, output_width

            # Verificar se cabe
            if ((page_width * 2) <= output_width and 
                page_height <= output_height):
                print("✅ Formato especificado é adequado")
                formato_selecionado = {
                    'nome': 'Especificado pelo usuário',
                    'tamanho_saida': (output_width/2.83494, output_height/2.83494)
                }
            else:
                print("❌ Formato especificado não comporta as páginas")

        # Se não tem formato adequado ainda, buscar nos configurados
        if formato_selecionado is None:
            print("\n=== Buscando Formato Configurado ===")
            formato_config = encontrar_formato_configurado()
            
            if formato_config:
                formato_selecionado = formato_config
                print(f"✅ Formato configurado selecionado: {formato_config['nome']}")
            else:
                print("ℹ️ Nenhum formato configurado adequado encontrado")

        # Se ainda não tem formato, calcular automático
        if formato_selecionado is None:
            print("\n=== Calculando Formato Automático ===")
            # Adiciona 2mm de margem e arredonda para múltiplos de 5mm
            output_width_mm = math.ceil(required_width_mm + 2)
            output_height_mm = math.ceil(required_height_mm + 2)
            
            output_width_mm = math.ceil(output_width_mm / 5) * 5
            output_height_mm = math.ceil(output_height_mm / 5) * 5
            
            formato_selecionado = {
                'nome': 'Calculado automaticamente',
                'tamanho_saida': (output_width_mm, output_height_mm)
            }
            print(f"⚙️ Formato calculado: {output_width_mm}x{output_height_mm} mm")

        # Definir dimensões finais
        saida_w, saida_h = formato_selecionado['tamanho_saida']
        output_width = saida_w * 2.83494
        output_height = saida_h * 2.83494
        
        print(f"\n=== Formato Final Selecionado ===")
        print(f"Nome: {formato_selecionado['nome']}")
        print(f"Dimensões: {saida_w}x{saida_h} mm")
        print(f"Pontos: {output_width:.2f}x{output_height:.2f} pt")

        # Verificação final rigorosa
        required_width_pt = page_width * 2
        if (required_width_pt > output_width or 
            page_height > output_height):
            error_msg = [
                "\n❌ ERRO CRÍTICO: As páginas não cabem no formato selecionado",
                f"- Formato selecionado: {saida_w}x{saida_h} mm",
                f"- Necessário mínimo: {required_width_mm:.2f}x{required_height_mm:.2f} mm",
                "\nSugestões:",
                "1. Especifique um formato maior manualmente",
                "2. Adicione um formato adequado no DEFAULT_CONFIG",
                f"3. Reduza o tamanho da página para no máximo {saida_w/2:.1f}x{saida_h:.1f} mm"
            ]
            raise ValueError("\n".join(error_msg))

        # Calcular posições centralizadas
        center_x_offset = (output_width - required_width_pt) / 2
        center_y_offset = (output_height - page_height) / 2

        # Gerar marcas de corte
        crop_pdf = generate_crop_marks(
            output_width, output_height,
            [
                (center_x_offset, center_y_offset, page_width, page_height),
                (center_x_offset + page_width, center_y_offset, page_width, page_height)
            ],
            is_1up=False
        )

        output_pdf = Pdf.new()
        
        # Quantidade total de folhas (cada folha tem 4 páginas: frente e verso)
        num_folhas = math.ceil(num_pages / 4)
        total_paginas_frente = num_folhas * 2  # metade do documento será frente (páginas ímpares)

        print(f"\n=== Processando Páginas ===")
        print(f"Total de páginas: {num_pages}")
        print(f"Folhas necessárias: {num_folhas}")

        for i in range(num_folhas):
            # === FRENTE ===
            new_page = output_pdf.add_blank_page(page_size=(output_width, output_height))

            # Página esquerda (normal): 1, 3, 5...
            pagina_normal = i * 2
            if pagina_normal < num_pages:
                new_page.add_overlay(
                    pdf.pages[pagina_normal],
                    Rectangle(
                        center_x_offset,
                        center_y_offset,
                        center_x_offset + page_width,
                        center_y_offset + page_height
                    )
                )

            # Página direita (invertida): 51, 53, 55... (virada 180°)
            pagina_invertida = total_paginas_frente + i * 2
            if pagina_invertida < num_pages:
                inverted_page = pdf.pages[pagina_invertida]
                inverted_page.Rotate = 180
                new_page.add_overlay(
                    inverted_page,
                    Rectangle(
                        center_x_offset + page_width,
                        center_y_offset,
                        center_x_offset + (page_width * 2),
                        center_y_offset + page_height
                    )
                )

            # Marcas de corte
            new_page.add_overlay(crop_pdf.pages[0])

            # === VERSO ===
            verso_page = output_pdf.add_blank_page(page_size=(output_width, output_height))

            # Página esquerda (invertida): 52, 54, 56... (virada 180°)
            pagina_invertida_verso = total_paginas_frente + i * 2 + 1
            if pagina_invertida_verso < num_pages:
                inverted_page_verso = pdf.pages[pagina_invertida_verso]
                inverted_page_verso.Rotate = 180
                verso_page.add_overlay(
                    inverted_page_verso,
                    Rectangle(
                        center_x_offset,
                        center_y_offset,
                        center_x_offset + page_width,
                        center_y_offset + page_height
                    )
                )

            # Página direita (normal): 2, 4, 6...
            pagina_normal_verso = i * 2 + 1
            if pagina_normal_verso < num_pages:
                verso_page.add_overlay(
                    pdf.pages[pagina_normal_verso],
                    Rectangle(
                        center_x_offset + page_width,
                        center_y_offset,
                        center_x_offset + (page_width * 2),
                        center_y_offset + page_height
                    )
                )

            # Marcas de corte
            verso_page.add_overlay(crop_pdf.pages[0])


        output_pdf.save(output_path)
        print(f"\n✅ PDF criado com sucesso em: {output_path}")
        print(f"Formato final: {saida_w}x{saida_h} mm")
        return True

def generate_crop_marks(
    page_width, page_height, rects, 
    espacamento_mm=0, 
    deslocamento_mm=0, 
    grossura_mm=0.5, 
    mark_length_mm=5,
    offset_marks_mm=2,  
    is_1up=False
):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    # Conversão de unidades (1 mm ≈ 2.83494 pontos)
    mark_length = mark_length_mm * 2.83494
    deslocamento_pt = deslocamento_mm * 2.83494
    grossura_pt = grossura_mm * 2.83494
    espacamento_pt = espacamento_mm * 2.83494
    ajuste_offset = offset_marks_mm * 2.83494

    # Configuração da espessura da linha
    c.setLineWidth(grossura_pt)
    c.setStrokeColorRGB(0, 0, 0)

    for i, (x, y, w, h) in enumerate(rects):
        # Aplica deslocamento (funciona para ambos os modos)
        x += deslocamento_pt
        y += deslocamento_pt

        # Marcas VERTICAIS com ajuste
        c.line(x, y - 2 * 2.83494, x, y - mark_length)  # Inferior esquerdo
        c.line(x, y + h + ajuste_offset, x, y + h + mark_length)  # Superior esquerdo ajustado
        c.line(x + w, y - 2 * 2.83494, x + w, y - mark_length)  # Inferior direito
        c.line(x + w, y + h + ajuste_offset, x + w, y + h + mark_length)  # Superior direito ajustado

        # Marcas HORIZONTAIS com ajuste
        if is_1up:
            c.line(x - ajuste_offset, y + h, x - mark_length, y + h)  # Superior esquerda ajustada
            c.line(x + w + ajuste_offset, y + h, x + w + mark_length, y + h)  # Superior direita ajustada
            c.line(x - ajuste_offset, y, x - mark_length, y)  # Inferior esquerda ajustada
            c.line(x + w + ajuste_offset, y, x + w + mark_length, y)  # Inferior direita ajustada
        else:
            if i == 0:  # Primeira página (marca esquerda)
                c.line(x - ajuste_offset, y + h, x - mark_length, y + h)  # Superior
                c.line(x - ajuste_offset, y, x - mark_length, y)  # Inferior
            
            if i == len(rects) - 1:  # Última página (marca direita)
                c.line(x + w + ajuste_offset, y + h, x + w + mark_length, y + h)  # Superior
                c.line(x + w + ajuste_offset, y, x + w + mark_length, y)  # Inferior
    
    c.save()
    packet.seek(0)
    return Pdf.open(packet)

    pass

class Application:
    def __init__(self, root):
        self.root = root
        self.arquivos_selecionados = []
        self.imprimir_automaticamente = tk.BooleanVar(value=True)
        self.tema_atual = tk.StringVar(value=CONFIG.get('tema', 'darkly'))
        self.style = Style(theme=self.tema_atual.get())
        self.setup_ui()
        self.setup_drag_and_drop()
    
    def setup_ui(self):
        self.root.title("Sistema de Processamento de Arquivos")
        self.root.geometry("900x700")
        self.center_window()
        
        self.criar_menu_superior()
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        # Aba Principal
        self.aba_principal = ttk.Frame(self.notebook)
        self.notebook.add(self.aba_principal, text="Capa")
        self.criar_interface_principal(self.aba_principal)
        
        # Aba Miolo
        self.aba_miolo = MioloTab(self.notebook, self)
        self.notebook.add(self.aba_miolo, text="Miolo")
        
        # Aba Configurações
        self.aba_config = ConfigTab(self.notebook, self)
        self.notebook.add(self.aba_config, text="Configurações")
    
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
        self.lbl_status.config(text=f"{len(txt_files)} arquivo(s) selecionado(s)")
        
        self.listbox_arquivos.delete(0, tk.END)
        for arquivo in txt_files:
            self.listbox_arquivos.insert(tk.END, os.path.basename(arquivo))

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 2
        self.root.geometry(f'+{x}+{y}')

    def criar_menu_superior(self):
        menubar = tk.Menu(self.root)
        
        menu_tema = tk.Menu(menubar, tearoff=0)
        for tema in TEMAS_DISPONIVEIS:
            menu_tema.add_radiobutton(
                label=tema.capitalize(),
                variable=self.tema_atual,
                value=tema,
                command=self.mudar_tema
            )
        menubar.add_cascade(label="Temas", menu=menu_tema)
        
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
        bg = self.style.colors.get('bg') if self.style.theme.type == 'dark' else self.style.colors.get('light')
        fg = self.style.colors.get('fg') if self.style.theme.type == 'dark' else self.style.colors.get('dark')
        selectbg = self.style.colors.get('primary')
        selectfg = self.style.colors.get('bg') if self.style.theme.type == 'dark' else self.style.colors.get('light')
        
        self.listbox_arquivos.config(
            bg=bg,
            fg=fg,
            selectbackground=selectbg,
            selectforeground=selectfg
        )

    def criar_interface_principal(self, parent):
        header_frame = ttk.Frame(parent, padding=20)
        header_frame.pack(fill=X, pady=(0, 15))
        
        ttk.Label(
            header_frame, 
            text="PROCESSADOR DE CAPAS UICLAP", 
            font=('Helvetica', 14, 'bold'),
            bootstyle=(PRIMARY),
            
        ).pack()

        instrucoes_frame = ttk.Frame(parent)
        instrucoes_frame.pack(fill=X, pady=(0, 20))
        
        ttk.Label(
            instrucoes_frame,
            text="Selecione arquivos TXT de uma destas formas:",
            font=('Helvetica', 10)
        ).pack(anchor=W)
        
        btn_selecionar = ttk.Button(
            instrucoes_frame,
            text="Selecionar Arquivos TXT",
            command=self.selecionar_arquivos_manual,
            bootstyle=(PRIMARY, OUTLINE),
            width=22
        )
        btn_selecionar.pack(side=LEFT, padx=(0, 10), pady=5)
        
        ttk.Label(
            instrucoes_frame,
            text="OU arraste arquivos TXT para esta janela",
            font=('Helvetica', 10)
        ).pack(side=LEFT, pady=5)
        
        selection_frame = ttk.Labelframe(
            parent, 
            text="Arquivos Selecionados", 
            bootstyle=PRIMARY,
            padding=15
        )
        selection_frame.pack(fill=BOTH, expand=True)
        
        self.listbox_arquivos = tk.Listbox(
            selection_frame,
            selectmode=EXTENDED,
            height=10,
            font=('Helvetica', 10),
            exportselection=False
        )
        self.listbox_arquivos.pack(fill=BOTH, expand=True, padx=5, pady=(0, 10))
        
        
        self.checkbox_imprimir = ttk.Checkbutton(
            selection_frame, 
            text="Imprimir automaticamente após processar",
            variable=self.imprimir_automaticamente,
            bootstyle="primary-round-toggle"
        )
        self.checkbox_imprimir.pack(anchor=W, pady=(10, 0))
        
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=X, pady=(15, 5))
        
        self.btn_processar = ttk.Button(
            btn_frame, 
            text="PROCESSAR ARQUIVOS SELECIONADOS", 
            command=self.iniciar_processamento,
            bootstyle=(PRIMARY, OUTLINE),
            width=40
        )
        self.btn_processar.pack(pady=10, ipady=5)
        
        self.status_frame = ttk.Frame(parent, padding=(10, 5))
        self.status_frame.pack(fill=X, padx=10, pady=(0, 10))
        
        self.lbl_status = ttk.Label(
            self.status_frame, 
            text="Pronto para processar", 
            anchor=W,
            font=('Helvetica', 9),
            bootstyle=SECONDARY
        )
        self.lbl_status.pack(fill=X)

    def iniciar_processamento(self):
        if not verificar_data_limite():
            return
            
        if not self.arquivos_selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um arquivo para processar!")
            return
        
        self.btn_processar.config(state=DISABLED)
        
        thread = threading.Thread(target=self.processar_arquivos)
        thread.start()
        
        self.verificar_thread(thread)

    def verificar_thread(self, thread):
        if thread.is_alive():
            self.root.after(100, lambda: self.verificar_thread(thread))
        else:
            self.btn_processar.config(state=NORMAL)

    def processar_arquivos(self):
        pasta_temp = os.path.join(PASTA_TRABALHO, f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(pasta_temp, exist_ok=True)
        
        try:
            progress = ProgressDialog(self.root)
            progress.update_progress(0, "Preparando processamento...")
            
            start_time = time.time()
            
            def update_timer():
                elapsed = int(time.time() - start_time)
                progress.update_timer(elapsed)
                if progress.winfo_exists():
                    progress.after(1000, update_timer)
            
            update_timer()
            
            pdfs_para_mesclar = {}
            total_arquivos = sum(len(ler_arquivo_txt(txt)) for txt in self.arquivos_selecionados)
            arquivos_processados = 0

            for arquivo_txt in self.arquivos_selecionados:
                dados = ler_arquivo_txt(arquivo_txt)
                
                for linha in dados:
                    try:
                        caminho_origem = os.path.join(linha["CAMINHO"], linha["ARQUIVO"])
                        if not os.path.exists(caminho_origem):
                            continue

                        quantidade = linha["QTD"]
                        identific = linha["IDENTIFIC"]
                        nomefic = linha["NOMEFIC"]
                        filacolor = linha["FILACOLOR"]

                        nome_arquivo, extensao = os.path.splitext(linha["ARQUIVO"])
                        temp_file = os.path.join(pasta_temp, f"{nomefic}_temp{extensao}")
                        
                        shutil.copy(caminho_origem, temp_file)

                        if temp_file.lower().endswith(".pdf"):
                            pdf_marcado = adicionar_marca_dagua(
                                temp_file,
                                linha["LAMINAÇÃO"], 
                                linha["CODIGOACAB"], 
                                identific,
                                linha["FILACOLOR"]
                            )

                            if pdf_marcado:
                                # CORREÇÃO: Usar a tupla (identific, filacolor) como chave
                                chave = (identific, filacolor)
                                if chave not in pdfs_para_mesclar:
                                    pdfs_para_mesclar[chave] = []
                                pdfs_para_mesclar[chave].extend([pdf_marcado] * quantidade)

                        arquivos_processados += 1
                        progress_text = f"Processando {linha['ARQUIVO']}... ({arquivos_processados}/{total_arquivos})"
                        progress.update_progress((arquivos_processados / total_arquivos) * 80, progress_text)

                    except Exception as e:
                        logging.error(f"Erro ao processar {linha['ARQUIVO']}: {e}")
                        continue

            total_identifics = len(pdfs_para_mesclar)
            identifics_processados = 0

            # CORREÇÃO: Iterar sobre as chaves (identific, filacolor)
            for chave, lista_pdfs in pdfs_para_mesclar.items():
                identific, filacolor = chave  # Desempacotar a chave
                progress_text = f"Agrupando PDFs ({identifics_processados + 1}/{total_identifics})..."
                progress.update_progress(80 + ((identifics_processados / total_identifics) * 20), progress_text)

                pasta_destino = os.path.dirname(self.arquivos_selecionados[0])
                # Incluir filacolor no nome do arquivo de saída se desejar
                nome_saida = os.path.join(pasta_destino, f"{identific}.pdf")
                self.mesclar_pdfs(lista_pdfs, nome_saida, filacolor)
                identifics_processados += 1

            progress.update_progress(100, "Processamento concluído!")
            time.sleep(1)
            progress.destroy()
            
            messagebox.showinfo("Sucesso", "Processamento dos arquivos selecionados concluído!")
            
        except Exception as e:
            logging.error(f"Erro durante o processamento: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {e}")
        finally:
            try:
                shutil.rmtree(pasta_temp, ignore_errors=True)
                logging.info(f"Pasta temporária removida: {pasta_temp}")
            except Exception as e:
                logging.error(f"Erro ao limpar pasta temporária: {e}")

    def mesclar_pdfs(self, lista_pdfs, nome_saida,filacolor):
        try:
            merger = PdfMerger()
            for pdf in lista_pdfs:
                merger.append(pdf)
            
            merger.write(nome_saida)
            merger.close()
            logging.info(f"Arquivo mesclado salvo: {nome_saida}")
            
            if self.imprimir_automaticamente.get():
                self.enviar_para_impressora(nome_saida, os.path.basename(nome_saida),filacolor)
            
        except Exception as e:
            logging.error(f"Erro ao mesclar PDFs: {e}")
            messagebox.showerror("Erro", f"Falha ao mesclar arquivos: {e}")

    def enviar_para_impressora(self, arquivo_pdf, nome_arquivo,filacolor):
        try:
            comando = f'lpr -S 192.168.200.84 -P "{filacolor}" -J "{nome_arquivo.replace(filacolor, "")}" "{nome_arquivo}" "{arquivo_pdf}"'
            print(f"Enviando para impressora: lpr -S 192.168.200.84 -P {filacolor} -J {nome_arquivo} {arquivo_pdf}")
            subprocess.run(comando, shell=True, check=True)
            logging.info(f"Arquivo {nome_arquivo} enviado para impressora")
        except Exception as e:
            logging.error(f"Erro ao enviar para impressora: {e}")
            messagebox.showerror("Erro", f"Falha ao enviar para impressora: {e}")

    def mostrar_sobre(self):
        about_window = ttk.Toplevel(self.root)
        about_window.title("Sobre")
        about_window.geometry("400x250")
        about_window.resizable(False, False)
        
        main_frame = ttk.Frame(about_window, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        ttk.Label(
            main_frame, 
            text="🖨️", 
            font=('Helvetica', 24),
            bootstyle=PRIMARY
        ).pack(pady=(0, 10))
        
        ttk.Label(
            main_frame, 
            text="Processador de Capas UICLAP", 
            font=('Helvetica', 12, 'bold'),
            bootstyle=PRIMARY
        ).pack(pady=(0, 10))
        
        ttk.Label(
            main_frame, 
            text="Versão 2.0", 
            font=('Helvetica', 9)
        ).pack()
        
        ttk.Label(
            main_frame, 
            text="Desenvolvido para POD", 
            font=('Helvetica', 9)
        ).pack(pady=(0, 20))
        
        ttk.Label(
            main_frame, 
            text="© 2023 Todos os direitos reservados", 
            font=('Helvetica', 8),
            bootstyle=SECONDARY
        ).pack(side=BOTTOM)
        
        ttk.Button(
            main_frame, 
            text="OK", 
            command=about_window.destroy,
            bootstyle=PRIMARY
        ).pack(pady=(10, 0))

# --- INICIALIZAÇÃO ---
CONFIG = carregar_configuracoes()

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