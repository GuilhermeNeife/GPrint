import os 
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
from PyPDF2 import PdfMerger
import subprocess
import logging
from datetime import datetime
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style

# --- CONSTANTES ---
DATA_LIMITE = "30/07/2027"
PASTA_BASE = r"N:\____REPOSITORIO GERAL POD\UICLAP\LOTES"
PASTA_TRABALHO = os.path.join(os.environ.get('TEMP', os.path.expanduser("~")), "UICLAP_TEMP")  # Pasta temporária do sistema
LIMITE_ARQUIVOS_TEMP = 100

# Lista de TODOS os temas do ttkbootstrap
TEMAS_DISPONIVEIS = [
    "cosmo", "flatly", "journal", "litera", "lumen", "minty", "pulse", 
    "sandstone", "united", "yeti", "morph", "darkly", "superhero",
    "solar", "cyborg", "vapor"
]

# --- CONFIGURAÇÕES ---
CONFIG = {
    'cor_fosco': (0, 0, 1),
    'cor_brilho': (1, 0.7, 0.1),
    'fonte': 'helv',
    'tamanho_fonte': 11,
    'deslocamento_padrao': -30,
    'deslocamento_capa_28x56': 82,
    'deslocamento_x': 70 * 2.83465
}

# --- FUNÇÕES AUXILIARES ---
def verificar_data_limite():
    try:
        data_limite = datetime.strptime(DATA_LIMITE, "%d/%m/%Y")
        data_atual = datetime.now()
        if data_atual > data_limite:
            messagebox.showerror("Licença Expirada", "O limite de uso deste programa foi excedido.\n\nPor favor, entre em contato com o responsável para renovação.")
            return False
        return True
    except Exception as e:
        logging.error(f"Erro ao verificar data limite: {e}")
        return True

def listar_subpastas(pasta_base):
    try:
        return sorted([nome for nome in os.listdir(pasta_base) 
                      if os.path.isdir(os.path.join(pasta_base, nome))], reverse=True)
    except Exception as e:
        logging.error(f"Erro ao listar subpastas: {e}")
        messagebox.showerror("Erro", f"Não foi possível listar as subpastas: {e}")
        return []

def listar_pilhas(pasta):
    try:
        return sorted([f for f in os.listdir(pasta) 
                      if f.endswith(".txt") and "CAPA" in f.upper()])
    except Exception as e:
        logging.error(f"Erro ao listar arquivos em {pasta}: {e}")
        messagebox.showerror("Erro", f"Não foi possível listar os arquivos: {e}")
        return []

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
                            "IDENTIFIC": campos[13].strip(),
                            "NOMEFIC": campos[11].strip()
                        })
                    except (ValueError, IndexError) as e:
                        logging.warning(f"Linha {idx} mal formatada no arquivo {caminho_txt}: {linha}")
        return dados
    except Exception as e:
        logging.error(f"Erro ao ler arquivo {caminho_txt}: {e}")
        messagebox.showerror("Erro", f"Não foi possível ler o arquivo {os.path.basename(caminho_txt)}")
        return []

def adicionar_marca_dagua(pdf_path, laminacao, formato_capa, identific):
    try:
        doc = fitz.open(pdf_path)
        nome_arquivo = os.path.splitext(os.path.basename(pdf_path))[0]

        cor_marca = CONFIG['cor_fosco'] if laminacao.upper() == "FOSCO" else CONFIG['cor_brilho']

        deslocamento_y = CONFIG['deslocamento_padrao']
        if formato_capa == "Capas_28X56":
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

        novo_caminho = pdf_path.replace(".pdf", "_marcado.pdf")
        doc.save(novo_caminho)
        doc.close()
        return novo_caminho

    except Exception as e:
        logging.error(f"Erro ao adicionar marca d'água em {pdf_path}: {e}")
        messagebox.showerror("Erro", f"Falha ao marcar o arquivo {os.path.basename(pdf_path)}")
        return None

def verificar_espaco_disco(pasta, tamanho_necessario_mb=100):
    try:
        stat = shutil.disk_usage(pasta)
        espaco_livre_mb = stat.free / (1024 * 1024)
        return espaco_livre_mb >= tamanho_necessario_mb
    except Exception as e:
        logging.error(f"Erro ao verificar espaço em disco: {e}")
        return True

def limpar_pasta_segura(caminho_pasta, tentativas=3, delay=1):
    for tentativa in range(tentativas):
        try:
            for root, dirs, files in os.walk(caminho_pasta):
                for file in files:
                    caminho_arquivo = os.path.join(root, file)
                    try:
                        os.chmod(caminho_arquivo, 0o777)
                    except:
                        pass
            
            shutil.rmtree(caminho_pasta)
            os.makedirs(caminho_pasta, exist_ok=True)
            return True
            
        except Exception as e:
            logging.warning(f"Tentativa {tentativa+1} falhou ao limpar {caminho_pasta}: {e}")
            if tentativa < tentativas - 1:
                time.sleep(delay)
    
    logging.error(f"Falha ao limpar pasta após {tentativas} tentativas: {caminho_pasta}")
    return False

class ProgressDialog(ttk.Toplevel):
    def __init__(self, parent, title="Processando"):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x150")
        self.resizable(False, False)
        
        self.center_window()

        # Barra de progresso
        self.progress = ttk.Progressbar(
            self, 
            orient=HORIZONTAL, 
            length=460, 
            mode=DETERMINATE,
            bootstyle=PRIMARY
        )
        self.progress.pack(pady=(20, 10), padx=20)

        # Label com texto de status
        self.label = ttk.Label(
            self, 
            text="0% concluído", 
            anchor=W, 
            width=60,
            bootstyle=SECONDARY
        )
        self.label.pack(padx=20)

        # Label do timer
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

class Application:
    def __init__(self, root):
        self.root = root
        self.pilhas_selecionadas = []
        self.imprimir_automaticamente = tk.BooleanVar(value=True)
        self.tema_atual = tk.StringVar(value="darkly")
        self.style = Style(theme=self.tema_atual.get())
        self.setup_ui()

    def setup_ui(self):
        self.root.title("Processador de Capas UICLAP")
        
        # Configuração da janela
        self.root.geometry("1000x700")
        self.center_window()
        
        # Menu superior
        self.criar_menu_superior()
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Conteúdo principal
        self.criar_interface_principal(main_frame)
        
        # Inicializar
        self.atualizar_pastas()

    def center_window(self):
        """Centraliza a janela principal na tela"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        
        # Calcula a posição x e y para centralizar
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def criar_menu_superior(self):
        menubar = tk.Menu(self.root)
        
        # Menu Tema
        menu_tema = tk.Menu(menubar, tearoff=0)
        for tema in TEMAS_DISPONIVEIS:
            menu_tema.add_radiobutton(
                label=tema.capitalize(),
                variable=self.tema_atual,
                value=tema,
                command=self.mudar_tema
            )
        menubar.add_cascade(label="Temas", menu=menu_tema)
        
        # Menu Ajuda
        menu_ajuda = tk.Menu(menubar, tearoff=0)
        menu_ajuda.add_command(label="Sobre", command=self.mostrar_sobre)
        menubar.add_cascade(label="Ajuda", menu=menu_ajuda)
        
        self.root.config(menu=menubar)

    def mudar_tema(self):
        self.style.theme_use(self.tema_atual.get())
        self.aplicar_cores_tema()
        messagebox.showinfo("Tema Alterado", f"Tema mudado para: {self.tema_atual.get().capitalize()}")

    def aplicar_cores_tema(self):
        # Atualiza cores da listbox baseado no tema
        bg = self.style.colors.get('bg') if self.style.theme.type == 'dark' else self.style.colors.get('light')
        fg = self.style.colors.get('fg') if self.style.theme.type == 'dark' else self.style.colors.get('dark')
        selectbg = self.style.colors.get('primary')
        selectfg = self.style.colors.get('bg') if self.style.theme.type == 'dark' else self.style.colors.get('light')
        
        self.listbox_pilhas.config(
            bg=bg,
            fg=fg,
            selectbackground=selectbg,
            selectforeground=selectfg
        )

    def criar_interface_principal(self, parent):
        # Frame de cabeçalho
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=X, pady=(0, 15))
        
        ttk.Label(
            header_frame, 
            text="PROCESSADOR DE CAPAS UICLAP", 
            font=('Helvetica', 14, 'bold'),
            bootstyle=(PRIMARY, 'inverse')
        ).pack()
        
        # Frame de seleção
        selection_frame = ttk.Labelframe(
            parent, 
            text="Seleção de Arquivos", 
            bootstyle=PRIMARY,
            padding=15
        )
        selection_frame.pack(fill=BOTH, expand=True)
        
        # Combobox de pastas
        ttk.Label(
            selection_frame, 
            text="Pasta do Lote:", 
            font=('Helvetica', 10)
        ).grid(row=0, column=0, sticky=W, pady=(0, 10))
        
        self.combobox_pasta = ttk.Combobox(
            selection_frame, 
            state="readonly", 
            font=('Helvetica', 10)
        )
        self.combobox_pasta.grid(row=0, column=1, padx=5, pady=(0, 10), sticky=EW)
        
        # Listbox de pilhas
        ttk.Label(
            selection_frame, 
            text="Selecione as Pilhas (Clique + Shift/Ctrl para múltiplas):", 
            font=('Helvetica', 10)
        ).grid(row=1, column=0, sticky=W, pady=(0, 5))
        
        self.listbox_pilhas = tk.Listbox(
            selection_frame,
            selectmode=EXTENDED,
            height=10,
            font=('Helvetica', 10),
            exportselection=False
        )
        self.listbox_pilhas.grid(row=2, column=0, columnspan=2, padx=5, pady=(0, 10), sticky=NSEW)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(
            selection_frame,
            orient=VERTICAL,
            command=self.listbox_pilhas.yview,
            bootstyle=PRIMARY
        )
        scrollbar.grid(row=2, column=2, sticky=NS)
        self.listbox_pilhas.config(yscrollcommand=scrollbar.set)
        
        # Checkbox de impressão
        self.checkbox_imprimir = ttk.Checkbutton(
            selection_frame, 
            text="Imprimir automaticamente após processar",
            variable=self.imprimir_automaticamente,
            bootstyle="primary-round-toggle"
        )
        self.checkbox_imprimir.grid(row=3, column=0, columnspan=3, pady=(10, 0), sticky=W)
        
        # Botão de processamento
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=X, pady=(15, 5))
        
        self.btn_processar = ttk.Button(
            btn_frame, 
            text="PROCESSAR PILHAS SELECIONADAS", 
            command=self.iniciar_processamento,
            bootstyle=(PRIMARY, OUTLINE),
            width=40
        )
        self.btn_processar.pack(pady=10, ipady=5)
        
        # Barra de status
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
        
        # Configurar eventos
        self.combobox_pasta.bind("<<ComboboxSelected>>", self.atualizar_lista_pilhas)
        
        # Aplicar tema inicial
        self.aplicar_cores_tema()

    def atualizar_pastas(self):
        pastas = listar_subpastas(PASTA_BASE)
        self.combobox_pasta['values'] = pastas
        if pastas:
            self.combobox_pasta.current(0)
            self.atualizar_lista_pilhas()

    def atualizar_lista_pilhas(self, event=None):
        pasta_selecionada = os.path.join(PASTA_BASE, self.combobox_pasta.get())
        arquivos = listar_pilhas(pasta_selecionada)
        
        self.listbox_pilhas.delete(0, END)
        for arquivo in arquivos:
            self.listbox_pilhas.insert(END, arquivo)
        
        self.lbl_status.config(text=f"Pasta selecionada: {self.combobox_pasta.get()}")

    def iniciar_processamento(self):
        if not verificar_data_limite():
            return
            
        self.pilhas_selecionadas = [self.listbox_pilhas.get(i) for i in self.listbox_pilhas.curselection()]
        
        if not self.pilhas_selecionadas:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma pilha para processar!")
            return
        
        # Desabilita o botão durante o processamento
        self.btn_processar.config(state=DISABLED)
        
        # Inicia o processamento em thread separada
        thread = threading.Thread(target=self.processar_pilhas)
        thread.start()
        
        # Verifica quando a thread termina para reabilitar o botão
        self.verificar_thread(thread)

    def verificar_thread(self, thread):
        if thread.is_alive():
            self.root.after(100, lambda: self.verificar_thread(thread))
        else:
            self.btn_processar.config(state=NORMAL)

    def processar_pilhas(self):
        # Cria pasta temporária no diretório TEMP do sistema
        pasta_temp = os.path.join(PASTA_TRABALHO, f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(pasta_temp, exist_ok=True)
        
        try:
            pasta = os.path.join(PASTA_BASE, self.combobox_pasta.get())
            nometxt = self.combobox_pasta.get()
            
            progress = ProgressDialog(self.root)
            progress.update_progress(0, "Preparando processamento...")
            
            # Adicionar timer
            start_time = time.time()
            
            def update_timer():
                elapsed = int(time.time() - start_time)
                progress.update_timer(elapsed)
                if progress.winfo_exists():
                    progress.after(1000, update_timer)
            
            update_timer()  # Inicia o timer
            
            pdfs_para_mesclar = {}
            total_arquivos = sum(len(ler_arquivo_txt(os.path.join(pasta, txt))) for txt in self.pilhas_selecionadas)
            arquivos_processados = 0

            for arquivo_txt in self.pilhas_selecionadas:
                caminho_txt = os.path.join(pasta, arquivo_txt)
                dados = ler_arquivo_txt(caminho_txt)
                
                for linha in dados:
                    try:
                        caminho_origem = os.path.join(linha["CAMINHO"], linha["ARQUIVO"])
                        if not os.path.exists(caminho_origem):
                            continue

                        quantidade = linha["QTD"]
                        identific = linha["IDENTIFIC"]
                        nomefic = linha["NOMEFIC"]

                        # Trabalha na pasta temporária
                        nome_arquivo, extensao = os.path.splitext(linha["ARQUIVO"])
                        temp_file = os.path.join(pasta_temp, f"{nomefic}_temp{extensao}")
                        
                        shutil.copy(caminho_origem, temp_file)

                        if temp_file.lower().endswith(".pdf"):
                            pdf_marcado = adicionar_marca_dagua(
                                temp_file, 
                                linha["LAMINAÇÃO"], 
                                linha["FORMATO CAPA"], 
                                identific
                            )

                            if pdf_marcado:
                                if identific not in pdfs_para_mesclar:
                                    pdfs_para_mesclar[identific] = []
                                pdfs_para_mesclar[identific].extend([pdf_marcado] * quantidade)

                        arquivos_processados += 1
                        progress_text = f"Processando {linha['ARQUIVO']}... ({arquivos_processados}/{total_arquivos})"
                        progress.update_progress((arquivos_processados / total_arquivos) * 80, progress_text)

                    except Exception as e:
                        logging.error(f"Erro ao processar {linha['ARQUIVO']}: {e}")
                        continue

            # Mescla os PDFs e salva na pasta final
            total_identifics = len(pdfs_para_mesclar)
            identifics_processados = 0

            for identific, lista_pdfs in pdfs_para_mesclar.items():
                progress_text = f"Agrupando PDFs ({identifics_processados + 1}/{total_identifics})..."
                progress.update_progress(80 + ((identifics_processados / total_identifics) * 20), progress_text)

                nome_saida = os.path.join(pasta, f"{identific}.pdf")
                self.mesclar_pdfs(lista_pdfs, nome_saida)
                identifics_processados += 1

            progress.update_progress(100, "Processamento concluído!")
            time.sleep(1)
            progress.destroy()
            
            messagebox.showinfo("Sucesso", "Processamento das pilhas selecionadas concluído!")
            
        except Exception as e:
            logging.error(f"Erro durante o processamento: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {e}")
        finally:
            # Limpa TODAS as pastas temporárias, mesmo se ocorrer erro
            try:
                shutil.rmtree(pasta_temp, ignore_errors=True)
                logging.info(f"Pasta temporária removida: {pasta_temp}")
            except Exception as e:
                logging.error(f"Erro ao limpar pasta temporária: {e}")

    def mesclar_pdfs(self, lista_pdfs, nome_saida):
        try:
            merger = PdfMerger()
            for pdf in lista_pdfs:
                merger.append(pdf)
            
            merger.write(nome_saida)
            merger.close()
            logging.info(f"Arquivo mesclado salvo: {nome_saida}")
            
            if self.imprimir_automaticamente.get():
                self.enviar_para_impressora(nome_saida, os.path.basename(nome_saida))
            
        except Exception as e:
            logging.error(f"Erro ao mesclar PDFs: {e}")
            messagebox.showerror("Erro", f"Falha ao mesclar arquivos: {e}")

    def enviar_para_impressora(self, arquivo_pdf, nome_arquivo):
        try:
            comando = f'lpr -S 192.168.200.84 -P "Capas_32x60" -J "{nome_arquivo}" "{arquivo_pdf}"'
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

if __name__ == "__main__":
    logging.basicConfig(
        filename='processamento_capas.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    if verificar_data_limite():
        root = ttk.Window(title="Processador de Capas", themename="darkly")
        app = Application(root)
        root.mainloop()