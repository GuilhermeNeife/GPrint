import os 
import shutil
import time
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import fitz  # PyMuPDF
from PyPDF2 import PdfMerger
import subprocess
import logging
from datetime import datetime
from tkinter import simpledialog
import threading

# Limite de data (formato: DD/MM/AAAA)
DATA_LIMITE = "30/07/2027"  # Altere para a data desejada

# Configuração de logging
logging.basicConfig(
    filename='processamento_capas.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Caminho base onde estão os lotes
PASTA_BASE = r"N:\____REPOSITORIO GERAL POD\UICLAP\LOTES"
PASTA_TEMP = os.path.join(os.path.expanduser("~"), "Desktop", "temp_impressao")
LIMITE_ARQUIVOS_TEMP = 100  # Limite máximo de arquivos na pasta TEMP

# Configurações padrão
CONFIG = {
    'cor_fosco': (0, 0, 1),      # Azul
    'cor_brilho': (1, 0.7, 0.1), # Amarelo
    'fonte': 'helv',
    'tamanho_fonte': 11,
    'deslocamento_padrao': -30,
    'deslocamento_capa_28x56': 82,
    'deslocamento_x': 70 * 2.83465
}

def verificar_data_limite():
    try:
        data_limite = datetime.strptime(DATA_LIMITE, "%d/%m/%Y")
        data_atual = datetime.now()
        
        if data_atual > data_limite:
            messagebox.showerror(
                "Licença Expirada",
                f"O limite de uso deste programa foi excedido.\n\n"
                "Por favor, entre em contato com o responsável para renovação."
            )
            return False
        return True
    except Exception as e:
        logging.error(f"Erro ao verificar data limite: {e}")
        return True

class ProgressDialog(tk.Toplevel):
    """Janela de progresso personalizada"""
    def __init__(self, parent, title="Processando"):
        super().__init__(parent)
        self.title(title)
        self.configure(bg='white')

        self.geometry("500x120")
        self.resizable(False, False)

        # Centralizar
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 500) // 2
        y = (screen_height - 120) // 2
        self.geometry(f"+{x}+{y}")

        # Barra de progresso
        self.progress = ttk.Progressbar(self, orient="horizontal", length=460, mode="determinate")
        self.progress.pack(pady=(20, 10), padx=20)

        # Label com texto de status
        self.label = tk.Label(self, text="0% concluído", anchor="w", width=60, bg='white', fg='#2c3e50')
        self.label.pack(padx=20)

        # Impede interação fora da janela
        self.grab_set()
        self.focus_force()

    def update_progress(self, value, text=None):
        self.progress['value'] = value
        if text:
            self.label.config(text=text)
        else:
            self.label.config(text=f"{int(value)}% concluído")
        self.update_idletasks()

def limpar_pasta_segura(caminho_pasta, tentativas=3, delay=1):
    """Tenta limpar a pasta de forma segura com múltiplas tentativas"""
    for tentativa in range(tentativas):
        try:
            # Primeiro tenta fechar qualquer handle aberto
            for root, dirs, files in os.walk(caminho_pasta):
                for file in files:
                    caminho_arquivo = os.path.join(root, file)
                    try:
                        os.chmod(caminho_arquivo, 0o777)  # Tenta liberar permissões
                    except:
                        pass
            
            # Depois tenta remover
            shutil.rmtree(caminho_pasta)
            os.makedirs(caminho_pasta, exist_ok=True)
            return True
            
        except Exception as e:
            logging.warning(f"Tentativa {tentativa+1} falhou ao limpar {caminho_pasta}: {e}")
            if tentativa < tentativas - 1:
                time.sleep(delay)  # Espera antes de tentar novamente
    
    logging.error(f"Falha ao limpar pasta após {tentativas} tentativas: {caminho_pasta}")
    return False

def gerenciar_pasta_temp():
    """Garante que a pasta TEMP existe e mantém no máximo LIMITE_ARQUIVOS_TEMP arquivos"""
    try:
        # Cria a pasta se não existir
        os.makedirs(PASTA_TEMP, exist_ok=True)
        
        # Lista todos os arquivos na pasta TEMP com informações de data de modificação
        arquivos = []
        for nome in os.listdir(PASTA_TEMP):
            caminho = os.path.join(PASTA_TEMP, nome)
            if os.path.isfile(caminho):
                mod_time = os.path.getmtime(caminho)
                arquivos.append((mod_time, caminho, nome))
        
        # Ordena do mais antigo para o mais novo
        arquivos.sort()
        
        # Remove os arquivos mais antigos se exceder o limite
        while len(arquivos) >= LIMITE_ARQUIVOS_TEMP:
            _, caminho_mais_antigo, nome_mais_antigo = arquivos.pop(0)
            try:
                os.remove(caminho_mais_antigo)
                logging.info(f"Arquivo antigo removido: {nome_mais_antigo}")
            except Exception as e:
                logging.error(f"Erro ao remover arquivo antigo {nome_mais_antigo}: {e}")
                
    except Exception as e:
        logging.error(f"Erro ao gerenciar pasta temp: {e}")

def copiar_para_temp(arquivo_origem):
    """Copia o arquivo para a pasta TEMP, gerenciando o limite de arquivos"""
    try:
        gerenciar_pasta_temp()  # Garante que estamos dentro do limite
        
        nome_arquivo = os.path.basename(arquivo_origem)
        destino = os.path.join(PASTA_TEMP, nome_arquivo)
        
        # Se o arquivo já existe na pasta TEMP, remove antes de copiar
        if os.path.exists(destino):
            os.remove(destino)
            
        shutil.copy2(arquivo_origem, destino)
        logging.info(f"Arquivo copiado para pasta temp: {nome_arquivo}")
        return destino
    except Exception as e:
        logging.error(f"Erro ao copiar para pasta temp: {e}")
        return None

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
        arquivos = ["Processar Todas"] + sorted(
            [f for f in os.listdir(pasta) 
             if f.endswith(".txt") and "CAPA" in f.upper()]
        )
        return arquivos
    except Exception as e:
        logging.error(f"Erro ao listar arquivos em {pasta}: {e}")
        messagebox.showerror("Erro", f"Não foi possível listar os arquivos: {e}")
        return ["Processar Todas"]

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

        # Define a cor com base na laminação
        cor_marca = CONFIG['cor_fosco'] if laminacao.upper() == "FOSCO" else CONFIG['cor_brilho']

        # Define o deslocamento com base no formato da capa
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
    """Verifica se há espaço suficiente em disco"""
    try:
        stat = shutil.disk_usage(pasta)
        espaco_livre_mb = stat.free / (1024 * 1024)
        return espaco_livre_mb >= tamanho_necessario_mb
    except Exception as e:
        logging.error(f"Erro ao verificar espaço em disco: {e}")
        return True

def iniciar_processamento():
    thread = threading.Thread(target=processar_pilha)
    thread.start()

def processar_pilha():
    pasta = os.path.join(PASTA_BASE, combobox_pasta.get())
    nometxt = combobox_pasta.get()
    pilha = combobox_pilha.get()

    if not verificar_espaco_disco(pasta):
        messagebox.showwarning("Aviso", "Espaço em disco insuficiente para processamento!")
        return

    # Verifica se já existe pasta de destino para esta pilha
    if pilha != "Processar Todas":
        pasta_destino = os.path.join(pasta, os.path.splitext(pilha)[0])
        if os.path.exists(pasta_destino):
            resposta = messagebox.askyesno(
                "Pilha já processada",
                f"A pilha '{pilha}' já foi processada anteriormente.\n\n"
                "Deseja reprocessar? (Os arquivos existentes serão sobrescritos)",
                icon='warning'
            )
            if not resposta:
                return
            
            if not limpar_pasta_segura(pasta_destino):
                messagebox.showerror(
                    "Erro",
                    f"Não foi possível limpar a pasta de destino.\n"
                    f"Por favor, feche todos os arquivos em:\n{pasta_destino}\n"
                    "e tente novamente."
                )
                return

    # Cria diálogo de progresso
    progress = ProgressDialog(root)
    progress.update_progress(0, "Preparando processamento...")

    arquivos_txt = listar_pilhas(pasta)
    arquivos_para_processar = arquivos_txt if pilha == "Processar Todas" else [pilha]

    pdfs_para_mesclar = {}
    total_arquivos = sum(len(ler_arquivo_txt(os.path.join(pasta, txt))) for txt in arquivos_para_processar)
    arquivos_processados = 0

    for arquivo_txt in arquivos_para_processar:
        caminho_txt = os.path.join(pasta, arquivo_txt)
        dados = ler_arquivo_txt(caminho_txt)

        pasta_destino = os.path.join(pasta, os.path.splitext(arquivo_txt)[0])
        os.makedirs(pasta_destino, exist_ok=True)

        for linha in dados:
            try:
                caminho_origem = os.path.join(linha["CAMINHO"], linha["ARQUIVO"])
                if not os.path.exists(caminho_origem):
                    logging.warning(f"Arquivo não encontrado: {caminho_origem}")
                    continue

                quantidade = linha["QTD"]
                identific = linha["IDENTIFIC"]
                nomefic = linha["NOMEFIC"]

                # Processa o arquivo original
                nome_arquivo, extensao = os.path.splitext(linha["ARQUIVO"])
                primeiro_destino = os.path.join(pasta_destino, f"{nomefic}_1{extensao}")

                shutil.copy(caminho_origem, primeiro_destino)

                if primeiro_destino.lower().endswith(".pdf"):
                    pdf_marcado = adicionar_marca_dagua(
                        primeiro_destino, 
                        linha["LAMINAÇÃO"], 
                        linha["FORMATO CAPA"], 
                        identific
                    )

                    if pdf_marcado:
                        # Criar cópias do arquivo marcado
                        for i in range(1, quantidade):
                            copia_destino = os.path.join(pasta_destino, f"{nomefic}_{i+1}{extensao}")
                            shutil.copy(pdf_marcado, copia_destino)

                        # Adicionar à lista para mesclagem
                        if identific not in pdfs_para_mesclar:
                            pdfs_para_mesclar[identific] = []
                        pdfs_para_mesclar[identific].extend([pdf_marcado] * quantidade)

                # Atualiza a contagem antes de atualizar o progresso
                arquivos_processados += 1
                progress_text = f"Processando {linha['ARQUIVO']}... ({arquivos_processados}/{total_arquivos})"
                progress.update_progress((arquivos_processados / total_arquivos) * 80, progress_text)

            except Exception as e:
                logging.error(f"Erro ao processar {linha['ARQUIVO']}: {e}")
                continue

    # Mesclar PDFs por IDENTIFIC
    total_identifics = len(pdfs_para_mesclar)
    identifics_processados = 0

    for identific, lista_pdfs in pdfs_para_mesclar.items():
        progress_text = f"Agrupando PDFs ({identifics_processados + 1}/{total_identifics})..."
        progress.update_progress(80 + ((identifics_processados / total_identifics) * 20), progress_text)

        nome_saida = os.path.join(pasta, f"{os.path.splitext(arquivo_txt.replace(' ', '_'))[0]}_{nometxt.replace(' ', '')}.pdf")
        mesclar_pdfs(lista_pdfs, nome_saida)

        identifics_processados += 1

    progress.update_progress(100, "Processamento concluído!")
    progress.destroy()

def mesclar_pdfs(lista_pdfs, nome_saida):
    try:
        merger = PdfMerger()
        for pdf in lista_pdfs:
            merger.append(pdf)
        
        merger.write(nome_saida)
        merger.close()
        logging.info(f"Arquivo mesclado salvo: {nome_saida}")
        time.sleep(5)
        print("Enviando para impressora...")
        print(f"Nome do arquivo: {os.path.basename(nome_saida)}")
        print(f"Localização do arquivo: {nome_saida}")

        if messagebox.askyesno("Enviar para Impressora", f"Deseja enviar '{os.path.basename(nome_saida)}' para a máquina de impressão?"):
            enviar_para_impressora(nome_saida, os.path.basename(nome_saida))
            
    except Exception as e:
        logging.error(f"Erro ao mesclar PDFs: {e}")
        messagebox.showerror("Erro", f"Falha ao mesclar arquivos: {e}")

def enviar_para_impressora(arquivo_pdf, nome_arquivo):
    try:
        # Copia para a pasta temp primeiro
        caminho_temp = copiar_para_temp(arquivo_pdf)
        if not caminho_temp:
            raise Exception("Falha ao copiar para pasta temporária")
        
        comando = f'lpr -S 192.168.200.84 -P "Capas_32x60" -J "{nome_arquivo}" "{caminho_temp}"'
        print(f"Enviando para impressora: {nome_arquivo}")
        print(f"Caminho temporário: {caminho_temp}")

        subprocess.run(comando, shell=True, check=True)
        logging.info(f"Arquivo {nome_arquivo} enviado para impressora")
        messagebox.showinfo("Sucesso", f"Arquivo {nome_arquivo} enviado para impressão!")
        
    except subprocess.CalledProcessError as e:
        logging.error(f"Erro ao enviar para impressora: {e}")
        messagebox.showerror("Erro", f"Falha ao enviar para impressora: {e}")
    except Exception as e:
        logging.error(f"Erro inesperado ao enviar para impressora: {e}")
        messagebox.showerror("Erro", f"Erro inesperado: {e}")

def criar_interface():
    global root, combobox_pasta, combobox_pilha, lbl_status
    
    root = tk.Tk()
    root.title("Processador de Capas UICLAP")
    
    # Configuração da janela principal
    window_width = 650
    window_height = 450
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.resizable(False, False)
    
    # Cores
    cor_fundo = "white"
    cor_primaria = "#2c3e50"  # Azul escuro
    cor_secundaria = "#3498db"  # Azul claro
    cor_destaque = "#e74c3c"  # Vermelho
    cor_texto = "#333333"
    
    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configurar estilos
    style.configure('TFrame', background=cor_fundo)
    style.configure('TLabel', background=cor_fundo, foreground=cor_texto, font=('Arial', 10))
    style.configure('TButton', font=('Arial', 10, 'bold'), background=cor_primaria, 
                   foreground='white', borderwidth=1)
    style.map('TButton', 
              background=[('active', cor_secundaria)],
              foreground=[('active', 'white')])
    style.configure('TCombobox', fieldbackground='white')
    style.configure('TLabelframe', background=cor_fundo, foreground=cor_primaria)
    style.configure('TLabelframe.Label', background=cor_fundo, foreground=cor_primaria)
    
    # Frame principal
    main_frame = ttk.Frame(root, padding=(20, 15))
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Cabeçalho
    header_frame = ttk.Frame(main_frame)
    header_frame.pack(fill=tk.X, pady=(0, 15))
    
    lbl_titulo = tk.Label(header_frame, text="PROCESSADOR DE CAPAS UICLAP", 
                         font=('Arial', 14, 'bold'), bg=cor_fundo, fg=cor_primaria)
    lbl_titulo.pack()
    
    # Frame de conteúdo
    content_frame = ttk.Frame(main_frame)
    content_frame.pack(fill=tk.BOTH, expand=True)
    
    # Frame de seleção
    selection_frame = ttk.LabelFrame(content_frame, text=" Seleção de Arquivos ", 
                                   padding=(15, 10, 15, 15))
    selection_frame.pack(fill=tk.BOTH, padx=5, pady=5)
    
    # Widgets de seleção
    lbl_pasta = ttk.Label(selection_frame, text="Pasta do Lote:")
    lbl_pasta.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
    
    combobox_pasta = ttk.Combobox(selection_frame, values=listar_subpastas(PASTA_BASE), 
                                 state="readonly", width=50)
    combobox_pasta.grid(row=0, column=1, padx=5, pady=(0, 5), sticky=tk.EW)
    
    lbl_pilha = ttk.Label(selection_frame, text="Arquivo de Pilha:")
    lbl_pilha.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
    
    combobox_pilha = ttk.Combobox(selection_frame, state="readonly", width=50)
    combobox_pilha.grid(row=1, column=1, padx=5, pady=(0, 5), sticky=tk.EW)
    
    # Botão de processamento
    btn_frame = ttk.Frame(content_frame)
    btn_frame.pack(fill=tk.X, pady=(15, 5))
    
    btn_processar = tk.Button(btn_frame, text="INICIAR PROCESSAMENTO", 
                            command=iniciar_processamento,
                            bg=cor_primaria, fg='white',
                            font=('Arial', 10, 'bold'),
                            activebackground=cor_secundaria,
                            activeforeground='white',
                            relief=tk.RAISED,
                            bd=2)
    btn_processar.pack(pady=10, ipadx=20, ipady=5)
    
    # Barra de status
    status_frame = tk.Frame(root, bg=cor_primaria)
    status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
    
    lbl_status = tk.Label(status_frame, text="Pronto para processar", 
                         anchor=tk.W, bg=cor_primaria, fg='white', font=('Arial', 9))
    lbl_status.pack(fill=tk.X, padx=5)
    
    # Menu
    menubar = tk.Menu(root, tearoff=0, bg=cor_fundo, fg=cor_texto)
    
    # Menu Arquivo
    file_menu = tk.Menu(menubar, tearoff=0, bg=cor_fundo, fg=cor_texto)
    file_menu.add_command(label="Sair", command=root.quit)
    menubar.add_cascade(label="Arquivo", menu=file_menu)
    
    # Menu Ajuda
    help_menu = tk.Menu(menubar, tearoff=0, bg=cor_fundo, fg=cor_texto)
    help_menu.add_command(label="Configurações", command=mostrar_configuracoes)
    help_menu.add_command(label="Sobre", command=mostrar_sobre)
    menubar.add_cascade(label="Ajuda", menu=help_menu)
    
    root.config(menu=menubar)
    
    # Eventos
    def atualizar_pilhas(event):
        pasta_selecionada = os.path.join(PASTA_BASE, combobox_pasta.get())
        combobox_pilha["values"] = listar_pilhas(pasta_selecionada)
        combobox_pilha.current(0)
        lbl_status.config(text=f"Pasta selecionada: {combobox_pasta.get()}")
    
    combobox_pasta.bind("<<ComboboxSelected>>", atualizar_pilhas)
    
    return root

def mostrar_configuracoes():
    """Mostra diálogo de configurações moderno"""
    config_window = tk.Toplevel()
    config_window.title("Configurações")
    config_window.transient(root)
    config_window.grab_set()
    
    # Centralizar
    window_width = 400
    window_height = 300
    screen_width = config_window.winfo_screenwidth()
    screen_height = config_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    config_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    config_window.resizable(False, False)
    
    # Cores
    cor_fundo = "white"
    cor_primaria = "#2c3e50"
    
    # Frame principal
    main_frame = tk.Frame(config_window, bg=cor_fundo)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
    
    lbl_titulo = tk.Label(main_frame, text="Configurações do Sistema", 
                         font=('Arial', 12, 'bold'), bg=cor_fundo, fg=cor_primaria)
    lbl_titulo.pack(pady=(0, 15))
    
    # Área de configurações (adicionar widgets conforme necessário)
    config_frame = tk.Frame(main_frame, bg=cor_fundo)
    config_frame.pack(fill=tk.BOTH, expand=True)
    
    lbl_info = tk.Label(config_frame, text="Em desenvolvimento...", 
                       bg=cor_fundo, font=('Arial', 10))
    lbl_info.pack(pady=20)
    
    # Botões
    btn_frame = tk.Frame(main_frame, bg=cor_fundo)
    btn_frame.pack(fill=tk.X, pady=(10, 0))
    
    btn_fechar = tk.Button(btn_frame, text="Fechar", command=config_window.destroy,
                          bg=cor_primaria, fg='white', font=('Arial', 10))
    btn_fechar.pack(side=tk.RIGHT, padx=5)

def mostrar_sobre():
    """Mostra diálogo 'Sobre' moderno"""
    about_window = tk.Toplevel()
    about_window.title("Sobre")
    about_window.transient(root)
    about_window.grab_set()
    
    # Centralizar
    window_width = 400
    window_height = 250
    screen_width = about_window.winfo_screenwidth()
    screen_height = about_window.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    about_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    about_window.resizable(False, False)
    
    # Cores
    cor_fundo = "white"
    cor_primaria = "#2c3e50"
    cor_destaque = "#e74c3c"
    
    # Frame principal
    main_frame = tk.Frame(about_window, bg=cor_fundo)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
    
    # Logo/Ícone (pode adicionar uma imagem aqui)
    lbl_icone = tk.Label(main_frame, text="🖨️", font=('Arial', 24), bg=cor_fundo)
    lbl_icone.pack(pady=(0, 10))
    
    # Informações
    lbl_titulo = tk.Label(main_frame, text="Processador de Capas UICLAP", 
                         font=('Arial', 12, 'bold'), bg=cor_fundo, fg=cor_primaria)
    lbl_titulo.pack(pady=(0, 10))
    
    lbl_versao = tk.Label(main_frame, text="Versão 1.1", font=('Arial', 9), bg=cor_fundo)
    lbl_versao.pack()
    
    lbl_dev = tk.Label(main_frame, text="Desenvolvido para POD", font=('Arial', 9), bg=cor_fundo)
    lbl_dev.pack(pady=(0, 20))
    
    lbl_copy = tk.Label(main_frame, text="© 2023 Todos os direitos reservados", 
                       font=('Arial', 8), bg=cor_fundo, fg=cor_destaque)
    lbl_copy.pack(side=tk.BOTTOM)
    
    # Botão OK
    btn_ok = tk.Button(main_frame, text="OK", command=about_window.destroy,
                      bg=cor_primaria, fg='white', font=('Arial', 10))
    btn_ok.pack(pady=(10, 0))

if __name__ == "__main__":
    if verificar_data_limite():
        root = criar_interface()
        root.mainloop()