import os
import shutil
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, END
import fitz  # PyMuPDF
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
from PIL import Image
import io
import math
import queue
import sys

# --- CONSTANTES ---
DATA_LIMIT = "31/07/2027"
PASTA_BASE = r"N:\____REPOSITORIO GERAL POD\UICLAP\LOTES"
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

DEFAULT_CONFIG = {
    'cor_fosco': (0, 0, 1),
    'cor_brilho': (1, 0.7, 0.1),
    'fonte': 'helv',
    'tamanho_fonte': 11,
    'deslocamento_padrao': -30,
    'tema': 'darkly'
}

# --- FUNÇÕES AUXILIARES (DESENHO) ---

def inserir_marca_dagua(page, texto, tamanho_fonte=8, cor=(0,0,0), 
                        deslocamento_vertical=2.5*MM, deslocamento_horizontal=0):
    """
    Insere texto como marca d'água.
    - deslocamento_vertical: positivo = acima da trimbox, negativo = abaixo.
    - deslocamento_horizontal: positivo = direita, negativo = esquerda (relativo ao centro).
    """
    # Obtém a área de corte (trimbox) ou o retângulo da página
    if page.trimbox and page.trimbox.width > 0:
        bbox = page.trimbox
    else:
        bbox = page.rect

    # Centro horizontal da trimbox
    centro_x = (bbox.x0 + bbox.x1) / 2
    # Posição x com deslocamento
    x_texto = centro_x + deslocamento_horizontal

    # Posição vertical: bbox.y0 - deslocamento_vertical (positivo sobe)
    y_texto = bbox.y0 - deslocamento_vertical

    # Estimar largura aproximada do texto (fator 0.6 * tamanho_fonte por caractere é uma aproximação)
    largura_estimada = len(texto) * tamanho_fonte * 0.6
    # Limites horizontais: garantir que o texto não ultrapasse as bordas da página
    x_min = largura_estimada / 2
    x_max = page.rect.width - largura_estimada / 2
    x_texto = max(x_min, min(x_texto, x_max))

    # Limites verticais
    altura_pagina = page.rect.height
    y_min = tamanho_fonte * 0.8          # margem superior (evita texto cortado)
    y_max = altura_pagina - tamanho_fonte * 1.2   # margem inferior
    y_texto = max(y_min, min(y_texto, y_max))

    # Insere o texto (centralizado no ponto x_texto, y_texto)
    page.insert_text((x_texto, y_texto), texto,
                     fontsize=tamanho_fonte, fontname='helv',
                     color=cor, render_mode=0)

def get_border_pixmap(page, rect, zoom=2.0):
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, clip=rect)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img

def espelhar_borda(img, flip_horizontal=False, flip_vertical=False):
    if flip_horizontal: img = img.transpose(Image.FLIP_LEFT_RIGHT)
    if flip_vertical: img = img.transpose(Image.FLIP_TOP_BOTTOM)
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

def aplicar_sangria_espelhada_externa(page, nova, trim_original, destino, sangria_pt, bordas):
    zoom = 2.0
    
    if 'top' in bordas:
        top_rect = fitz.Rect(trim_original.x0, trim_original.y0,
                              trim_original.x1, trim_original.y0 + sangria_pt)
        img_top = get_border_pixmap(page, top_rect, zoom)
        img_top = espelhar_borda(img_top, flip_vertical=True)
        img_top = img_top.resize((int(top_rect.width), int(top_rect.height)), Image.Resampling.LANCZOS)
        top_dest = fitz.Rect(destino.x0, destino.y0 - sangria_pt, destino.x1, destino.y0)
        inserir_imagem_na_pagina(nova, img_top, top_dest)
    
    if 'bottom' in bordas:
        bottom_rect = fitz.Rect(trim_original.x0, trim_original.y1 - sangria_pt,
                                 trim_original.x1, trim_original.y1)
        img_bottom = get_border_pixmap(page, bottom_rect, zoom)
        img_bottom = espelhar_borda(img_bottom, flip_vertical=True)
        img_bottom = img_bottom.resize((int(bottom_rect.width), int(bottom_rect.height)), Image.Resampling.LANCZOS)
        bottom_dest = fitz.Rect(destino.x0, destino.y1, destino.x1, destino.y1 + sangria_pt)
        inserir_imagem_na_pagina(nova, img_bottom, bottom_dest)
    
    if 'left' in bordas:
        left_rect = fitz.Rect(trim_original.x0, trim_original.y0,
                               trim_original.x0 + sangria_pt, trim_original.y1)
        img_left = get_border_pixmap(page, left_rect, zoom)
        img_left = espelhar_borda(img_left, flip_horizontal=True)
        img_left = img_left.resize((int(left_rect.width), int(left_rect.height)), Image.Resampling.LANCZOS)
        left_dest = fitz.Rect(destino.x0 - sangria_pt, destino.y0, destino.x0, destino.y1)
        inserir_imagem_na_pagina(nova, img_left, left_dest)
    
    if 'right' in bordas:
        right_rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y0,
                                trim_original.x1, trim_original.y1)
        img_right = get_border_pixmap(page, right_rect, zoom)
        img_right = espelhar_borda(img_right, flip_horizontal=True)
        img_right = img_right.resize((int(right_rect.width), int(right_rect.height)), Image.Resampling.LANCZOS)
        right_dest = fitz.Rect(destino.x1, destino.y0, destino.x1 + sangria_pt, destino.y1)
        inserir_imagem_na_pagina(nova, img_right, right_dest)
    
    # Cantos
    if 'top' in bordas and 'left' in bordas:
        tl_rect = fitz.Rect(trim_original.x0, trim_original.y0,
                             trim_original.x0 + sangria_pt, trim_original.y0 + sangria_pt)
        img_tl = get_border_pixmap(page, tl_rect, zoom)
        img_tl = espelhar_borda(img_tl, flip_horizontal=True, flip_vertical=True)
        img_tl = img_tl.resize((int(sangria_pt), int(sangria_pt)), Image.Resampling.LANCZOS)
        tl_dest = fitz.Rect(destino.x0 - sangria_pt, destino.y0 - sangria_pt, destino.x0, destino.y0)
        inserir_imagem_na_pagina(nova, img_tl, tl_dest)
    
    if 'top' in bordas and 'right' in bordas:
        tr_rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y0,
                             trim_original.x1, trim_original.y0 + sangria_pt)
        img_tr = get_border_pixmap(page, tr_rect, zoom)
        img_tr = espelhar_borda(img_tr, flip_horizontal=True, flip_vertical=True)
        img_tr = img_tr.resize((int(sangria_pt), int(sangria_pt)), Image.Resampling.LANCZOS)
        tr_dest = fitz.Rect(destino.x1, destino.y0 - sangria_pt, destino.x1 + sangria_pt, destino.y0)
        inserir_imagem_na_pagina(nova, img_tr, tr_dest)
    
    if 'bottom' in bordas and 'left' in bordas:
        bl_rect = fitz.Rect(trim_original.x0, trim_original.y1 - sangria_pt,
                             trim_original.x0 + sangria_pt, trim_original.y1)
        img_bl = get_border_pixmap(page, bl_rect, zoom)
        img_bl = espelhar_borda(img_bl, flip_horizontal=True, flip_vertical=True)
        img_bl = img_bl.resize((int(sangria_pt), int(sangria_pt)), Image.Resampling.LANCZOS)
        bl_dest = fitz.Rect(destino.x0 - sangria_pt, destino.y1, destino.x0, destino.y1 + sangria_pt)
        inserir_imagem_na_pagina(nova, img_bl, bl_dest)
    
    if 'bottom' in bordas and 'right' in bordas:
        br_rect = fitz.Rect(trim_original.x1 - sangria_pt, trim_original.y1 - sangria_pt,
                             trim_original.x1, trim_original.y1)
        img_br = get_border_pixmap(page, br_rect, zoom)
        img_br = espelhar_borda(img_br, flip_horizontal=True, flip_vertical=True)
        img_br = img_br.resize((int(sangria_pt), int(sangria_pt)), Image.Resampling.LANCZOS)
        br_dest = fitz.Rect(destino.x1, destino.y1, destino.x1 + sangria_pt, destino.y1 + sangria_pt)
        inserir_imagem_na_pagina(nova, img_br, br_dest)

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

# --- FUNÇÕES DE MONTAGEM (com suporte a marca d'água) ---

def corta_empilha_rotacionado(input_path, output_path, output_width, output_height,
                              sangria_mm=3, add_marcas=True, marca_dagua_texto=None):
    """
    Montagem corta-empilha com rotação de 180° nas páginas da segunda metade.
    Espaçamento entre páginas fixo em 0.
    """
    doc = fitz.open(input_path)
    out = fitz.open()
    sangria_pt = sangria_mm * MM
    total = len(doc)
    metade1 = (total + 1) // 2
    metade2 = total - metade1
    num_folhas = max(metade1, metade2)
    espaco_pt = 0  # fixo

    # Página de referência para dimensões
    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    for i in range(num_folhas):
        # Determina índices das páginas (esquerda e direita) conforme padrão corta-empilha
        if i % 2 == 0:
            idx_esq = i if i < metade1 else None
            idx_dir = (metade1 + i) if (metade1 + i) < total else None
        else:
            idx_esq = (metade1 + i) if (metade1 + i) < total else None
            idx_dir = i if i < metade1 else None

        page_esq = doc[idx_esq] if idx_esq is not None else None
        page_dir = doc[idx_dir] if idx_dir is not None else None

        # Verifica se a página pertence à segunda metade (deve rotacionar)
        rot_esq = (idx_esq is not None and idx_esq >= metade1) if page_esq else False
        rot_dir = (idx_dir is not None and idx_dir >= metade1) if page_dir else False

        # Obtém áreas de corte e dimensões
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

        # Cria nova página no documento de saída
        nova = out.new_page(width=output_width, height=output_height)

        # Calcula posição Y (vertical) considerando sangria
        if sangria_pt > 0:
            y_esq, cabe_esq = calcular_y_com_sangria(output_height, h_esq, sangria_pt)
            y_dir, cabe_dir = calcular_y_com_sangria(output_height, h_dir, sangria_pt)
            sangria_aplicar_esq = cabe_esq and page_esq is not None
            sangria_aplicar_dir = cabe_dir and page_dir is not None
        else:
            y_esq = (output_height - h_esq) / 2
            y_dir = (output_height - h_dir) / 2
            sangria_aplicar_esq = sangria_aplicar_dir = False

        # Largura total ocupada pelas duas páginas + sangria
        largura_total = w_esq + w_dir + espaco_pt + 2 * sangria_pt
        x_inicio = (output_width - largura_total) / 2

        # Retângulos de destino
        x1 = x_inicio + sangria_pt
        destino1 = fitz.Rect(x1, y_esq, x1 + w_esq, y_esq + h_esq)
        destino1 = ajustar_para_media_box(destino1, output_width, output_height)

        x2 = x1 + w_esq + espaco_pt
        destino2 = fitz.Rect(x2, y_dir, x2 + w_dir, y_dir + h_dir)
        destino2 = ajustar_para_media_box(destino2, output_width, output_height)

        # Insere as páginas (com rotação se necessário)
        if page_esq and pagina_tem_conteudo(page_esq):
            nova.show_pdf_page(destino1, doc, idx_esq, clip=trim_esq, rotate=180 if rot_esq else 0)
        if page_dir and pagina_tem_conteudo(page_dir):
            nova.show_pdf_page(destino2, doc, idx_dir, clip=trim_dir, rotate=180 if rot_dir else 0)

        # Define trimbox (área de corte) – geralmente a primeira página
        nova.set_trimbox(destino1)

        # Aplica extensão de sangria (se houver)
        if sangria_pt > 0:
            if page_esq and pagina_tem_conteudo(page_esq) and sangria_aplicar_esq:
                aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1, sangria_pt,
                                                  bordas=['top', 'bottom', 'left'])
            if page_dir and pagina_tem_conteudo(page_dir) and sangria_aplicar_dir:
                aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2, sangria_pt,
                                                  bordas=['top', 'bottom', 'right'])

            # Define bleedbox (área de sangria) total
            bleed_total = fitz.Rect(
                destino1.x0 - sangria_pt,
                min(destino1.y0, destino2.y0) - sangria_pt,
                destino2.x1 + sangria_pt,
                max(destino1.y1, destino2.y1) + sangria_pt
            )
            bleed_total = ajustar_para_media_box(bleed_total, output_width, output_height)
            if bleed_total.width > 0 and bleed_total.height > 0:
                nova.set_bleedbox(bleed_total)

        # Desenha marcas de corte se solicitado
        if add_marcas:
            desenhar_marcas(nova, destino1, tam=5*MM, gap=2*MM, lado='esquerda')
            desenhar_marcas(nova, destino2, tam=5*MM, gap=2*MM, lado='direita')
            x_centro = destino1.x1
            desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1,
                                            tam=5*MM, gap=2*MM)

        # <-- NOVO: Insere marca d'água se texto fornecido
        if marca_dagua_texto:
            inserir_marca_dagua(
                nova,
                    marca_dagua_texto,
                    tamanho_fonte=8,
                    deslocamento_vertical=2.5*MM,      # sobe 2,5 mm (positivo = acima)
                    deslocamento_horizontal=-20*MM       # 5 mm para a direita (negativo para esquerda)
                )

    out.save(output_path)
    doc.close()
    out.close()
    return True

def dois_na_folha_rotacionado(input_path, output_path, output_width, output_height,
                              espacamento_mm=0, sangria_mm=3, add_marcas=True, marca_dagua_texto=None):   
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

        # <-- NOVO: Insere marca d'água
        if marca_dagua_texto:
            inserir_marca_dagua(
                nova,
                    marca_dagua_texto,
                    tamanho_fonte=8,
                    deslocamento_vertical=2.5*MM,      # sobe 2,5 mm (positivo = acima)
                    deslocamento_horizontal=-20*MM       # 5 mm para a direita (negativo para esquerda)
                )

    out.save(output_path)
    doc.close()
    out.close()
    return True

def um_na_folha(input_path, output_path, output_width, output_height,
                sangria_mm=3, add_marcas=True, marca_dagua_texto=None):   # <-- NOVO
    doc = fitz.open(input_path)
    out = fitz.open()
    sangria_pt = sangria_mm * MM
    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

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

        if pagina_tem_conteudo(page):
            nova.show_pdf_page(destino, doc, page.number, clip=trim_original)

        nova.set_trimbox(destino)

        if sangria_pt > 0 and sangria_aplicar and pagina_tem_conteudo(page):
            bleed = fitz.Rect(
                destino.x0 - sangria_pt,
                destino.y0 - sangria_pt,
                destino.x1 + sangria_pt,
                destino.y1 + sangria_pt
            )
            bleed = ajustar_para_media_box(bleed, output_width, output_height)
            if bleed.width > 0 and bleed.height > 0:
                nova.set_bleedbox(bleed)
            aplicar_sangria_espelhada_externa(page, nova, trim_original, destino, sangria_pt,
                                              bordas=['top', 'bottom', 'left', 'right'])

        if add_marcas:
            desenhar_marcas(nova, destino, tam=5*MM, gap=2*MM)

        if marca_dagua_texto:
            inserir_marca_dagua(
                nova,
                    marca_dagua_texto,
                    tamanho_fonte=8,
                    deslocamento_vertical=2.5*MM,      # sobe 2,5 mm (positivo = acima)
                    deslocamento_horizontal=-20*MM       # 5 mm para a direita (negativo para esquerda)
                )

    out.save(output_path)
    doc.close()
    out.close()
    return True

def dois_na_folha(input_path, output_path, output_width, output_height,
                  espacamento_mm=0, sangria_mm=3, add_marcas=True, marca_dagua_texto=None):   # <-- NOVO
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
            nova.show_pdf_page(destino1, doc, page.number, clip=trim_original)
            nova.show_pdf_page(destino2, doc, page.number, clip=trim_original)

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

        # <-- NOVO: Insere marca d'água
        if marca_dagua_texto:
            inserir_marca_dagua(
                nova,
                    marca_dagua_texto,
                    tamanho_fonte=8,
                    deslocamento_vertical=2.5*MM,      # sobe 2,5 mm (positivo = acima)
                    deslocamento_horizontal=-20*MM       # 5 mm para a direita (negativo para esquerda)
                )

    out.save(output_path)
    doc.close()
    out.close()
    return True

def Corta_empilha(input_path, output_path, output_width, output_height,
                  sangria_mm=3, add_marcas=True, marca_dagua_texto=None):   # <-- NOVO
    doc = fitz.open(input_path)
    out = fitz.open()
    sangria_pt = sangria_mm * MM
    total = len(doc)
    metade1 = (total + 1) // 2
    metade2 = total - metade1
    num_folhas = max(metade1, metade2)
    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    for i in range(num_folhas):
        if i % 2 == 0:
            idx_esq = i if i < metade1 else None
            idx_dir = (metade1 + i) if (metade1 + i) < total else None
        else:
            idx_esq = (metade1 + i) if (metade1 + i) < total else None
            idx_dir = i if i < metade1 else None

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

        if page_esq and pagina_tem_conteudo(page_esq):
            nova.show_pdf_page(destino1, doc, idx_esq, clip=trim_esq)
        if page_dir and pagina_tem_conteudo(page_dir):
            nova.show_pdf_page(destino2, doc, idx_dir, clip=trim_dir)

        nova.set_trimbox(destino1)

        if sangria_pt > 0:
            if page_esq and pagina_tem_conteudo(page_esq) and sangria_aplicar_esq:
                aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1, sangria_pt,
                                                  bordas=['top', 'bottom', 'left'])
            if page_dir and pagina_tem_conteudo(page_dir) and sangria_aplicar_dir:
                aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2, sangria_pt,
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
            desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1, tam=5*MM, gap=2*MM)

        # <-- NOVO: Insere marca d'água
        if marca_dagua_texto:
            inserir_marca_dagua(
                nova,
                    marca_dagua_texto,
                    tamanho_fonte=8,
                    deslocamento_vertical=2.5*MM,      # sobe 2,5 mm (positivo = acima)
                    deslocamento_horizontal=-20*MM       # 5 mm para a direita (negativo para esquerda)
                )

    out.save(output_path)
    doc.close()
    out.close()
    return True

def grampo(input_path, output_path, output_width, output_height,
           sangria_mm=3, add_marcas=True, marca_dagua_texto=None):   # <-- NOVO
    doc = fitz.open(input_path)
    out = fitz.open()
    sangria_pt = sangria_mm * MM
    total = len(doc)
    num_folhas = (total + 1) // 2
    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210*MM, 297*MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

    for i in range(num_folhas):
        if i % 2 == 0:
            idx_esq_base1 = total - i
            idx_dir_base1 = i + 1
        else:
            idx_esq_base1 = i + 1
            idx_dir_base1 = total - i

        idx_esq = idx_esq_base1 - 1 if idx_esq_base1 >= 1 and idx_esq_base1 <= total else None
        idx_dir = idx_dir_base1 - 1 if idx_dir_base1 >= 1 and idx_dir_base1 <= total else None

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

        if page_esq and pagina_tem_conteudo(page_esq):
            nova.show_pdf_page(destino1, doc, idx_esq, clip=trim_esq)
        if page_dir and pagina_tem_conteudo(page_dir):
            nova.show_pdf_page(destino2, doc, idx_dir, clip=trim_dir)

        nova.set_trimbox(destino1)

        if sangria_pt > 0:
            if page_esq and pagina_tem_conteudo(page_esq) and sangria_aplicar_esq:
                aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1, sangria_pt,
                                                  bordas=['top', 'bottom', 'left'])
            if page_dir and pagina_tem_conteudo(page_dir) and sangria_aplicar_dir:
                aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2, sangria_pt,
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
            desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1, tam=5*MM, gap=2*MM)

        # <-- NOVO: Insere marca d'água
        if marca_dagua_texto:
            inserir_marca_dagua(
                nova,
                    marca_dagua_texto,
                    tamanho_fonte=8,
                    deslocamento_vertical=2.5*MM,      # sobe 2,5 mm (positivo = acima)
                    deslocamento_horizontal=-20*MM       # 5 mm para a direita (negativo para esquerda)
                )

    out.save(output_path)
    doc.close()
    out.close()
    return True

# --- SISTEMA E INTERFACE MELHORADA ---

class MioloTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.processos = []
        self.processando = False
        self.thread = None
        self.log_queue = queue.Queue()
        self.dialog_queue = queue.Queue()      # fila para pedir diálogo
        self.dialog_response = queue.Queue()   # fila para receber resposta
        self.impressora_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Aguardando arquivo TXT...")
        self.setup_ui()
        self.atualizar_log_queue()
        self.check_dialog_queue()               # inicia verificação periódica
        
        logging.info("Sistema iniciado")

    def setup_ui(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        config_frame = ttk.LabelFrame(main_frame, text="Configurações de Impressão", bootstyle="INFO")
        config_frame.pack(fill='x', pady=5)
        
        ttk.Label(config_frame, text="Impressora:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        impressoras = [
            "OCÉ 1 - 192.168.200.80",
            "OCÉ 2 - 192.168.200.81", 
            "OCÉ 3 - 192.168.200.85"
        ]
        
        self.combo_impressora = ttk.Combobox(
            config_frame, 
            textvariable=self.impressora_var,
            values=impressoras,
            state="readonly", 
            width=30
        )
        self.combo_impressora.current(0)
        self.combo_impressora.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        
        ttk.Label(config_frame, text="IP:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.ip_label = ttk.Label(config_frame, text="192.168.200.80", font=('Consolas', 9))
        self.ip_label.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        
        self.combo_impressora.bind('<<ComboboxSelected>>', self.atualizar_ip_label)
        
        file_frame = ttk.LabelFrame(main_frame, text="Arquivo de Lote (TXT)", bootstyle="INFO")
        file_frame.pack(fill='x', pady=5)
        
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill='x', pady=5)
        
        ttk.Button(btn_frame, text="📂 Selecionar TXT", command=self.selecionar_txt, bootstyle="INFO").pack(side='left', padx=5)
        ttk.Button(btn_frame, text="🔄 Cancelar", command=self.cancelar_processamento, bootstyle="DANGER", state='disabled').pack(side='right', padx=5)
        
        self.file_label = ttk.Label(file_frame, text="Nenhum arquivo selecionado", relief="groove", anchor="center", padding=10)
        self.file_label.pack(fill='x', pady=5)
        
        preview_frame = ttk.LabelFrame(main_frame, text="📋 Processos Carregados", bootstyle="INFO")
        preview_frame.pack(fill='both', expand=True, pady=5)
        
        columns = ('Linha', 'Arquivo', 'Formato', 'Acabamento', 'Modo', 'QTD', 'QTD_Total')
        self.tree = ttk.Treeview(preview_frame, columns=columns, show='headings', height=6)
        
        self.tree.heading('Linha', text='Linha')
        self.tree.heading('Arquivo', text='Arquivo')
        self.tree.heading('Formato', text='Formato')
        self.tree.heading('Acabamento', text='Cód. Acab.')
        self.tree.heading('Modo', text='Modo Montagem')
        self.tree.heading('QTD', text='Qtd Envios')
        self.tree.heading('QTD_Total', text='Qtd Total')
        
        self.tree.column('Linha', width=50)
        self.tree.column('Arquivo', width=200)
        self.tree.column('Formato', width=80)
        self.tree.column('Acabamento', width=70)
        self.tree.column('Modo', width=120)
        self.tree.column('QTD', width=60)
        self.tree.column('QTD_Total', width=70)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        log_frame = ttk.LabelFrame(main_frame, text="📝 Processamento ", bootstyle="INFO")
        log_frame.pack(fill='both', expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=12, wrap='word', state='disabled', 
                                font=('Consolas', 9), bg='#2b2b2b', fg='#ffffff')
        self.log_text.pack(fill='both', expand=True, side='left')
        
        log_scroll = ttk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side='right', fill='y')
        
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill='x', pady=5)
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate', bootstyle="INFO-striped")
        self.progress.pack(fill='x', side='left', expand=True, padx=(0,10))
        
        self.progress_label = ttk.Label(progress_frame, text="0/0")
        self.progress_label.pack(side='right')
        
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill='x', pady=5)
        
        ttk.Label(status_frame, textvariable=self.status_var, bootstyle="SUCCESS").pack(side='left')
        
        self.btn_processar = ttk.Button(status_frame, text="▶ Processar e Enviar", 
                                        command=self.iniciar_processamento, 
                                        bootstyle="PRIMARY", width=20)
        self.btn_processar.pack(side='right')

    def atualizar_ip_label(self, event=None):
        impressora = self.impressora_var.get()
        if ' - ' in impressora:
            ip = impressora.split(' - ')[1].strip()
            self.ip_label.config(text=ip)
        else:
            self.ip_label.config(text=impressora)

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

    def check_dialog_queue(self):
        """Verifica se há pedidos de diálogo da thread de processamento."""
        try:
            req = self.dialog_queue.get_nowait()   # (req_id, mensagem, arquivo, copia)
            resposta = messagebox.askyesno(
                "Erro no envio",
                f"{req[1]}\n\nDeseja tentar novamente?"
            )
            self.dialog_response.put((req[0], resposta))
        except queue.Empty:
            pass
        finally:
            self.after(200, self.check_dialog_queue)  # verifica a cada 200ms

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

    def selecionar_txt(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos TXT", "*.txt")])
        if file_path:
            self.arquivo_txt = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.ler_arquivo_txt_miolo()

    def ler_arquivo_txt_miolo(self):
        try:
            self.processos = []
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            with open(self.arquivo_txt, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                
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
                    
                    # CORREÇÃO: Adicionar QTD_total (índice 10)
                    qtd = int(campos[2].strip()) if campos[2].strip().isdigit() else 1
                    qtd_total_str = campos[10].strip() if len(campos) > 10 else "0"
                    try:
                        qtd_total = int(qtd_total_str)
                    except:
                        qtd_total = 0
                    
                    processo = {
                        "LINHA": idx+1,
                        "CAMINHO": campos[0].strip(),
                        "ARQUIVO": campos[1].strip(),
                        "QTD": qtd,
                        "CODIGOACAB": cod_acab,
                        "QTD_total": qtd_total,
                        "NOMEFIC": campos[11].strip(),
                        "FORMATO": campos[12].strip()
                    }
                    
                    if cod_acab == "1":
                        if qtd == 1:
                            modo = "corta_empilha_rotacionado"
                        else:
                            modo = "dois_na_folha"
                    elif cod_acab == "2":
                        modo = "grampo"
                    elif cod_acab == "3":
                        modo = "um_na_folha"
                    elif cod_acab == "4":
                        modo = "Corta_empilha"
                    elif cod_acab == "5":
                        modo = "dois_na_folha_rotacionado"
                    else:
                        modo = "Corta_empilha"
                    
                    formato_str = processo["FORMATO"].split()[0]
                    try:
                        formato_limpo = formato_str.lower().replace(',', '.')
                        partes = formato_limpo.split('x')
                        dim1 = float(partes[0])
                        dim2 = float(partes[1])
                        
                        if cod_acab == "3":
                            largura_mm = min(dim1, dim2) * 10
                            altura_mm = max(dim1, dim2) * 10
                            espacamento = 0
                        else:
                            largura_mm = max(dim1, dim2) * 10
                            altura_mm = min(dim1, dim2) * 10
                            espacamento = 3 if f"{int(dim1)}x{int(dim2)}" in ["22x32", "32x22"] else 0
                        
                        processo.update({
                            "LARGURA_MM": largura_mm,
                            "ALTURA_MM": altura_mm,
                            "ESPACAMENTO": espacamento,
                            "MODO": modo,
                            "DESC_ACAB": desc_acab
                        })
                    except:
                        processo.update({
                            "LARGURA_MM": 480, 
                            "ALTURA_MM": 320, 
                            "ESPACAMENTO": 0,
                            "MODO": modo,
                            "DESC_ACAB": desc_acab
                        })
                    
                    self.processos.append(processo)
                    
                    # Inserir na treeview com as novas colunas
                    self.tree.insert('', 'end', values=(
                        f"L{processo['LINHA']}",
                        os.path.basename(processo['ARQUIVO']),
                        processo['FORMATO'].split()[0],
                        processo['CODIGOACAB'],
                        processo['MODO'],
                        processo['QTD'],
                        processo['QTD_total']
                    ))
            
            self.status_var.set(f"✅ {len(self.processos)} processos carregados")
            self.log(f"Arquivo carregado: {os.path.basename(self.arquivo_txt)}", "SUCESSO")
            self.log(f"Total de processos: {len(self.processos)}", "INFO")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo: {e}")
            self.log(f"ERRO ao ler arquivo: {e}", "ERRO")

    def cancelar_processamento(self):
        if self.processando:
            self.processando = False
            self.log("⚠️ Cancelando processamento...", "INFO")

    def iniciar_processamento(self):
        if not self.processos:
            messagebox.showwarning("Aviso", "Carregue um arquivo TXT primeiro.")
            return
        
        if self.processando:
            messagebox.showwarning("Aviso", "Já existe um processamento em andamento.")
            return
        
        self.processando = True
        self.btn_processar.config(state='disabled', text='⏳ Processando...')
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.processos)
        
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
        """Envia arquivo via protocolo LPR com parâmetro -J para nome do job"""
        try:
            # Comando LPR com -J para identificar o job
            comando = f'lpr -S {ip} -P "{formato}" -J "{job_name}" "{caminho_arquivo}"'
            
            self.log(f"📤 Enviando cópia {copia_numero}/{total_copias}...", "ENVIO")
            self.log(f"   IP: {ip}, Fila: {formato}, Job: {job_name}", "ENVIO")
            
            resultado = subprocess.run(
                comando, 
                shell=True, 
                capture_output=True, 
                text=True,
                timeout=30
            )
            
            # Log detalhado no console
            print(f"\n--- LPR DEBUG (cópia {copia_numero}) ---")
            print(f"Comando: {comando}")
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
            self.log(f"❌ Timeout no envio da cópia {copia_numero} (30s)", "ERRO")
            return False
        except Exception as e:
            self.log(f"❌ Exceção no envio: {str(e)}", "ERRO")
            print(f"Erro detalhado: {e}")
            return False

    def _executar(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        output_dir = os.path.join(desktop, "PDFs_PROCESSADOS")
        os.makedirs(output_dir, exist_ok=True)
        
        success_count = 0
        error_count = 0
        envio_count = 0
        envio_erro_count = 0
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        sessao_temp = os.path.join(PASTA_TRABALHO, f"sessao_{timestamp}")
        os.makedirs(sessao_temp, exist_ok=True)

        impressora_selecionada = self.impressora_var.get()
        ip_extraido = impressora_selecionada.split(' - ')[1].strip() if ' - ' in impressora_selecionada else impressora_selecionada
        
        self.log(f"\n{'='*50}", "INFO")
        self.log(f"🚀 INICIANDO PROCESSAMENTO", "INFO")
        self.log(f"📠 Impressora: {impressora_selecionada}", "INFO")
        self.log(f"📠 IP: {ip_extraido}", "INFO")
        self.log(f"📁 Pasta de saída: {output_dir}", "INFO")
        self.log(f"{'='*50}", "INFO")

        for i, p in enumerate(self.processos):
            if not self.processando:
                self.log("⚠️ Processamento cancelado pelo usuário", "INFO")
                break
                
            try:
                self.progress['value'] = i
                self.progress_label.config(text=f"{i+1}/{len(self.processos)}")
                
                self.log(f"\n--- Processando L{p['LINHA']}: {p['ARQUIVO']} ---", "INFO")
                self.log(f"📄 Formato: {p['FORMATO']}", "INFO")
                self.log(f"🔧 Modo: {p['MODO']}", "INFO")
                self.log(f"📊 QTD para envio: {p['QTD']}, QTD Total: {p['QTD_total']}", "INFO")
                
                origem = os.path.join(p["CAMINHO"], p["ARQUIVO"])
                if not os.path.exists(origem):
                    self.log(f"Arquivo não encontrado: {origem}", "ERRO")
                    error_count += 1
                    continue
                
                local = os.path.join(sessao_temp, f"L{p['LINHA']}_{p['ARQUIVO']}")
                shutil.copy2(origem, local)
                self.log(f"📋 Arquivo copiado para temporário", "INFO")
                
                output_path = os.path.join(output_dir, f"{p['NOMEFIC']}.pdf")
                w_pt, h_pt = p["LARGURA_MM"] * MM, p["ALTURA_MM"] * MM
                
                self.log(f"⚙️ Gerando PDF montado...", "INFO")
                
                # <-- NOVO: Passa o NOMEFIC como marca d'água
                if p["MODO"] == "dois_na_folha":
                    dois_na_folha(local, output_path, w_pt, h_pt,
                                  espacamento_mm=p["ESPACAMENTO"],
                                  marca_dagua_texto=p['NOMEFIC'])
                elif p["MODO"] == "grampo":
                    grampo(local, output_path, w_pt, h_pt,
                           marca_dagua_texto=p['NOMEFIC'])
                elif p["MODO"] == "Corta_empilha":
                    Corta_empilha(local, output_path, w_pt, h_pt,
                                  marca_dagua_texto=p['NOMEFIC'])
                elif p["MODO"] == "corta_empilha_rotacionado":
                    corta_empilha_rotacionado(local, output_path, w_pt, h_pt,
                                               marca_dagua_texto=p['NOMEFIC'])
                elif p["MODO"] == "dois_na_folha_rotacionado":
                    dois_na_folha_rotacionado(local, output_path, w_pt, h_pt,
                                               espacamento_mm=p["ESPACAMENTO"],
                                               marca_dagua_texto=p['NOMEFIC'])
                else:
                    um_na_folha(local, output_path, w_pt, h_pt,
                                marca_dagua_texto=p['NOMEFIC'])
                
                self.log(f"✅ PDF gerado: {os.path.basename(output_path)}", "SUCESSO")
                
                # Determinar quantidade de envios
                qtd_envios = p["QTD"]
                # Se for modo "dois_na_folha", cada envio corresponde a 2 cópias? Vamos usar a lógica:
                # Se modo == "dois_na_folha", então qtd_envios deve ser o número de folhas (cada folha tem 2 cópias)
                # Assumindo que QTD é o número de exemplares desejados, para dois_na_folha:
                if p["MODO"] == "dois_na_folha":
                    qtd_envios = math.ceil(p["QTD"] / 2)
                    self.log(f"🔢 Modo 2 na folha: {p['QTD']} exemplares -> {qtd_envios} envios", "INFO")
                # 🔒 LIMITADOR
                if qtd_envios > 10:
                    self.log(f"⚠️ Limite de envios excedido ({qtd_envios}). Enviando apenas 1 para teste.", "INFO")
                    qtd_envios = 1
                    
                formato_base = p["FORMATO"]
                job_name = p["NOMEFIC"]  # usar o nome do arquivo como job name
                
                # Loop de envios
                for envio_idx in range(qtd_envios):
                    if not self.processando:
                        self.log("⚠️ Cancelando envios...", "INFO")
                        break
                    
                    copia_num = envio_idx + 1
                    max_tentativas = 2
                    tentativa = 1
                    sucesso_envio = False

                    while tentativa <= max_tentativas and not sucesso_envio and self.processando:
                        if self.enviar_por_lpr(output_path, ip_extraido, formato_base, 
                                              job_name, copia_num, qtd_envios):
                            sucesso_envio = True
                            envio_count += 1
                        else:
                            if tentativa < max_tentativas:
                                # Gera um ID único para esta requisição
                                req_id = f"{p['LINHA']}_{copia_num}_{tentativa}"
                                msg = f"Falha ao enviar '{p['ARQUIVO']}' (cópia {copia_num})"
                                self.dialog_queue.put((req_id, msg, p['ARQUIVO'], copia_num))
                                time.sleep(3)
                                # Aguarda resposta do usuário (timeout 60s)
                                try:
                                    resposta = self.dialog_response.get(timeout=60)
                                    if resposta[0] == req_id and resposta[1]:
                                        tentativa += 1
                                        continue          # tenta novamente
                                    else:
                                        break              # usuário desistiu
                                except queue.Empty:
                                    self.log(f"⏱️ Timeout aguardando resposta do usuário", "ERRO")
                                    break
                            else:
                                # Esgotou as tentativas
                                break

                    if not sucesso_envio:
                        envio_erro_count += 1
                        self.log(f"❌ Falha definitiva no envio da cópia {copia_num} de '{p['ARQUIVO']}'", "ERRO")

                    # Pausa de 3 segundos entre envios (mesmo que tenha falhado)
                    if envio_idx < qtd_envios - 1:
                        time.sleep(3)
                
                success_count += 1
                
            except Exception as e:
                self.log(f"❌ Erro: {str(e)}", "ERRO")
                error_count += 1
                logging.error(f"Erro detalhado: {e}", exc_info=True)

        shutil.rmtree(sessao_temp, ignore_errors=True)
        
        self.log(f"\n{'='*50}", "INFO")
        self.log(f"📊 RESUMO FINAL", "INFO")
        self.log(f"{'='*50}", "INFO")
        self.log(f"✅ Montagens com sucesso: {success_count}", "SUCESSO")
        self.log(f"❌ Montagens com erro: {error_count}", "ERRO" if error_count > 0 else "INFO")
        self.log(f"🚀 Envios com sucesso: {envio_count}", "SUCESSO")
        self.log(f"❌ Envios com erro: {envio_erro_count}", "ERRO" if envio_erro_count > 0 else "INFO")
        self.log(f"{'='*50}", "INFO")
        
        self.status_var.set(f"Concluído: {success_count} OK, {error_count} Erros | Enviados: {envio_count}")
        
        self.processando = False
        self.btn_processar.config(state='normal', text='▶ Processar e Enviar')
        self.progress['value'] = 0
        self.progress_label.config(text="0/0")
        
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
                          f"📁 Pasta: {output_dir}")

class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Processamento de Arquivos")
        self.root.geometry("1100x850")
        #abrir maximizado
        self.root.state('zoomed')
        Style(theme="darkly")
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=5, pady=5)
        notebook.add(MioloTab(notebook, self), text="Miolo")

if __name__ == "__main__":
    try:
        data_limite = datetime.strptime(DATA_LIMIT, "%d/%m/%Y")
        if datetime.now() <= data_limite:
            root = TkinterDnD.Tk()
            app = Application(root)
            root.mainloop()
        else: 
            messagebox.showerror("Erro", "Licença expirada.")
    except Exception as e: 
        print(f"Erro fatal: {e}")
        messagebox.showerror("Erro Fatal", str(e))