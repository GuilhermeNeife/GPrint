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

import fitz  # PyMuPDF
from PyPDF2 import PdfMerger
from pikepdf import Pdf, Rectangle

from reportlab.lib.pagesizes import A3, A4
from reportlab.pdfgen import canvas

from PIL import Image

DATA_LIMITE = "31/07/2027"
# --- CONSTANTES ---
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

# --- FUNÇÕES DE LÓGICA DE PDF (mantidas iguais) ---

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

# --- FUNÇÕES DE MONTAGEM (mantidas iguais) ---
def corta_empilha_rotacionado(input_path, output_path, output_width, output_height, sangria_mm=3, add_marcas=True):
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

    out.save(output_path)
    doc.close()
    out.close()
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
def um_na_folha(input_path, output_path, output_width, output_height, sangria_mm=3, add_marcas=True):
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

    out.save(output_path)
    doc.close()
    out.close()
    return True

def dois_na_folha(input_path, output_path, output_width, output_height, espacamento_mm=0, sangria_mm=3, add_marcas=True):
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

    out.save(output_path)
    doc.close()
    out.close()
    return True

def Corta_empilha(input_path, output_path, output_width, output_height, sangria_mm=3, add_marcas=True):
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

    out.save(output_path)
    doc.close()
    out.close()
    return True

def grampo(input_path, output_path, output_width, output_height, sangria_mm=3, add_marcas=True):
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

    out.save(output_path)
    doc.close()
    out.close()
    return True

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
        self.processos = []
        self.processando = False
        self.thread = None
        self.log_queue = queue.Queue()
        self.impressora_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Aguardando arquivo TXT...")
        self.setup_ui()
        self.atualizar_log_queue()
        
        logging.info("Sistema iniciado")

    def setup_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")

        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(3, weight=1)

        # ================= CONFIG =================
        config_frame = ttk.LabelFrame(main_frame, text="⚙️ Configurações de Impressão", bootstyle=(PRIMARY))
        config_frame.grid(row=0, column=0, sticky="ew", pady=5)

        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="Impressora:").grid(row=0, column=0, padx=5, pady=5, sticky="w")

        impressoras = [
            "OCÉ 1 - 192.168.200.80",
            "OCÉ 2 - 192.168.200.81",
            "OCÉ 3 - 192.168.200.85",
            "ATEXCO - 192.168.200.121"
        ]

        self.combo_impressora = ttk.Combobox(
            config_frame,
            textvariable=self.impressora_var,
            values=impressoras,
            state="readonly"
        )
        self.combo_impressora.current(0)
        self.combo_impressora.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(config_frame, text="IP selecionado:").grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.ip_label = ttk.Label(
            config_frame,
            text="192.168.200.80",
            font=("Consolas", 10, "bold"),
            bootstyle=(PRIMARY)
        )
        self.ip_label.grid(row=1, column=1, padx=5, pady=2, sticky="w")

        self.combo_impressora.bind('<<ComboboxSelected>>', self.atualizar_ip_label)

        # ================= ARQUIVO =================
        file_frame = ttk.LabelFrame(main_frame, text="📂 Arquivo de Lote (TXT)", bootstyle=(PRIMARY))
        file_frame.grid(row=1, column=0, sticky="ew", pady=5)

        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill='x', pady=5)

        self.btn_select = ttk.Button(
            btn_frame,
            text="Selecionar TXT",
            command=self.selecionar_txt,
            bootstyle=(PRIMARY)
        )
        self.btn_select.pack(side='left', padx=5)



        self.file_label = ttk.Label(
            file_frame,
            text="Nenhum arquivo selecionado",
            anchor="center",
            padding=10,
            bootstyle="secondary"
        )
        self.file_label.pack(fill='x', padx=5, pady=5)

        # ================= TABELA =================
        preview_frame = ttk.LabelFrame(main_frame, text="📋 Processos", bootstyle=(PRIMARY))
        preview_frame.grid(row=2, column=0, sticky="nsew", pady=5)

        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        columns = ('Linha', 'Arquivo', 'Formato', 'Acabamento', 'Modo', 'QTD', 'Total')

        self.tree = ttk.Treeview(preview_frame, columns=columns, show='headings')

        for col in columns:
            self.tree.heading(col, text=col)

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

        # ================= PROGRESS =================
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=0, sticky="ew", pady=5)

        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            bootstyle="success-striped"
        )
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.progress_label = ttk.Label(progress_frame, text="0/0")
        self.progress_label.grid(row=0, column=1)

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
                
                if p["MODO"] == "dois_na_folha":
                    dois_na_folha(local, output_path, w_pt, h_pt, espacamento_mm=p["ESPACAMENTO"])
                elif p["MODO"] == "grampo":
                    grampo(local, output_path, w_pt, h_pt)
                elif p["MODO"] == "Corta_empilha":
                    Corta_empilha(local, output_path, w_pt, h_pt)
                elif p["MODO"] == "corta_empilha_rotacionado":
                    corta_empilha_rotacionado(local, output_path, w_pt, h_pt)
                elif p["MODO"] == "dois_na_folha_rotacionado":
                    dois_na_folha_rotacionado(local, output_path, w_pt, h_pt, espacamento_mm=p["ESPACAMENTO"])
                else:
                    um_na_folha(local, output_path, w_pt, h_pt)
                
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
                    if self.enviar_por_lpr(output_path, ip_extraido, formato_base, 
                                          job_name, copia_num, qtd_envios):
                        envio_count += 1
                    else:
                        envio_erro_count += 1
                    
                    # Pequena pausa entre envios (opcional)
                    time.sleep(0.5)
                
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
    def selecionar_txt(self):
        """Seleciona um arquivo TXT usando diálogo de arquivo"""
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
                            modo = "Corta_empilha"
                        else:
                            modo = "dois_na_folha"
                    elif cod_acab == "2":
                        modo = "grampo"
                    elif cod_acab == "3":
                        modo = "um_na_folha"
                    elif cod_acab == "4":
                        modo = "corta_empilha_rotacionado"
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

            

    def atualizar_log(self, mensagem):
        """Atualiza o log de mensagens"""
        self.txt_view.config(state='normal')
        self.txt_view.insert(tk.END, mensagem + "\n")
        self.txt_view.see(tk.END)
        self.txt_view.config(state='disabled')
        self.winfo_toplevel().update()  # Atualiza a interface



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
        #abrir maximizado
        self.root.state('zoomed')
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
        about_window.title("Sobre o Sistema")
        about_window.geometry("400x300")
        about_window.resizable(False, False)
        #Não minimizar
        about_window.overrideredirect(False)
        

        # Centralizar janela
        about_window.update_idletasks()
        x = (about_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (about_window.winfo_screenheight() // 2) - (300 // 2)
        about_window.geometry(f"+{x}+{y}")

        main_frame = ttk.Frame(about_window, padding=25)
        main_frame.pack(fill=BOTH, expand=True)

        # ===== Nome =====
        ttk.Label(
            main_frame,
            text="Processador de Arquivos",
            font=('Segoe UI', 14, 'bold'),
            bootstyle="primary"
        ).pack()

        # ===== Versão =====
        ttk.Label(
            main_frame,
            text="Versão 3.0",
            font=('Segoe UI', 10),
            bootstyle="secondary"
        ).pack(pady=(2, 10))

        # ===== Linha separadora =====
        ttk.Separator(main_frame).pack(fill='x', pady=10)

        # ===== Descrição =====
        ttk.Label(
            main_frame,
            text="Sistema desenvolvido para automação\n"
                "de processamento e envio de arquivos POD.",
            font=('Segoe UI', 10),
            justify="center"
        ).pack(pady=5)

        # ===== Rodapé =====
        ttk.Label(
            main_frame,
            text="© 2026 • Todos os direitos reservados",
            font=('Segoe UI', 8),
            bootstyle="secondary"
        ).pack(side=BOTTOM, pady=(15, 5))

        # ===== Botão =====
        ttk.Button(
            main_frame,
            text="Fechar",
            command=about_window.destroy,
            bootstyle="primary",
            width=15
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