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

DATA_LIMITE = "31/10/2027"
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
# --- FUNÇÃO PRINCIPAL DE MONTAGEM (com correção na chamada da sangria) ---
def corta_empilha_rotacionado(input_path, output_path, output_width, output_height,
                              sangria_mm=3, add_marcas=True,
                              texto_offset_normal_mm=0.5, texto_offset_180_mm=0.5):
    """
    Montagem corta‑empilha inteligente (frente e verso):
    - Caderno único garantido (total ajustado para múltiplo de 4).
    - Margem vertical adaptativa (5 mm, expandida para acomodar textos).
    - Textos de identificação centralizados na margem de 5 mm.
    - Textos desenhados por último → sobrepõem sangrias e marcas.
    """

    doc = fitz.open(input_path)

    if len(doc) == 0:
        print("Erro: O documento PDF está vazio.")
        doc.close()
        return False

    import os
    nome_arquivo = os.path.basename(input_path)

    def _dimensoes_referencia(doc):
        try:
            ref = doc.load_page(0)
            if ref is not None:
                return ref.rect.width, ref.rect.height
        except Exception:
            pass
        try:
            ref = doc[0]
            if ref is not None:
                return ref.rect.width, ref.rect.height
        except Exception:
            pass
        print("Aviso: Usando tamanho A4 como referência para páginas em branco.")
        return 210 * MM, 297 * MM

    total_original = len(doc)
    if total_original % 2 != 0:
        w_ref, h_ref = _dimensoes_referencia(doc)
        doc.new_page(width=w_ref, height=h_ref)
        print(f"Total ímpar ({total_original}). Adicionada 1 página em branco (total: {len(doc)}).")

    if len(doc) % 4 != 0:
        w_ref, h_ref = _dimensoes_referencia(doc)
        for _ in range(2):
            doc.new_page(width=w_ref, height=h_ref)
        print(f"Total não múltiplo de 4. Adicionadas 2 páginas em branco (total: {len(doc)}).")

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
        font_size = 5                 # ajustado para caber na margem de 5 mm
        altura_texto = 4 * MM         # 4 mm de altura
        margem_min = 0.5 * MM

        # --- Página esquerda ---
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

        # --- Página direita ---
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

        if sangria_pt > 0:
            cabe_esq = (y_esq >= sangria_pt and y_esq + h_esq <= output_height - sangria_pt)
            cabe_dir = (y_dir >= sangria_pt and y_dir + h_dir <= output_height - sangria_pt)
        else:
            cabe_esq = cabe_dir = True

        largura_total = w_esq + w_dir + espaco_horizontal_pt + 2 * sangria_pt
        x_inicio = 0.0 if largura_total > output_width else (output_width - largura_total) / 2
        x1 = x_inicio + sangria_pt

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

        # 2. Sangria
        if sangria_pt > 0:
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

        # 4. Trimbox (definida após tudo, mas não afeta a renderização)
        trim_union = fitz.Rect(
            destino_esq.x0, min(destino_esq.y0, destino_dir.y0),
            destino_dir.x1, max(destino_esq.y1, destino_dir.y1)
        ).intersect(mediabox)
        if trim_union.is_empty or trim_union.get_area() == 0:
            trim_union = mediabox
        nova.set_trimbox(trim_union)

        # 5. TEXTOS DE IDENTIFICAÇÃO (POR ÚLTIMO → SOBREPÕE TUDO)
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
                  sangria_mm=3, add_marcas=True, comprimir_entrada=True):
    """
    Montagem corta-empilha com caderno único.
    Se comprimir_entrada=True, o PDF original é comprimido antes da imposição.
    """

    # Se solicitado, comprime o PDF de entrada num arquivo temporário
    if comprimir_entrada:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_input = tmp.name
        comprimir_pdf(input_path, tmp_input)
        doc = fitz.open(tmp_input)
    else:
        doc = fitz.open(input_path)

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

    ref_page = doc[0]
    trim_ref = obter_area_corte_segura(ref_page, 210 * MM, 297 * MM)
    w_ref, h_ref = trim_ref.width, trim_ref.height

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

        # Inserir páginas com rotação
        if page_esq and pagina_tem_conteudo(page_esq):
            nova.show_pdf_page(destino1, doc, idx_esq, clip=trim_esq, rotate=rot_esq)
        if page_dir and pagina_tem_conteudo(page_dir):
            nova.show_pdf_page(destino2, doc, idx_dir, clip=trim_dir, rotate=rot_dir)

        # Trimbox
        trim_union = fitz.Rect(destino1.x0, min(destino1.y0, destino2.y0),
                               destino2.x1, max(destino1.y1, destino2.y1))
        trim_union = ajustar_para_media_box(trim_union, output_width, output_height)
        nova.set_trimbox(trim_union)

        # Sangria (com correção da BleedBox)
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
            media_box = fitz.Rect(0, 0, output_width, output_height)
            bleed_total = bleed_total.intersect(media_box)

            if (not bleed_total.is_empty
                and bleed_total.width > 0
                and bleed_total.height > 0
                and bleed_total.contains(trim_union)):
                nova.set_bleedbox(bleed_total)

        # Marcas
        if add_marcas:
            desenhar_marcas(nova, destino1, tam=5 * MM, gap=2 * MM, lado='esquerda')
            desenhar_marcas(nova, destino2, tam=5 * MM, gap=2 * MM, lado='direita')
            x_centro = destino1.x1
            desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1,
                                            tam=5 * MM, gap=2 * MM)
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
    if comprimir_entrada:
        try:
            os.unlink(tmp_input)
        except:
            pass
        print("PDF comprimido usado como entrada e depois removido.")

    print("Montagem concluída.")
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
            "ATEXCO - 192.168.200.121",
            "OCÉ 1 - 192.168.200.80",
            "OCÉ 2 - 192.168.200.81",
            "OCÉ 3 - 192.168.200.85",
            "RICOH - 192.168.200.84"
            
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


        self.impressao_auto_var = tk.BooleanVar(value=True)

        self.toggle_impressao = ttk.Checkbutton(
            config_frame,
            text="Imprimir automaticamente após processar",
            variable=self.impressao_auto_var,
            bootstyle="success-round-toggle"
        )

        self.toggle_impressao.grid(
            row=2,
            column=0,
            columnspan=2,
            padx=5,
            pady=10,
            sticky="w"
        )

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

            import shutil
            import tempfile
            import os
            import subprocess
            import time
            import uuid

            # =========================
            # LOG INICIAL
            # =========================

            self.log(
                f"📤 Enviando cópia {copia_numero}/{total_copias}...",
                "ENVIO"
            )

            self.log(
                f"   IP: {ip}, Fila: {formato}, Job: {job_name}",
                "ENVIO"
            )

            # =========================
            # CRIA PDF TEMPORÁRIO ÚNICO
            # =========================

            pasta_temp = tempfile.gettempdir()

            nome_temp = f"{uuid.uuid4()}.pdf"

            arquivo_temp = os.path.join(
                pasta_temp,
                nome_temp
            )

            # copia PDF original
            shutil.copy2(caminho_arquivo, arquivo_temp)

            # pequena pausa
            time.sleep(1)

            # =========================
            # COMANDO LPR
            # =========================

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

            # =========================
            # REMOVE TEMPORÁRIO
            # =========================

            try:
                os.remove(arquivo_temp)
            except:
                pass

            # =========================
            # DEBUG
            # =========================

            print(f"\n--- LPR DEBUG (cópia {copia_numero}) ---")
            print(f"Arquivo original: {caminho_arquivo}")
            print(f"Arquivo temp: {arquivo_temp}")
            print(f"Comando: {resultado.args}")
            print(f"Return code: {resultado.returncode}")
            print(f"STDOUT: {resultado.stdout}")
            print(f"STDERR: {resultado.stderr}")
            print(f"-----------------\n")

            # =========================
            # LOGS
            # =========================

            if resultado.stdout:
                self.log(
                    f"   Saída: {resultado.stdout}",
                    "ENVIO"
                )

            if resultado.stderr:
                self.log(
                    f"   Erro: {resultado.stderr}",
                    "ERRO"
                )

            # =========================
            # RESULTADO
            # =========================

            if resultado.returncode == 0:

                self.log(
                    f"✅ Cópia {copia_numero} enviada com sucesso",
                    "SUCESSO"
                )

                return True

            else:

                self.log(
                    f"❌ Falha no envio da cópia {copia_numero} "
                    f"(código {resultado.returncode})",
                    "ERRO"
                )

                return False

        except subprocess.TimeoutExpired:

            self.log(
                f"❌ Timeout no envio da cópia {copia_numero} (120s)",
                "ERRO"
            )

            return False

        except Exception as e:

            self.log(
                f"❌ Exceção no envio: {str(e)}",
                "ERRO"
            )

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
                job_name = f'{p["NOMEFIC"]}'  # usar o nome do arquivo como job name
                
                # Loop de envios
                # LOOP DE ENVIOS
                for envio_idx in range(qtd_envios):

                    if not self.processando:
                        self.log("⚠️ Cancelando envios...", "INFO")
                        break

                    copia_num = envio_idx + 1

                    # AGUARDA ARQUIVO LIBERAR
                    timeout_arquivo = 10
                    inicio = time.time()

                    while True:
                        try:
                            with open(output_path, "rb") as f:
                                f.read(1)

                            # arquivo liberado
                            break

                        except Exception as e:
                            print(f"Arquivo ainda bloqueado: {e}")

                        if time.time() - inicio > timeout_arquivo:
                            raise Exception("Arquivo não liberou para leitura")

                        time.sleep(0.5)

                    # ENVIA
                    # =========================
                    # VERIFICA TOGGLE
                    # =========================

                    if self.impressao_auto_var.get():

                        self.log(
                            "🖨️ Impressão automática ativada",
                            "INFO"
                        )

                        # ENVIA
                        if self.enviar_por_lpr(
                            output_path,
                            ip_extraido,
                            formato_base,
                            job_name,
                            copia_num,
                            qtd_envios
                        ):

                            envio_count += 1

                        else:

                            envio_erro_count += 1

                    else:

                        self.log(
                            "⚠️ Impressão automática desativada",
                            "INFO"
                        )

                    time.sleep(2.5)
                
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

            # 🔥 TENTA DIFERENTES ENCODINGS
            lines = None
            encodings = ['utf-8', 'cp1252', 'latin-1']

            for enc in encodings:
                try:
                    with open(self.arquivo_txt, 'r', encoding=enc) as f:
                        lines = f.readlines()

                    self.log(f"Arquivo lido com encoding: {enc}", "INFO")
                    break

                except UnicodeDecodeError:
                    continue

            # SE NÃO CONSEGUIR LER
            if lines is None:
                raise Exception("Não foi possível identificar o encoding do arquivo.")

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

                # QTD
                qtd = int(campos[2].strip()) if campos[2].strip().isdigit() else 1

                # QTD TOTAL
                qtd_total_str = campos[10].strip() if len(campos) > 10 else "0"

                try:
                    qtd_total = int(qtd_total_str)
                except:
                    qtd_total = 0

                processo = {
                    "LINHA": idx + 1,
                    "CAMINHO": campos[0].strip(),
                    "ARQUIVO": campos[1].strip(),
                    "QTD": qtd,
                    "CODIGOACAB": cod_acab,
                    "QTD_total": qtd_total,
                    "NOMEFIC": campos[11].strip(),
                    "FORMATO": campos[12].strip(),
                    "IDENTIFIC": campos[13].strip()
                }

                # DEFINIÇÃO DE MODO
                if cod_acab == "1":
                    modo = "Corta_empilha" if qtd == 1 else "dois_na_folha"

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

                # EXTRAÇÃO DO FORMATO
                formato_str = processo["FORMATO"].split('_')[0]

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

                # TREEVIEW
                self.tree.insert('', 'end', values=(
                    f"L{processo['LINHA']}",
                    os.path.basename(processo['ARQUIVO']),
                    processo['FORMATO'].split('_')[0],
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

        #aba erros
        self.aba_erros = ErrosTab(self.notebook, self)
        self.notebook.add(self.aba_erros, text="Erros")
        
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
            text="Enviar para impressora automaticamente após processar",
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

            # Pasta final onde os arquivos mesclados ficarão (mesmo diretório dos TXTs selecionados)
            pasta_destino = os.path.dirname(self.arquivos_selecionados[0])

            # CORREÇÃO: Iterar sobre as chaves (identific, filacolor)
            for chave, lista_pdfs in pdfs_para_mesclar.items():
                identific, filacolor = chave  # Desempacotar a chave
                progress_text = f"Agrupando PDFs ({identifics_processados + 1}/{total_identifics})..."
                progress.update_progress(80 + ((identifics_processados / total_identifics) * 20), progress_text)

                # 1. Criar o arquivo mesclado DENTRO da pasta temporária
                nome_temp = os.path.join(pasta_temp, f"{identific}.pdf")
                self.mesclar_pdfs(lista_pdfs, nome_temp, filacolor)

                # 2. Agora mover o arquivo da pasta temporária para a pasta final
                nome_final = os.path.join(pasta_destino, f"{identific}.pdf")
                shutil.move(nome_temp, nome_final)
                logging.info(f"Arquivo movido de {nome_temp} para {nome_final}")

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

    def mesclar_pdfs(self, lista_pdfs, nome_saida, filacolor):
        try:
            merger = PdfMerger()
            for pdf in lista_pdfs:
                merger.append(pdf)

            merger.write(nome_saida)
            merger.close()
            logging.info(f"Arquivo mesclado salvo: {nome_saida}")

            if self.imprimir_automaticamente.get():
                self.enviar_para_impressora(nome_saida, os.path.basename(nome_saida), filacolor)

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
# ═══════════════ CLASSE DA ABA "ERROS" ═══════════════




class ErrosTab(ttk.Frame):

    def __init__(self, parent, app=None):
        super().__init__(parent)
        self.app = app
        self.lotes = []
        self.pdfs = []
        self.fila = set()
        self.fila_chaves = []
        self.setup_ui()
        self.carregar_lotes()

    def setup_ui(self):
        main = ttk.Frame(self, padding=5)
        main.pack(fill=tk.BOTH, expand=True)

        # Lote
        f_lote = ttk.LabelFrame(main, text="Lote (digite ou selecione)", padding=5)
        f_lote.pack(fill=tk.X, pady=(0, 5))
        self.combo_lote = ttk.Combobox(f_lote, width=65)
        self.combo_lote.pack(fill=tk.X, padx=5, pady=2)
        self.combo_lote.bind("<KeyRelease>", self.filtrar_lotes)
        self.combo_lote.bind("<<ComboboxSelected>>", self.carregar_pdfs)

        # Lista de PDFs
        f_lista = ttk.LabelFrame(main, text="PDFs do lote (seleção múltipla)", padding=5)
        f_lista.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        f_filtro = ttk.Frame(f_lista)
        f_filtro.pack(fill=tk.X, pady=2)
        self.filtro_entry = ttk.Entry(f_filtro, width=40)
        self.filtro_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.filtro_entry.bind("<KeyRelease>", self.filtrar_pdfs)
        ttk.Button(f_filtro, text="Limpar filtro", command=self.limpar_filtro).pack(side=tk.RIGHT, padx=5)

        frame_pdf = ttk.Frame(f_lista)
        frame_pdf.pack(fill=tk.BOTH, expand=True)
        scroll_pdf = ttk.Scrollbar(frame_pdf, orient=tk.VERTICAL)
        self.lista_pdfs = tk.Listbox(frame_pdf, width=80, height=6,
                                     selectmode=tk.MULTIPLE,
                                     yscrollcommand=scroll_pdf.set)
        scroll_pdf.config(command=self.lista_pdfs.yview)
        self.lista_pdfs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_pdf.pack(side=tk.RIGHT, fill=tk.Y)

        botoes_pdf = ttk.Frame(f_lista)
        botoes_pdf.pack(fill=tk.X, pady=2)
        ttk.Button(botoes_pdf, text="➕ Adicionar selecionado(s)",
                   command=self.adicionar_da_lista).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_pdf, text="📁 Adicionar todos visíveis",
                   command=self.adicionar_todos).pack(side=tk.LEFT, padx=5)

        # Fila de processamento
        f_fila = ttk.LabelFrame(main, text="Fila de processamento", padding=5)
        f_fila.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        self.label_contador_fila = ttk.Label(f_fila, text="0 item(ns) na fila")
        self.label_contador_fila.pack(anchor=tk.W)

        frame_fila = ttk.Frame(f_fila)
        frame_fila.pack(fill=tk.BOTH, expand=True)
        scroll_fila = ttk.Scrollbar(frame_fila, orient=tk.VERTICAL)
        self.lista_fila = tk.Listbox(frame_fila, width=80, height=6,
                                     selectmode=tk.MULTIPLE,
                                     yscrollcommand=scroll_fila.set)
        scroll_fila.config(command=self.lista_fila.yview)
        self.lista_fila.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_fila.pack(side=tk.RIGHT, fill=tk.Y)

        botoes_fila = ttk.Frame(f_fila)
        botoes_fila.pack(fill=tk.X, pady=2)
        ttk.Button(botoes_fila, text="Remover selecionado(s)",
                   command=self.remover_item).pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_fila, text="Limpar fila",
                   command=self.limpar_lista).pack(side=tk.LEFT, padx=5)

        # Rodapé (progresso + botão)
        rodape = ttk.Frame(main, padding=5)
        rodape.pack(side=tk.BOTTOM, fill=tk.X)
        f_progresso = ttk.Frame(rodape)
        f_progresso.pack(fill=tk.X, pady=2)
        self.progress = ttk.Progressbar(f_progresso, length=400, mode='determinate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.label_status = ttk.Label(f_progresso, text="Pronto")
        self.label_status.pack(side=tk.RIGHT, padx=5)
        self.btn_processar = ttk.Button(rodape, text="🚀 Processar Tudo",
                                        command=self.processar_tudo)
        self.btn_processar.pack(pady=5)

    # ───── Lógica de negócio (idêntica à classe App original) ─────
    def carregar_lotes(self):
        self.lotes = sorted(
            [p for p in os.listdir(BASE1) if os.path.isdir(os.path.join(BASE1, p))],
            reverse=True
        )
        self.combo_lote['values'] = self.lotes

    def filtrar_lotes(self, event=None):
        texto = self.combo_lote.get().lower()
        if len(texto) < 4:
            self.combo_lote['values'] = self.lotes
            return
        filtrados = [l for l in self.lotes if texto in l.lower()]
        self.combo_lote['values'] = filtrados
        if filtrados:
            self.combo_lote.event_generate('<Down>')

    def carregar_pdfs(self, event=None):
        lote = self.combo_lote.get()
        caminho = os.path.join(BASE1, lote)
        try:
            self.pdfs = [f for f in os.listdir(caminho) if f.lower().endswith(".pdf")]
            self.filtro_entry.delete(0, tk.END)
            self.atualizar_lista(self.pdfs)
        except:
            self.pdfs = []
            self.atualizar_lista([])

    def atualizar_lista(self, lista):
        self.lista_pdfs.delete(0, tk.END)
        for item in lista:
            self.lista_pdfs.insert(tk.END, item)

    def filtrar_pdfs(self, event=None):
        filtro = self.filtro_entry.get().lower()
        filtrados = [f for f in self.pdfs if filtro in f.lower()]
        self.atualizar_lista(filtrados)

    def limpar_filtro(self):
        self.filtro_entry.delete(0, tk.END)
        self.atualizar_lista(self.pdfs)

    def _add_fila(self, chave, nome, lote):
        if chave in self.fila:
            return False
        self.fila.add(chave)
        self.fila_chaves.append(chave)
        self.lista_fila.insert(tk.END, f"{lote} - {nome}")
        return True

    def adicionar_da_lista(self):
        lote = self.combo_lote.get()
        selecionados = self.lista_pdfs.curselection()
        if not lote or not selecionados:
            return
        adicionados = 0
        ja_estavam = 0
        for idx in selecionados:
            nome = self.lista_pdfs.get(idx)
            chave = f"{lote}|{nome}"
            if self._add_fila(chave, nome, lote):
                adicionados += 1
            else:
                ja_estavam += 1
        self._atualizar_contador_fila()
        if adicionados:
            self.label_status.config(text=f"{adicionados} PDF(s) adicionado(s) à fila")
        if ja_estavam:
            messagebox.showinfo("Duplicados ignorados",
                                f"{ja_estavam} item(ns) já estavam na fila e foram ignorados.")

    def adicionar_todos(self):
        lote = self.combo_lote.get()
        if not lote:
            return
        todos = [self.lista_pdfs.get(i) for i in range(self.lista_pdfs.size())]
        if not todos:
            return
        if not messagebox.askyesno("Confirmar",
                                   f"Adicionar todos os {len(todos)} PDFs visíveis na fila?"):
            return
        adicionados = 0
        for nome in todos:
            chave = f"{lote}|{nome}"
            if self._add_fila(chave, nome, lote):
                adicionados += 1
        self._atualizar_contador_fila()
        self.label_status.config(text=f"{adicionados} PDF(s) adicionado(s) à fila")
        if adicionados < len(todos):
            messagebox.showinfo("Aviso", "Alguns itens já estavam na fila e foram ignorados.")

    def remover_item(self):
        selecionados = self.lista_fila.curselection()
        if not selecionados:
            return
        for idx in sorted(selecionados, reverse=True):
            chave = self.fila_chaves[idx]
            self.fila.remove(chave)
            del self.fila_chaves[idx]
            self.lista_fila.delete(idx)
        self._atualizar_contador_fila()
        self.label_status.config(text="Item(ns) removido(s) da fila")

    def limpar_lista(self):
        if not self.fila:
            return
        if messagebox.askyesno("Limpar fila", "Deseja realmente limpar toda a fila?"):
            self.lista_fila.delete(0, tk.END)
            self.fila.clear()
            self.fila_chaves.clear()
            self._atualizar_contador_fila()
            self.label_status.config(text="Fila limpa")

    def _atualizar_contador_fila(self):
        total = len(self.fila_chaves)
        self.label_contador_fila.config(text=f"{total} item(ns) na fila")

    def processar_tudo(self):
        if not self.fila_chaves:
            messagebox.showwarning("Fila vazia", "Adicione PDFs à fila antes de processar.")
            return
        total = len(self.fila_chaves)
        if not messagebox.askyesno("Iniciar processamento",
                                   f"Processar {total} arquivo(s) da fila?"):
            return

        self.btn_processar.config(state=tk.DISABLED)
        self.progress["maximum"] = total
        self.progress["value"] = 0

        resultados = []
        count = 0

        for chave in self.fila_chaves:
            lote, nome = chave.split("|")
            self.label_status.config(text=f"Processando {count+1}/{total}: {nome}")
            self.update()  # atualiza a interface do frame

            origem = os.path.join(BASE1, lote, nome)
            pasta_erros = os.path.join(BASE2, lote, "ERROS")
            os.makedirs(pasta_erros, exist_ok=True)
            destino = os.path.join(pasta_erros, nome)

            try:
                shutil.copy2(origem, destino)
                time.sleep(0.2)

                reduzido = os.path.join(pasta_erros, nome.replace(".pdf", "_REDUZIDO.pdf"))
                comprimir_pdf(destino, reduzido)

                nome_upper = nome.upper()
                if "[BROCHURA]" in nome_upper or "[GRAMPO]" in nome_upper:
                    tipo = 'brochura' if "[BROCHURA]" in nome_upper else 'grampo'
                    montado = os.path.join(pasta_erros, nome.replace(".pdf", "_MONTADO.pdf"))
                    self.label_status.config(text=f"Montando {nome} ({tipo})...")
                    self.update()
                    montar_pdf(reduzido, montado, tipo)
                    resultados.append(f"{lote} - {nome} ✔ OK (montagem {tipo})")
                else:
                    resultados.append(f"{lote} - {nome} ✔ OK")

            except Exception as e:
                resultados.append(f"{lote} - {nome} ❌ ERRO ({e})")

            count += 1
            self.progress["value"] = count
            self.update()

        self.label_status.config(text="Processamento finalizado ✔")
        self.btn_processar.config(state=tk.NORMAL)

        caminho_txt = salvar_relatorio(resultados)
        messagebox.showinfo("Finalizado",
                            f"Processamento concluído.\nRelatório salvo em:\n{caminho_txt}")

        if messagebox.askyesno("Abrir relatório", "Deseja abrir o arquivo de relatório agora?"):
            os.startfile(caminho_txt)

# ═══════════════════ CONSTANTES ═══════════════════

BASE1 = r"\\192.168.200.252\dados_clientes\UICLAP\Shared drives\UICLAP-Docuprint\Producao\Lotes-Novo"
BASE2 = r"N:\____REPOSITORIO GERAL POD\UICLAP\LOTES"

LARGURA_CM = 45
ALTURA_CM = 32
CM_TO_PT = 28.3465
SANGRIA_CM = 0.3

LARGURA_PT = LARGURA_CM * CM_TO_PT
ALTURA_PT = ALTURA_CM * CM_TO_PT
SANGRIA_PT = SANGRIA_CM * CM_TO_PT

TAM_MARCA = 5 * CM_TO_PT / 10   # 5 mm
GAP_MARCA = 2 * CM_TO_PT / 10   # 2 mm

def salvar_relatorio(resultados):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    caminho = os.path.join(desktop, "relatorio_processamento.txt")
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    if os.path.exists(caminho) and os.path.getsize(caminho) > 0:
        with open(caminho, "a", encoding="utf-8") as f:
            f.write("\n")
    with open(caminho, "a", encoding="utf-8") as f:
        f.write("=" * 70 + "\n")
        f.write(f"PROCESSAMENTO REALIZADO EM: {agora}\n")
        f.write("=" * 70 + "\n")
        for linha in resultados:
            f.write(linha + "\n")
    return caminho

# ───── Funções de montagem (copiadas integralmente) ─────
def inserir_imagem_na_pagina(page, img, rect, qualidade=95):
    with io.BytesIO() as output:
        img.save(output, format='PNG', quality=qualidade, optimize=True)
        data = output.getvalue()
    page.insert_image(rect, stream=data, keep_proportion=False)


def desenhar_marca_central_vertical(page, x, y0, y1, tam=TAM_MARCA, gap=GAP_MARCA):
    s = page.new_shape()
    s.draw_line((x, y0 - tam), (x, y0 - gap))
    s.draw_line((x, y1 + tam), (x, y1 + gap))
    s.finish(width=0.5)
    s.commit()

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

        img = get_border_pixmap(page, rect, zoom)
        if rotate == 90:
            img = img.transpose(Image.ROTATE_90)
        elif rotate == 180:
            img = img.transpose(Image.ROTATE_180)
        elif rotate == 270:
            img = img.transpose(Image.ROTATE_270)

        img = espelhar_borda(img, flip_horizontal=flip_h, flip_vertical=flip_v)
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

        inserir_imagem_na_pagina(nova, img, dest_rect)

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

            img = get_border_pixmap(page, rect, zoom)
            if rotate == 90:
                img = img.transpose(Image.ROTATE_90)
            elif rotate == 180:
                img = img.transpose(Image.ROTATE_180)
            elif rotate == 270:
                img = img.transpose(Image.ROTATE_270)
            img = espelhar_borda(img, flip_horizontal=True, flip_vertical=True)
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
            inserir_imagem_na_pagina(nova, img, dest_rect)

def pagina_tem_conteudo(page):
    if page.get_contents():
        return True
    if page.get_text("text").strip():
        return True
    if page.get_images():
        return True
    return False

def obter_area_corte_segura(page, w_ref, h_ref):
    if page.trimbox and page.trimbox.width > 0 and page.trimbox.height > 0:
        return page.trimbox
    rect = page.rect
    if rect.width > 0 and rect.height > 0:
        return rect
    return fitz.Rect(0, 0, w_ref, h_ref)

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

def ajustar_para_media_box(rect, largura, altura):
    x0 = max(0, min(rect.x0, largura - 1))
    y0 = max(0, min(rect.y0, altura - 1))
    x1 = max(x0 + 1, min(rect.x1, largura))
    y1 = max(y0 + 1, min(rect.y1, altura))
    return fitz.Rect(x0, y0, x1, y1)

def montar_pdf(entrada, saida, tipo):
    doc = fitz.open(entrada)
    out = fitz.open()
    if len(doc) > 0:
        w_ref, h_ref = doc[0].rect.width, doc[0].rect.height
    else:
        w_ref, h_ref = 595, 842

    largura_pt = LARGURA_PT
    altura_pt = ALTURA_PT
    sangria_pt = SANGRIA_PT
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

            nova = out.new_page(width=largura_pt, height=altura_pt)
            if sangria_pt > 0:
                y_esq, cabe_esq = calcular_y_com_sangria(altura_pt, h_esq, sangria_pt)
                y_dir, cabe_dir = calcular_y_com_sangria(altura_pt, h_dir, sangria_pt)
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
            destino1 = ajustar_para_media_box(destino1, largura_pt, altura_pt)
            x2 = x1 + w_esq + espaco_pt
            destino2 = fitz.Rect(x2, y_dir, x2 + w_dir, y_dir + h_dir)
            destino2 = ajustar_para_media_box(destino2, largura_pt, altura_pt)

            if page_esq and pagina_tem_conteudo(page_esq):
                try:
                    nova.show_pdf_page(destino1, doc, idx_esq, clip=trim_esq, rotate=180 if rot_esq else 0)
                except Exception:
                    paginas_vazias.append(str(idx_esq+1))
            elif page_esq is None:
                pass
            else:
                paginas_vazias.append(str(idx_esq+1))
            if page_dir and pagina_tem_conteudo(page_dir):
                try:
                    nova.show_pdf_page(destino2, doc, idx_dir, clip=trim_dir, rotate=180 if rot_dir else 0)
                except Exception:
                    paginas_vazias.append(str(idx_dir+1))
            elif page_dir is None:
                pass
            else:
                paginas_vazias.append(str(idx_dir+1))

            if sangria_pt > 0:
                if page_esq and pagina_tem_conteudo(page_esq) and sangria_aplicar_esq:
                    aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1, sangria_pt,
                                                      bordas=['top','bottom','left'],
                                                      rotate=180 if rot_esq else 0)
                if page_dir and pagina_tem_conteudo(page_dir) and sangria_aplicar_dir:
                    aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2, sangria_pt,
                                                      bordas=['top','bottom','right'],
                                                      rotate=180 if rot_dir else 0)
                bleed_total = fitz.Rect(
                    destino1.x0 - sangria_pt,
                    min(destino1.y0, destino2.y0) - sangria_pt,
                    destino2.x1 + sangria_pt,
                    max(destino1.y1, destino2.y1) + sangria_pt
                )
                bleed_total = ajustar_para_media_box(bleed_total, largura_pt, altura_pt)
                if bleed_total.width > 0 and bleed_total.height > 0:
                    nova.set_bleedbox(bleed_total)
            nova.set_trimbox(destino1)
            if espaco_pt == 0:
                desenhar_marcas(nova, destino1, lado='esquerda')
                desenhar_marcas(nova, destino2, lado='direita')
                x_centro = destino1.x1
                desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1)

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

            nova = out.new_page(width=largura_pt, height=altura_pt)
            if sangria_pt > 0:
                y_esq, cabe_esq = calcular_y_com_sangria(altura_pt, h_esq, sangria_pt)
                y_dir, cabe_dir = calcular_y_com_sangria(altura_pt, h_dir, sangria_pt)
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
            destino1 = ajustar_para_media_box(destino1, largura_pt, altura_pt)
            x2 = x1 + w_esq + espaco_pt
            destino2 = fitz.Rect(x2, y_dir, x2 + w_dir, y_dir + h_dir)
            destino2 = ajustar_para_media_box(destino2, largura_pt, altura_pt)

            if page_esq and pagina_tem_conteudo(page_esq):
                try:
                    nova.show_pdf_page(destino1, doc, idx_esq - 1, clip=trim_esq)
                except Exception:
                    paginas_vazias.append(str(idx_esq))
            elif page_esq is None:
                pass
            else:
                paginas_vazias.append(str(idx_esq))
            if page_dir and pagina_tem_conteudo(page_dir):
                try:
                    nova.show_pdf_page(destino2, doc, idx_dir - 1, clip=trim_dir)
                except Exception:
                    paginas_vazias.append(str(idx_dir))
            elif page_dir is None:
                pass
            else:
                paginas_vazias.append(str(idx_dir))

            if sangria_pt > 0:
                if page_esq and pagina_tem_conteudo(page_esq) and sangria_aplicar_esq:
                    aplicar_sangria_espelhada_externa(page_esq, nova, trim_esq, destino1, sangria_pt,
                                                      bordas=['top','bottom','left'])
                if page_dir and pagina_tem_conteudo(page_dir) and sangria_aplicar_dir:
                    aplicar_sangria_espelhada_externa(page_dir, nova, trim_dir, destino2, sangria_pt,
                                                      bordas=['top','bottom','right'])
                bleed_total = fitz.Rect(
                    destino1.x0 - sangria_pt,
                    min(destino1.y0, destino2.y0) - sangria_pt,
                    destino2.x1 + sangria_pt,
                    max(destino1.y1, destino2.y1) + sangria_pt
                )
                bleed_total = ajustar_para_media_box(bleed_total, largura_pt, altura_pt)
                if bleed_total.width > 0 and bleed_total.height > 0:
                    nova.set_bleedbox(bleed_total)
            nova.set_trimbox(destino1)
            if espaco_pt == 0:
                desenhar_marcas(nova, destino1, lado='esquerda')
                desenhar_marcas(nova, destino2, lado='direita')
                x_centro = destino1.x1
                desenhar_marca_central_vertical(nova, x_centro, destino1.y0, destino1.y1)

    out.save(saida, garbage=4, deflate=True, clean=True)
    doc.close()
    out.close()
    if paginas_vazias:
        print(f"Páginas vazias (não incluídas): {', '.join(paginas_vazias)}")

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