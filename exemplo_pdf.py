from reportlab.lib import colors  # Trabalhar com cores
from reportlab.lib.pagesizes import A4  # (210*mm,297*mm)
from reportlab.lib.units import cm, inch, mm  # Padrão = Points
from reportlab.pdfbase import pdfmetrics  # Usamos p/ registrar a fonte
from reportlab.pdfbase.ttfonts import TTFont  # Criar a fonte em si
from reportlab.pdfgen import canvas  # Onde 'desenhamos' no PDF


def desenhar_regua(pdf_canvas, pagina_tamanho):
    """Desenha uma régua para facilitar entender a posição de cada elemento

    Args:
        pdf_canvas: O Canvas que você está trabalhando
        pagina_tamanho: O tamanho da página atual
    """
    pdf_canvas.setFontSize(8)
    pagina_tamanho_y = int(pagina_tamanho[1] / cm)
    pagina_tamanho_x = int(pagina_tamanho[0] / cm)

    for y in range(pagina_tamanho_y + 1):
        pdf_canvas.drawString(0 * cm, y * cm, f"y{y}")

    for x in range(pagina_tamanho_x + 1):
        pdf_canvas.drawString(x * cm, 0 * cm, f"x{x}")


nome_arquivo = "Relatório Teste.pdf"
titulo_arquivo = "Relatório Teste"
tamanho_pagina_arquivo = A4

pdf_canvas = canvas.Canvas(nome_arquivo, pagesize=tamanho_pagina_arquivo)
pdf_canvas.setTitle(titulo_arquivo)

logo_bylearn = "img\\rodape_pagina.png"
pdf_canvas.drawImage(logo_bylearn, 0, 0, 210 * mm, 297 * mm, mask="auto")

desenhar_regua(pdf_canvas, tamanho_pagina_arquivo)

pdf_canvas.save()
