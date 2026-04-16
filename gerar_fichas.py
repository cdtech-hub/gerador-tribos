"""
gerar_fichas.py — Gerador de Fichas PDF dos Campistas
Gera 1 ficha por página com todos os dados + assinatura do Gustavo
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import simpleSplit
import io

TRIBO_COR = {
    'Simeão':'#FFC000','Rubem':'#FF0000','Judá':'#FF6600','Levi':'#FF99CC',
    'Benjamim':'#0070C0','Issacar':'#00B050','Gad':'#8B4513','Zabulom':'#000000','EFRAIM':'#BFBFBF',
}
TRIBO_TXT = {
    'Simeão':'#000000','Rubem':'#FFFFFF','Judá':'#FFFFFF','Levi':'#000000',
    'Benjamim':'#FFFFFF','Issacar':'#FFFFFF','Gad':'#FFFFFF','Zabulom':'#FFFFFF','EFRAIM':'#000000',
}

AZUL     = colors.HexColor('#1F4E79')
CINZALBL = colors.HexColor('#666666')
PRETO    = colors.HexColor('#111111')
BORDA    = colors.HexColor('#BBBBBB')
W, H     = A4
LM=15*mm; RM=15*mm; TM=12*mm; BM=12*mm
CW = W - LM - RM


def _celula(c, x, y, w, h, label, valor):
    c.setStrokeColor(BORDA); c.setLineWidth(0.3)
    c.rect(x, y-h, w, h, fill=0, stroke=1)
    c.setFillColor(CINZALBL); c.setFont('Helvetica', 6)
    c.drawString(x+2*mm, y-3.5*mm, label.upper())
    c.setFillColor(PRETO); c.setFont('Helvetica', 8)
    val = str(valor)
    while val and c.stringWidth(val, 'Helvetica', 8) > w-4*mm:
        val = val[:-1]
    c.drawString(x+2*mm, y-h+2*mm, val)


def _secao(c, y, titulo):
    c.setFillColor(AZUL)
    c.rect(LM, y-5.5*mm, CW, 6*mm, fill=1, stroke=0)
    c.setFillColor(colors.white); c.setFont('Helvetica-Bold', 8)
    c.drawString(LM+3*mm, y-3.8*mm, titulo)
    return y - 6*mm


def _assinatura_gustavo(c, x, y):
    c.setFont('Times-Italic', 28)
    c.setFillColor(colors.HexColor('#1a1a1a'))
    c.drawString(x, y, 'Gustavo Pereira')
    sw = c.stringWidth('Gustavo Pereira', 'Times-Italic', 28)
    c.setStrokeColor(colors.HexColor('#1a1a1a')); c.setLineWidth(1.0)
    p = c.beginPath()
    p.moveTo(x, y-4*mm)
    p.curveTo(x+sw*0.3, y-6*mm, x+sw*0.7, y-3*mm, x+sw+5*mm, y-5*mm)
    c.drawPath(p, stroke=1, fill=0)


def _draw_ficha(c, camp):
    y = H - TM

    # Cabeçalho
    c.setFillColor(AZUL)
    c.rect(LM, y-18*mm, CW*0.72, 18*mm, fill=1, stroke=0)
    c.setFillColor(colors.white); c.setFont('Helvetica-Bold', 13)
    c.drawCentredString(LM+CW*0.36, y-8*mm, 'ACAMPAMENTO GUARDIÕES DO AMOR MAIOR')
    c.setFont('Helvetica', 8.5)
    c.drawCentredString(LM+CW*0.36, y-13*mm, 'TERMO DE AUTORIZAÇÃO E RESPONSABILIDADE')

    tribo = camp.get('tribo', '')
    tcor  = colors.HexColor(TRIBO_COR.get(tribo, '#1F4E79'))
    ttxt  = colors.HexColor(TRIBO_TXT.get(tribo, '#FFFFFF'))
    fx=LM+CW*0.73; fw=CW*0.27
    c.setFillColor(tcor)
    c.rect(fx, y-18*mm, fw, 18*mm, fill=1, stroke=0)
    c.setFillColor(ttxt); c.setFont('Helvetica-Bold', 16)
    c.drawCentredString(fx+fw/2, y-8*mm, f'FICHA Nº {camp.get("fc",0):03d}')
    c.setFont('Helvetica-Bold', 9)
    c.drawCentredString(fx+fw/2, y-14*mm, tribo.upper())
    y -= 18*mm

    # 1. Dados Pessoais
    y = _secao(c, y, '1. DADOS PESSOAIS DO CAMPISTA')
    c.setFillColor(CINZALBL); c.setFont('Helvetica', 6)
    c.drawString(LM+2*mm, y-2*mm, 'NOME COMPLETO')
    y -= 3*mm
    c.setFillColor(PRETO); c.setFont('Helvetica-Bold', 10)
    c.drawString(LM+2*mm, y-4*mm, camp.get('nome', ''))
    c.setStrokeColor(BORDA); c.setLineWidth(0.3)
    c.line(LM, y-5.5*mm, LM+CW, y-5.5*mm)
    y -= 6.5*mm

    h1 = 9*mm
    _celula(c,LM,          y,CW*0.28,h1,f"{camp.get('doc_tipo','CPF')} / Documento",camp.get('doc_num',''))
    _celula(c,LM+CW*0.28,  y,CW*0.22,h1,'Data de Nascimento',camp.get('nasc',''))
    _celula(c,LM+CW*0.50,  y,CW*0.22,h1,'Sexo',camp.get('sexo',''))
    _celula(c,LM+CW*0.72,  y,CW*0.28,h1,'Estado Civil',camp.get('est_civil',''))
    y -= h1
    _celula(c,LM,          y,CW*0.35,h1,'Celular / WhatsApp',camp.get('celular',''))
    _celula(c,LM+CW*0.35,  y,CW*0.65,h1,'E-mail',camp.get('email',''))
    y -= h1
    _celula(c,LM,          y,CW*0.35,h1,'Cidade',camp.get('cidade',''))
    _celula(c,LM+CW*0.35,  y,CW*0.65,h1,'Paróquia / Comunidade',camp.get('paroquia',''))
    y -= h1
    _celula(c,LM,          y,CW*0.35,h1,'Profissão',camp.get('profissao',''))
    _celula(c,LM+CW*0.35,  y,CW*0.20,h1,'Camiseta',camp.get('camiseta',''))
    _celula(c,LM+CW*0.55,  y,CW*0.22,h1,'Peso (kg)',camp.get('peso',''))
    _celula(c,LM+CW*0.77,  y,CW*0.23,h1,'Altura (m)',camp.get('altura',''))
    y -= h1 + 1*mm

    # 2. Contatos
    y = _secao(c, y, '2. CONTATOS DE EMERGÊNCIA')
    _celula(c,LM,         y,CW*0.55,h1,'Contato 1 — Nome',camp.get('c1nome',''))
    _celula(c,LM+CW*0.55, y,CW*0.45,h1,'Telefone',camp.get('c1tel',''))
    y -= h1
    _celula(c,LM,         y,CW*0.55,h1,'Contato 2 — Nome',camp.get('c2nome',''))
    _celula(c,LM+CW*0.55, y,CW*0.45,h1,'Telefone',camp.get('c2tel',''))
    y -= h1 + 1*mm

    # 3. Saúde
    y = _secao(c, y, '3. INFORMAÇÕES DE SAÚDE')
    c.setFillColor(CINZALBL); c.setFont('Helvetica', 6)
    c.drawString(LM+2*mm, y-2.5*mm, 'ALERGIAS A ALIMENTOS, MEDICAMENTOS OU OUTROS')
    y -= 3.5*mm
    c.setFillColor(PRETO); c.setFont('Helvetica', 8.5)
    ale = camp.get('alergias', '') or 'Não'
    c.drawString(LM+2*mm, y-4*mm, ale[:120])
    c.setStrokeColor(BORDA); c.setLineWidth(0.3)
    c.line(LM, y-5.5*mm, LM+CW, y-5.5*mm)
    y -= 7*mm

    # 4. Autorização
    y = _secao(c, y, '4. AUTORIZAÇÃO DE USO DE IMAGEM E VOZ')
    texto = ('Autorizo, de forma gratuita, irrevogável e por prazo indeterminado, o uso da imagem, voz e nome nas fotografias, vídeos e materiais '
             'produzidos durante o evento, para fins exclusivamente institucionais, educativos e religiosos, incluindo redes sociais do grupo GAM, site oficial, '
             'apresentações internas e material informativo. Fica vedado o uso comercial ou lucrativo da imagem. Os dados pessoais fornecidos serão utilizados '
             'exclusivamente para organização, segurança e comunicação do evento, conforme a LGPD (Lei 13.709/2018).')
    c.setFillColor(PRETO); c.setFont('Helvetica', 7.5)
    for linha in simpleSplit(texto, 'Helvetica', 7.5, CW-4*mm):
        c.drawString(LM+2*mm, y-4*mm, linha)
        y -= 4.5*mm
    y -= 1*mm

    # 5. Assinatura
    y = _secao(c, y, '5. ASSINATURA')
    c.setFillColor(PRETO); c.setFont('Helvetica', 8)
    c.drawString(LM+10*mm, y-8*mm, 'Local / Data: GOIÂNIA, _____/_____/2026')
    y -= 20*mm
    c.setStrokeColor(CINZALBL); c.setLineWidth(0.5)
    c.line(LM+5*mm, y, LM+CW*0.42, y)
    c.setFont('Helvetica', 7.5); c.setFillColor(CINZALBL)
    c.drawCentredString(LM+CW*0.22, y-4*mm, 'Assinatura do Participante')
    _assinatura_gustavo(c, LM+CW*0.50, y+2*mm)
    c.setFont('Helvetica', 6.5); c.setFillColor(CINZALBL)
    c.drawCentredString(LM+CW*0.76, y-4*mm, 'Responsável pela Organização — GUSTAVO PEREIRA DE ALMEIDA')
    c.setStrokeColor(BORDA); c.setLineWidth(0.8)
    c.rect(LM, BM, CW, H-TM-BM, fill=0, stroke=1)


def gerar_fichas_pdf(campistas):
    """
    Recebe lista de dicts com campos do campista.
    Retorna bytes do PDF gerado.
    """
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    c.setTitle('Fichas Campistas — VIII Acampamento Guardiões do Amor Maior')
    for camp in sorted(campistas, key=lambda x: x.get('fc', 0)):
        _draw_ficha(c, camp)
        c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()
