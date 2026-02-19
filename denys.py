import streamlit as st
import datetime
import pandas as pd
import zipfile
from io import BytesIO
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors

# ==========================================
# 1. CONFIGURAÇÕES E ESTILOS
# ==========================================
st.set_page_config(page_title="Gestão de Comissões - Pense e Conecte", layout="wide")

def aplicar_estilos():
    st.markdown("""
    <style>
    .section { background-color: #ffffff; padding: 20px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #e6f2ff; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .footer { margin-top: 40px; padding: 20px; background-color: #e6f2ff; text-align: center; border-radius: 8px; font-size: 14px; }
    .stButton>button { background-color: #4da6ff; color: white; border-radius: 8px; font-weight: bold; width: 100%; }
    h1 { color: #4da6ff; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. FUNÇÕES DE LIMPEZA
# ==========================================
def limpar_moeda(valor):
    try:
        if pd.isna(valor) or valor == "": return 0.0
        if isinstance(valor, (int, float)): return float(valor)
        s = str(valor).replace("R$", "").strip()
        if "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
        elif "," in s: s = s.replace(",", ".")
        return float(s)
    except: return 0.0

def normalizar_percentual(valor):
    try:
        if pd.isna(valor) or valor == "": return 0.0
        if isinstance(valor, (int, float)): v = float(valor)
        else:
            v = str(valor).replace("%", "").replace(",", ".").strip()
            v = float(v)
        if v < 1 and v > 0: v *= 100
        return round(v, 2)
    except: return 0.0

# ==========================================
# 3. GERAÇÃO DE PDF (VISÃO CONSULTOR)
# ==========================================
def gerar_pdf_universal(vendedor, dados, periodo, tipo_modulo):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    largura, altura = landscape(letter)
    
    c.setFont("Helvetica-Bold", 16)
    titulo = f"RELATÓRIO DE COMISSÃO - {tipo_modulo.upper()}"
    c.drawCentredString(largura/2, altura - 50, titulo)
    c.setFont("Helvetica", 10)
    c.drawCentredString(largura/2, altura - 65, f"Consultor: {vendedor}  |  Período: {periodo}")
    c.setStrokeColor(colors.HexColor("#4da6ff"))
    c.line(40, altura - 80, largura - 40, altura - 80)

    y = altura - 110
    c.setFillColor(colors.HexColor("#e6f2ff"))
    c.rect(40, y, largura - 80, 20, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 10)
    
    c.drawString(45, y + 6, "PLACA")
    c.drawString(130, y + 6, "ASSOCIADO")
    
    if tipo_modulo == "Adesões":
        c.drawRightString(520, y + 6, "VALOR BRUTO")
        c.drawRightString(630, y + 6, "DESC. RAST.")
        c.drawRightString(740, y + 6, "SUBTOTAL")
    else:
        c.drawRightString(480, y + 6, "VALOR BASE")
        c.drawRightString(600, y + 6, "% APLICADA")
        c.drawRightString(740, y + 6, "COMISSÃO")
    
    y -= 25
    total_final = 0
    c.setFont("Helvetica", 9)
    
    for _, row in dados.iterrows():
        if y < 80:
            c.showPage()
            y = altura - 80
            c.setFont("Helvetica", 9)
        
        c.drawString(45, y, str(row['PLACA']))
        c.drawString(130, y, str(row['ASSOCIADO'])[:45])
        
        if tipo_modulo == "Adesões":
            c.drawRightString(520, y, f"R$ {row['VALOR_BASE']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c.drawRightString(630, y, f"R$ {row['DESC']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c.drawRightString(740, y, f"R$ {row['LIQUIDO']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
            c.drawRightString(480, y, f"R$ {row['VALOR_BASE']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c.drawRightString(600, y, f"{row['PERC_FINAL']}%")
            c.drawRightString(740, y, f"R$ {row['VALOR_FINAL']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        total_final += row['VALOR_FINAL']
        y -= 18

    y -= 40
    c.setStrokeColor(colors.black)
    c.line(450, y + 25, largura - 40, y + 25)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(450, y, f"ASSOCIADOS ATIVOS:")
    c.drawRightString(740, y, f"{len(dados)}")
    y -= 20
    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(colors.HexColor("#1e90ff"))
    c.drawString(450, y, f"TOTAL A RECEBER:")
    c.drawRightString(740, y, f"R$ {total_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 4. GERAÇÃO DE PDF GERENCIAL (DETALHADO)
# ==========================================
def gerar_pdf_gerencial_adesao(dados_resumo, periodo, perc_imp):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    largura, altura = landscape(letter)
    
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura/2, altura - 50, "RELATÓRIO GERENCIAL - DETALHAMENTO DE ADESÕES")
    c.setFont("Helvetica", 10)
    c.drawCentredString(largura/2, altura - 65, f"Período: {periodo} | Desconto NF: {perc_imp}%")
    c.line(40, altura - 80, largura - 40, altura - 80)
    
    y = altura - 110
    c.setFillColor(colors.HexColor("#333333"))
    c.rect(40, y, largura - 80, 20, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(45, y + 6, "CONSULTOR")
    c.drawRightString(300, y + 6, "BRUTO TOTAL")
    c.drawRightString(410, y + 6, "RASTREADORES")
    c.drawRightString(520, y + 6, "BASE IMPOSTO")
    c.drawRightString(630, y + 6, f"IMPOSTO ({perc_imp}%)")
    c.drawRightString(740, y + 6, "VALOR PAGO")
    
    y -= 25
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9)
    
    t_bruto = t_rast = t_base = t_imp = t_pago = 0
    
    for item in dados_resumo:
        bruto = item['bruto_original']
        rast = item['rastreadores_total']
        base = item['subtotal_antes_imposto']
        imposto = base * (perc_imp/100)
        pago = item['total']
        
        c.drawString(45, y, item['nome'])
        c.drawRightString(300, y, f"R$ {bruto:,.2f}")
        c.drawRightString(410, y, f"R$ {rast:,.2f}")
        c.drawRightString(520, y, f"R$ {base:,.2f}")
        c.drawRightString(630, y, f"R$ {imposto:,.2f}")
        c.drawRightString(740, y, f"R$ {pago:,.2f}")
        
        t_bruto += bruto; t_rast += rast; t_base += base; t_imp += imposto; t_pago += pago
        y -= 18
        
    c.line(40, y, largura - 40, y)
    y -= 20
    c.setFont("Helvetica-Bold", 10)
    c.drawString(45, y, "TOTAIS CONSOLIDADOS")
    c.drawRightString(300, y, f"R$ {t_bruto:,.2f}")
    c.drawRightString(410, y, f"R$ {t_rast:,.2f}")
    c.drawRightString(520, y, f"R$ {t_base:,.2f}")
    c.drawRightString(630, y, f"R$ {t_imp:,.2f}")
    c.drawRightString(740, y, f"R$ {t_pago:,.2f}")
    
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 4.1 GERAÇÃO DE PDF EXCEÇÕES (AGRUPADO)
# ==========================================
def gerar_pdf_excecoes(dados, periodo):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    largura, altura = landscape(letter)
    
    def desenhar_cabecalho(y_pos):
        c.setFillColor(colors.HexColor("#e6f2ff"))
        c.rect(40, y_pos, largura - 80, 20, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(45, y_pos + 6, "PLACA")
        c.drawString(130, y_pos + 6, "ASSOCIADO")
        c.drawRightString(480, y_pos + 6, "VALOR BASE")
        c.drawRightString(600, y_pos + 6, "% APLICADA")
        c.drawRightString(740, y_pos + 6, "COMISSÃO")

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura/2, altura - 50, "RELATÓRIO DE EXCEÇÕES - RECORRÊNCIA")
    c.setFont("Helvetica", 10)
    c.drawCentredString(largura/2, altura - 65, f"Período: {periodo}")
    c.setStrokeColor(colors.HexColor("#4da6ff"))
    c.line(40, altura - 80, largura - 40, altura - 80)

    y = altura - 110
    desenhar_cabecalho(y)
    y -= 30
    
    total_geral = 0
    consultores = sorted(dados["CONSULTOR"].unique())
    
    for cons in consultores:
        df_cons = dados[dados["CONSULTOR"] == cons]
        subtotal = df_cons["VALOR_FINAL"].sum()
        
        if y < 60:
            c.showPage()
            y = altura - 80
            desenhar_cabecalho(y)
            y -= 30
            
        c.setFillColor(colors.HexColor("#f0f0f0"))
        c.rect(40, y, largura - 80, 15, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(45, y + 4, f"CONSULTOR: {cons}")
        y -= 20
        
        c.setFont("Helvetica", 9)
        for _, row in df_cons.iterrows():
            if y < 40:
                c.showPage()
                y = altura - 80
                desenhar_cabecalho(y)
                y -= 30
                
            c.drawString(45, y, str(row['PLACA']))
            c.drawString(130, y, str(row['ASSOCIADO'])[:45])
            c.drawRightString(480, y, f"R$ {row['VALOR_BASE']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c.drawRightString(600, y, f"{row['PERC_FINAL']}%")
            c.drawRightString(740, y, f"R$ {row['VALOR_FINAL']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            y -= 15
            
        c.setFont("Helvetica-Bold", 9)
        c.drawRightString(650, y, "SUBTOTAL:")
        c.drawRightString(740, y, f"R$ {subtotal:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.setStrokeColor(colors.grey)
        c.line(40, y-2, largura-40, y-2)
        y -= 25
        total_geral += subtotal

    if y < 40:
        c.showPage()
        y = altura - 80

    y -= 10
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.HexColor("#1e90ff"))
    c.drawString(450, y, "TOTAL GERAL EXCEÇÕES:")
    c.drawRightString(740, y, f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 4.2 GERAÇÃO DE PDF GERENTE (RECORRÊNCIA)
# ==========================================
def gerar_pdf_gerente_recorrencia(dados, periodo):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    largura, altura = landscape(letter)
    
    def desenhar_cabecalho(y_pos):
        c.setFillColor(colors.HexColor("#e6f2ff"))
        c.rect(40, y_pos, largura - 80, 20, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(45, y_pos + 6, "PLACA")
        c.drawString(130, y_pos + 6, "ASSOCIADO")
        c.drawRightString(480, y_pos + 6, "VALOR BASE")
        c.drawRightString(600, y_pos + 6, "% GERENTE")
        c.drawRightString(740, y_pos + 6, "COMISSÃO")

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura/2, altura - 50, "RELATÓRIO DE COMISSÃO - GERENTE (RECORRÊNCIA)")
    c.setFont("Helvetica", 10)
    c.drawCentredString(largura/2, altura - 65, f"Período: {periodo}")
    c.setStrokeColor(colors.HexColor("#4da6ff"))
    c.line(40, altura - 80, largura - 40, altura - 80)

    y = altura - 110
    desenhar_cabecalho(y)
    y -= 30
    
    total_geral = 0
    consultores = sorted(dados["CONSULTOR"].unique())
    
    for cons in consultores:
        df_cons = dados[dados["CONSULTOR"] == cons]
        subtotal = df_cons["VALOR_FINAL"].sum()
        
        if y < 60:
            c.showPage()
            y = altura - 80
            desenhar_cabecalho(y)
            y -= 30
            
        c.setFillColor(colors.HexColor("#f0f0f0"))
        c.rect(40, y, largura - 80, 15, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(45, y + 4, f"ORIGEM CONSULTOR: {cons}")
        y -= 20
        
        c.setFont("Helvetica", 9)
        for _, row in df_cons.iterrows():
            if y < 40:
                c.showPage()
                y = altura - 80
                desenhar_cabecalho(y)
                y -= 30
                
            c.drawString(45, y, str(row['PLACA']))
            c.drawString(130, y, str(row['ASSOCIADO'])[:45])
            c.drawRightString(480, y, f"R$ {row['VALOR_BASE']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c.drawRightString(600, y, f"{row['PERC_FINAL']}%")
            c.drawRightString(740, y, f"R$ {row['VALOR_FINAL']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            y -= 15
            
        c.setFont("Helvetica-Bold", 9)
        c.drawRightString(650, y, "SUBTOTAL:")
        c.drawRightString(740, y, f"R$ {subtotal:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.setStrokeColor(colors.grey)
        c.line(40, y-2, largura-40, y-2)
        y -= 25
        total_geral += subtotal

    if y < 40:
        c.showPage()
        y = altura - 80

    y -= 10
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.HexColor("#1e90ff"))
    c.drawString(450, y, "TOTAL GERAL GERENTE:")
    c.drawRightString(740, y, f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 4.3 GERAÇÃO DE PDF GERENCIAL (RECORRÊNCIA)
# ==========================================
def gerar_pdf_gerencial_recorrencia(dados_resumo, periodo, perc_imp):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    largura, altura = landscape(letter)
    
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura/2, altura - 50, "RELATÓRIO GERENCIAL - DETALHAMENTO DE RECORRÊNCIA")
    c.setFont("Helvetica", 10)
    c.drawCentredString(largura/2, altura - 65, f"Período: {periodo} | Desconto NF: {perc_imp}%")
    c.line(40, altura - 80, largura - 40, altura - 80)
    
    y = altura - 110
    c.setFillColor(colors.HexColor("#333333"))
    c.rect(40, y, largura - 80, 20, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(45, y + 6, "CONSULTOR")
    c.drawRightString(250, y + 6, "PLACAS ATIVAS")
    c.drawRightString(380, y + 6, "COM. CONSULTOR (BRUTO)")
    c.drawRightString(510, y + 6, f"IMPOSTO ({perc_imp}%)")
    c.drawRightString(630, y + 6, "COM. CONSULTOR (LÍQUIDO)")
    c.drawRightString(740, y + 6, "COMISSÃO GERENTE")
    
    y -= 25
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9)
    
    t_placas = t_com_bruta = t_imposto = t_com_liquida = t_com_gerente = 0
    
    for item in dados_resumo:
        placas = item['qtd_placas']
        com_bruta = item['comissao_bruta_consultor']
        imposto = item['imposto_valor']
        com_liquida = item['comissao_liquida_consultor']
        com_gerente = item['comissao_gerente']
        
        c.drawString(45, y, item['nome'])
        c.drawRightString(250, y, f"{placas}")
        c.drawRightString(380, y, f"R$ {com_bruta:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.drawRightString(510, y, f"R$ {imposto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.drawRightString(630, y, f"R$ {com_liquida:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.drawRightString(740, y, f"R$ {com_gerente:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        t_placas += placas; t_com_bruta += com_bruta; t_imposto += imposto; t_com_liquida += com_liquida; t_com_gerente += com_gerente
        y -= 18
        
    c.line(40, y, largura - 40, y)
    y -= 20
    c.setFont("Helvetica-Bold", 10)
    c.drawString(45, y, "TOTAIS CONSOLIDADOS")
    c.drawRightString(250, y, f"{t_placas}")
    c.drawRightString(380, y, f"R$ {t_com_bruta:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawRightString(510, y, f"R$ {t_imposto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawRightString(630, y, f"R$ {t_com_liquida:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawRightString(740, y, f"R$ {t_com_gerente:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 4.4 GERAÇÃO DE PDF EVOLUÇÃO GERENCIAL
# ==========================================
def gerar_pdf_evolucao_gerencial(dados_resumo, dados_gerais, periodo):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    largura, altura = landscape(letter)

    # --- Header ---
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura/2, altura - 50, "RELATÓRIO GERENCIAL DE EVOLUÇÃO DA CARTEIRA")
    c.setFont("Helvetica", 10)
    c.drawCentredString(largura/2, altura - 65, f"Período: {periodo}")
    c.line(40, altura - 80, largura - 40, altura - 80)
    y = altura - 100

    # --- Summary Section ---
    c.setFont("Helvetica-Bold", 12)
    c.drawString(45, y, "Resumo da Evolução:")
    y -= 20
    
    c.setFont("Helvetica", 10)
    c.drawString(55, y, f"✅ Associados Mantidos: {dados_gerais['mantidas']['qtd']} (R$ {dados_gerais['mantidas']['valor']:,.2f})".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawString(350, y, f"🚀 Associados Novos: {dados_gerais['novas']['qtd']} (R$ {dados_gerais['novas']['valor']:,.2f})".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawString(600, y, f"🔻 Associados Inativos: {dados_gerais['canceladas']['qtd']} (R$ {dados_gerais['canceladas']['valor']:,.2f})".replace(",", "X").replace(".", ",").replace("X", "."))
    y -= 30

    # --- Detailed Table Header ---
    c.setFillColor(colors.HexColor("#333333"))
    c.rect(40, y, largura - 80, 20, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(45, y + 6, "CONSULTOR")
    c.drawRightString(220, y + 6, "ATIVOS")
    c.drawRightString(350, y + 6, "BASE RECORRENCIA")
    c.drawRightString(480, y + 6, "COMISSÃO CONSULTOR")
    c.drawRightString(610, y + 6, "COMISSÃO GERENTE")
    c.drawRightString(740, y + 6, "IMPOSTOS")
    y -= 25

    # --- Table Rows ---
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9)
    
    t_ativos = t_fat = t_com_cons = t_imp = t_com_ger = 0

    for item in dados_resumo:
        if y < 80:
            c.showPage()
            y = altura - 80
            c.setFont("Helvetica", 9)

        # Formatar nome: Primeiro + Último
        nome_completo = item['consultor'].strip()
        partes_nome = nome_completo.split()
        nome_formatado = f"{partes_nome[0]} {partes_nome[-1]}" if len(partes_nome) > 1 else nome_completo

        c.drawString(45, y, nome_formatado)
        c.drawRightString(220, y, str(item['qtd_ativos']))
        c.drawRightString(350, y, f"R$ {item['faturamento_base']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.drawRightString(480, y, f"R$ {item['total_recorrencia_consultor']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.drawRightString(610, y, f"R$ {item['total_recorrencia_gerente']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.drawRightString(740, y, f"R$ {item['total_impostos']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        t_ativos += item['qtd_ativos']
        t_fat += item['faturamento_base']
        t_com_cons += item['total_recorrencia_consultor']
        t_imp += item['total_impostos']
        t_com_ger += item['total_recorrencia_gerente']
        y -= 18

    # --- Totals Footer ---
    c.line(40, y, largura - 40, y)
    y -= 20
    c.setFont("Helvetica-Bold", 10)
    c.drawString(45, y, "TOTAIS")
    c.drawRightString(220, y, str(t_ativos))
    c.drawRightString(350, y, f"R$ {t_fat:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawRightString(480, y, f"R$ {t_com_cons:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawRightString(610, y, f"R$ {t_com_ger:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawRightString(740, y, f"R$ {t_imp:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    # Grand total revenue in footer as requested
    y -= 40
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.HexColor("#1e90ff"))
    c.drawString(400, y, "FATURAMENTO TOTAL (BASE MENSALIDADE):")
    c.drawRightString(740, y, f"R$ {t_fat:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 4.5 GERAÇÃO DE PDF INATIVOS (EVOLUÇÃO)
# ==========================================
def gerar_pdf_inativos(dados, periodo):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))
    largura, altura = landscape(letter)

    def desenhar_cabecalho(y_pos):
        c.setFillColor(colors.HexColor("#e6f2ff"))
        c.rect(40, y_pos, largura - 80, 20, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(45, y_pos + 6, "PLACA")
        c.drawString(130, y_pos + 6, "ASSOCIADO")
        c.drawRightString(480, y_pos + 6, "VALOR MENSAL")
        c.drawRightString(600, y_pos + 6, "VALOR RECORRÊNCIA")
        c.drawRightString(740, y_pos + 6, "PERCENTUAL")
    
    # --- PDF Header ---
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura/2, altura - 50, "RELATÓRIO DE ASSOCIADOS INATIVADOS")
    c.setFont("Helvetica", 10)
    c.drawCentredString(largura/2, altura - 65, f"Período: {periodo}")
    c.line(40, altura - 80, largura - 40, altura - 80)

    y = altura - 110
    
    total_geral_mensal = 0
    total_geral_recorrencia = 0
    consultores = sorted(dados["CONSULTOR"].unique())

    for cons in consultores:
        df_cons = dados[dados["CONSULTOR"] == cons]
        subtotal_mensal = df_cons["VALOR MENSAL"].sum()
        subtotal_recorrencia = df_cons["VALOR RECORRENCIA"].sum()

        if y < 100:
            c.showPage()
            y = altura - 80

        # Consultant Header
        c.setFillColor(colors.HexColor("#f0f0f0"))
        c.rect(40, y, largura - 80, 18, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(45, y + 5, f"CONSULTOR: {cons}")
        y -= 25
        
        desenhar_cabecalho(y)
        y -= 25

        c.setFont("Helvetica", 9)
        for _, row in df_cons.iterrows():
            if y < 40:
                c.showPage()
                y = altura - 80
                desenhar_cabecalho(y)
                y -= 25
            
            c.drawString(45, y, str(row['PLACA']))
            c.drawString(130, y, str(row['ASSOCIADO'])[:45])
            c.drawRightString(480, y, f"R$ {row['VALOR MENSAL']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c.drawRightString(600, y, f"R$ {row['VALOR RECORRENCIA']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            c.drawRightString(740, y, f"{row['PORCENTAGEM RECORRENCIA']}%")
            y -= 15
        
        # Subtotal
        c.setStrokeColor(colors.grey)
        c.line(400, y, largura-40, y)
        y -= 15
        c.setFont("Helvetica-Bold", 9)
        c.drawRightString(480, y, f"R$ {subtotal_mensal:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.drawRightString(600, y, f"R$ {subtotal_recorrencia:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        c.drawString(300, y, "SUBTOTAL CONSULTOR:")
        y -= 25
        
        total_geral_mensal += subtotal_mensal
        total_geral_recorrencia += subtotal_recorrencia

    # Grand Total
    if y < 60:
        c.showPage()
        y = altura - 80

    y -= 10
    c.setStrokeColor(colors.black)
    c.line(40, y, largura - 40, y)
    y -= 20
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.HexColor("#ff4d4d"))
    c.drawString(250, y, "TOTAL GERAL INATIVADO:")
    c.drawRightString(480, y, f"R$ {total_geral_mensal:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c.drawRightString(600, y, f"R$ {total_geral_recorrencia:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 5. PROCESSAMENTOS
# ==========================================
def processar_adesao(arquivo, d_ini, d_fim, aplicar_imp, perc_imp):
    periodo_txt = f"{d_ini.strftime('%d/%m/%Y')} a {d_fim.strftime('%d/%m/%Y')}"
    try:
        df_raw = pd.read_excel(arquivo, header=None)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return False
    
    # Procura por CONSULTOR ou VOLUNTÁRIO
    linha_cab = next((i for i in range(len(df_raw)) if any(c in df_raw.iloc[i].astype(str).str.upper().values for c in ["CONSULTOR", "VOLUNTÁRIO"])), None)
    
    if linha_cab is None:
        st.error("Erro: Coluna 'CONSULTOR' ou 'VOLUNTÁRIO' não encontrada.")
        return False
    df = pd.read_excel(arquivo, header=linha_cab)
    df.columns = df.columns.str.strip().str.upper()
    if "VOLUNTÁRIO" in df.columns: df.rename(columns={"VOLUNTÁRIO": "CONSULTOR"}, inplace=True)
    df_adesao = df[df["TIPO COMISSÃO"].str.contains("ADES|ADHES|VENDA|ATIV", na=False, case=False)].copy()
    
    if df_adesao.empty:
        st.warning("Nenhuma comissão de adesão foi encontrada no arquivo.")
        return False

    st.session_state.pdfs = {}
    st.session_state.resumo = []

    for vend in df_adesao["CONSULTOR"].dropna().unique():
        dv = df_adesao[df_adesao["CONSULTOR"] == vend].copy()
        dv["VALOR_BASE"] = dv["VALOR BASE ADESÃO"].apply(limpar_moeda)
        dv["DESC"] = dv["DESCONTO RASTREADOR"].apply(limpar_moeda)
        dv["LIQUIDO"] = dv["VALOR_BASE"] - dv["DESC"]
        dv["VALOR_FINAL"] = (dv["LIQUIDO"] * (1 - perc_imp/100 if aplicar_imp else 1.0)).round(2)
        
        st.session_state.pdfs[vend] = gerar_pdf_universal(vend, dv, periodo_txt, "Adesões")
        st.session_state.resumo.append({
            "nome": vend, 
            "total": dv["VALOR_FINAL"].sum(), 
            "bruto_original": dv["VALOR_BASE"].sum(),
            "rastreadores_total": dv["DESC"].sum(),
            "subtotal_antes_imposto": dv["LIQUIDO"].sum(),
            "qtd": len(dv)
        })

    if aplicar_imp:
        st.session_state.pdfs["_GERENCIAL_TRIBUTOS"] = gerar_pdf_gerencial_adesao(st.session_state.resumo, periodo_txt, perc_imp)
    
    st.session_state.pronto = True
    return True

def processar_recorrencia(arquivo, mes, ano, aplicar_imp, perc_imp):
    periodo_txt = f"{mes}/{ano}"
    df, err = _ler_e_preparar_planilha_recorrencia(arquivo)
    if err:
        st.error(err)
        return False

    # 5. Lógica de Negócio (Ajustes de Percentual)
    df["% CONSULTOR"] = (df["PORCENTAGEM RECORRENCIA"] / 2)
    df["VALOR COMISSAO"] = (df["BASE RECORRENCIA"] * df["% CONSULTOR"] / 100 * (1 - perc_imp/100 if aplicar_imp else 1.0)).round(2)
    
    df["% GERENTE"] = (df["PORCENTAGEM RECORRENCIA"] / 2)
    df["VALOR COMISSAO GERENTE"] = (df["BASE RECORRENCIA"] * df["% GERENTE"] / 100).round(2)

    # 6. Gerar Relatórios
    st.session_state.pdfs = {}
    st.session_state.resumo = []
    resumo_gerencial = []

    # Padrão
    df_padrao = df[df["% CONSULTOR"] > 0].copy()
    for vend in df_padrao["CONSULTOR"].unique():
        dv = df_padrao[df_padrao["CONSULTOR"] == vend].copy()
        
        comissao_bruta_consultor = (dv["BASE RECORRENCIA"] * dv["% CONSULTOR"] / 100).sum()
        imposto_valor = comissao_bruta_consultor * (perc_imp / 100 if aplicar_imp else 0.0)
        resumo_gerencial.append({
            "nome": vend,
            "qtd_placas": len(dv),
            "comissao_bruta_consultor": comissao_bruta_consultor,
            "imposto_valor": imposto_valor,
            "comissao_liquida_consultor": dv["VALOR COMISSAO"].sum(),
            "comissao_gerente": dv["VALOR COMISSAO GERENTE"].sum(),
        })
        
        dv_pdf = dv.rename(columns={
            "BASE RECORRENCIA": "VALOR_BASE",
            "% CONSULTOR": "PERC_FINAL",
            "VALOR COMISSAO": "VALOR_FINAL"
        })
        
        st.session_state.pdfs[vend] = gerar_pdf_universal(vend, dv_pdf, periodo_txt, "Recorrência")
        st.session_state.resumo.append({
            "nome": vend,
            "total": dv_pdf["VALOR_FINAL"].sum(),
            "qtd": len(dv_pdf)
        })

    # Gerente Recorrência
    df_gerente = df[df["VALOR COMISSAO GERENTE"] > 0].copy()
    if not df_gerente.empty:
        dv_pdf_gerente = df_gerente.copy().rename(columns={
            "BASE RECORRENCIA": "VALOR_BASE",
            "% GERENTE": "PERC_FINAL",
            "VALOR COMISSAO GERENTE": "VALOR_FINAL"
        })
        st.session_state.pdfs["_GERENTE_RECORRENCIA"] = gerar_pdf_gerente_recorrencia(dv_pdf_gerente, periodo_txt)

    if resumo_gerencial:
        st.session_state.pdfs["_GERENCIAL_RECORRENCIA"] = gerar_pdf_gerencial_recorrencia(resumo_gerencial, periodo_txt, perc_imp if aplicar_imp else 0.0)

    if not st.session_state.resumo and "_GERENTE_RECORRENCIA" not in st.session_state.pdfs:
        st.warning("Nenhuma comissão de recorrência (padrão ou exceção) foi encontrada no arquivo.")
        return False

    st.session_state.pronto = True
    return True

# ==========================================
# 5.1 FUNÇÃO AUXILIAR DE LEITURA (RECORRÊNCIA)
# ==========================================
def _ler_e_preparar_planilha_recorrencia(arquivo):
    try:
        df_raw = pd.read_excel(arquivo, header=None)
    except Exception as e:
        return None, f"Erro ao ler o arquivo: {e}"

    # 1. Localizar cabeçalho
    indice_inicio = df_raw[
        df_raw.apply(lambda row: row.astype(str).str.upper().str.contains("PLACA").any(), axis=1)
    ].index

    if len(indice_inicio) == 0:
        return None, "Erro: Cabeçalho 'PLACA' não encontrado."
    
    df = df_raw.iloc[indice_inicio[0]:].copy()
    df.columns = df.iloc[0]
    df = df[1:]
    df.columns = df.columns.astype(str).str.upper().str.strip()

    # 2. Renomear colunas
    mapeamento = {
        "VALOR MENSALIDADE": "BASE RECORRENCIA",
        "BASE MENSALIDADE": "BASE RECORRENCIA",
        "MENSALIDADE": "BASE RECORRENCIA",
        "VALOR BASE RECORRÊNCIA": "BASE RECORRENCIA",
        "PORCENTAGEM": "PORCENTAGEM RECORRENCIA",
        "PERCENTUAL": "PORCENTAGEM RECORRENCIA",
        "% RECORRÊNCIA": "PORCENTAGEM RECORRENCIA",
        "VOLUNTÁRIO": "CONSULTOR"
    }
    df.rename(columns=mapeamento, inplace=True)

    # 3. Validar colunas
    obrigatorias = ["PLACA", "ASSOCIADO", "CONSULTOR", "BASE RECORRENCIA", "PORCENTAGEM RECORRENCIA"]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        return None, f"Colunas faltando: {faltando}"

    # 4. Limpeza e Conversão
    df["PLACA"] = df["PLACA"].astype(str).str.upper().str.strip()
    df["BASE RECORRENCIA"] = df["BASE RECORRENCIA"].apply(limpar_moeda)
    df["PORCENTAGEM RECORRENCIA"] = df["PORCENTAGEM RECORRENCIA"].apply(normalizar_percentual)
    
    df = df.dropna(subset=["PLACA", "BASE RECORRENCIA", "PORCENTAGEM RECORRENCIA", "CONSULTOR"])
    
    df["VALOR MENSAL"] = df["BASE RECORRENCIA"]

    return df, None

# ==========================================
# 5.1 PROCESSAMENTO EVOLUÇÃO
# ==========================================
def processar_evolucao(arq_ant, arq_atu, aplicar_imp, perc_imp):
    df_ant, err_ant = _ler_e_preparar_planilha_recorrencia(arq_ant)
    if err_ant: st.error(f"Erro Mês Anterior: {err_ant}"); return False
    
    df_atu, err_atu = _ler_e_preparar_planilha_recorrencia(arq_atu)
    if err_atu: st.error(f"Erro Mês Atual: {err_atu}"); return False

    def calcular_comissoes(df):
        if df is None or df.empty:
            return df

        df["% CONSULTOR"] = df["PORCENTAGEM RECORRENCIA"] / 2
        df["% GERENTE"] = df["PORCENTAGEM RECORRENCIA"] / 2

        df["COMISSAO_CONSULTOR_BRUTA"] = (df["BASE RECORRENCIA"] * df["% CONSULTOR"] / 100).round(2)
        df["VALOR COMISSAO"] = (df["COMISSAO_CONSULTOR_BRUTA"] * (1 - perc_imp/100 if aplicar_imp else 1.0)).round(2)
        df["VALOR COMISSAO GERENTE"] = (df["BASE RECORRENCIA"] * df["% GERENTE"] / 100).round(2)
        df["IMPOSTO"] = (df["COMISSAO_CONSULTOR_BRUTA"] * (perc_imp/100 if aplicar_imp else 0.0)).round(2)
        return df

    df_atu = calcular_comissoes(df_atu)
    df_ant = calcular_comissoes(df_ant)

    placas_ant = set(df_ant["PLACA"])
    placas_atu = set(df_atu["PLACA"])
    
    comum = placas_ant.intersection(placas_atu)
    novas = placas_atu - placas_ant
    canceladas = placas_ant - placas_atu
    
    resumo_gerencial_evo = []
    consultores_atuais = sorted(df_atu["CONSULTOR"].unique())
    for cons in consultores_atuais:
        df_cons = df_atu[df_atu["CONSULTOR"] == cons]
        resumo_gerencial_evo.append({
            "consultor": cons,
            "qtd_ativos": len(df_cons),
            "total_recorrencia_consultor": df_cons["VALOR COMISSAO"].sum(),
            "total_impostos": df_cons["IMPOSTO"].sum(),
            "total_recorrencia_gerente": df_cons["VALOR COMISSAO GERENTE"].sum(),
            "faturamento_base": df_cons["BASE RECORRENCIA"].sum()
        })

    df_inativos = df_ant[df_ant["PLACA"].isin(canceladas)].copy()
    df_inativos = df_inativos.rename(columns={"VALOR COMISSAO": "VALOR RECORRENCIA"})
    
    st.session_state.evo_resumo = {
        "mantidas": {"qtd": len(comum), "valor": df_atu[df_atu["PLACA"].isin(comum)]["BASE RECORRENCIA"].sum()},
        "novas": {"qtd": len(novas), "valor": df_atu[df_atu["PLACA"].isin(novas)]["BASE RECORRENCIA"].sum()},
        "canceladas": {"qtd": len(canceladas), "valor": df_ant[df_ant["PLACA"].isin(canceladas)]["BASE RECORRENCIA"].sum()}
    }
    
    st.session_state.pdfs_evolucao = {}
    periodo_evo_txt = f"Comparativo: {arq_ant.name} vs {arq_atu.name}"
    
    st.session_state.pdfs_evolucao["_EVOLUCAO_GERENCIAL"] = gerar_pdf_evolucao_gerencial(
        resumo_gerencial_evo,
        st.session_state.evo_resumo,
        periodo_evo_txt
    )

    if not df_inativos.empty:
        st.session_state.pdfs_evolucao["_EVOLUCAO_INATIVOS"] = gerar_pdf_inativos(
            df_inativos,
            periodo_evo_txt
        )

    st.session_state.pronto_evolucao = True
    return True

# ==========================================
# 6. UI PRINCIPAL
# ==========================================
def main():
    aplicar_estilos()
    st.title("🚀 Gestão de Comissões - Pense e Conecte")

    with st.sidebar:
        st.header("⚙️ Configurações")
        modulo = st.radio("Selecione o Módulo:", ["Adesões", "Recorrência", "Evolução"])
        
        if modulo == "Adesões":
            d_ini = st.date_input("Data Início", datetime.date.today() - datetime.timedelta(days=30))
            d_fim = st.date_input("Data Fim", datetime.date.today())
            imp = st.checkbox("Habilitar Desconto NF", value=True)
            perc_imp = st.number_input("% de Imposto", 0.0, 100.0, 12.0) if imp else 0.0
            arquivo = st.file_uploader("Suba sua planilha Excel", type=["xlsx"])
        elif modulo == "Recorrência":
            mes = st.selectbox("Mês de Referência", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"])
            ano = st.number_input("Ano", 2024, 2030, 2026)
            imp = st.checkbox("Habilitar Desconto NF", value=True)
            perc_imp = st.number_input("% de Imposto", 0.0, 100.0, 12.0) if imp else 0.0
            arquivo = st.file_uploader("Suba sua planilha Excel", type=["xlsx"])
        else:
            st.info("Comparativo entre meses")
            arq_ant = st.file_uploader("Mês Anterior", type=["xlsx"], key="ant")
            arq_atu = st.file_uploader("Mês Atual", type=["xlsx"], key="atu")
            imp = st.checkbox("Habilitar Desconto NF", value=True, key="imp_evo")
            perc_imp = st.number_input("% de Imposto", 0.0, 100.0, 12.0, key="perc_imp_evo") if imp else 0.0
            arquivo = None # Placeholder
            
        if st.button("PROCESSAR"):
            if modulo == "Evolução":
                if arq_ant and arq_atu:
                    processar_evolucao(arq_ant, arq_atu, imp, perc_imp)
                    st.rerun()
                else:
                    st.warning("Selecione os dois arquivos para comparação.")
            elif not arquivo:
                st.warning("Por favor, selecione um arquivo para processar.")
            else:
                success = False
                if modulo == "Adesões":
                    success = processar_adesao(arquivo, d_ini, d_fim, imp, perc_imp)
                else:
                    success = processar_recorrencia(arquivo, mes, ano, imp, perc_imp)
                if success:
                    st.rerun()

    # Exibição Evolução
    if st.session_state.get("pronto_evolucao") and modulo == "Evolução":
        r = st.session_state.evo_resumo
        st.subheader("📊 Análise de Evolução da Carteira")
        c1, c2, c3 = st.columns(3)
        c1.metric("✅ Mantidas (Base de Recorrência)", f"{r['mantidas']['qtd']} placas", f"R$ {r['mantidas']['valor']:,.2f}")
        c2.metric("🚀 Novas (Base de Recorrência)", f"{r['novas']['qtd']} placas", f"R$ {r['novas']['valor']:,.2f}")
        c3.metric("🔻 Canceladas (Base de Recorrência)", f"{r['canceladas']['qtd']} placas", f"R$ {r['canceladas']['valor']:,.2f}")
        
        st.divider()
        st.subheader("📄 Relatórios de Evolução")
        pdfs = st.session_state.get("pdfs_evolucao", {})
        
        if "_EVOLUCAO_GERENCIAL" in pdfs:
            st.download_button(
                "📂 Baixar Relatório Gerencial de Evolução",
                pdfs["_EVOLUCAO_GERENCIAL"],
                "Evolucao_Gerencial.pdf"
            )
        
        if "_EVOLUCAO_INATIVOS" in pdfs:
            st.download_button(
                "📂 Baixar Relatório de Inativos",
                pdfs["_EVOLUCAO_INATIVOS"],
                "Evolucao_Inativos.pdf"
            )
        st.divider()
        if st.button("Limpar Comparação"):
            st.session_state.pronto_evolucao = False
            if "pdfs_evolucao" in st.session_state:
                del st.session_state.pdfs_evolucao
            st.rerun()

    elif st.session_state.get("pronto"):
        c1, c2 = st.columns([2, 1])
        with c1:
            st.subheader("📋 Relatórios dos Consultores")
            for item in st.session_state.resumo:
                with st.expander(f"👤 {item['nome']} - R$ {item['total']:.2f}"):
                    st.download_button(f"Baixar PDF {item['nome']}", st.session_state.pdfs[item['nome']], f"Comissao_{item['nome']}.pdf")
        
        with c2:
            st.subheader("🛠️ Área Administrativa")
            if "_GERENCIAL_TRIBUTOS" in st.session_state.pdfs:
                st.success("📊 Detalhamento de Tributos pronto")
                st.download_button("📂 BAIXAR CONFERÊNCIA GERENCIAL", st.session_state.pdfs["_GERENCIAL_TRIBUTOS"], "Gerencial_Adesao_Completo.pdf")
            
            if "_GERENCIAL_RECORRENCIA" in st.session_state.pdfs:
                st.success("📊 Detalhamento Gerencial de Recorrência pronto")
                st.download_button("📂 BAIXAR CONFERÊNCIA GERENCIAL (RECORRÊNCIA)", st.session_state.pdfs["_GERENCIAL_RECORRENCIA"], "Gerencial_Recorrencia_Completo.pdf")
            
            if "_GERENTE_RECORRENCIA" in st.session_state.pdfs:
                st.info("💼 Comissão Gerente (Recorrência)")
                st.download_button("📂 Relatório Gerente", st.session_state.pdfs["_GERENTE_RECORRENCIA"], "Comissao_Gerente_Recorrencia.pdf")
            
            st.divider()
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for k, v in st.session_state.pdfs.items(): 
                    if not k.startswith("_"): zf.writestr(f"Comissao_{k}.pdf", v.getvalue())
            st.download_button("📦 Baixar Todos ZIP", zip_buf.getvalue(), "comissoes_zip.zip")
            if st.button("Limpar"):
                st.session_state.pronto = False
                st.session_state.pronto_evolucao = False
                if "pdfs_evolucao" in st.session_state:
                    del st.session_state.pdfs_evolucao
                st.rerun()

if __name__ == "__main__":
    main()


# ==============================
# Rodapé do Sistema
# ==============================

st.markdown("---")
st.markdown(
    """
    <div style="text-align:center; font-size:13px; color:gray;">
        Desenvolvido por <b>Pablo Lopes</b> | <a href="https://wa.me/5547991608941" target="_blank" style="color:gray; text-decoration:none;">(47) 99160-8941</a><br>
        Versão 1.0
    </div>
    """,
    unsafe_allow_html=True
)