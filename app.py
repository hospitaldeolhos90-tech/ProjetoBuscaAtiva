import streamlit as st

import streamlit.components.v1 as components

import pandas as pd

import openpyxl

from openpyxl.styles import Alignment, Font

import io

import re

import math

import zipfile

import base64

import csv

import time

from datetime import datetime

import xml.etree.ElementTree as ET



# --- 1. CONFIGURAÇÃO MASTER ---

try:

    st.set_page_config(

        page_title="Hova | Master Intelligence", 

       page_icon="💠",

        layout="wide",

        initial_sidebar_state="auto"

    )

except:

    pass



# --- 2. INICIALIZAÇÃO DE MEMÓRIA E DICIONÁRIOS ---

if 'lista_pendencias' not in st.session_state:

    st.session_state['lista_pendencias'] = []

if 'movimento_gravado' not in st.session_state:

    st.session_state['movimento_gravado'] = False

if 'base_deduplicada' not in st.session_state:

    st.session_state['base_deduplicada'] = None



# --- MEGA DICIONÁRIO HOVA (BOMBA DE DADOS) ---

if 'dicionario_tipos' not in st.session_state:

    st.session_state['dicionario_tipos'] = pd.DataFrame([

        # Consultas e Retornos

        {"SIGLA": "C", "TRADUCAO": "Consulta Especializada"},

        {"SIGLA": "R", "TRADUCAO": "Retorno"},

        {"SIGLA": "RC", "TRADUCAO": "Retorno de Cirurgia"},

        {"SIGLA": "RT", "TRADUCAO": "Reconsulta"},

        {"SIGLA": "RL", "TRADUCAO": "Retorno de Lente"},

        {"SIGLA": "RO", "TRADUCAO": "Revisão de Óculos"},

        {"SIGLA": "CRA", "TRADUCAO": "Consulta Reavaliação Anual"},

        {"SIGLA": "CDI", "TRADUCAO": "Consulta Diagnóstico Inicial"},

        {"SIGLA": "C_INTER", "TRADUCAO": "Consulta de Intercorrência"},

        {"SIGLA": "PO", "TRADUCAO": "Pós Operatório"},

        {"SIGLA": "AC", "TRADUCAO": "Avaliação de Cirurgia"},

        {"SIGLA": "AGL", "TRADUCAO": "Avaliação de Glaucoma"},

        {"SIGLA": "AR", "TRADUCAO": "Avaliação com Retinólogo"},

        {"SIGLA": "PREOP", "TRADUCAO": "Pré-Operatório"},

        {"SIGLA": "NAN", "TRADUCAO": "CÉLULA VAZIA NO DOCTORS"},

        {"SIGLA": "NONE", "TRADUCAO": "CÉLULA VAZIA NO DOCTORS"},
        

        

        # Variações de Avaliações / Glaucoma / Catarata

        {"SIGLA": "1 CAT", "TRADUCAO": "1ª Avaliação"},

        {"SIGLA": "2 CAT", "TRADUCAO": "2ª Avaliação"},

        {"SIGLA": "3 CAT", "TRADUCAO": "3ª Avaliação"},

        {"SIGLA": "1 CCAT", "TRADUCAO": "1ª Avaliação"},

        {"SIGLA": "2 CCAT", "TRADUCAO": "2ª Avaliação"},

        {"SIGLA": "3 CCAT", "TRADUCAO": "3ª Avaliação"},

        {"SIGLA": "1ªCAT", "TRADUCAO": "1ª Avaliação"},

        {"SIGLA": "2ªCAT", "TRADUCAO": "2ª Avaliação"},

        {"SIGLA": "3ªCAT", "TRADUCAO": "3ª Avaliação"},

        

        # Exames

        {"SIGLA": "AGF", "TRADUCAO": "Angiofluoresceinografia"},

        {"SIGLA": "RF", "TRADUCAO": "Retinografia Fluorescente"},

        {"SIGLA": "RD", "TRADUCAO": "Retinografia Digital"},

        {"SIGLA": "TOPO", "TRADUCAO": "Topografia"},

        {"SIGLA": "CVC", "TRADUCAO": "Campo Visual Computadorizado"},

        {"SIGLA": "IOL", "TRADUCAO": "Biometria Óptica (IOL Master)"},

        {"SIGLA": "BIO", "TRADUCAO": "Biometria Ultrassônica"},

        {"SIGLA": "OCT", "TRADUCAO": "Tomografia de Coerência Óptica"},

        {"SIGLA": "ANGIO-OCT", "TRADUCAO": "Angio-OCT"},

        {"SIGLA": "PTC", "TRADUCAO": "Pentacam"},

        {"SIGLA": "MBG", "TRADUCAO": "Meibomiografia"},

        {"SIGLA": "MEIBO", "TRADUCAO": "Meibomiografia"},

        {"SIGLA": "TSC", "TRADUCAO": "Teste de Schirmer"},

        {"SIGLA": "TSH", "TRADUCAO": "Teste de Sobrecarga Hídrica"},

        {"SIGLA": "MR", "TRADUCAO": "Mapeamento de Retina"},

        {"SIGLA": "US", "TRADUCAO": "Ultrassom (Ecografia)"},

        {"SIGLA": "CDPO", "TRADUCAO": "Curva Diária de Pressão"},

        {"SIGLA": "GONIO", "TRADUCAO": "Gonioscopia"},

        {"SIGLA": "PAQ", "TRADUCAO": "Paquimetria"},

        {"SIGLA": "MICRO", "TRADUCAO": "Microscopia Especular"},

        {"SIGLA": "AV", "TRADUCAO": "Acuidade Visual"},

        {"SIGLA": "EX", "TRADUCAO": "Tonometria"},

        {"SIGLA": "TO", "TRADUCAO": "Teste do Olhinho"},

        {"SIGLA": "ISHIHARA", "TRADUCAO": "Teste de Ishihara"},



        # PROCEDIMENTOS ADICIONAIS / ENFERMAGEM

        {"SIGLA": "VVPP", "TRADUCAO": "Vitrectomia Posterior"},

        {"SIGLA": "ESTRABISMO", "TRADUCAO": "Cirurgia de Estrabismo"},

        {"SIGLA": "CICLOFOTO", "TRADUCAO": "Ciclofotocoagulação"},

        {"SIGLA": "TRAB", "TRADUCAO": "Trabeculectomia"},

        {"SIGLA": "ANTIVEG", "TRADUCAO": "Aplicação Intravítrea"},

        

        # Lentes de Contato

        {"SIGLA": "TL", "TRADUCAO": "Teste de Lente"},

        {"SIGLA": "TLR", "TRADUCAO": "Teste de Lente Rígida"},

        {"SIGLA": "TLESCL", "TRADUCAO": "Teste de Lente Escleral"},

        {"SIGLA": "BL", "TRADUCAO": "Busca de Lente"},

        {"SIGLA": "CLGN", "TRADUCAO": "Compra Lente Gelatinosa"},

        {"SIGLA": "CLGP", "TRADUCAO": "Compra Lente Gelatinosa Positiva"},

        {"SIGLA": "CLIO", "TRADUCAO": "Compra Lente Intra Ocular"},

        {"SIGLA": "CLMF", "TRADUCAO": "Compra Lente Multifocal"},

        {"SIGLA": "CLR", "TRADUCAO": "Compra Lente Rígida"},

        {"SIGLA": "CLT", "TRADUCAO": "Compra Lente Tórica"},

        {"SIGLA": "CLTERAPEUTICA", "TRADUCAO": "Compra Lente Terapêutica"},

        

        # Procedimentos / Laser

        {"SIGLA": "YAG", "TRADUCAO": "Capsulotomia Yag Laser"},

        {"SIGLA": "FOTO", "TRADUCAO": "Fotocoagulação a Laser"},

        {"SIGLA": "IRI", "TRADUCAO": "Iridectomia"},

        {"SIGLA": "FOTOTRAB", "TRADUCAO": "Fototrabeculoplastia"},

        {"SIGLA": "ILIO", "TRADUCAO": "Procedimento ILIO"},

        {"SIGLA": "LP", "TRADUCAO": "Luz Pulsada"},

        

        # Instruções e Cirurgias

        {"SIGLA": "IC", "TRADUCAO": "Instrução Cirúrgica"},

        {"SIGLA": "IC FACO", "TRADUCAO": "Instrução de Cirurgia de Faco"},

        {"SIGLA": "IC_FACO", "TRADUCAO": "Instrução de Cirurgia de Faco"},

        {"SIGLA": "IC VITRE", "TRADUCAO": "Instrução de Vitrectomia"},

        {"SIGLA": "IC_VITRE", "TRADUCAO": "Instrução de Vitrectomia"},

        {"SIGLA": "IC PTERIGIO", "TRADUCAO": "Instrução de Pterígio"},

        {"SIGLA": "IC_PTERIGIO", "TRADUCAO": "Instrução de Pterígio"},

        {"SIGLA": "IC CICLO", "TRADUCAO": "Instrução de Cirurgia de Ciclofoto"},

        {"SIGLA": "IC_CICLO", "TRADUCAO": "Instrução de Cirurgia de Ciclofoto"},

        {"SIGLA": "IC CALAZIO", "TRADUCAO": "Instrução de Calázio"},

        {"SIGLA": "IC_CALAZIO", "TRADUCAO": "Instrução de Calázio"},

        {"SIGLA": "RIC", "TRADUCAO": "Risco Cirúrgico"},

        {"SIGLA": "RET.CIR", "TRADUCAO": "Retoque de Cirurgia"},

        

        # Administrativo / Outros

        {"SIGLA": "CO", "TRADUCAO": "Conversa"},

        {"SIGLA": "EEX", "TRADUCAO": "Entrega de Exames"},

        {"SIGLA": "TFD", "TRADUCAO": "Entrega de TFD"},

        {"SIGLA": "PDC", "TRADUCAO": "Pendência de Marcação"},

        {"SIGLA": "MD", "TRADUCAO": "Medicação"},

        {"SIGLA": "CE", "TRADUCAO": "Compra de Estojo"},

        {"SIGLA": "RP", "TRADUCAO": "Retirar Ponto"},

        {"SIGLA": "RPG", "TRADUCAO": "Restante Pagamento"},

        {"SIGLA": "PAG.A", "TRADUCAO": "Pagamento Antecipado"},

        {"SIGLA": "TX", "TRADUCAO": "Taxas Administrativas"},

        {"SIGLA": "EXLAB", "TRADUCAO": "Exames Laboratoriais"},

        {"SIGLA": "RE", "TRADUCAO": "2ª Via de Exame"},

        {"SIGLA": "REP.EX", "TRADUCAO": "Repetir Exame"}

    ])



# --- REGRAS DE OURO DA CLÍNICA ---

tipos_ignorados = ['R', 'CRA', 'EEX', '1', '2', '3', '1ª', '2ª', '3ª', 'CAT', 'CCAT', '1CAT', '2CAT', '3CAT', '1CCAT', '2CCAT', '3CCAT', '1ªCAT', '2ªCAT', '3ªCAT']



agendas_enviar = [
    'Altair Rosa', 'Denise Matos', 'Felipe Ferreira', 'Francesca de As', 
    'Gabriel Conde', 'Gabriel Lemos', 'Gustavo Sampaio', 'Mariluci Mendes', 
    'Luis Antonio', 'Rodrigo Costa', 'Vera Lucia', 
    'Centro Cirúrgico', 'Exames', 'Enfermagem', 'Catarata',
    'Mateus Barbosa', 'Maxwell dos Reis', 'Lucas', 'Victor'
]



# IRI e FOTOTRAB adicionados aqui para ativar o aviso de dilatação

exames_dilatam = ['AGF', 'RF', 'RD', 'MR', 'OCT', 'YAG', 'FOTO', 'ANGIO-OCT', 'IRI', 'FOTOTRAB']



def get_base64(bin_file):

    try:

        with open(bin_file, 'rb') as f: return base64.b64encode(f.read()).decode()

    except: return None



def colorir_status(row):

    status = str(row.get('STATUS DO PACIENTE', ''))

    if '✅ CONFIRMADO' in status:

        return ['background-color: #d4edda; color: #155724'] * len(row)

    elif '❌ CANCELADO' in status:

        return ['background-color: #f8d7da; color: #721c24'] * len(row)

    elif '💬 MENSAGEM RECEBIDA' in status:

        return ['background-color: #cce5ff; color: #004085'] * len(row)

    elif '⏳ SEM RESPOSTA' in status:

        return ['background-color: #fff3cd; color: #856404'] * len(row)

    return [''] * len(row)



# --- 3. DESIGN "HYBRID ELITE" ---

st.markdown(f"""

    <style>

    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap');

    

    @keyframes fadeInUp {{ from {{ opacity: 0; transform: translateY(20px); }} to {{ opacity: 1; transform: translateY(0); }} }}

    @keyframes auroraBG {{ 0% {{ background-position: 0% 50%; }} 50% {{ background-position: 100% 50%; }} 100% {{ background-position: 0% 50%; }} }}



    .stApp {{ background: linear-gradient(-45deg, #f0f7f4, #e0efeb, #ffffff, #cde0dc); background-size: 400% 400% !important; animation: auroraBG 15s ease infinite !important; font-family: 'Outfit', sans-serif; }}

    [data-testid="stSidebar"] {{ background-color: #1e3d3a !important; border-right: 2px solid rgba(0, 255, 204, 0.2); }}

    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {{ color: #ffffff !important; font-family: 'Outfit', sans-serif; }}



    .premium-title {{ color: #1e3d3a; font-weight: 600 !important; font-size: 2.2rem; text-align: center; text-transform: uppercase; margin-top: 50px !important; margin-bottom: 40px; display: flex; align-items: center; justify-content: center; gap: 15px; animation: fadeInUp 0.6s ease-out; }}

    @media (max-width: 768px) {{ .premium-title {{ font-size: 1.4rem !important; margin-top: 20px !important; }} .master-card {{ padding: 25px 15px !important; border-radius: 25px !important; }} div.stButton > button {{ height: 60px !important; font-size: 1.1rem !important; }} .stTabs [data-baseweb="tab"] {{ height: 60px !important; }} .stTabs [data-baseweb="tab"] p {{ font-size: 0.8rem !important; }} }}



    .stTabs [data-baseweb="tab-list"] {{ display: flex !important; width: 100% !important; gap: 8px !important; background-color: transparent !important; padding: 5px; }}

    .stTabs [data-baseweb="tab"] {{ flex-grow: 1 !important; height: 85px !important; background-color: #1e3d3a !important; border-radius: 15px 15px 0 0 !important; transition: all 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275) !important; }}

    .stTabs [data-baseweb="tab"] p {{ font-size: 1.1rem !important; font-weight: 700 !important; color: #FFFFFF !important; }}

    .stTabs [data-baseweb="tab"]:hover {{ transform: translateY(-10px) !important; background-color: #2f6c68 !important; }}

    .stTabs [aria-selected="true"] {{ background-color: #2f6c68 !important; border-bottom: 6px solid #00ffcc !important; }}



    .master-card {{ background: rgba(255, 255, 255, 0.7); backdrop-filter: blur(20px); border-radius: 40px; padding: 60px 80px; box-shadow: 0 40px 80px rgba(0,0,0,0.08); text-align: center; border: 1px solid rgba(255, 255, 255, 0.4); animation: fadeInUp 1s ease-out; }}

    div.stButton > button, div.stDownloadButton > button {{ background: linear-gradient(135deg, #1e3d3a 0%, #2f6c68 100%) !important; color: white !important; border-radius: 60px !important; height: 60px !important; width: 100% !important; font-size: 1.1rem !important; font-weight: 700 !important; text-transform: uppercase !important; border: none !important; transition: 0.4s all cubic-bezier(0.175, 0.885, 0.32, 1.275) !important; letter-spacing: 1px; }}

    div.stButton > button:hover, div.stDownloadButton > button:hover {{ transform: scale(1.07) translateY(-5px) !important; box-shadow: 0 20px 45px rgba(30, 61, 58, 0.35) !important; }}



    .response-box {{ background: white; padding: 25px; border-radius: 22px; border-top: 8px solid #1e3d3a; box-shadow: 0 10px 30px rgba(0,0,0,0.04); margin-bottom: 25px; text-align: left; animation: fadeInUp 0.8s ease-out; }}

    .penc-row {{ background: #f8fafc; padding: 15px; border-radius: 12px; margin-bottom: 10px; font-size: 0.9rem; color: #1e3d3a; font-weight: 700; border-left: 5px solid #1e3d3a; border: 1px solid #e2e8f0; }}

    .footer-master {{ text-align: center; color: #1e3d3a; font-weight: 600; letter-spacing: 10px; margin-top: 80px; margin-bottom: 30px; font-size: 1rem; opacity: 0.8; animation: fadeInUp 1.2s ease-out; }}

    .led-green {{ color: #00ffcc; text-shadow: 0 0 10px #00ffcc; }}

    </style>

""", unsafe_allow_html=True)



# --- 4. FUNÇÕES DE LÓGICA E DEDUPLICAÇÃO ---

def limpar_num(num): return "".join(filter(str.isdigit, str(num)))[-8:]



def eh_celular(num): 

    n = "".join(filter(str.isdigit, str(num)))

    while n.startswith('0'): n = n[1:]

    if len(n) >= 10: return n[-9] == '9' or n[-8] in ['7', '8', '9']

    elif len(n) >= 8: return n[-8] in ['7', '8', '9']

    return False



def eh_numero_falso(num):

    n = str(num).strip()

    if len(n) < 8: return True

    if len(set(n)) == 1: return True 

    if '12345678' in n: return True

    return False



def formatar_telefone_real(num):

    txt = str(num).strip().upper()

    if txt in ['NAN', 'NONE', '<NA>', '']: return ""

    if "(000)" in txt or "___" in txt: return ""

    digitos = re.sub(r'\D', '', txt)

    if not digitos: return ""

    if len(set(digitos)) == 1 and digitos[0] == '0': return ""

    while digitos.startswith('0') and len(digitos) > 9: digitos = digitos[1:]

    return digitos



def pegar_nome_curto(nome_completo):

    try:

        if not nome_completo or str(nome_completo).strip() == "": return ""

        nome_limpo = re.sub(r'[^A-Za-zÀ-ÖØ-öø-ÿ\s]', '', str(nome_completo)).strip()

        partes = nome_limpo.split()

        if len(partes) == 0: return ""

        if len(partes) == 1: return partes[0].capitalize()

        p1 = partes[0].capitalize()

        if partes[1].upper() in ['DE', 'DA', 'DO', 'DOS', 'DAS', 'E']:

            p2 = partes[1].lower()

            if len(partes) > 2: return f"{p1} {p2} {partes[2].capitalize()}"

            else: return f"{p1} {p2}"

        else: return f"{p1} {partes[1].capitalize()}"

    except: return ""



def traduzir_tipo(tipo_raw, dict_df):

    tipo = str(tipo_raw).strip().upper()

    if tipo.startswith("CI_") or tipo.startswith("CI ") or tipo.startswith("CI.") or tipo.startswith("IC ") or tipo.startswith("IC_"):

        medico_cirurgia = tipo[3:].replace(".", " ").replace("_", " ").title()

        return f"Cirurgia Dr(a). {medico_cirurgia}"

    

    if "/" in tipo:

        partes = tipo.split("/")

        traducoes = []

        for p in partes:

            p = p.strip()

            match = dict_df[dict_df['SIGLA'].str.upper() == p]

            if not match.empty: traducoes.append(match.iloc[0]['TRADUCAO'])

            else: traducoes.append(p)

        if len(traducoes) > 1: return ", ".join(traducoes[:-1]) + " e " + traducoes[-1]

        return traducoes[0] if traducoes else "DESCONHECIDO"

        

    match = dict_df[dict_df['SIGLA'].str.upper() == tipo]

    return match.iloc[0]['TRADUCAO'] if not match.empty else "DESCONHECIDO"



def resolver_nome_completo(nome_arquivo):

    n = str(nome_arquivo).upper()

    if "GABRIELL" in n: return "Gabriel Lemos"

    if "GABRIEL" in n: return "Gabriel Conde"

    if "ALTAIR" in n: return "Altair Rosa"

    if "DENISE" in n: return "Denise Matos"

    if "FELIPE" in n: return "Felipe Ferreira"

    if "FRANCESCA" in n or "FRAN" in n: return "Francesca de As"

    if "GUSTAVO" in n: return "Gustavo Sampaio"

    if "MARILUCI" in n: return "Mariluci Mendes"

    if "LUIS" in n: return "Luis Antonio"

    if "RODRIGO" in n: return "Rodrigo Costa"

    if "VERA" in n: return "Vera Lucia"

    if "CIRURGIA" in n: return "Centro Cirúrgico"

    if "LENTE" in n: return "Setor de Lentes de Contato"

    if "ENFERMAG" in n: return "Enfermagem"

    if "EXAME" in n: return "Exames"

    if "GLAUCOMA" in n: return "Glaucoma"

    if "LAUDOS" in n: return "Laudos"

    if "VAGNER" in n: return "Vagner Diniz"

    if "LUCAS" in n: return "Lucas"

    if "CATARATA" in n: return "Catarata"

    if "FARMACIA" in n: return "Farmacia"

    if "MATEUS" in n: return "Mateus Barbosa"

    if "MAXWELL" in n: return "Maxwell dos Reis"

    if "VICTOR" in n: return "Victor"

    partes = n.rsplit('.', 1)[0].replace(" ", "_").split("_")

    return partes[-1].title()



def formatar_brasileiro_sem_hora(txt):

    txt = str(txt).strip()

    if txt.lower() in ['nan', 'none', 'nat', '<na>', '']: return ""

    if re.match(r'^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}', txt): txt = txt.split(' ')[0]

    if re.match(r'^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}', txt): txt = txt.split(' ')[0]

    match_iso = re.fullmatch(r'(\d{4})-(\d{2})-(\d{2})', txt)

    if match_iso: return f"{match_iso.group(3)}/{match_iso.group(2)}/{match_iso.group(1)}"

    return txt



def organizar_mensagens_lista(lista_msgs):

    todas_msgs = set()

    for val in lista_msgs:

        val_str = str(val).strip()

        if val_str and val_str.lower() not in ['nan', 'none', 'nat', '<na>']:

            for l in val_str.split('\n'):

                l = l.strip()

                if l:

                    l = l.replace(' 00:00:00', '')

                    l = re.sub(r'\s*\b\d{2}:\d{2}(:\d{2})?\b', '', l).strip()

                    todas_msgs.add(l)

                    

    def date_key(txt):

        match = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4})', txt)

        if match:

            try: return pd.to_datetime(match.group(1), format='%d/%m/%Y', dayfirst=True)

            except: return pd.Timestamp.min

        return pd.Timestamp.min

        

    return "\n".join(sorted(list(todas_msgs), key=date_key))


# --- FUNÇÃO DE PRÉ-VISUALIZAÇÃO (RAIO-X) ---
def load_excel_with_ui(file_obj, key_prefix):
    if file_obj.name.endswith('.csv'):
        try:
            df = pd.read_csv(file_obj, sep=';', dtype=str)
            if len(df.columns) < 2:
                file_obj.seek(0)
                df = pd.read_csv(file_obj, sep=',', dtype=str)
        except:
            file_obj.seek(0)
            df = pd.read_csv(file_obj, sep=',', dtype=str)
        st.markdown(f"**Pré-visualização ({file_obj.name}):**")
        st.dataframe(df.head(5), use_container_width=True)
        return df
    
    xl = pd.ExcelFile(file_obj)
    if len(xl.sheet_names) > 1:
        sheet = st.selectbox(f"Qual aba processar? ({file_obj.name})", xl.sheet_names, key=f"sheet_{key_prefix}")
    else:
        sheet = xl.sheet_names[0]
    df = pd.read_excel(file_obj, sheet_name=sheet, dtype=str)
    st.markdown(f"**Pré-visualização da aba '{sheet}':**")
    st.dataframe(df.head(5), use_container_width=True)
    return df

# --- NOVOS MOTORES: TELEFONES E CONDUTAS ---
def processar_telefones_avancado(row, cols_telefone):
    telefones_encontrados = []
    for c in cols_telefone:
        val = str(row.get(c, '')).strip()
        if val and val.lower() not in ['nan', 'none', '<na>', '']:
            telefones_encontrados.append(val.upper())
            
    if not telefones_encontrados:
        return ["", ""]
        
    principal = ""
    # 1. Prioridade para quem tem "W" (WhatsApp)
    for t in telefones_encontrados:
        if 'W' in t:
            principal = t
            break
            
    if not principal:
        principal = telefones_encontrados[0]
        
    adicionais = []
    for t in telefones_encontrados:
        if t != principal:
            adicionais.append(t)
            
    def limpar_e_formatar(t):
        so_num = re.sub(r'\D', '', t)
        if not so_num or eh_numero_falso(so_num): return ""
        # Injeta 9 no celular de MG
        if len(so_num) == 10 and so_num.startswith('31') and so_num[2] in ['7', '8', '9']:
            so_num = f"319{so_num[2:]}"
        # Carimba fixo
        elif len(so_num) >= 10 and so_num[2] not in ['7', '8', '9']:
            return f"{so_num} (FIXO)"
        return so_num

    prin_fmt = limpar_e_formatar(principal)
    adic_fmt = []
    for a in adicionais:
        f = limpar_e_formatar(a)
        if f and f != prin_fmt and f not in adic_fmt:
            adic_fmt.append(f)
            
    return [prin_fmt, " / ".join(adic_fmt)]

def formatar_conduta(c):
    c = str(c).upper().strip()
    c = re.sub(r'(?i)(PERTO|LONGE)\s*[+-]?\d+[,.]\d+', '', c)
    c = re.sub(r'(?i)MANTER\s*ÓCULOS|USAR\s*ÓCULOS|ENTREGUE\s*RECEITA', '', c)
    if "RX ÓCULOS" in c or "RX OCULOS" in c:
        c = c.replace("RX ÓCULOS", "").replace("RX OCULOS", "")
        if not c.strip(): return "CONSULTA ANUAL"
    if c.strip() == "C": return "CONSULTA ESPECIALIZADA"
    if "C/US/MR" in c or "US/MR/C" in c or "C/MR/US" in c: return "CONSULTA ANUAL"
    return c.strip()

def rank_conduta(c):
    c = str(c).upper()
    score = 999
    if 'MÊS' in c or 'MES' in c:
        nums = re.findall(r'\d+', c)
        if nums: score = int(nums[0]) * 30
    elif 'TRIMESTRAL' in c: score = 90
    elif 'SEMESTRAL' in c: score = 180
    elif 'ANUAL' in c: score = 365
    desc_bonus = 0 if len(c) > 10 else 50
    return score + desc_bonus

# --- SIDEBAR ---

with st.sidebar:

    st.markdown("<div style='text-align: center; padding-bottom: 20px;'><h2 style='color: white; letter-spacing: 2px;'>OPERACIONAL</h2></div>", unsafe_allow_html=True)

    st.markdown("---")

    qtd_ajustes = len(st.session_state['lista_pendencias'])

    st.markdown(f"""

        <div style='background: rgba(255,255,255,0.05); padding: 15px; border-radius: 15px; border: 1px solid rgba(255,255,255,0.1);'>

            <p style='margin:0; font-size: 0.9rem;'><span class='led-green'>●</span> SERVIDOR: <b>ESTÁVEL</b></p>

            <p style='margin:0; font-size: 0.9rem;'><span class='led-green'>●</span> CONEXÃO REDE: <b>OK</b></p>

            <p style='margin:0; font-size: 0.9rem;'>AJUSTES LOGADOS: <b>{qtd_ajustes}</b></p>

            <p style='margin-top:10px; font-size: 0.75rem; opacity: 0.7;'>Início: {datetime.now().strftime('%H:%M')}</p>

        </div>

    """, unsafe_allow_html=True)

    st.markdown("---")

    if st.button("🆕 REINICIAR APP"):

        st.session_state['lista_pendencias'] = []

        st.session_state['movimento_gravado'] = False

        st.session_state['base_deduplicada'] = None 

        if 'nome_arquivo_aba3' in st.session_state: del st.session_state['nome_arquivo_aba3']

        if 'abas_planilha' in st.session_state: del st.session_state['abas_planilha']

        st.rerun()



# --- HEADER LOGO ---
logo = get_base64('logo.png')
if logo:
    st.markdown(
        f"""
        <div style="text-align:center; padding:20px 0px;">
            <img src="data:image/png;base64,{logo}" width="380" style="filter: drop-shadow(0px 4px 6px rgba(0,0,0,0.1));">
        </div>
        """, 
        unsafe_allow_html=True
    )


# --- ABAS OPERACIONAIS ---

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
    "Triagem", "Divisor", "Conciliador", "Rependentes", "Busca Ativa", "Gravador", "Confirmação Agenda", "Salva-Vidas", "Central de Limpeza", "Agentes IA"
])



with tab1:
    st.markdown('<div class="master-card">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title">TRIAGEM DE CONTATOS</div>""", unsafe_allow_html=True)
    f_tri = st.file_uploader("Suba a Planilha Mestre", type=["xlsx", "csv"], key="tri_file")
    if f_tri:
        df = load_excel_with_ui(f_tri, "triagem")
        if st.button("INICIAR TRIAGEM", type="primary"):
            col = next((c for c in df.columns if 'TELEFONE' in str(c).upper()), df.columns[2])
            
            # Nova Triagem em 3 partes
            df_vazio = df[df[col].isna() | (df[col].astype(str).str.strip() == '') | (df[col].astype(str).str.lower() == 'nan')].copy()
            df_resto = df[~df.index.isin(df_vazio.index)].copy()
            df_c = df_resto[df_resto[col].apply(eh_celular)].copy()
            df_f = df_resto[~df_resto[col].apply(eh_celular)].copy()
            
            c1, c2, c3 = st.columns(3)
            buf_c = io.BytesIO(); df_c.to_excel(buf_c, index=False); c1.download_button("BAIXAR CELULARES", buf_c.getvalue(), "1_CELULARES.xlsx", type="primary", use_container_width=True)
            buf_f = io.BytesIO(); df_f.to_excel(buf_f, index=False); c2.download_button("BAIXAR FIXOS", buf_f.getvalue(), "2_FIXOS.xlsx", use_container_width=True)
            buf_v = io.BytesIO(); df_vazio.to_excel(buf_v, index=False); c3.download_button("BAIXAR S/ CONTATO", buf_v.getvalue(), "3_EM_BRANCO.xlsx", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)



with tab2:
    st.markdown('<div class="master-card">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title">DIVISOR DE LOTES</div>""", unsafe_allow_html=True)
    c_d1, c_d2 = st.columns([3, 1])
    f_div = c_d1.file_uploader("Planilha de Celulares", type=["xlsx", "csv"], key="div_file")
    tam = c_d2.number_input("Qtd p/ Lote", min_value=1, value=40)
    
    if f_div:
        df_l = load_excel_with_ui(f_div, "divisor")
        if st.button("GERAR LOTES CSV", type="primary"):
            try:
                lotes = math.ceil(len(df_l)/tam)
                b_zip = io.BytesIO()
                
                with zipfile.ZipFile(b_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for i in range(lotes):
                        fatia = df_l.iloc[i*tam : (i+1)*tam].copy()
                        csv_buf = io.StringIO()
                        fatia.to_csv(csv_buf, index=False, sep=';', quoting=csv.QUOTE_MINIMAL)
                        zf.writestr(f"LOTE_PT{i+1}.csv", csv_buf.getvalue().encode('utf-8-sig'))
                
                b_zip.seek(0)
                st.success(f"Dividido com sucesso em {lotes} lotes.")
                st.download_button("BAIXAR PACOTE DE LOTES (ZIP)", b_zip.getvalue(), "lotes_hova_csv.zip", mime="application/zip", type="primary")
            except Exception as e:
                st.error(f"Erro ao dividir a planilha: {e}")
    st.markdown('</div>', unsafe_allow_html=True)



with tab3:
    st.markdown('<div class="master-card">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title">CONCILIADOR DE RESPOSTAS</div>""", unsafe_allow_html=True)
    c_op1, c_op2 = st.columns(2)
    op = c_op1.text_input("NOME DO OPERADOR:", value="ESTER").upper()
    col_destino = c_op2.text_input("NOME DA COLUNA (Ex: MSG2026):", value="MSG2026").upper()
    st.markdown("---")
    c_c1, c_c2 = st.columns(2)
    b_geral = c_c1.file_uploader("Base Geral (Planilha Doctor's)", type=["xlsx"], key="base_geral")
    r_zap = c_c2.file_uploader("Relatórios ZapRocket (CSV)", type=["csv"], accept_multiple_files=True)
    
    selected_sheet = None
    if b_geral:
        cache_key = f"abas_{b_geral.name}_{b_geral.size}"
        if cache_key not in st.session_state:
            try:
                b_geral.seek(0)
                with zipfile.ZipFile(b_geral, 'r') as z:
                    xml_content = z.read('xl/workbook.xml')
                    root = ET.fromstring(xml_content)
                    ns = {'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    sheets = root.findall('.//xmlns:sheet', ns)
                    sheet_names = [s.get('name') for s in sheets]
                    st.session_state[cache_key] = sheet_names
            except Exception as e:
                b_geral.seek(0)
                xls = pd.ExcelFile(b_geral)
                st.session_state[cache_key] = xls.sheet_names
        selected_sheet = st.selectbox("Escolha qual Aba (Mês) você quer atualizar:", st.session_state[cache_key])
        
        # Pré-visualização
        b_geral.seek(0)
        df_prev_concil = pd.read_excel(b_geral, sheet_name=selected_sheet, dtype=str)
        st.markdown(f"**Pré-visualização da Aba '{selected_sheet}':**")
        st.dataframe(df_prev_concil.head(3), use_container_width=True)
    
    if b_geral and r_zap and selected_sheet:
        if st.button("CONCILIAR DADOS", type="primary"):
            with st.spinner(f"Injetando respostas na aba '{selected_sheet}'..."):
                try:
                    b_geral.seek(0)
                    wb = openpyxl.load_workbook(b_geral)
                    ws = wb[selected_sheet]  
                    header = [str(cell.value).strip().upper() if cell.value else "" for cell in ws[1]]
                    alvo_limpo = col_destino.replace(" ", "")
                    idx_msg = next((i for i, h in enumerate(header) if alvo_limpo in h.replace(" ", "")), None)
                    idx_tel = next((i for i, h in enumerate(header) if "TELEFONE" in h.replace(" ", "")), None)
                    
                    if idx_msg is None: st.error(f"Não achei a coluna '{col_destino}'.")
                    elif idx_tel is None: st.error(f"Não achei a coluna 'TELEFONE'.")
                    else:
                        idx_msg += 1; idx_tel += 1; s_map = {}
                        for r in r_zap:
                            r.seek(0)
                            texto_csv = r.read().decode('utf-8', errors='ignore')
                            separador = ';' if ';' in texto_csv.splitlines()[0] else ','
                            for ln in texto_csv.splitlines():
                                p = ln.split(separador)
                                if len(p) > 1:
                                    num = limpar_num(p[0])
                                    linha_inteira = str(p).lower()
                                    dt = re.search(r'(\d{2}/\d{2}/\d{4})', linha_inteira)
                                    data_final = dt.group(1) if dt else datetime.now().strftime('%d/%m/%Y')
                                    
                                    if "sent" in linha_inteira or "enviado" in linha_inteira:
                                        s_map[num] = f"{data_final} ENVIEI ZAP ROCKET {op}"
                                    else:
                                        s_map[num] = f"SEM WHATS {op} {data_final}" # Invertido para data vir primeiro
                                        
                        max_col_needed = max(idx_tel, idx_msg)
                        for row_cells in ws.iter_rows(min_row=2, max_col=max_col_needed):
                            cell_tel = row_cells[idx_tel - 1]
                            n_l = limpar_num(str(cell_tel.value))
                            if n_l in s_map:
                                cell_msg = row_cells[idx_msg - 1]
                                m_n = s_map[n_l]
                                m_a = str(cell_msg.value or "")
                                if m_n not in m_a:
                                    # Ajuste Quebra de Linha
                                    cell_msg.value = m_n if m_a in ["","None","nan"] else f"{m_a}\n{m_n}"
                                    
                        out_f = io.BytesIO(); wb.save(out_f)
                        st.session_state['base_conciliada'] = out_f.getvalue()
                        st.success(f"Aba '{selected_sheet}' conciliada com sucesso.")
                except Exception as e: st.error(f"Erro inesperado: {e}")

    if 'base_conciliada' in st.session_state:
        st.download_button("BAIXAR PLANILHA CONCILIADA", st.session_state['base_conciliada'], "Base_Final_Hova.xlsx", type="primary")
    st.markdown('</div>', unsafe_allow_html=True)


with tab4:
    st.markdown('<div class="master-card">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title">FILTRO DE REENVIO</div>""", unsafe_allow_html=True)
    cr1, cr2 = st.columns(2)
    
    f_ori = cr1.file_uploader("Lote Original", type=["xlsx", "csv"], key="original_lote")
    f_ret = cr2.file_uploader("Retorno Zap", type=["csv"], accept_multiple_files=True, key="retorno_lote")
    
    if f_ori and f_ret:
        df_o = load_excel_with_ui(f_ori, "reenvio")
        if st.button("FILTRAR PENDENTES", type="primary"):
            try:
                col_t = next((c for c in df_o.columns if 'TELEFONE' in str(c).upper()), df_o.columns[0])
                enviados = set()
                
                for r in f_ret:
                    r.seek(0)
                    linhas = r.read().decode('utf-8', errors='ignore').splitlines()
                    for ln in linhas: 
                        p = ln.split(',')
                        enviados.add(limpar_num(p[0]))
                
                df_pt2 = df_o[~df_o[col_t].apply(limpar_num).isin(enviados)].copy()
                csv_buf = io.StringIO()
                df_pt2.to_csv(csv_buf, index=False, sep=';', quoting=csv.QUOTE_MINIMAL)
                
                st.success(f"Filtro aplicado! Restaram {len(df_pt2)} pacientes.")
                st.download_button("BAIXAR LISTA DE REENVIO", csv_buf.getvalue().encode('utf-8-sig'), "REENVIO_FALTANTES.csv", mime="text/csv", type="primary")
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
    st.markdown('</div>', unsafe_allow_html=True)



with tab5:

    st.markdown('<div class="master-card" style="padding: 40px;">', unsafe_allow_html=True)

    st.markdown("""<div class="premium-title">BUSCA ATIVA E RESPOSTAS RÁPIDAS</div>""", unsafe_allow_html=True)

    

    # 1. Configurações Globais (O "Preenchimento Mágico")

    col_ia, col_nome = st.columns([1, 2])

    ias_nomes = ["Ester", "Clara", "Iris", "Piter", "Theia", "Lumina", "Aurora", "Verônica", "Glauco"]

    ia_sel = col_ia.selectbox("IA ATIVA:", options=ias_nomes, index=0)

    nome_paciente_global = col_nome.text_input("NOME DO PACIENTE:", placeholder="Digite o nome para preencher todas as mensagens...")

    nome_display = nome_paciente_global.strip().title() if nome_paciente_global else "[NOME DO PACIENTE]"

    

    st.markdown("---")



    # 2. Banco de Textos Formatados Dinamicamente

    banco = {

        "Aviso de Áudio 🎧": f"Olá, *{nome_display}*! 👁️\nAqui é a *{ia_sel}*, do Hospital de Olhos Vale do Aço.\nVou te enviar um áudio explicativo com informações importantes sobre seu atendimento.\nPor favor, ouça com atenção. 🎧\nQualquer dúvida, estou à disposição. 🤍",

        

        "Contato Incorreto / Cadastro Antigo 🚫": f"Olá! 👁️\nPedimos desculpas pelo contato anterior. 🙏\nSeu número estava vinculado a um cadastro antigo — já corrigimos para que não aconteça novamente.\nAgradecemos o aviso!\nSe você ou sua família precisarem de cuidado com a visão, entre em contato pelo nosso canal oficial:\n👉 https://wa.me/553138011800\nEstamos à disposição. 🤍\nHospital de Olhos Vale do Aço",

        

        "Lembrete Indevido (Já Consultou) 😊": f"Olá, *{nome_display}*! 👁️\nPedimos desculpas pelo envio. 🙏\nComo nosso sistema de lembretes é automático, acabou enviando a mensagem mesmo você já tendo consultado recentemente.\nAgradecemos por avisar! Já atualizamos aqui para que você não receba novos lembretes por enquanto. 😊\nEstamos à disposição. 🤍\nHospital de Olhos Vale do Aço",

        

        "Redirecionamento para Canal Oficial 🤖": f"Olá! Obrigado pelo contato. 👁️\nEste número é apenas para lembretes automáticos e não recebe mensagens.\nPara falar com nossa equipe, clique aqui:\n👉 https://wa.me/553138011800\n\nNossa equipe pode ajudar com:\n✔️ Agendamentos\n✔️ Reagendamentos\n✔️ Dúvidas sobre exames\n✔️ Informações sobre consultas\n✔️ Qualquer outra necessidade\n\nEstamos à disposição! 🤍\nHospital de Olhos Vale do Aço",

        

        "Óbito (Acolhimento Familiar) 🕊️": f"Olá, 🕊️\nAqui é a *{ia_sel}*, do Hospital de Olhos Vale do Aço.\nRecebemos a notícia do falecimento de *{nome_display}* e queremos expressar nossos sentimentos à família.\nPedimos desculpas pela mensagem automática que foi enviada — não sabíamos o que havia acontecido.\nJá corrigimos nosso cadastro e não enviaremos mais mensagens.\nDesejamos que Deus conforte o coração de toda a família neste momento difícil.\nNossos sentimentos,\n*{ia_sel}*\nHospital de Olhos Vale do Aço\n🕊️ ❤️",

        

        "Retorno Preventivo 📅": f"Olá, *{nome_display}*! Tudo bem? 👁️\n\nAqui é a *{ia_sel}*, do *Hospital de Olhos Vale do Aço*.\n\nEstava revisando seu histórico e percebi que já faz um tempo desde sua última consulta. A saúde dos seus olhinhos merece atenção regular — está na hora de realizar uma nova avaliação preventiva.\n\n📲 *Para agendar:*\n1. Toque no link azul abaixo\n2. O WhatsApp abre automaticamente\n3. Nossa equipe vai te atender\n\n👇 *Clique aqui para agendar:*\nhttps://wa.me/553138011800\n\n🛑 *Aviso:* Este número só envia lembretes. Para agendar, use o link acima.\n\n💡 Já consultou recentemente? Pode desconsiderar esta mensagem.\n\nPodemos te esperar? 🤍\n\n*{ia_sel} — Hospital de Olhos Vale do Aço*",

        

        "Pendência de Documentos 📁": f"Olá, *{nome_display}*. Aqui é o(a) {ia_sel}. Notamos que você tem um agendamento, mas ainda constam pendências de exames laboratoriais ou risco cirúrgico em seu prontuário. Por favor, envie uma foto dos resultados por aqui para validarmos sua vaga.",

        

        "Cancelar ou Remarcar ❌": f"Olá,\n\nPara cancelar ou remarcar, não pode ser por aqui.\n\nToque no link azul abaixo para falar com a equipe:\n👉 https://wa.me/553138011800\n\n(É só tocar que abre sozinho)\n\nA equipe vai te ajudar a cancelar ou remarcar! 🤍\n\n{ia_sel} — Hospital de Olhos Vale do Aço"

    }

    

    c_msgs, c_log = st.columns([1.8, 1])

    

    with c_msgs:

        with st.expander("📂 VER TODAS AS MENSAGENS PRONTAS", expanded=True):

            for titulo, texto in banco.items():

                st.markdown(f'<div class="response-box" style="margin-bottom: 10px;"><b>{titulo}</b>', unsafe_allow_html=True)

                st.code(texto, language="text") 

                

                if st.button(f"Salvar no Log", key=f"btn_log_{titulo}"):

                    if nome_paciente_global:

                        st.session_state['lista_pendencias'].append({

                            "Data": datetime.now().strftime("%d/%m"), 

                            "Paciente": nome_paciente_global.upper(), 

                            "Motivo": titulo, 

                            "IA": ia_sel

                        })

                        st.toast(f"Log: {nome_paciente_global.upper()} registrado!")

                    else:

                        st.warning("⚠️ Digite o Nome do Paciente lá em cima para salvar no LOG.")

                st.markdown('</div>', unsafe_allow_html=True)

                

    with c_log:

        st.markdown("<div style='background:#ffffff; padding:25px; border-radius:25px; border:1px solid #e2e8f0;'>", unsafe_allow_html=True)

        st.markdown("<h4 style='color:#1e3d3a; margin-top:0;'>📝 LOG DA EQUIPE</h4>", unsafe_allow_html=True)

        if st.session_state['lista_pendencias']:

            for it in st.session_state['lista_pendencias']:

                st.markdown(f'<div class="penc-row">{it["Paciente"]}<br><small>{it["Motivo"]} | {it["IA"]} | {it["Data"]}</small></div>', unsafe_allow_html=True)

            df_aj = pd.DataFrame(st.session_state['lista_pendencias']); b_aj = io.BytesIO(); df_aj.to_excel(b_aj, index=False)

            st.download_button("📥 EXPORTAR LOGS (EXCEL)", b_aj.getvalue(), "ajustes_hova.xlsx")

            if st.button("LIMPAR LOGS"): st.session_state['lista_pendencias'] = []; st.rerun()

        else: st.info("Vazio.")

        st.markdown('</div>', unsafe_allow_html=True)



with tab6:
    st.markdown('<div class="master-card">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title">🤖 APRENDER E REPETIR MOVIMENTO</div>""", unsafe_allow_html=True)
    col_learn, col_apply = st.columns(2)
    with col_learn:
        st.subheader("1. Gravar Movimento")
        f_orig = st.file_uploader("Planilha ANTES", type=["xlsx"], key="orig")
        f_done = st.file_uploader("Planilha DEPOIS", type=["xlsx"], key="done")
        if f_orig and f_done and st.button("🔴 GRAVAR MEU MOVIMENTO"):
            st.session_state['movimento_gravado'] = True
            st.success("✅ Movimento lido! Já aprendi a blindar seus números.")

    with col_apply:
        st.subheader("2. Repetir Sequência")
        f_batch = st.file_uploader("Suba as outras planilhas", type=["xlsx", "html", "htm"], accept_multiple_files=True, key="batch")
        if f_batch and st.session_state['movimento_gravado'] and st.button("🤖 REPETIR MOVIMENTO EM TODAS"):
            b_zip = io.BytesIO()
            with zipfile.ZipFile(b_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, f in enumerate(f_batch):
                    f.seek(0)
                    if f.name.lower().endswith(('.html', '.htm')):
                        dfs = pd.read_html(f.read())[0].astype(str)
                    else:
                        df = pd.read_excel(f, dtype=str)
                    for col in df.columns:
                        df[col] = df[col].apply(lambda x: f'="{x}"' if pd.notnull(x) else x)
                    nome_original = f.name.rsplit('.', 1)[0].upper().replace(" ", "_")
                    csv_buf = io.StringIO(); csv_buf.write("sep=;\n")
                    df.to_csv(csv_buf, index=False, sep=';', quoting=csv.QUOTE_NONE, escapechar="\\", encoding='utf-8-sig')
                    zf.writestr(f"{nome_original}.csv", csv_buf.getvalue())
            
            b_zip.seek(0)
            st.session_state['zip_tab6'] = b_zip.getvalue()
            st.success("🤖 Sequência repetida com sucesso!")
            
        if 'zip_tab6' in st.session_state:
            st.download_button("📥 BAIXAR RESULTADO", st.session_state['zip_tab6'], "hova_movimento_concluido.zip", mime="application/zip")

    st.markdown('</div>', unsafe_allow_html=True)

with tab7:
    st.markdown('<div class="master-card">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title"> CONFIRMAÇÃO MASTER (MÓDULO CIRÚRGICO ATIVO)</div>""", unsafe_allow_html=True)
    
    with st.expander("💬 NOVO MODELO ESTRATÉGICO ZAPROCKET (VARIÁVEIS AUTOMÁTICAS)"):
        st.info("💡 **DICA:** Copie o texto abaixo e cole direto no ZapRocket. As variáveis `{{column}}` já estão configuradas para puxar os dados corretos da planilha!")
        st.code("""Olá, {{column_4}}! Tudo bem? 👁️

Aqui é a PITER, sua assistente do Hospital de Olhos Vale do Aço.

📌 *Lembrete automático do seu atendimento:*

📋 *Atendimento:* {{column_6}}
👨‍⚕️ *Médico(a):* {{column_7}} 
📅 *Data:* {{column_1}}
⏰ *Chegada:* {{column_2}}

🚨 *PREPARO OBRIGATÓRIO:* {{column_8}}

📁 *O QUE TRAZER:* {{column_9}}

⚠️ *CONFIRMAÇÃO:*

✅ *Confirmar:* Responda *SIM* _(Sistema registra automaticamente)_

❌ *Cancelar/Remarcar:* Não responda aqui  
Temos fila de espera. Clique no link:  
👉 https://wa.me/553138011800

🤖 Mensagem automática. Sistema em atualização.  
Dúvidas? Use o link acima.

🤍 *PITER — Hospital de Olhos Vale do Aço*""", language="text")

    f_agendas = st.file_uploader("Suba as planilhas do Doctors (Consultas, Exames, Cirurgias)", type=["html", "htm", "xls", "xlsx"], accept_multiple_files=True, key="agenda_batch")
    
    if f_agendas:
        if st.button("🚀 PROCESSAR COM INTELIGÊNCIA CIRÚRGICA"):
            siglas_nao_encontradas = set()
            lista_consolidada = []
            audit_log = [] # INICIA O RADAR DE AUDITORIA AQUI!
            
            # --- LEITURA DAS PLANILHAS ---
            for f in f_agendas:
                try:
                    f.seek(0)
                    data_extraida = ""
                    match_data = re.search(r'(\d{2}[-_\.]\d{2}(?:[-_\.]\d{2,4})?)', f.name)
                    if match_data:
                        data_extraida = match_data.group(1).replace('_', '/').replace('.', '/').replace('-', '/')

                    if f.name.lower().endswith(('.html', '.htm')):
                        dfs = pd.read_html(f.read(), header=0)
                        df_conf = max(dfs, key=len).astype(str)
                    else:
                        df_conf = pd.read_excel(f, dtype=str)

                    df_conf.columns = [str(c).upper().strip() for c in df_conf.columns]
                    if 'PACIENTE' not in df_conf.columns and len(df_conf) > 0:
                        row_str = " ".join([str(x).upper() for x in df_conf.iloc[0].values])
                        if 'PACIENTE' in row_str:
                            df_conf.columns = [str(x).upper().strip() for x in df_conf.iloc[0].values]
                            df_conf = df_conf[1:]
                            
                    # BLINDAGEM 4: AIRBAG DE COLUNA FALTANTE
                    if 'PACIENTE' in df_conf.columns:
                        df_conf = df_conf.rename(columns={'PACIENTE': 'NOME'})
                    
                    colunas_obrigatorias = ['NOME', 'TELEFONE', 'HORA']
                    colunas_faltantes = [c for c in colunas_obrigatorias if c not in df_conf.columns]
                    if colunas_faltantes:
                        st.error(f"⚠️ OPA! A planilha '{f.name}' está sem a(s) coluna(s): {', '.join(colunas_faltantes)}. Por favor, marque essa coluna lá no Doctors e exporte novamente.")
                        continue

                    nome_medico_completo = resolver_nome_completo(f.name)

                    if 'NOME' in df_conf.columns:
                        # 1. CAÇA "EM ANÁLISE" E GERA LOG
                        termo_analise = r'(?i)(EM AN[AÁ]LISE|EM ANALISE)'
                        mask_analise_nome = df_conf['NOME'].astype(str).str.contains(termo_analise, na=False, regex=True)
                        col_obs = next((c for c in df_conf.columns if 'OBS' in str(c).upper()), None)
                        mask_analise_obs = pd.Series(False, index=df_conf.index)
                        if col_obs:
                            mask_analise_obs = df_conf[col_obs].astype(str).str.contains(termo_analise, na=False, regex=True)
                        mask_analise_total = mask_analise_nome | mask_analise_obs
                        df_em_analise = df_conf[mask_analise_total]
                        
                        if not df_em_analise.empty:
                            for _, row_analise in df_em_analise.iterrows():
                                nome_analise = str(row_analise['NOME']).strip().upper()
                                # Log da Aba 5
                                st.session_state['lista_pendencias'].append({
                                    "Data": datetime.now().strftime("%d/%m"), 
                                    "Paciente": nome_analise, 
                                    "Motivo": "⚠️ GUIA EM ANÁLISE", 
                                    "IA": "Sistema"
                                })
                                # Log da Auditoria (Aba 7)
                                audit_log.append({
                                    "DATA": data_extraida, "NOME": nome_analise, 
                                    "TELEFONE": row_analise.get('TELEFONE', ''), 
                                    "MÉDICO DA AGENDA": nome_medico_completo, 
                                    "MOTIVO DO CORTE": "⚠️ GUIA EM ANÁLISE"
                                })
                            st.toast(f"⚠️ {len(df_em_analise)} paciente(s) com Guia em Análise interceptado(s) e movido(s) para a Aba 5!", icon="🚨")

                        # 3. GUILHOTINA DE CANCELADOS
                        termos_cancel = r'(?i)(CANCELADO|CANCELADA|DESMARCOU|DESMARCADO|DESISTIU|FALTOU)'
                        mask_cancel_nome = df_conf['NOME'].astype(str).str.contains(termos_cancel, na=False, regex=True)
                        df_cancelados = df_conf[mask_cancel_nome]
                        
                        for _, r_canc in df_cancelados.iterrows():
                            audit_log.append({
                                "DATA": data_extraida, "NOME": str(r_canc['NOME']).strip().upper(), 
                                "TELEFONE": r_canc.get('TELEFONE', ''), 
                                "MÉDICO DA AGENDA": nome_medico_completo, 
                                "MOTIVO DO CORTE": "❌ CANCELADO/FALTOU NO DOCTOR'S"
                            })

                        # Corta efetivamente os Em analise e Cancelados da base
                        df_conf = df_conf[~(mask_analise_total | mask_cancel_nome)]
                        df_conf = df_conf.dropna(subset=['NOME'])
                        
                        # SALVA O TEXTO BRUTO ANTES DA GUILHOTINA CORTAR!
                        df_conf['OBS_EXTRAIDA'] = df_conf['NOME'].astype(str)
                        
                        padrao_corte = r'(?i)(#|-|,|/|\(|MOTIVO|IDADE|CPF|\bTEM\b|\bCONFIRMAD[OA]S?\b|\bGUIA\b|\bSALV[OA]\b|\bREMARCAD[OA]\b|\bAV\b|\bOD\b|\bOE\b|\bAO\b|\bJUSTIFICATIVA\b|\bPEGAR\b|\bPEDIDO\b|\bDOC\b|\bPACIENTE\b|\bSOLICITADO\b|\bJUNT[OA] COM\b|Nº|N°|\d|\?|<|>|\bJ[AÁ] EST[AÁ]\b|\bEST[AÁ] PAGO\b|\bAPRESENTAR\b|\bA MESMA\b|\bESPOSA\b|\bMARIDO\b|\bPAI\b|\bM[AÃ]E\b|\bFILH[OA]S?\b|\bIRM[AÃ]O?\b|\bNET[OA]S?\b|\bTIO\b|\bTIA\b|\bAV[OÓÔ]\b|\bSOGRA\b|\bSOGRO\b|\bACOMPANHANTE\b|\bLIGAÇ[AÃ]O\b|\bATEND\.?|\bTBM\b|\bTAMB[EÉ]M\b|\bRECADO\b|\bREALIZA\b|\bHEMODI[AÁ]LISE\b|\bESTAVAM?\b|\bCADASTRO\b|\bRELATA\b|\bVEIO\b|\bJUSTIFICAR\b|\bC/\b|\bCOM\b|\bOU\b|\bOBS\b|\bDRA?\b)'
                        df_conf['NOME'] = df_conf['NOME'].apply(lambda x: re.split(padrao_corte, str(x).replace('*', ''))[0].strip() if pd.notnull(x) else "")
                    if 'TELEFONE' in df_conf.columns:
                        df_conf = df_conf[df_conf['TELEFONE'].str.strip() != '-']

                    col_tipo = next((c for c in df_conf.columns if 'TIPO' in c), None)
                    df_conf['TIPO_RAW'] = df_conf[col_tipo] if col_tipo else 'C' 

                    for index, row in df_conf.iterrows():
                        tipo_raw = str(row.get('TIPO_RAW', '')).strip().upper()
                        traducao = traduzir_tipo(tipo_raw, st.session_state['dicionario_tipos'])
                        if traducao == "DESCONHECIDO" and tipo_raw != "":
                            siglas_nao_encontradas.add(tipo_raw)
                        df_conf.at[index, 'TIPO_FINAL'] = traducao

                    df_conf['MEDICO_ORIGINAL'] = nome_medico_completo
                    df_conf['DATA'] = data_extraida
                    
                    lista_consolidada.append(df_conf)
                
                except Exception as e:
                    st.error(f"Erro no arquivo {f.name}: {e}")

            if siglas_nao_encontradas:
                st.error(f"🛑 ATENÇÃO! O robô encontrou siglas na planilha que não estão no Dicionário: **{', '.join(siglas_nao_encontradas)}**")
                st.warning("👉 Suba a tela, abra o 'DICIONÁRIO DE ABREVIAÇÕES', adicione essas siglas, aperte Enter para salvar e clique no botão Processar de novo.")
            
            elif lista_consolidada:
                df_full = pd.concat(lista_consolidada, ignore_index=True)
                df_full['NOME'] = df_full['NOME'].str.strip().str.upper()
                df_full['TELEFONE_LIMPO'] = df_full['TELEFONE'].apply(limpar_num)
                
                # BLINDAGEM: TELEFONES VAZIOS E FALSOS (COM AUDITORIA)
                mask_bad_phone = (df_full['TELEFONE_LIMPO'] == '') | df_full['TELEFONE_LIMPO'].apply(eh_numero_falso)
                df_bad_phone = df_full[mask_bad_phone]
                for _, r_bad in df_bad_phone.iterrows():
                    audit_log.append({
                        "DATA": r_bad.get('DATA', ''), "NOME": r_bad['NOME'], 
                        "TELEFONE": r_bad.get('TELEFONE', ''), 
                        "MÉDICO DA AGENDA": r_bad.get('MEDICO_ORIGINAL', ''), 
                        "MOTIVO DO CORTE": "🚫 SEM TELEFONE OU NÚMERO FALSO"
                    })
                df_full = df_full[~mask_bad_phone]
                
                # BLINDAGEM: DUPLICADOS (COM AUDITORIA)
                mask_dupes = df_full.duplicated(subset=['TELEFONE_LIMPO', 'NOME', 'DATA', 'HORA', 'TIPO_FINAL', 'MEDICO_ORIGINAL'], keep='first')
                df_dupes = df_full[mask_dupes]
                for _, r_dup in df_dupes.iterrows():
                    audit_log.append({
                        "DATA": r_dup.get('DATA', ''), "NOME": r_dup['NOME'], 
                        "TELEFONE": r_dup.get('TELEFONE', ''), 
                        "MÉDICO DA AGENDA": r_dup.get('MEDICO_ORIGINAL', ''), 
                        "MOTIVO DO CORTE": "♻️ DUPLICIDADE EXATA (MESMO EXAME/HORA)"
                    })
                df_full = df_full[~mask_dupes]
                
                resultados_finais = []

                if 'HORA' in df_full.columns and 'NOME' in df_full.columns:
                    
                    for (telefone, nome_pct, data_pct), group in df_full.groupby(['TELEFONE_LIMPO', 'NOME', 'DATA']):
                        
                        # --- ÂNCORA DE HORÁRIO INTACTA ---
                        group['HORA_TEMP'] = pd.to_datetime(group['HORA'], format='%H:%M', errors='coerce')
                        group = group.sort_values(by='HORA_TEMP')
                        menor_hora_do_dia = group.iloc[0]['HORA']
                        
                        group_medicos_validos = group[group['MEDICO_ORIGINAL'].isin(agendas_enviar)]
                        if group_medicos_validos.empty:
                            # Auditoria: Médico não habilitado
                            for _, r_med in group.iterrows():
                                audit_log.append({
                                    "DATA": data_pct, "NOME": nome_pct, "TELEFONE": r_med.get('TELEFONE',''), 
                                    "MÉDICO DA AGENDA": r_med.get('MEDICO_ORIGINAL',''), 
                                    "MOTIVO DO CORTE": "👨‍⚕️ MÉDICO/SETOR NÃO HABILITADO PARA ENVIO"
                                })
                            continue

                        # --- FLAGS DO MOTOR DE PREPARO ---
                        f_acompanhante_grupo1 = False
                        f_acompanhante_grupo2 = False
                        
                        f_jejum_absoluto = False
                        f_jejum_observacao = False
                        f_isento_jejum = False
                        
                        f_colirio_24h = False
                        f_colirio_48h = False
                        f_colirio_mydriacyl_1h = False
                        f_traz_colirio = False
                        
                        f_lente_24h = False
                        f_lente_72h = False
                        f_lente_7d = False
                        
                        f_traz_oculos = False
                        f_traz_estojo = False
                        
                        f_banho_refra = False
                        f_duracao_agf = False
                        f_duracao_cdpo = False
                        
                        f_docs_cirurgia = False
                        
                        f_is_ic = False
                        f_ic_com_exames = False
                        
                        f_ilio_oculto = False
                        f_pre_op_catarata = False



                        tem_valido = False
                        tem_cra = False
                        tem_cirurgia = False
                        tipo_cirurgia = "GERAL"
                        medico_cirurgia = ""
                        tipos_para_nome = set()
                        
                        # Pegando TODAS as observações do paciente no dia de forma segura
                        col_obs_nome = 'OBSERVAÇÃO' if 'OBSERVAÇÃO' in group.columns else ('OBS' if 'OBS' in group.columns else None)
                        observacao_do_dia = " ".join(group[col_obs_nome].astype(str)).upper() if col_obs_nome else " "
                        
                        for _, row in group_medicos_validos.iterrows():
                            tipo_original = str(row['TIPO_RAW']).upper().replace('_', ' ')
                            tipo_final = str(row['TIPO_FINAL'])
                            medico = str(row['MEDICO_ORIGINAL']).upper()
                            
                            obs_isolada = str(row.get('OBSERVAÇÃO', row.get('OBSERVACAO', row.get('OBS', '')))).upper()
                            obs_escondida = str(row.get('OBS_EXTRAIDA', '')).upper()
                            
                            # O robô só lê as observações se for Agenda de Cirurgia ou Enfermagem!
                            if medico == 'CENTRO CIRÚRGICO' or 'CIRURGIA' in tipo_original or 'CI ' in tipo_original or 'CI_' in tipo_original or medico == 'ENFERMAGEM':
                                tipo_raw = f"{tipo_original} {obs_isolada} {obs_escondida}"
                            else:
                                tipo_raw = tipo_original
                                
                            if medico == 'CATARATA' or 'PRED P' in tipo_raw:
                                f_pre_op_catarata = True
                                
                            partes = tipo_raw.replace('/', ' ').split()
                            
                            # --- 1. TRATAMENTO DO ILIO (Oculta nome, mantém hora) ---
                            if tipo_raw == 'ILIO':
                                f_ilio_oculto = True
                                continue 

                            # --- 2. TRATAMENTO DA IC (Instrução Cirúrgica) ---
                            is_ic_current_row = False
                            if medico == 'ENFERMAGEM' or tipo_raw == 'IC' or tipo_raw.startswith(('IC ', 'IC_', 'IC.')):
                                f_is_ic = True
                                is_ic_current_row = True
                                if any(x in tipo_raw or x in observacao_do_dia for x in ['ANTIVEG', 'ANTIVEGF', 'TTO', 'INTRAVITREA']):
                                    f_isento_jejum = True
                                    tipo_cirurgia = 'INTRAVITREA'
                            
                            # --- NOVA GUILHOTINA DE AVALIAÇÕES E RETORNOS ---
                            # Se a tradução final for uma dessas, o robô ignora e não manda mensagem!
                            traducoes_fantasmas = [
                                "Retorno", "Consulta Reavaliação Anual", "Entrega de Exames",
                                "1ª Avaliação", "2ª Avaliação", "3ª Avaliação", "Avaliação de Glaucoma"
                            ]
                            
                            is_fantasma = tipo_final in traducoes_fantasmas
                            
                            if not is_fantasma: 
                                tem_valido = True
                                tipos_para_nome.add(tipo_final)
                                
                            if tipo_final == "Consulta Reavaliação Anual" or "Retorno" in tipo_final: 
                                tem_cra = True

                            # --- REGRAS DE EXAMES E CIRURGIAS ---
                            if not is_ic_current_row:
                                # Agora o robô procura a sigla solta no meio da "Super Frase"
                                if any(p in ['AGF', 'RF', 'FOTO', 'YAG'] for p in partes): 
                                    f_acompanhante_grupo1 = True
                                if any(p in ['MR', 'RD', 'OCT', 'ANGIO-OCT'] for p in partes) or any(p in exames_dilatam for p in partes): 
                                    f_acompanhante_grupo2 = True
                                    f_traz_estojo = True

                                if any(p in ['TOPO', 'PTC', 'PAQ'] for p in partes): f_lente_72h = True
                                if any(p in ['AV', 'CVC'] for p in partes): f_traz_oculos = True
                                if any(p in ['AV', 'MR'] for p in partes): f_traz_estojo = True
                                if 'AGF' in partes: f_duracao_agf = True
                                if 'CDPO' in partes: f_duracao_cdpo = True
                                if 'RC' in partes: f_traz_colirio = True

                               # O robô só decreta cirurgia se a sigla estiver no TIPO oficial, ignorando a palavra "cirurgia" solta nas observações!
                                if 'CIRURGIA' in tipo_original or 'CI ' in tipo_original or 'CI_' in tipo_original or 'CI.' in tipo_original or medico == 'CENTRO CIRÚRGICO':
                                    tem_cirurgia = True
                                    f_acompanhante_grupo1 = True 
                                    f_docs_cirurgia = True
                                    
                                    if "Cirurgia Dr(a)." in tipo_final:
                                        medico_cirurgia = tipo_final.replace("Cirurgia ", "").strip()
                                    else:
                                        medico_cirurgia = medico

                                    if 'FACO' in tipo_raw or 'CATARATA' in tipo_raw: 
                                        tipo_cirurgia = 'FACO'
                                        f_colirio_24h = True
                                        if "JEJUM" in observacao_do_dia: f_jejum_observacao = True
                                    elif 'VITRE' in tipo_raw or 'RETINA' in tipo_raw or 'VVPP' in tipo_raw: 
                                        tipo_cirurgia = 'RETINA'
                                        f_jejum_absoluto = True
                                        f_colirio_mydriacyl_1h = True
                                    elif 'TRAB' in tipo_raw or 'CICLOFOTO' in tipo_raw or 'GLAUCOMA' in tipo_raw:
                                        tipo_cirurgia = 'GLAUCOMA'
                                        f_jejum_absoluto = True
                                    elif 'ESTRABISMO' in tipo_raw:
                                        tipo_cirurgia = 'ESTRABISMO'
                                        f_jejum_absoluto = True
                                    elif 'ANTIVEG' in tipo_raw or 'INTRAVITREA' in tipo_raw: 
                                        tipo_cirurgia = 'INTRAVITREA'
                                        f_isento_jejum = True
                                    elif 'ANEL' in tipo_raw: 
                                        tipo_cirurgia = 'ANEL'
                                        f_colirio_48h = True
                                        f_lente_72h = True
                                        if "JEJUM" in observacao_do_dia: f_jejum_observacao = True
                                    elif 'PRK' in tipo_raw or 'LASIK' in tipo_raw or 'REFRATIVA' in tipo_raw: 
                                        tipo_cirurgia = 'REFRATIVA'
                                        f_banho_refra = True
                                        f_lente_7d = True
                                        if "JEJUM" in observacao_do_dia: f_jejum_observacao = True
                                    elif 'TRANSPLANTE' in tipo_raw or 'CORNEA' in tipo_raw: 
                                        tipo_cirurgia = 'TRANSPLANTE'
                                        f_jejum_absoluto = True
                                        f_lente_24h = True
                                    elif 'PTERIGIO' in tipo_raw or 'CALAZIO' in tipo_raw or 'TUMOR' in tipo_raw or 'LP' in tipo_raw or 'TTO' in tipo_raw: 
                                        tipo_cirurgia = 'SUPERFICIAL'
                                        if "JEJUM" in observacao_do_dia: f_jejum_observacao = True

                                    if tipo_cirurgia == 'GERAL':
                                        if 'MARILUCE' in medico_cirurgia.upper():
                                            tipo_cirurgia = 'TRANSPLANTE'
                                            f_jejum_absoluto = True
                                            f_lente_24h = True
                                        else:
                                            if "JEJUM" in observacao_do_dia: f_jejum_observacao = True

                                    # 🚨 REGRA DO DR. GABRIEL LEMOS 🚨
                                    if 'GABRIEL L' in medico_cirurgia.upper():
                                        # Ele só pede jejum se o procedimento NÃO for isento por natureza
                                        if not f_isento_jejum:
                                            f_jejum_absoluto = True
                                            f_jejum_observacao = False

                        if not tem_valido:
                            # Auditoria para o paciente que tem APENAS um procedimento invisível
                            motivo_corte = "👻 PROCEDIMENTO ISOLADO (Apenas ILIO). Verificar manualmente." if f_ilio_oculto else "👻 PROCEDIMENTO INVISÍVEL (Ex: Apenas Retorno ou CAT)"
                            audit_log.append({
                                "DATA": data_pct, "NOME": nome_pct, "TELEFONE": telefone, 
                                "MÉDICO DA AGENDA": group_medicos_validos.iloc[0]['MEDICO_ORIGINAL'], 
                                "MOTIVO DO CORTE": motivo_corte
                            })
                            continue

                        # --- O CARRINHO DE COMPRAS DE REGRAS (MONTAGEM DO PREPARO) ---
                        preparo_lista = []
                        
                        if f_pre_op_catarata: categoria_lote = "4_PRE_OP_CATARATA"
                        elif f_is_ic: categoria_lote = "1_ENFERMAGEM_E_IC"
                        elif tem_cirurgia: categoria_lote = "1_CIRURGIAS"
                        elif f_acompanhante_grupo2 or f_duracao_agf: categoria_lote = "2_EXAMES_QUE_DILATAM"
                        else: categoria_lote = "3_CONSULTAS_NORMAIS"

                        # Regra específica da IC (Enfermagem)
                        if f_is_ic:
                            preparo_lista.append("⚠️ ATENÇÃO: O paciente não precisa comparecer presencialmente para esta instrução, pode ser algum familiar ou responsável pelo mesmo.")
                            preparo_lista.append("💰 SOBRE O PAGAMENTO: Se o seu atendimento for PARTICULAR, CONVÊNIO DE PAGAMENTO À VISTA ou NÃO AUTORIZADO PELO PLANO, o pagamento deverá ser feito obrigatoriamente nesta data de atendimento.")

                       # --- SEÇÃO DE ALIMENTAÇÃO E JEJUM (REGRAS BLINDADAS) ---
                        if f_isento_jejum:
                            preparo_lista.append("✔️ Alimentação: NÃO é necessário fazer jejum. Pode se alimentar normalmente.")
                        
                        elif f_duracao_agf:
                            preparo_lista.append("✔️ Alimentação: Não há necessidade de jejum absoluto. É recomendável apenas uma alimentação mais leve 2h antes do exame.")
                        
                        elif (f_jejum_absoluto or f_jejum_observacao):
                            preparo_lista.append("✔️ Jejum Absoluto: É necessário fazer jejum de 08 horas. Não coma nem beba nada (incluindo água) nas 8 horas antes do procedimento.")

                        # Acompanhante e Dilatação / Medicamentos da AGF
                        if f_duracao_agf:
                            preparo_lista.append("✔️ Medicamentos: Pacientes em uso de medicamentos para pressão alta (entre outros) devem tomar a medicação normalmente.")
                            preparo_lista.append("✔️ Cuidados (Diabéticos): Se você é diabético, lembre-se de trazer seu lanche para não ocorrer hipoglicemia.")
                            preparo_lista.append("✔️ Duração: Exame demorado. Venha com disponibilidade de horário.")
                            preparo_lista.append("✔️ Acompanhante: Presença OBRIGATÓRIA de acompanhante adulto.")
                        else:
                            if f_acompanhante_grupo1: 
                                preparo_lista.append("✔️ Acompanhante: Presença OBRIGATÓRIA de 1 acompanhante adulto. Você não poderá dirigir após o procedimento.")
                            elif f_acompanhante_grupo2:
                                preparo_lista.append("✔️ Dilatação da Pupila: Seu atendimento exige a dilatação da pupila, o que causará embaçamento visual temporário. Recomendamos que você evite dirigir logo após o procedimento.")
                                preparo_lista.append("✔️ Acompanhante: A presença de acompanhante é OPCIONAL caso você se sinta inseguro para retornar devido ao embaçamento visual (permitido APENAS 1 acompanhante). ATENÇÃO: Para pacientes menores de 18 anos, o acompanhante é OBRIGATÓRIO.")

                        # Colírios
                        if f_colirio_24h: preparo_lista.append("✔️ Colírio: Iniciar o uso do colírio prescrito 1 dia antes da cirurgia.")
                        if f_colirio_48h: preparo_lista.append("✔️ Colírio: Iniciar o uso 48h antes da cirurgia, conforme a receita.")
                        if f_colirio_mydriacyl_1h: preparo_lista.append("✔️ Colírio: Iniciar o uso do colírio Mydriacyl 1 hora antes da cirurgia, conforme a receita médica.")
                        if f_traz_colirio: preparo_lista.append("✔️ Atenção: Por favor, traga todos os colírios que você está usando no tratamento.")
                        
                        # Lentes e Óculos
                        if f_lente_24h: preparo_lista.append("✔️ Lentes: Usuários de Lentes de Contato devem suspender o uso 1 dia (24h) antes do procedimento.")
                        if f_lente_72h: preparo_lista.append("✔️ Lentes: Usuários de Lentes de Contato devem suspender o uso 72h (3 dias) antes do procedimento.")
                        if f_lente_7d: preparo_lista.append("✔️ Lentes: Usuários de Lentes de Contato devem suspender o uso 7 dias antes do procedimento.")
                        if f_traz_oculos: preparo_lista.append("✔️ Óculos: É OBRIGATÓRIO trazer os óculos e a receita médica no dia do exame.")
                        if f_traz_estojo: preparo_lista.append("✔️ Lentes: Se você usa Lentes de Contato, traga o seu estojo para retirar as lentes aqui no hospital.")
                        
                        # Cuidados Especiais
                        if f_banho_refra: preparo_lista.append("✔️ Higiene: Tomar banho e lavar muito bem a cabeça e o rosto. Não usar maquiagem, cremes, perfume ou produtos no cabelo.")
                        if f_duracao_cdpo: 
                            preparo_lista.append("✔️ Atenção à Duração: Você ficará no hospital o dia todo. Medições da pressão ocular às 08:00, 11:00 e 13:30.")
                            preparo_lista.append("✔️ Preparo: A Curva de Pressão não exige dilatação da pupila.")

                        if tem_cirurgia and not f_banho_refra:
                            preparo_lista.append("✔️ Cuidados: Não use maquiagem ou cremes no rosto. Deixe objetos de valor em casa. Vista roupas confortáveis.")
                            
                        if f_ilio_oculto:
                            preparo_lista.append("✔️ Atenção: Siga rigorosamente as orientações de preparo que foram repassadas pela nossa equipe no momento do seu agendamento.")

                        if not preparo_lista:
                            preparo_lista.append("✔️ Preparo: A dilatação da pupila será realizada conforme necessidade e protocolo médico.")
                            preparo_lista.append("✔️ Chegada: Por favor, chegue com 15 minutos de antecedência para a realização do seu cadastro.")

                        preparo = "\n".join(preparo_lista)
                        
                        nome_paciente = str(group_medicos_validos.iloc[0]['NOME']).upper()
                        tel_valido = group_medicos_validos.iloc[0]['TELEFONE']
                        data_final = group_medicos_validos.iloc[0].get('DATA', '')
                        
                        medico_base = group_medicos_validos.iloc[-1]['MEDICO_ORIGINAL'] if tem_cirurgia else group_medicos_validos.iloc[0]['MEDICO_ORIGINAL']
                        if medico_base not in ['Centro Cirúrgico', 'Setor de Lentes de Contato', 'Exames', 'Glaucoma', 'Enfermagem']:
                            medico_formatado = f"Dr(a). {medico_base}"
                        else:
                            medico_formatado = medico_base

                        lista_nomes_reais = sorted(list(tipos_para_nome))
                        tipo_final_saida = lista_nomes_reais[0] if lista_nomes_reais else "Atendimento Especializado"
                        
                        if len(group_medicos_validos) > 1:
                            if tem_cirurgia:
                                nomes_cirurgia = {'FACO': 'Cirurgia de Catarata (FACO)', 'ANEL': 'Implante de Anel', 'RETINA': 'Cirurgia de Retina', 'INTRAVITREA': 'Aplicação Intravítrea', 'REFRATIVA': 'Cirurgia Refrativa', 'TRANSPLANTE': 'Transplante de Córnea', 'SUPERFICIAL': 'Procedimento Cirúrgico (Pterígio/Calázio/Tumor)', 'GLAUCOMA': 'Cirurgia de Glaucoma', 'ESTRABISMO': 'Cirurgia de Estrabismo', 'GERAL': 'Procedimento Cirúrgico'}
                                base_nome = nomes_cirurgia.get(tipo_cirurgia, 'Procedimento Cirúrgico')
                                tipo_final_saida = "Exames Prévios + " + base_nome
                            elif tem_cra:
                                outros_tipos = [t for t in lista_nomes_reais if "Reavaliação" not in t and "Retorno" not in t]
                                if outros_tipos:
                                    tipo_final_saida = "Reavaliação Anual + " + " + ".join(outros_tipos)
                                else:
                                    tipo_final_saida = "Reavaliação Anual + Exames"
                            else:
                                tipo_final_saida = " + ".join(lista_nomes_reais)
                        elif tem_cirurgia:
                            nomes_cirurgia = {'FACO': 'Cirurgia de Catarata (FACO)', 'ANEL': 'Implante de Anel', 'RETINA': 'Cirurgia de Retina', 'INTRAVITREA': 'Aplicação Intravítrea', 'REFRATIVA': 'Cirurgia Refrativa', 'TRANSPLANTE': 'Transplante de Córnea', 'SUPERFICIAL': 'Procedimento Cirúrgico (Pterígio/Calázio/Tumor)', 'GLAUCOMA': 'Cirurgia de Glaucoma', 'ESTRABISMO': 'Cirurgia de Estrabismo', 'GERAL': 'Procedimento Cirúrgico'}
                            tipo_final_saida = nomes_cirurgia.get(tipo_cirurgia, 'Procedimento Cirúrgico')

                        dr_escondido = None
                        if tem_cirurgia:
                            for t in tipos_para_nome:
                                if "Cirurgia Dr(a)." in str(t):
                                    dr_escondido = str(t).replace("Cirurgia ", "").strip()
                                    break

                        if tem_cirurgia and dr_escondido:
                            medico_formatado = dr_escondido

                        # --- DOCUMENTOS DINÂMICOS ---
                        docs = ["✔️ Documento de Identidade com foto (RG, CNH ou outro).", "✔️ Carteirinha do convênio (física ou aplicativo no celular)."]
                        if f_is_ic:
                            if f_ic_com_exames:
                                docs.append("✔️ Resultados de Exames de Laboratório, Eletrocardiograma (ECG) e Laudo de Risco Cirúrgico.")
                        elif tem_cirurgia:
                            docs.append("✔️ Pedido Médico.")
                            if f_docs_cirurgia:
                                docs.append("✔️ Termo de Consentimento assinado por extenso em TODAS as folhas (para entregar à equipe de enfermagem).")
                        else:
                            docs.append("✔️ Pedido Médico (se houver).")

                        documentos = "\n".join(docs)
                        
                        if f_pre_op_catarata:
                            tipo_final_saida = "Pré-Operatório para Cirurgia de Catarata"
                            preparo = "⏱️ IMPORTANTE: Venha com disponibilidade de tempo — o pré-operatório costuma ser demorado."
                            documentos = "✔️ Cartão do SUS\n✔️ Documento de Identificação com foto"
                            medico_formatado = "Equipe de Catarata"

                        resultados_finais.append({
                            'CATEGORIA': categoria_lote,
                            'DATA': data_final,
                            'HORA': menor_hora_do_dia,
                            'NOME': nome_paciente,  
                            'NOME_CURTO': pegar_nome_curto(nome_paciente),
                            'TELEFONE': tel_valido,
                            'TIPO': tipo_final_saida,
                            'MEDICO': medico_formatado,
                            'PREPARO': preparo,
                            'DOCUMENTOS': documentos
                        })

                df_limpo = pd.DataFrame(resultados_finais)

                # --- GERADOR DO RELATÓRIO DE AUDITORIA ---
                if audit_log:
                    df_audit = pd.DataFrame(audit_log)
                    b_audit = io.BytesIO()
                    with pd.ExcelWriter(b_audit, engine='openpyxl') as writer:
                        df_audit.to_excel(writer, index=False, sheet_name='Pacientes Cortados')
                        ws_audit = writer.sheets['Pacientes Cortados']
                        for col in ws_audit.columns:
                            ws_audit.column_dimensions[col[0].column_letter].width = 30
                    st.session_state['audit_report'] = b_audit.getvalue()
                    st.session_state['audit_count'] = len(df_audit)

                if not df_limpo.empty:
                    for col in df_limpo.columns:
                        df_limpo[col] = df_limpo[col].astype(str).replace(r'[;\n]', ' ', regex=True).str.strip()
                        df_limpo[col] = df_limpo[col].apply(lambda x: "" if x.lower() == "nan" else x)
                    
                    df_limpo['PREPARO'] = df_limpo['PREPARO'].str.replace('✔️', '\n✔️').str.strip()
                    df_limpo['DOCUMENTOS'] = df_limpo['DOCUMENTOS'].str.replace('✔️', '\n✔️').str.strip()

                    b_zip = io.BytesIO()
                    with zipfile.ZipFile(b_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                        categorias = df_limpo['CATEGORIA'].unique()
                        
                        for cat in sorted(categorias):
                            df_cat = df_limpo[df_limpo['CATEGORIA'] == cat].drop(columns=['CATEGORIA'])
                            qtd_lotes = math.ceil(len(df_cat) / 45)
                            
                            for i in range(qtd_lotes):
                                df_lote = df_cat.iloc[i * 45 : (i + 1) * 45].copy()
                                if not df_lote.empty:
                                    df_lote = pd.concat([df_lote.iloc[[0]], df_lote], ignore_index=True)
                                sufixo = f"_PT{i+1}" if qtd_lotes > 1 else ""
                                csv_agenda = io.StringIO()
                                df_lote.to_csv(csv_agenda, index=False, sep=';', quoting=csv.QUOTE_MINIMAL)
                                zf.writestr(f"{cat}{sufixo}.csv", csv_agenda.getvalue().encode('utf-8-sig'))

                    b_zip.seek(0)
                    st.session_state['zip_tab7'] = b_zip.getvalue()
                    
                    st.success(f"✅ Master System Atualizado! Todas as Regras e Blindagens (IC/ILIO/CI) Operando com Sucesso.")
                else:
                    st.warning("⚠️ Não sobrou nenhum paciente válido após aplicar os filtros da gerência.")

        # --- BOTÕES DE DOWNLOAD (ZAPROCKET E AUDITORIA) ---
        col_dw1, col_dw2 = st.columns(2)
        if 'zip_tab7' in st.session_state:
            col_dw1.download_button("📥 1. BAIXAR PACOTES (ZAPROCKET)", st.session_state['zip_tab7'], "agendas_hova_master.zip", mime="application/zip", use_container_width=True)
            
        if 'audit_report' in st.session_state:
            col_dw2.download_button(f"🚨 2. VER {st.session_state['audit_count']} PACIENTES CORTADOS", st.session_state['audit_report'], "Relatorio_Auditoria_Cortados.xlsx", type="primary", use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

with tab8:
    st.markdown('<div class="master-card" style="border: 2px solid #ff4b4b;">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title" style="color: #ff4b4b;"> SALVA-VIDAS (CORRETOR DE ERROS)</div>""", unsafe_allow_html=True)
    st.markdown("Use esta aba para comparar as planilhas que você **já enviou (incompletas)** com a planilha **correta (gerada na Aba 7)**.")
    
    col_err1, col_err2 = st.columns(2)
    f_errada = col_err1.file_uploader("1. Planilhas INCOMPLETAS (As que você já disparou no ZapRocket)", type=["csv"], accept_multiple_files=True, key="salva_errada")
    f_correta = col_err2.file_uploader("2. Planilha CORRETA (Gerada agora com TODAS as agendas)", type=["csv"], accept_multiple_files=True, key="salva_correta")
    
    if f_errada and f_correta:
        if st.button("🚨 CRUZAR DADOS E DESCOBRIR ERROS", type="primary"):
            try:
                lista_err = []
                for f in f_errada:
                    f.seek(0)
                    df_t = pd.read_csv(f, sep=';', dtype=str)
                    if len(df_t.columns) < 2: f.seek(0); df_t = pd.read_csv(f, sep=',', dtype=str)
                    lista_err.append(df_t)
                df_env = pd.concat(lista_err, ignore_index=True)
                df_env.columns = [str(c).upper().strip() for c in df_env.columns]
                
                lista_cor = []
                for f in f_correta:
                    f.seek(0)
                    df_c = pd.read_csv(f, sep=';', dtype=str)
                    if len(df_c.columns) < 2: f.seek(0); df_c = pd.read_csv(f, sep=',', dtype=str)
                    lista_cor.append(df_c)
                df_cor = pd.concat(lista_cor, ignore_index=True)
                df_cor.columns = [str(c).upper().strip() for c in df_cor.columns]
                
                if 'NOME' not in df_env.columns or 'NOME' not in df_cor.columns:
                    st.error("As planilhas precisam ter a coluna NOME para o cruzamento funcionar.")
                else:
                    df_env['NOME_CHAVE'] = df_env['NOME'].astype(str).str.strip().str.upper()
                    df_cor['NOME_CHAVE'] = df_cor['NOME'].astype(str).str.strip().str.upper()
                    
                    nomes_enviados = set(df_env['NOME_CHAVE'].dropna().tolist())
                    df_faltaram = df_cor[~df_cor['NOME_CHAVE'].isin(nomes_enviados)].copy()
                    if 'NOME_CHAVE' in df_faltaram.columns: df_faltaram = df_faltaram.drop(columns=['NOME_CHAVE'])
                    
                    if 'HORA' in df_cor.columns and 'HORA' in df_env.columns:
                        df_intersection = pd.merge(df_cor, df_env[['NOME_CHAVE', 'HORA']], on='NOME_CHAVE', suffixes=('_CORRETA', '_ENVIADA'))
                        df_mudaram = df_intersection[df_intersection['HORA_CORRETA'] != df_intersection['HORA_ENVIADA']].copy()
                        df_mudaram = df_mudaram.rename(columns={'HORA_CORRETA': 'HORA'})
                        cols_to_keep = [c for c in df_cor.columns if c != 'NOME_CHAVE']
                        df_mudaram = df_mudaram[cols_to_keep]
                    else:
                        df_mudaram = pd.DataFrame()

                    b_zip_salva = io.BytesIO()
                    with zipfile.ZipFile(b_zip_salva, "w", zipfile.ZIP_DEFLATED) as zf:
                        if not df_faltaram.empty:
                            qtd_lotes_f = math.ceil(len(df_faltaram) / 45)
                            for i in range(qtd_lotes_f):
                                df_lote_f = df_faltaram.iloc[i * 45 : (i + 1) * 45].copy()
                                sufixo = f"_PT{i+1}" if qtd_lotes_f > 1 else ""
                                csv_f = io.StringIO()
                                df_lote_f.to_csv(csv_f, index=False, sep=';', quoting=csv.QUOTE_MINIMAL)
                                zf.writestr(f"1_FALTARAM_ENVIAR{sufixo}.csv", csv_f.getvalue().encode('utf-8-sig'))
                        
                        if not df_mudaram.empty:
                            qtd_lotes_m = math.ceil(len(df_mudaram) / 45)
                            for i in range(qtd_lotes_m):
                                df_lote_m = df_mudaram.iloc[i * 45 : (i + 1) * 45].copy()
                                sufixo = f"_PT{i+1}" if qtd_lotes_m > 1 else ""
                                csv_m = io.StringIO()
                                df_lote_m.to_csv(csv_m, index=False, sep=';', quoting=csv.QUOTE_MINIMAL)
                                zf.writestr(f"2_CORRECAO_DE_HORARIO{sufixo}.csv", csv_m.getvalue().encode('utf-8-sig'))
                            
                    b_zip_salva.seek(0)
                    
                    st.success("✅ Varredura concluída com sucesso!")
                    st.warning(f"**Identificamos:**\n- {len(df_faltaram)} pacientes que NÃO receberam mensagem.\n- {len(df_mudaram)} pacientes que receberam o HORÁRIO ERRADO.")
                    
                    st.session_state['base_salva_vidas'] = b_zip_salva.getvalue()

            except Exception as e:
                st.error(f"Erro ao cruzar as planilhas: {e}")

    if 'base_salva_vidas' in st.session_state:
        st.download_button("🚑 BAIXAR PLANILHAS DE CORREÇÃO", st.session_state['base_salva_vidas'], "HOVA_SALVA_VIDAS.zip", mime="application/zip", type="primary")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# --- ABA 9: CENTRAL DE LIMPEZA E UNIFICAÇÃO (NOVO MOTOR + DESEMPATE) ---
# ==========================================
with tab9:
    st.markdown('<div class="master-card">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title"> CENTRAL DE LIMPEZA E UNIFICAÇÃO</div>""", unsafe_allow_html=True)
    
    st.markdown("### ⚙️ ESCOLHA O MOTOR DE PROCESSAMENTO:")
    modo_aba10 = st.radio(
        "Selecione o tipo de limpeza:",
        ["🧹 MODO 1: Limpeza Padrão WhatsApp (Ímã de Dados, Lavar Mensagens, Formatar Telefones)", 
         "🤝 MODO 2: Unificador Mensal (Fundir linhas vazias, mantendo Data mais recente e Telefones limpos)"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    
    if "MODO 2" in modo_aba10:
        st.info("💡 **Aviso:** Neste modo, o robô vai juntar o Paciente X que está duplicado, aplicar o **Novo Motor de Telefonia** (coluna TEL. ADIC.), formatar a conduta e usar o **Desempate de Prazo**!")
        
        col_n2, col_name2, col_d2 = st.columns(3)
        nm_num2 = col_n2.text_input("Nome da coluna NUM/Prontuário:", "NUM", key="m2_n")
        nm_nome2 = col_name2.text_input("Nome da coluna NOME:", "PACIENTE", key="m2_name")
        nm_data2 = col_d2.text_input("Nome da coluna DATA:", "DATA VISITA", key="m2_d")
        
        f_geral2 = st.file_uploader("Suba o Relatório Mensal Duplicado (Excel ou CSV)", type=["xlsx", "csv"], key="dup_file_m2")
        
        if f_geral2:
            df_ba = load_excel_with_ui(f_geral2, "m2_view")
            if st.button("🚀 INICIAR UNIFICAÇÃO MENSAL", type="primary"):
                try:
                    with st.spinner("Buscando a data mais recente, limpando telefones e compactando a planilha..."):
                        colunas_originais = list(df_ba.columns)
                        df_ba['_ORIG_INDEX'] = range(len(df_ba))
                        
                        for c in colunas_originais:
                            df_ba[c] = df_ba[c].apply(formatar_brasileiro_sem_hora)
                        
                        cols_upper = [str(c).upper().strip() for c in colunas_originais]
                        idx_num = next((i for i, c in enumerate(cols_upper) if nm_num2.upper() in c), None)
                        idx_nome = next((i for i, c in enumerate(cols_upper) if nm_nome2.upper() in c), None)
                        idx_data = next((i for i, c in enumerate(cols_upper) if nm_data2.upper() in c), None)
                        idx_conduta = next((i for i, c in enumerate(cols_upper) if 'CONDUTA' in c), None)
                        
                        if None in (idx_num, idx_nome, idx_data):
                            st.error("❌ MODO 2 requer as colunas NUM, NOME e DATA VISITA para funcionar. Verifique os nomes nas caixas!")
                        else:
                            real_num = colunas_originais[idx_num]
                            real_nome = colunas_originais[idx_nome]
                            real_data = colunas_originais[idx_data]
                            real_conduta = colunas_originais[idx_conduta] if idx_conduta is not None else None
                            
                            df_ba[real_nome] = df_ba[real_nome].astype(str).str.strip().str.upper()
                            
                            if real_conduta:
                                df_ba[real_conduta] = df_ba[real_conduta].apply(formatar_conduta)
                                df_ba['_SCORE_CONDUTA'] = df_ba[real_conduta].apply(rank_conduta)
                            else:
                                df_ba['_SCORE_CONDUTA'] = 999
                            
                            df_clean = df_ba.copy()
                            df_clean = df_clean.replace(r'^\s*$', pd.NA, regex=True)
                            df_clean = df_clean.replace(['nan', 'NaN', 'None', '<NA>', ''], pd.NA)
                            
                            df_clean['_GRP_NUM'] = df_clean[real_num].astype(str).str.strip().replace(r'\.0$', '', regex=True).str.upper()
                            df_clean['_GRP_NOME'] = df_clean[real_nome].astype(str).str.strip().str.upper()
                            
                            df_clean['_TEMP_DATE'] = pd.to_datetime(df_clean[real_data], format='%d/%m/%Y', errors='coerce', dayfirst=True)
                            
                            # DESEMPATE: Menor Score da Conduta ganha, e Data mais nova ganha.
                            df_clean = df_clean.sort_values(by=['_GRP_NUM', '_GRP_NOME', '_SCORE_CONDUTA', '_TEMP_DATE'], ascending=[True, True, True, False])
                            
                            total_linhas_antes = len(df_clean)
                            grp_cols = ['_GRP_NUM', '_GRP_NOME']
                            
                            def pegar_primeiro_valido(x):
                                validos = [v for v in x if pd.notna(v) and str(v).strip() != '' and str(v).lower() not in ['nan', 'none', '<na>']]
                                return validos[0] if validos else ""
                            
                            cols_to_agg = colunas_originais + ['_ORIG_INDEX']
                            df_unified = df_clean.groupby(grp_cols, dropna=False)[cols_to_agg].agg(pegar_primeiro_valido).reset_index()
                            df_unified = df_unified.sort_values(by='_ORIG_INDEX')
                            df_unified = df_unified[colunas_originais]
                            
                            # APLICA MOTOR AVANÇADO DE TELEFONES NA BASE UNIFICADA
                            cols_telefone = [c for c in colunas_originais if 'TEL' in str(c).upper() or 'CEL' in str(c).upper()]
                            if cols_telefone:
                                col_tel_main = cols_telefone[0]
                                tel_res = df_unified.apply(lambda row: processar_telefones_avancado(row, cols_telefone), axis=1)
                                df_unified[col_tel_main] = tel_res[0]
                                df_unified['TEL. ADIC.'] = tel_res[1]
                            
                            linhas_removidas = total_linhas_antes - len(df_unified)
                            
                            b_zip_dup = io.BytesIO()
                            with zipfile.ZipFile(b_zip_dup, "w", zipfile.ZIP_DEFLATED) as zf:
                                buf_limpo = io.BytesIO()
                                with pd.ExcelWriter(buf_limpo, engine='openpyxl') as writer:
                                    df_unified.to_excel(writer, index=False, sheet_name='BASE_UNIFICADA')
                                    ws_new = writer.sheets['BASE_UNIFICADA']
                                    
                                    if f_geral2.name.endswith('.xlsx'):
                                        try:
                                            f_geral2.seek(0)
                                            wb_orig = openpyxl.load_workbook(f_geral2)
                                            ws_orig = wb_orig.active
                                            for col_letter, col_dim in ws_orig.column_dimensions.items():
                                                ws_new.column_dimensions[col_letter].width = col_dim.width
                                        except:
                                            for col in ws_new.columns: ws_new.column_dimensions[col[0].column_letter].width = 20
                                    else:
                                        for col in ws_new.columns: ws_new.column_dimensions[col[0].column_letter].width = 20
                                    
                                    phone_col_indices = []
                                    for cell in ws_new[1]:
                                        cell.font = Font(bold=True)
                                        if cell.value and ('TEL' in str(cell.value).upper() or 'CEL' in str(cell.value).upper()):
                                            phone_col_indices.append(cell.column)
                                            
                                    for row in ws_new.iter_rows(min_row=2):
                                        for cell in row:
                                            cell.alignment = Alignment(wrap_text=True, vertical='center')
                                            if cell.column in phone_col_indices and cell.value:
                                                if len(str(cell.value).replace('+55','')) > 0 and len(str(cell.value).replace('+55','')) < 10:
                                                    cell.font = Font(color="FF0000", bold=True)
                                    ws_new.freeze_panes = 'A2'
                                zf.writestr("RELATORIO_MENSAL_UNIFICADO.xlsx", buf_limpo.getvalue())
                                    
                            b_zip_dup.seek(0)
                            st.session_state['base_deduplicada_mensal'] = b_zip_dup.getvalue()
                            
                            st.success(f"✅ Fusão Concluída! O robô unificou {linhas_removidas} linhas, ativou o desempate de condutas e formatou os telefones!")

                except Exception as e:
                    st.error(f"Erro ao processar: {e}")

        if st.session_state.get('base_deduplicada_mensal') is not None:
            st.download_button("📥 BAIXAR RELATÓRIO MENSAL UNIFICADO", st.session_state['base_deduplicada_mensal'], "HOVA_Relatorio_Mensal.zip", mime="application/zip", type="primary")

    else:
        st.info("💡 **Aviso:** Neste modo, o robô vai caçar ativamente Emails e CPFs em colunas erradas, juntar históricos de mensagens, aplicar o **Novo Motor de Telefonia** e formatar Condutas.")
        col_n, col_name, col_t, col_d = st.columns(4)
        nm_num = col_n.text_input("Nome da coluna do NUM/Prontuário:", "NUM")
        nm_nome = col_name.text_input("Nome da coluna de NOME:", "PACIENTE")
        nm_tel = col_t.text_input("Nome da coluna de TELEFONE:", "TELEFONE")
        nm_data = col_d.text_input("Nome da coluna de DATA:", "DATA VISITA")
        
        f_geral = st.file_uploader("Suba a sua Planilha Geral (Excel ou CSV)", type=["xlsx", "csv"], key="dup_file_m1")
        
        if f_geral:
            df_ba = load_excel_with_ui(f_geral, "m1_view")
            if st.button("🧹 LIMPAR, ASPIRAR E FORMATAR BASE", type="primary"):
                try:
                    with st.spinner("Escaneando dados, ativando novo motor de telefonia e limpando a base..."):
                        df_ba['_ORIG_INDEX'] = range(len(df_ba))
                        colunas_originais = list(df_ba.columns)
                        colunas_originais.remove('_ORIG_INDEX') 
                        
                        for c in colunas_originais:
                            df_ba[c] = df_ba[c].apply(formatar_brasileiro_sem_hora)
                            df_ba[c] = df_ba[c].replace(['nan', 'NaN', 'None', '<NA>'], '')

                        cols_upper = [str(c).upper().strip() for c in colunas_originais]
                        idx_num = next((i for i, c in enumerate(cols_upper) if nm_num.upper() in c), None)
                        idx_nome = next((i for i, c in enumerate(cols_upper) if nm_nome.upper() in c), None)
                        idx_tel = next((i for i, c in enumerate(cols_upper) if nm_tel.upper() in c), None)
                        idx_data = next((i for i, c in enumerate(cols_upper) if nm_data.upper() in c), None)
                        idx_conduta = next((i for i, c in enumerate(cols_upper) if 'CONDUTA' in c), None)
                        
                        if None in (idx_num, idx_nome, idx_tel, idx_data):
                            st.error("❌ Não foi possível encontrar todas as 4 colunas base. Verifique as caixas acima!")
                        else:
                            real_num = colunas_originais[idx_num]
                            real_nome = colunas_originais[idx_nome]
                            real_tel = colunas_originais[idx_tel]
                            real_data = colunas_originais[idx_data]
                            real_conduta = colunas_originais[idx_conduta] if idx_conduta is not None else None
                            
                            df_ba[real_nome] = df_ba[real_nome].astype(str).str.strip().str.upper()
                            
                            # APLICA MOTOR AVANÇADO DE TELEFONES
                            cols_telefone = [c for c in colunas_originais if 'TEL' in str(c).upper() or 'CEL' in str(c).upper()]
                            if cols_telefone:
                                col_tel_main = cols_telefone[0]
                                tel_res = df_ba.apply(lambda row: processar_telefones_avancado(row, cols_telefone), axis=1)
                                df_ba[col_tel_main] = tel_res[0]
                                if 'TEL. ADIC.' not in df_ba.columns:
                                    df_ba['TEL. ADIC.'] = tel_res[1]
                                    colunas_originais.append('TEL. ADIC.')
                                else:
                                    df_ba['TEL. ADIC.'] = tel_res[1]
                                    
                            # APLICA MOTOR DE CONDUTA E SCORE
                            if real_conduta:
                                df_ba[real_conduta] = df_ba[real_conduta].apply(formatar_conduta)
                                df_ba['_SCORE_CONDUTA'] = df_ba[real_conduta].apply(rank_conduta)
                            else:
                                df_ba['_SCORE_CONDUTA'] = 999

                            msg_cols = [c for c in colunas_originais if 'MSG' in str(c).upper()]
                            col_email = next((c for c in colunas_originais if 'MAIL' in str(c).upper()), None)
                            col_cpf = next((c for c in colunas_originais if 'CPF' in str(c).upper()), None)
                            col_nasc = next((c for c in colunas_originais if 'NASCIMENTO' in str(c).upper()), None)
                            
                            regex_email = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
                            regex_cpf = re.compile(r'\b\d{3}\.\d{3}\.\d{3}-\d{2}\b')
                            regex_data = re.compile(r'\b\d{2}/\d{2}/\d{4}\b')

                            for index, row in df_ba.iterrows():
                                for col in colunas_originais:
                                    val = str(row[col]).strip()
                                    if val and val.lower() not in ['nan', 'none', '<na>']:
                                        if col_email and col != col_email:
                                            achados_email = regex_email.findall(val)
                                            if achados_email:
                                                if not str(df_ba.at[index, col_email]).strip() or str(df_ba.at[index, col_email]).lower() == 'nan':
                                                    df_ba.at[index, col_email] = achados_email[0]
                                                val = val.replace(achados_email[0], '').strip()
                                                df_ba.at[index, col] = val

                                        if col_cpf and col != col_cpf:
                                            achados_cpf = regex_cpf.findall(val)
                                            if achados_cpf:
                                                if not str(df_ba.at[index, col_cpf]).strip() or str(df_ba.at[index, col_cpf]).lower() == 'nan':
                                                    df_ba.at[index, col_cpf] = achados_cpf[0]
                                                val = val.replace(achados_cpf[0], '').strip()
                                                df_ba.at[index, col] = val

                                        if col_nasc and col != col_nasc and col != real_data:
                                            achados_data = regex_data.findall(val)
                                            for d in achados_data:
                                                try:
                                                    ano = int(d.split('/')[-1])
                                                    if ano < 2020:
                                                        if not str(df_ba.at[index, col_nasc]).strip() or str(df_ba.at[index, col_nasc]).lower() == 'nan':
                                                            df_ba.at[index, col_nasc] = d
                                                        val = val.replace(d, '').strip()
                                                        df_ba.at[index, col] = val
                                                except: pass

                            def lavar_mensagens_duplicadas(lista_msgs):
                                msgs_limpas = {}
                                for val_msg in lista_msgs:
                                    val_str = str(val_msg).strip()
                                    if val_str and val_str.lower() not in ['nan', 'none', 'nat', '<na>']:
                                        for linha_msg in val_str.split('\n'):
                                            linha_msg = linha_msg.strip()
                                            if linha_msg:
                                                linha_msg = linha_msg.replace(' 00:00:00', '')
                                                linha_msg = re.sub(r'\s*\b\d{2}:\d{2}(:\d{2})?\b', '', linha_msg).strip()
                                                frase_matematica = re.sub(r'[^A-Z0-9]', '', linha_msg.upper())
                                                
                                                if frase_matematica in msgs_limpas:
                                                    if linha_msg.count(' ') > msgs_limpas[frase_matematica].count(' '):
                                                        msgs_limpas[frase_matematica] = linha_msg
                                                else:
                                                    msgs_limpas[frase_matematica] = linha_msg
                                                    
                                def date_key(txt):
                                    match = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4})', txt)
                                    if match:
                                        try: return pd.to_datetime(match.group(1), format='%d/%m/%Y', dayfirst=True)
                                        except: return pd.Timestamp.min
                                    return pd.Timestamp.min
                                    
                                return "\n".join(sorted(list(msgs_limpas.values()), key=date_key))

                            df_ba['_GRP_NUM'] = df_ba[real_num].astype(str).str.strip().replace(r'\.0$', '', regex=True).str.upper()
                            df_ba['_GRP_NOME'] = df_ba[real_nome].astype(str).str.strip().str.upper()
                            df_ba['_GRP_TEL'] = df_ba[real_tel].apply(limpar_num)
                            df_ba['_TEMP_DATE'] = pd.to_datetime(df_ba[real_data], format='%d/%m/%Y', errors='coerce', dayfirst=True)
                            
                            # DESEMPATE MASTER: Menor Score da Conduta ganha, e Data mais nova ganha
                            df_ba = df_ba.sort_values(by=['_GRP_NUM', '_GRP_NOME', '_GRP_TEL', '_SCORE_CONDUTA', '_TEMP_DATE'], ascending=[True, True, True, True, False])
                            
                            grp_cols = ['_GRP_NUM', '_GRP_NOME', '_GRP_TEL']
                            duplicated_mask = df_ba.duplicated(subset=grp_cols, keep='first')
                            
                            df_limpo = df_ba[~duplicated_mask].copy()
                            df_duplicados = df_ba[duplicated_mask].copy()
                            
                            if not df_duplicados.empty:
                                df_clean = df_ba.copy()
                                for c in colunas_originais:
                                    if c not in msg_cols:
                                        df_clean[c] = df_clean[c].replace(r'^\s*$', pd.NA, regex=True)
                                        
                                grouped_first = df_clean.groupby(grp_cols, dropna=False).first().reset_index()
                                
                                df_limpo.set_index(grp_cols, inplace=True)
                                grouped_first.set_index(grp_cols, inplace=True)
                                
                                cols_to_update = [c for c in colunas_originais if c not in [real_num, real_nome, real_tel, real_data] + msg_cols]
                                df_limpo.update(grouped_first[cols_to_update])
                                
                                if msg_cols:
                                    grouped_msgs = df_ba.groupby(grp_cols, dropna=False)[msg_cols].agg(lambda x: lavar_mensagens_duplicadas(x.tolist()))
                                    df_limpo.update(grouped_msgs)
                                    
                                df_limpo.reset_index(inplace=True)
                            
                            df_limpo = df_limpo.sort_values(by='_ORIG_INDEX')
                            
                            df_limpo = df_limpo[colunas_originais]
                            df_duplicados = df_duplicados[colunas_originais]
                            
                            b_zip_dup = io.BytesIO()
                            with zipfile.ZipFile(b_zip_dup, "w", zipfile.ZIP_DEFLATED) as zf:
                                buf_limpo = io.BytesIO()
                                with pd.ExcelWriter(buf_limpo, engine='openpyxl') as writer:
                                    df_limpo.to_excel(writer, index=False, sheet_name='BASE_LIMPA')
                                    ws_new = writer.sheets['BASE_LIMPA']
                                    
                                    if f_geral.name.endswith('.xlsx'):
                                        try:
                                            f_geral.seek(0)
                                            wb_orig = openpyxl.load_workbook(f_geral)
                                            ws_orig = wb_orig.active
                                            for col_letter, col_dim in ws_orig.column_dimensions.items():
                                                ws_new.column_dimensions[col_letter].width = col_dim.width
                                        except:
                                            for col in ws_new.columns: ws_new.column_dimensions[col[0].column_letter].width = 20
                                    else:
                                        for col in ws_new.columns: ws_new.column_dimensions[col[0].column_letter].width = 20
                                    
                                    # CORES NO EXCEL: Identificando colunas de telefone para pintar de vermelho
                                    phone_col_indices = []
                                    for cell in ws_new[1]:
                                        cell.font = Font(bold=True)
                                        if cell.value and ('TEL' in str(cell.value).upper() or 'CEL' in str(cell.value).upper() or 'ADIC' in str(cell.value).upper()):
                                            phone_col_indices.append(cell.column)
                                            
                                    for row in ws_new.iter_rows(min_row=2):
                                        for cell in row:
                                            cell.alignment = Alignment(wrap_text=True, vertical='center')
                                            if cell.column in phone_col_indices and cell.value:
                                                if len(str(cell.value).replace('+55','')) > 0 and len(str(cell.value).replace('+55','')) < 10:
                                                    cell.font = Font(color="FF0000", bold=True)
                                    ws_new.freeze_panes = 'A2'
                                    
                                zf.writestr("1_BASE_LIMPA_E_ASPIRADA.xlsx", buf_limpo.getvalue())
                                
                                if not df_duplicados.empty:
                                    buf_dup = io.BytesIO()
                                    df_duplicados.to_excel(buf_dup, index=False)
                                    zf.writestr("2_DUPLICADOS_DESCARTADOS.xlsx", buf_dup.getvalue())
                                    
                            b_zip_dup.seek(0)
                            st.session_state['base_deduplicada_padrao'] = b_zip_dup.getvalue()
                            
                            st.success(f"✅ Limpeza Master Concluída! Novo Motor de WhatsApp e Condutas ativado, Emails e Telefones arrumados, e {len(df_duplicados)} duplicados pulverizados.")

                except Exception as e:
                    st.error(f"Erro ao processar: {e}")

        if st.session_state.get('base_deduplicada_padrao') is not None:
            st.download_button("📥 BAIXAR BASE LIMPA FORMATADA", st.session_state['base_deduplicada_padrao'], "HOVA_Base_Limpa.zip", mime="application/zip", type="primary")

    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# --- ABA 10: CENTRAL DE AGENTES IA (O PAINEL MODERNO DEFINITIVO - ESTANTE ELÁSTICA) ---
# ==========================================
with tab10 if 'tab10' in locals() else st.container():
    st.markdown('<div class="master-card" style="padding: 20px;">', unsafe_allow_html=True)
    st.markdown("""<div class="premium-title"> ESQUADRÃO HOVA MASTER INTELLIGENCE</div>""", unsafe_allow_html=True)

    import os
    import base64
    import streamlit.components.v1 as components

    # Transforma a foto JPG em código pra furar o bloqueio do Streamlit
    def get_foto_b64(nome):
        primeiro_nome = nome.split(' ')[0].lower()
        caminhos = [f"{primeiro_nome}.jpg", f"{primeiro_nome}.png", f"AGENTES IA/{primeiro_nome}.jpg", f"AGENTES IA/{primeiro_nome}.png"]
        for c in caminhos:
            if os.path.exists(c):
                try:
                    with open(c, "rb") as f:
                        ext = "png" if "png" in c.lower() else "jpeg"
                        return f"data:image/{ext};base64,{base64.b64encode(f.read()).decode()}"
                except: pass
        return f"https://ui-avatars.com/api/?name={primeiro_nome}&background=1e3d3a&color=00ffcc&size=150&rounded=true&bold=true"

    # Banco de Dados dos Agentes
    agentes_lista = [
        {"n": "LUMINA ALMEIDA", "s": "WEB", "z": "31 9723-6408", "g": "F", "t": "busca"},
        {"n": "PITER SANTOS", "s": "SLOT 11", "z": "31 9528-5492", "g": "M", "t": "busca"},
        {"n": "CLARA MARTINS", "s": "SLOT 07", "z": "31 9743-4631", "g": "F", "t": "busca"},
        {"n": "PRISMA RAMOS", "s": "SLOT 20", "z": "31 7221-8952", "g": "F", "t": "busca"},
        {"n": "ROGER OLIVEIRA", "s": "SLOT 14", "z": "31 9953-2096", "g": "M", "t": "confirm"},
        {"n": "NATALIA VIANA", "s": "SLOT 01", "z": "31 7150-8930", "g": "F", "t": "catarata"},
        {"n": "STELLA VEIRA", "s": "SLOT 10", "z": "31 9670-1479", "g": "F", "t": "confirm"},
        {"n": "ESTER TEIXEIRA", "s": "SLOT 16", "z": "31 97202-3913", "g": "F", "t": "confirm"},
        {"n": "AYLA FREITAS", "s": "SLOT 04", "z": "N/A", "g": "F", "t": "confirm"},
        {"n": "OSCAR SIQUEIRA", "s": "SLOT 13", "z": "N/A", "g": "M", "t": "busca"},
        {"n": "THEIA DIAS", "s": "SLOT 12", "z": "31 9788-9331", "g": "F", "t": "busca"}
    ]

    cards_html = ""
    for i, ag in enumerate(agentes_lista):
        nome = ag['n']
        art = "o" if ag['g'] == "M" else "a"
        
        # Textos
        msg_busca = f"Olá, *{{{{column_2}}}}*! Tudo bem? 👁️\n\nAqui é {art} *{nome}*, {art}ssistente do *Hospital de Olhos Vale do Aço*.\n\nEstava revisando seu histórico e percebi que já faz um tempo desde sua última consulta. A saúde dos seus olhinhos merece atenção regular — está na hora de realizar uma nova avaliação preventiva.\n\n📲 *Para agendar:*\n1. Toque no link azul abaixo\n2. O WhatsApp abre automaticamente\n3. Nossa equipe vai te atender\n\n👇 *Clique aqui para agendar:*\nhttps://wa.me/553138011800\n\n*Aviso:* Este número só envia lembretes. Para agendar, use o link acima.\n\n💡 Já consultou recentemente? Pode desconsiderar esta mensagem.\n\nPodemos te esperar? 🤍\n\n*{nome} — Hospital de Olhos Vale do Aço*"
        
        msg_conf = f"Olá, {{{{column_4}}}}! Tudo bem? 👁️\n\nAqui é {art} *{nome}*, su{art} assistente do Hospital de Olhos Vale do Aço.\n\n📌 *Lembrete automático do seu atendimento:*\n\n📋 *Atendimento:* {{{{column_6}}}}\n👨‍⚕️ *Médico(a):* {{{{column_7}}}}\n📅 *Data:* {{{{column_1}}}}\n⏰ *Chegada:* {{{{column_2}}}}\n\n🚨 *PREPARO OBRIGATÓRIO:*\n\n{{{{column_8}}}}\n\n📁 *O QUE TRAZER:*\n\n{{{{column_9}}}}\n\n⚠️ *CONFIRMAÇÃO:*\n\n✅ *Confirmar:* Responda *SIM* _(Sistema registra automaticamente)_\n\n❌ *Cancelar/Remarcar:* Não responda aqui\nTemos fila de espera. Clique no link:\n👉 https://wa.me/553138011800\n\n🤖 Mensagem automática. Sistema em atualização.\n\n🤍 *{nome} — Hospital de Olhos Vale do Aço*"
        
        msg_cat = f"Olá, *{{{{column_2}}}}*! 👁️\n\nAqui é a *NATALIA VIANA*, do Hospital de Olhos Vale do Aço.\n\nVou te enviar um áudio explicativo com informações importantes sobre seu atendimento. Por favor, ouça com atenção. 🎧\n\nQualquer dúvida, estou à disposição. 🤍\n\nCaso queira agendar, entre em contato pelo WhatsApp: https://wa.me/553138011800\n\nObs.: Caso você já esteja se preparando ou já tenha realizado o procedimento, por gentileza, desconsidere esta mensagem."

        img_b64 = get_foto_b64(nome)
        delay = i * 0.05  
        
        txt_busca_js = msg_busca.replace('\n', '\\n').replace('`', '\\`')
        txt_conf_js = msg_conf.replace('\n', '\\n').replace('`', '\\`')
        txt_cat_js = msg_cat.replace('\n', '\\n').replace('`', '\\`')

        # Destaque visual baseado na função principal
        btn_busca_class = "btn-solid" if ag['t'] == 'busca' else "btn-outline"
        btn_conf_class = "btn-solid" if ag['t'] == 'confirm' else "btn-outline"

        botoes_html = f"""
            <button class="btn-ag {btn_busca_class}" onclick="copiar(this, `{txt_busca_js}`)">📋 BUSCA ATIVA</button>
            <button class="btn-ag {btn_conf_class}" onclick="copiar(this, `{txt_conf_js}`)">📅 CONFIRMAÇÃO</button>
        """
        if ag['t'] == 'catarata':
            botoes_html += f"""<button class="btn-ag btn-solid btn-catarata" onclick="copiar(this, `{txt_cat_js}`)">🎧 ÁUDIO CATARATA</button>"""

        cards_html += f"""
        <div class="card-ag" style="animation-delay: {delay}s;">
            <div class="img-wrapper"><img src="{img_b64}"></div>
            <div class="nome-ag">{nome}</div>
            <div class="slot-ag">{ag['s']}</div>
            <div class="zap-ag">📱 {ag['z']}</div>
            <div class="botoes-container">
                {botoes_html}
            </div>
        </div>
        """

    # CSS - A MÁGICA DA ESTANTE ELÁSTICA (Grid auto-fill + minmax)
    html_final = f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap');
        :root {{ --verde-escuro: #1e3d3a; --verde-claro: #2f6c68; --destaque: #00ffcc; }}
        
        /* A regra de Ouro: Preenche a tela toda, da esquerda pra direita, adaptando as cartas */
        .grid-ag {{ 
            display: grid; 
            grid-template-columns: repeat(auto-fill, minmax(360px, 1fr)); 
            gap: 30px; padding: 20px; font-family: 'Outfit', sans-serif; 
            width: 100%;
        }}
        
        @keyframes fadeSlideUp {{ 0% {{ opacity: 0; transform: translateY(40px); }} 100% {{ opacity: 1; transform: translateY(0); }} }}
        
        .card-ag {{ 
            background: #ffffff; border-radius: 25px; padding: 35px 20px; text-align: center; 
            box-shadow: 0 10px 30px rgba(30,61,58,0.08); border: 1px solid rgba(30,61,58,0.1);
            opacity: 0; animation: fadeSlideUp 0.6s cubic-bezier(0.16, 1, 0.3, 1) forwards;
            transition: all 0.3s ease; position: relative; overflow: hidden;
            display: flex; flex-direction: column;
        }}
        .card-ag::before {{
            content: ''; position: absolute; top: 0; left: 0; right: 0; height: 6px;
            background: linear-gradient(90deg, var(--verde-escuro), var(--destaque));
            transform: scaleX(0); transition: transform 0.4s ease; transform-origin: left;
        }}
        .card-ag:hover {{ transform: translateY(-10px); box-shadow: 0 20px 40px rgba(30,61,58,0.15); border-color: rgba(30,61,58,0.3); z-index: 10; }}
        .card-ag:hover::before {{ transform: scaleX(1); }}
        
        .img-wrapper {{
            width: 150px; height: 150px; /* FOTO GIGANTE */
            margin: 0 auto 20px; border-radius: 50%; padding: 5px;
            background: linear-gradient(135deg, var(--verde-escuro), var(--verde-claro));
            box-shadow: 0 8px 25px rgba(30,61,58,0.25); transition: 0.3s ease;
        }}
        .card-ag:hover .img-wrapper {{ transform: scale(1.05) rotate(3deg); }}
        .img-wrapper img {{ width: 100%; height: 100%; object-fit: cover; border-radius: 50%; border: 4px solid #fff; background: white; }}
        
        .nome-ag {{ font-weight: 800; color: var(--verde-escuro); font-size: 1.4rem; letter-spacing: 0.5px; margin-bottom: 6px; }}
        .slot-ag {{ 
            background: var(--verde-escuro); color: #fff; padding: 5px 15px; border-radius: 20px; 
            font-size: 0.85rem; font-weight: bold; display: inline-block; margin-bottom: 15px;
        }}
        .zap-ag {{ color: #64748b; font-size: 1rem; font-weight: 700; display: block; margin-bottom: 25px; }}
        
        .botoes-container {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: auto; }}
        
        .btn-ag {{ 
            padding: 14px 5px; border-radius: 12px; cursor: pointer; 
            font-weight: 800; transition: 0.2s; font-size: 0.85rem; 
            text-transform: uppercase; display: flex; align-items: center; justify-content: center; 
            font-family: 'Outfit', sans-serif;
        }}
        
        .btn-solid {{ background: var(--verde-escuro); color: white; border: none; box-shadow: 0 6px 15px rgba(30,61,58,0.15); }}
        .btn-solid:hover {{ background: var(--verde-claro); transform: translateY(-2px); box-shadow: 0 8px 20px rgba(30,61,58,0.25); }}
        
        .btn-outline {{ background: transparent; color: var(--verde-escuro); border: 2px solid rgba(30,61,58,0.25); }}
        .btn-outline:hover {{ background: rgba(30,61,58,0.05); border-color: var(--verde-escuro); transform: translateY(-2px); }}
        
        .btn-catarata {{ grid-column: span 2; background: #e67e22; color: white; border: none; }}
        .btn-catarata:hover {{ background: #d35400; }}
        
        .btn-ag:active {{ transform: scale(0.95); }}
        .btn-ag.success {{ background: #000 !important; color: var(--destaque) !important; border-color: #000 !important; grid-column: span 2; font-size: 0.95rem; }}
    </style>

    <div class="grid-ag">{cards_html}</div>

    <script>
        function copiar(btn, texto) {{
            const txtCorrigido = texto.replace(/\\n/g, '\\n');
            navigator.clipboard.writeText(txtCorrigido).then(() => {{
                const txtVelho = btn.innerHTML;
                btn.innerHTML = "✅ COPIADO!";
                btn.classList.add('success');
                setTimeout(() => {{
                    btn.innerHTML = txtVelho;
                    btn.classList.remove('success');
                }}, 1500);
            }});
        }}
    </script>
    """
    
    components.html(html_final, height=1800, scrolling=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="footer-master">HOVA MASTER INTELLIGENCE — UNIDADE IPATINGA</div>', unsafe_allow_html=True)