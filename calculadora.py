import streamlit as st
import pandas as pd
import requests
import io
from datetime import datetime, timedelta, date
from decimal import Decimal
from fpdf import FPDF

# Configuração da página
st.set_page_config(page_title="Auditoria de Cheque Especial", layout="wide")
st.title("Sistema Pericial: Cálculo de Cheque Especial")

# --- CONTROLE DE MEMÓRIA (SESSION STATE) ---
if 'reset_contador' not in st.session_state:
    st.session_state.reset_contador = 0

def limpar_tabela():
    st.session_state.reset_contador += 1

# --- Dicionário com os Códigos Oficiais do SGS/BCB ---
CODIGOS_BCB = {
    "IGP-M": 189,
    "IPCA": 433,
    "INPC": 188,
    "INCC": 192
}

# --- FUNÇÃO: BUSCAR DADOS DO BANCO CENTRAL ---
@st.cache_data
def buscar_indice_bcb(codigo_bcb):
    try:
        url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo_bcb}/dados?formato=json"
        resposta = requests.get(url)
        df = pd.DataFrame(resposta.json())
        df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y')
        df['Mês/Ano'] = df['data'].dt.strftime('%Y-%m')
        df['Índice (%)'] = df['valor'].astype(float)
        return df[['Mês/Ano', 'Índice (%)']]
    except Exception as e:
        st.error("Erro ao conectar com o Banco Central.")
        return pd.DataFrame(columns=["Mês/Ano", "Índice (%)"])

# --- FUNÇÃO: GERAR ARQUIVO EXCEL EM MEMÓRIA ---
def gerar_excel(df_resumo, df_detalhado):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='Resumo da Divida', index=False)
        df_detalhado.to_excel(writer, sheet_name='Memoria de Calculo', index=False)
    return output.getvalue()

# --- FUNÇÃO: GERAR RELATÓRIO PDF FORMAL ---
def gerar_pdf(resumo_dados, df_detalhado, indice_nome, juros_tipo, taxa):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    pdf.set_font("Arial", style="B", size=16)
    pdf.set_text_color(0, 64, 128) 
    pdf.cell(0, 10, "RELATÓRIO PERICIAL DE REVISÃO FINANCEIRA", ln=True, align="C")
    pdf.set_text_color(0, 0, 0) 
    
    pdf.set_font("Arial", style="I", size=10)
    pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=True, align="C")
    pdf.ln(10)
    
    pdf.set_font("Arial", style="B", size=12)
    pdf.cell(0, 8, "1. PARÂMETROS DO CONTRATO E ATUALIZAÇÃO", ln=True)
    pdf.set_font("Arial", size=11)
    
    # Ajusta o texto do PDF caso seja "Sem Atualização"
    if indice_nome == "Sem Atualização (Apenas Juros)":
        pdf.cell(0, 6, "Índice de Correção Monetária: Não Aplicado (Apenas Juros)", ln=True)
    else:
        pdf.cell(0, 6, f"Índice de Correção Monetária: {indice_nome} (SGS/Banco Central)", ln=True)
        
    pdf.cell(0, 6, f"Método de Capitalização dos Juros: Juros {juros_tipo}", ln=True)
    pdf.cell(0, 6, f"Taxa de Juros Contratada: {taxa:.3f}% ao mês", ln=True)
    pdf.cell(0, 6, f"Período de Auditoria: {resumo_dados['Dias']} dias", ln=True)
    pdf.ln(8)
    
    pdf.set_font("Arial", style="B", size=12)
    pdf.cell(0, 8, "2. RESUMO DOS VALORES APURADOS", ln=True)
    pdf.set_font("Arial", size=11)
    pdf.cell(0, 6, f"Saldo Original Acumulado (Sem Juros): R$ {resumo_dados['Original']:.2f}", ln=True)
    pdf.cell(0, 6, f"Total de Juros Computados no Período: R$ {resumo_dados['Juros']:.2f}", ln=True)
    pdf.set_font("Arial", style="B", size=11)
    pdf.cell(0, 6, f"VALOR TOTAL RECALCULADO DA DÍVIDA: R$ {resumo_dados['Final']:.2f}", ln=True)
    pdf.ln(10)
    
    pdf.set_font("Arial", style="B", size=12)
    pdf.cell(0, 8, "3. EXTRATO DA MEMÓRIA DE CÁLCULO DIÁRIA", ln=True)
    pdf.ln(2)
    
    pdf.set_font("Arial", style="B", size=9)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(25, 6, "Data", border=1, align="C", fill=True)
    pdf.cell(35, 6, "S. Anterior", border=1, align="R", fill=True)
    pdf.cell(20, 6, "Corr. (R$)", border=1, align="R", fill=True)
    pdf.cell(28, 6, "Débitos (-)", border=1, align="R", fill=True)
    pdf.cell(28, 6, "Créditos (+)", border=1, align="R", fill=True)
    pdf.cell(35, 6, "S. Final Dia", border=1, align="R", fill=True)
    pdf.ln()
    
    pdf.set_font("Arial", size=8)
    for _, row in df_detalhado.iterrows():
        pdf.cell(25, 5, str(row["Data"]), border=1, align="C")
        pdf.cell(35, 5, f"{row['Saldo Anterior']:.2f}", border=1, align="R")
        pdf.cell(20, 5, f"{row['Correção (R$)']:.2f}", border=1, align="R")
        pdf.cell(28, 5, f"{row['Débitos (R$)']:.2f}", border=1, align="R")
        pdf.cell(28, 5, f"{row['Créditos (R$)']:.2f}", border=1, align="R")
        pdf.cell(35, 5, f"{row['Saldo Final Dia']:.2f}", border=1, align="R")
        pdf.ln()
        
    return bytes(pdf.output())

# --- INTERFACE WEB ---
Config_hoje = date.today()
primeiro_dia_mes = Config_hoje.replace(day=1)

with st.sidebar:
    st.header("Parâmetros do Cálculo")
    data_inicial = st.date_input("Data Inicial", primeiro_dia_mes, format="DD/MM/YYYY")
    data_final = st.date_input("Data Final", Config_hoje, format="DD/MM/YYYY")
    saldo_inicial = st.number_input("Saldo Inicial Negativo (R$)", value=0.00, step=100.0)
    
    st.markdown("---")
    st.header("Regras do Contrato")
    
    # NOVA OPÇÃO INCLUÍDA AQUI
    opcoes_indices = ["Sem Atualização (Apenas Juros)"] + list(CODIGOS_BCB.keys())
    indice_escolhido = st.selectbox("Índice de Atualização", opcoes_indices)
    
    tipo_juros = st.radio("Método de Juros", ["Compostos", "Simples"])
    taxa_juros = st.number_input("Taxa de Juros a.m. (%)", value=8.000, format="%.3f")

    st.markdown("---")
    
    # LÓGICA VISUAL: Esconde a tabela se escolher Sem Atualização
    if indice_escolhido == "Sem Atualização (Apenas Juros)":
        st.info("ℹ️ O recálculo será feito aplicando apenas a Taxa de Juros informada, sem correção monetária sobre o saldo diário.")
        df_indices = pd.DataFrame(columns=["Mês/Ano", "Índice (%)"]) # Cria tabela vazia oculta para o motor não travar
    else:
        st.header(f"Tabela Oficial - {indice_escolhido}")
        codigo_atual = CODIGOS_BCB[indice_escolhido]
        df_historico_completo = buscar_indice_bcb(codigo_atual)
        
        ano_inicio = str(data_inicial.year)
        if not df_historico_completo.empty:
            df_filtrado = df_historico_completo[df_historico_completo['Mês/Ano'] >= f"{ano_inicio}-01"].copy()
        else:
            df_filtrado = pd.DataFrame(columns=["Mês/Ano", "Índice (%)"])
        df_indices = st.data_editor(df_filtrado, num_rows="dynamic", hide_index=True)

st.subheader("Livro de Lançamentos Diários (Entradas e Saídas)")
st.write("Insira os valores correspondentes a cada dia.")

dias_totais = (data_final - data_inicial).days
datas_iniciais = [(data_inicial + timedelta(days=i)) for i in range(dias_totais + 1)]

df_lancamentos_iniciais = pd.DataFrame({
    "Data": datas_iniciais,
    "Débitos (-)": [0.00 for _ in range(len(datas_iniciais))],
    "Créditos (+)": [0.00 for _ in range(len(datas_iniciais))]
})

df_lancamentos = st.data_editor(
    df_lancamentos_iniciais, 
    key=f"tabela_lancamentos_{st.session_state.reset_contador}",
    num_rows="dynamic", 
    use_container_width=True,
    hide_index=True,
    column_config={
        "Data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
        "Débitos (-)": st.column_config.NumberColumn("Débitos (-)", format="R$ %.2f"),
        "Créditos (+)": st.column_config.NumberColumn("Créditos (+)", format="R$ %.2f")
    }
)

# --- BOTÕES DE AÇÃO (LADO A LADO) ---
st.markdown("<br>", unsafe_allow_html=True)
col_btn1, col_btn2 = st.columns([3, 1])

with col_btn1:
    btn_processar = st.button("PROCESSAR REVISÃO DE DÍVIDA", type="primary", use_container_width=True)
with col_btn2:
    st.button("🧹 Limpar Tabela", on_click=limpar_tabela, use_container_width=True)

# --- EXECUÇÃO DO CÁLCULO ---
if btn_processar:
    dic_indices = {row["Mês/Ano"]: Decimal(str(row["Índice (%)"] / 100)) for _, row in df_indices.iterrows()}
    dic_lancamentos = {}
    for _, row in df_lancamentos.iterrows():
        try:
            data_str = pd.to_datetime(row["Data"], format="%d/%m/%Y").strftime("%Y-%m-%d")
            dic_lancamentos[data_str] = {
                "debitos": Decimal(str(row["Débitos (-)"])),
                "creditos": Decimal(str(row["Créditos (+)"]))
            }
        except:
            pass

    memoria_calculo = []
    saldo_atual = Decimal(str(saldo_inicial))
    data_atual = data_inicial
    
    # Ajusta o nome da coluna no relatório dinamicamente
    col_taxa_nome = "Taxa de Atualização (%)" if indice_escolhido == "Sem Atualização (Apenas Juros)" else f"Taxa {indice_escolhido} (%)"
    
    while data_atual <= data_final:
        str_data = data_atual.strftime("%Y-%m-%d")
        mes_ano_anterior = (data_atual.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
        
        saldo_inicio_dia = saldo_atual
        valor_correcao = Decimal('0.00')
        percentual_aplicado = Decimal('0.00')
        
        if data_atual.day == 1 and mes_ano_anterior in dic_indices:
            percentual_aplicado = dic_indices[mes_ano_anterior]
            valor_correcao = saldo_atual * percentual_aplicado
            saldo_atual += valor_correcao
            
        lancamentos_dia = dic_lancamentos.get(str_data, {"debitos": Decimal('0.00'), "creditos": Decimal('0.00')})
        debitos = lancamentos_dia["debitos"]
        creditos = lancamentos_dia["creditos"]
        
        saldo_atual = saldo_atual + debitos - creditos
        
        memoria_calculo.append({
            "Data": data_atual.strftime("%d/%m/%Y"),
            "Saldo Anterior": float(saldo_inicio_dia),
            col_taxa_nome: float(percentual_aplicado * 100),
            "Correção (R$)": float(valor_correcao),
            "Débitos (R$)": float(debitos),
            "Créditos (R$)": float(creditos),
            "Saldo Final Dia": float(saldo_atual)
        })
        data_atual += timedelta(days=1)
    
    taxa_mensal_dec = Decimal(str(taxa_juros / 100))
    dias_totais_calc = (data_final - data_inicial).days
    
    if tipo_juros == "Compostos":
        taxa_periodo = (1 + taxa_mensal_dec) ** (Decimal(dias_totais_calc) / Decimal(30)) - 1
        valor_juros = saldo_atual * taxa_periodo
    else: 
        taxa_periodo = taxa_mensal_dec * (Decimal(dias_totais_calc) / Decimal(30))
        valor_juros = saldo_atual * taxa_periodo
        
    saldo_final_absoluto = saldo_atual + valor_juros
    
    # --- EXIBIÇÃO DOS RESULTADOS ---
    st.markdown("---")
    st.header("Resumo da Dívida")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Saldo Original Acumulado", f"R$ {float(saldo_atual):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    col2.metric(f"Juros Aplicados ({tipo_juros})", f"R$ {float(valor_juros):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    col3.metric("Dívida Final Recalculada", f"R$ {float(saldo_final_absoluto):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    col4.metric("Período (Dias)", dias_totais_calc)
    
    # --- BOTÕES DE EXPORTAÇÃO ---
    st.subheader("📥 Exportar Relatório Oficial")
    
    df_memoria = pd.DataFrame(memoria_calculo)
    df_resumo_export = pd.DataFrame({
        "Métrica": ["Saldo Original", "Juros Computados", "Dívida Final Recalculada", "Dias Totais"],
        "Valor": [float(saldo_atual), float(valor_juros), float(saldo_final_absoluto), dias_totais_calc]
    })
    
    dados_resumo_pdf = {"Original": float(saldo_atual), "Juros": float(valor_juros), "Final": float(saldo_final_absoluto), "Dias": dias_totais_calc}
    
    col_exp1, col_exp2 = st.columns(2)
    
    with col_exp1:
        dados_excel = gerar_excel(df_resumo_export, df_memoria)
        st.download_button(
            label="📊 Baixar Planilha Auditável (Excel)",
            data=dados_excel,
            file_name=f"Revisao_Divida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with col_exp2:
        dados_pdf = gerar_pdf(dados_resumo_pdf, df_memoria, indice_escolhido, tipo_juros, taxa_juros)
        st.download_button(
            label="📄 Baixar Laudo de Cálculos (PDF)",
            data=dados_pdf,
            file_name=f"Laudo_Pericial_Divida.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    st.markdown("---")
    st.header("Memória de Cálculo Detalhada")
    st.dataframe(
        df_memoria.style.format({
            "Saldo Anterior": "R$ {:.2f}",
            col_taxa_nome: "{:.2f}%",
            "Correção (R$)": "R$ {:.2f}",
            "Débitos (R$)": "R$ {:.2f}",
            "Créditos (R$)": "R$ {:.2f}",
            "Saldo Final Dia": "R$ {:.2f}"
        }), 
        use_container_width=True,
        height=400
    )
