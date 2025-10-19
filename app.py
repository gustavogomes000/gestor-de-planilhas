import streamlit as st
import pandas as pd
import os
import tempfile
from math import ceil
import zipfile
from io import BytesIO
import time


# Configurar a página
st.set_page_config(
    page_title="Andrezin - Ferramentas de Planilhas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS personalizado FINAL
st.markdown("""
<style>
    /* Fundo cinza permanente */
    .stApp {
        background-color: #f8f9fa !important;
    }
    
    /* Forçar tema claro */
    [data-testid="stAppViewContainer"] {
        background-color: #f8f9fa !important;
    }
    
    /* Remover elementos do Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {visibility: hidden;}
    
    /* Container principal */
    .main-container {
        background: white;
        border-radius: 20px;
        padding: 2rem;
        margin: 1rem auto;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        border: 1px solid #e9ecef;
        max-width: 1200px;
    }
    
    .main-header {
        font-size: 2.5rem;
        text-align: center;
        margin-bottom: 0.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #7D3C98 0%, #8E44AD 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .subtitle {
        text-align: center;
        color: #6c757d;
        font-size: 1.1rem;
        margin-bottom: 2rem;
        font-weight: 300;
    }
    
    /* UPLOADER COM FUNDO ROXO E LETRAS BRANCAS */
    .upload-section {
        background: linear-gradient(135deg, #7D3C98 0%, #8E44AD 100%);
        border-radius: 15px;
        padding: 2rem;
        margin: 1rem 0;
        color: white !important;
    }
    
    .upload-section h3 {
        color: white !important;
        margin-bottom: 1rem;
    }
    
    .upload-section .stFileUploader > label {
        color: white !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
    }
    
    .upload-section .stFileUploader > section > div {
        background-color: rgba(255, 255, 255, 0.1) !important;
        border: 2px dashed rgba(255, 255, 255, 0.5) !important;
        border-radius: 15px !important;
        color: white !important;
    }
    
    .upload-section .stFileUploader > section > div:hover {
        background-color: rgba(255, 255, 255, 0.2) !important;
        border-color: rgba(255, 255, 255, 0.8) !important;
    }
    
    .upload-section .stFileUploader > section > div > div > small {
        color: rgba(255, 255, 255, 0.9) !important;
        font-size: 0.8rem !important;
    }
    
    .upload-section .stFileUploader > section > div > div > button {
        background: rgba(255, 255, 255, 0.9) !important;
        color: #7D3C98 !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.4rem 1.2rem !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }
    
    .upload-section .stFileUploader > section > div > div > button:hover {
        background: white !important;
        color: #7D3C98 !important;
    }
    
    /* Cards modernos */
    .feature-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border-radius: 15px;
        padding: 2rem 1.5rem;
        margin: 1rem;
        border: 2px solid #e9ecef;
        text-align: center;
        transition: all 0.3s ease;
        cursor: pointer;
        height: 250px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        position: relative;
        overflow: hidden;
        box-shadow: 0 5px 15px rgba(0,0,0,0.08);
    }
    
    .feature-card:hover {
        transform: translateY(-8px) scale(1.02);
        box-shadow: 0 15px 30px rgba(125, 60, 152, 0.15);
        border-color: #7D3C98;
    }
    
    .feature-icon {
        font-size: 3.5rem;
        margin-bottom: 1.5rem;
        background: linear-gradient(135deg, #7D3C98 0%, #8E44AD 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .feature-title {
        font-size: 1.5rem;
        color: #2d3748;
        margin-bottom: 1rem;
        font-weight: 700;
    }
    
    .feature-description {
        font-size: 0.95rem;
        color: #4a5568;
        line-height: 1.5;
        font-weight: 400;
    }
    
    /* BOTÕES ROXOS */
    .stButton>button {
        background: linear-gradient(135deg, #7D3C98 0%, #8E44AD 100%) !important;
        color: white !important;
        border: none !important;
        padding: 0.7rem 1.5rem !important;
        border-radius: 10px !important;
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        width: 100% !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(125, 60, 152, 0.3) !important;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(125, 60, 152, 0.4) !important;
        background: linear-gradient(135deg, #8E44AD 0%, #9B59B6 100%) !important;
    }
    
    /* Botão voltar */
    .back-button {
        background: linear-gradient(135deg, #6c757d 0%, #495057 100%) !important;
        padding: 0.5rem 1rem !important;
        font-size: 0.85rem !important;
        margin-bottom: 1rem !important;
    }
    
    /* Barra de progresso roxa */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #7D3C98 0%, #8E44AD 100%) !important;
    }
    
    .progress-text {
        font-weight: 600;
        color: #7D3C98;
        min-width: 60px;
        font-size: 1rem;
    }
    
    /* Cards de informação ROXOS */
    .info-card {
        background: linear-gradient(135deg, #7D3C98 0%, #8E44AD 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        box-shadow: 0 5px 15px rgba(125, 60, 152, 0.3);
    }
    
    .success-card {
        background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        box-shadow: 0 5px 15px rgba(39, 174, 96, 0.3);
    }
    
    .warning-card {
        background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);
        color: white;
        padding: 1rem;
        border-radius: 12px;
        margin: 1rem 0;
        box-shadow: 0 5px 15px rgba(231, 76, 60, 0.3);
    }
    
    /* Seções expansíveis */
    .expandable-section {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 1rem 0;
        border-left: 4px solid #7D3C98;
        box-shadow: 0 3px 10px rgba(0,0,0,0.08);
    }
    
    /* Templates */
    .template-card {
        background: white;
        border: 2px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 0.5rem 0;
        cursor: pointer;
        transition: all 0.3s ease;
        text-align: center;
        height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    
    .template-card:hover {
        border-color: #7D3C98;
        transform: translateY(-3px);
        box-shadow: 0 5px 15px rgba(125, 60, 152, 0.1);
    }
    
    .template-card.selected {
        border-color: #7D3C98;
        background: linear-gradient(135deg, rgba(125, 60, 152, 0.1) 0%, rgba(142, 68, 173, 0.1) 100%);
        box-shadow: 0 5px 15px rgba(125, 60, 152, 0.15);
    }
    
    /* Animações */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .fade-in {
        animation: fadeIn 0.4s ease-out;
    }
    
    /* Status indicators */
    .status-indicator {
        display: inline-block;
        width: 10px;
        height: 10px;
        border-radius: 50%;
        margin-right: 6px;
    }
    
    .status-processing {
        background: #f39c12;
        animation: pulse 1.5s infinite;
    }
    
    .status-completed {
        background: #27ae60;
    }
    
    @keyframes pulse {
        0% { opacity: 1; }
        50% { opacity: 0.5; }
        100% { opacity: 1; }
    }
    
    /* Textos sempre legíveis */
    .stMetric {
        color: #2d3748 !important;
    }
    
    .stMetric > div[data-testid="stMetricLabel"] > div {
        color: #4a5568 !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }
    
    .stMetric > div[data-testid="stMetricValue"] > div {
        color: #2d3748 !important;
        font-weight: 700 !important;
        font-size: 1.1rem !important;
    }
    
    .stRadio > label {
        color: #2d3748 !important;
        font-weight: 500 !important;
    }
    
    .stNumberInput > label {
        color: #2d3748 !important;
        font-weight: 600 !important;
    }
    
    .stNumberInput > div > div > input {
        color: #2d3748 !important;
        background: white !important;
        border: 1px solid #e2e8f0 !important;
    }
    
    /* Textos gerais */
    .stMarkdown {
        color: #2d3748 !important;
    }
    
    h1, h2, h3, h4, h5, h6 {
        color: #2d3748 !important;
        margin-bottom: 0.75rem !important;
    }
    
    p, div, span {
        color: #2d3748 !important;
    }
    
    /* Ajustes para inputs */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stTextArea > div > div > textarea {
        background: white !important;
        color: #2d3748 !important;
        border: 1px solid #e2e8f0 !important;
    }
    
    /* Ajuste para os tabs */
    .stTabs [data-baseweb="tab-list"] {
        background: white !important;
        border-bottom: 1px solid #e2e8f0 !important;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: white !important;
        color: #2d3748 !important;
    }
    
    .stTabs [aria-selected="true"] {
        background: white !important;
        color: #7D3C98 !important;
    }
    
    /* Forçar cores em todos os elementos */
    div[data-testid="stVerticalBlock"] {
        background: transparent !important;
    }
    
    section[data-testid="stFileUploadDropzone"] {
        background: rgba(255, 255, 255, 0.1) !important;
    }
</style>
""", unsafe_allow_html=True)

# Função para dividir planilhas com barra de progresso
def split_spreadsheet_with_progress(file_path, progress_bar, progress_text, status_text, lines_per_file=None, num_files=None):
    try:
        file_ext = os.path.splitext(file_path)[1].lower()
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        
        # Criar diretório temporário
        temp_dir = tempfile.mkdtemp()
        
        # Determinar se é CSV ou Excel
        is_csv = file_ext in ['.csv', '.txt']
        
        status_text.text("📖 Lendo arquivo...")
        progress_bar.progress(5)
        progress_text.text("5%")

        # Ler o arquivo completo
        if is_csv:
            # Para CSV, contar linhas primeiro
            with open(file_path, 'r', encoding='utf-8') as f:
                total_lines = sum(1 for line in f) - 1
            
            df = pd.read_csv(file_path, encoding='utf-8')
        else:
            df = pd.read_excel(file_path, engine='openpyxl')
            total_lines = len(df)
        
        progress_bar.progress(15)
        progress_text.text("15%")
        status_text.text(f"📊 Arquivo lido: {total_lines} linhas encontradas")
        time.sleep(0.3)
        
        # Calcular linhas por arquivo
        if num_files:
            lines_per_file = ceil(total_lines / num_files)
        elif not lines_per_file:
            raise ValueError("Defina o número de linhas ou arquivos")
        
        total_parts = ceil(total_lines / lines_per_file)
        generated_files = []
        
        status_text.text("✂️ Iniciando divisão da planilha...")
        progress_bar.progress(20)
        progress_text.text("20%")
        
        for i in range(total_parts):
            start_idx = i * lines_per_file
            end_idx = min((i + 1) * lines_per_file, total_lines)
            part_df = df.iloc[start_idx:end_idx]
            
            # Atualizar progresso
            progress_percent = 20 + (i / total_parts) * 70
            progress_bar.progress(int(progress_percent))
            progress_text.text(f"{int(progress_percent)}%")
            status_text.text(f"📝 Criando parte {i+1} de {total_parts}...")
            
            # Salvar a parte
            if is_csv:
                output_file = os.path.join(temp_dir, f"{base_name}_parte_{i+1:02d}.csv")
                part_df.to_csv(output_file, index=False, encoding='utf-8')
            else:
                output_file = os.path.join(temp_dir, f"{base_name}_parte_{i+1:02d}.xlsx")
                part_df.to_excel(output_file, index=False, engine='openpyxl')
            
            generated_files.append(output_file)
            time.sleep(0.05)
        
        progress_bar.progress(95)
        progress_text.text("95%")
        status_text.text("📦 Preparando arquivos para download...")
        time.sleep(0.3)
        
        progress_bar.progress(100)
        progress_text.text("100%")
        status_text.text("✅ Processamento concluído!")
        
        return generated_files, temp_dir, total_parts
        
    except Exception as e:
        progress_bar.progress(0)
        progress_text.text("0%")
        status_text.text("❌ Erro no processamento")
        raise Exception(f"Erro ao dividir planilha: {str(e)}")

# Função para juntar planilhas com templates
def merge_spreadsheets_with_template(uploaded_files, progress_bar, progress_text, status_text, template_type):
    try:
        status_text.text("🔍 Verificando templates...")
        progress_bar.progress(10)
        progress_text.text("10%")
        
        # Carregar template base
        template_path = f"modelos/{template_type}ModeloExcel.xlsx"
        
        if not os.path.exists(template_path):
            # Usar estrutura da primeira planilha
            first_file = uploaded_files[0]
            file_ext = os.path.splitext(first_file.name)[1].lower()
            
            if file_ext in ['.csv', '.txt']:
                template_df = pd.read_csv(first_file, encoding='utf-8', nrows=0)
            else:
                template_df = pd.read_excel(first_file, engine='openpyxl', nrows=0)
        else:
            template_df = pd.read_excel(template_path, engine='openpyxl', nrows=0)
        
        progress_bar.progress(30)
        progress_text.text("30%")
        status_text.text("📚 Lendo planilhas...")
        
        dataframes = []
        total_files = len(uploaded_files)
        
        for i, uploaded_file in enumerate(uploaded_files):
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            
            progress_percent = 30 + (i / total_files) * 50
            progress_bar.progress(int(progress_percent))
            progress_text.text(f"{int(progress_percent)}%")
            status_text.text(f"📖 Processando arquivo {i+1} de {total_files}...")
            
            if file_ext in ['.csv', '.txt']:
                df = pd.read_csv(uploaded_file, encoding='utf-8')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            # Reordenar colunas conforme template
            df = df.reindex(columns=template_df.columns, fill_value="")
            dataframes.append(df)
            
            time.sleep(0.1)
        
        progress_bar.progress(85)
        progress_text.text("85%")
        status_text.text("🔗 Combinando planilhas...")
        
        # Combinar verticalmente
        merged_df = pd.concat(dataframes, axis=0, ignore_index=True)
        
        progress_bar.progress(95)
        progress_text.text("95%")
        status_text.text("🎨 Aplicando formatação...")
        time.sleep(0.3)
        
        progress_bar.progress(100)
        progress_text.text("100%")
        status_text.text("✅ Junção concluída!")
        
        return merged_df, template_df.columns.tolist()
        
    except Exception as e:
        progress_bar.progress(0)
        progress_text.text("0%")
        status_text.text("❌ Erro na junção")
        raise Exception(f"Erro ao juntar planilhas: {str(e)}")

# Página inicial
def main_page():
    st.markdown("""
    <div class="main-container fade-in">
        <div class="main-header">Planilhas Setup Auvo</div>
        <div class="subtitle">Organize sua gestão de setups</div>
    """, unsafe_allow_html=True)
    
    # Cards principais
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">📊</div>
            <div class="feature-title">Dividir Planilhas</div>
            <div class="feature-description">
                Transforme planilhas extensas em arquivos menores e gerenciáveis. 
                
            
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚀 Iniciar Divisão", key="split_btn", use_container_width=True):
            st.session_state.page = "split"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">🔗</div>
            <div class="feature-title">Juntar Planilhas</div>
            <div class="feature-description">
                Unifique múltiplas planilhas seguindo templates padronizados. 
                Perfeito para consolidação de dados e relatórios.
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚀 Iniciar Junção", key="merge_btn", use_container_width=True):
            st.session_state.page = "merge"
            st.rerun()
    
    # Seção de informações
    

# Página de divisão com progresso
def split_page():
    st.markdown("""
    <div class="main-container fade-in">
        <div class="main-header">Divisor de Planilhas</div>
        <div class="subtitle">Divida planilhas grandes em partes gerenciáveis</div>
    """, unsafe_allow_html=True)
    
    # Botão voltar
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Voltar", key="back_split", use_container_width=True, type="secondary"):
            st.session_state.page = "main"
            st.rerun()
    
    # Seção de upload COM FUNDO ROXO E LETRAS BRANCAS
    with st.container():
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        st.subheader("📁 Upload da Planilha")
        uploaded_file = st.file_uploader(
            "Selecione o arquivo para dividir",
            type=['xlsx', 'xls', 'csv', 'txt'],
            help="Arraste ou clique para selecionar o arquivo"
        )
        st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        # Informações do arquivo
        with st.container():
            st.markdown('<div class="expandable-section compact-section">', unsafe_allow_html=True)
            st.subheader("📋 Informações do Arquivo")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Nome", uploaded_file.name)
            with col2:
                st.metric("Tamanho", f"{uploaded_file.size / 1024:.1f} KB")
            with col3:
                st.metric("Tipo", uploaded_file.type.split('/')[-1].upper())
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Configurações de divisão
        with st.container():
            st.markdown('<div class="expandable-section compact-section">', unsafe_allow_html=True)
            st.subheader("⚙️ Configurações de Divisão")
            
            col1, col2 = st.columns(2)
            with col1:
                split_method = st.radio(
                    "Método de Divisão:",
                    ["Dividir por número de linhas", "Dividir em número de arquivos"],
                    key="split_method"
                )
            
            with col2:
                if split_method == "Dividir por número de linhas":
                    lines_per_file = st.number_input(
                        "Linhas por arquivo:",
                        min_value=1,
                        value=5000,
                        help="Cada arquivo terá esta quantidade de linhas",
                        key="lines_input"
                    )
                    num_files = None
                else:
                    num_files = st.number_input(
                        "Número de arquivos:",
                        min_value=2,
                        value=5,
                        help="A planilha será dividida nesta quantidade de partes",
                        key="files_input"
                    )
                    lines_per_file = None
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Área de processamento
        if st.button("🎯 Iniciar Processamento", type="primary", use_container_width=True):
            with st.container():
                st.markdown('<div class="expandable-section">', unsafe_allow_html=True)
                st.subheader("🔄 Processamento")
                
                # Barra de progresso com porcentagem
                progress_bar = st.progress(0)
                progress_col1, progress_col2 = st.columns([3, 1])
                with progress_col1:
                    status_text = st.empty()
                with progress_col2:
                    progress_text = st.empty()
                    progress_text.text("0%")
                
                try:
                    # Salvar arquivo temporariamente
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        tmp_path = tmp_file.name
                    
                    # Processar com barra de progresso
                    generated_files, temp_dir, total_parts = split_spreadsheet_with_progress(
                        tmp_path, progress_bar, progress_text, status_text, lines_per_file, num_files
                    )
                    
                    # Criar ZIP
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                        for file_path in generated_files:
                            zip_file.write(file_path, os.path.basename(file_path))
                    zip_buffer.seek(0)
                    
                    # Limpar arquivo temporário
                    os.unlink(tmp_path)
                    
                    # Resultados
                    st.markdown(f"""
                    <div class="success-card">
                        <h3 style="margin:0; color:white; font-size: 1.3rem;">✅ Processamento Concluído!</h3>
                        <p style="margin:0.5rem 0 0 0; color:white; font-size: 0.95rem;">
                        A planilha foi dividida em <strong>{total_parts}</strong> partes com sucesso.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Download
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="📥 Baixar Todas as Partes (ZIP)",
                            data=zip_buffer,
                            file_name=f"{os.path.splitext(uploaded_file.name)[0]}_partes.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                    
                    # Preview
                    st.subheader("👀 Preview dos Resultados")
                    tabs = st.tabs([f"Parte {i+1}" for i in range(min(3, len(generated_files)))])
                    
                    for i, tab in enumerate(tabs):
                        with tab:
                            if i < len(generated_files):
                                file_path = generated_files[i]
                                file_ext = os.path.splitext(file_path)[1].lower()
                                
                                try:
                                    if file_ext == '.csv':
                                        df_preview = pd.read_csv(file_path)
                                    else:
                                        df_preview = pd.read_excel(file_path, engine='openpyxl')
                                    
                                    st.write(f"**Parte {i+1}** - {len(df_preview)} linhas × {len(df_preview.columns)} colunas")
                                    st.dataframe(df_preview.head(6), use_container_width=True)
                                    
                                except Exception as e:
                                    st.error(f"Erro ao visualizar: {str(e)}")
                    
                    # Limpeza
                    for file_path in generated_files:
                        if os.path.exists(file_path):
                            os.unlink(file_path)
                    if os.path.exists(temp_dir):
                        os.rmdir(temp_dir)
                        
                except Exception as e:
                    st.error(f"❌ Erro no processamento: {str(e)}")
                
                st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

# Página de junção com templates
def merge_page():
    st.markdown("""
    <div class="main-container fade-in">
        <div class="main-header">Padronizar Planilhas</div>
        <div class="subtitle">Combine múltiplas planilhas com base nas planilhas modélos do Auvo</div>
    """, unsafe_allow_html=True)
    
    # Botão voltar
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Voltar", key="back_merge", use_container_width=True, type="secondary"):
            st.session_state.page = "main"
            st.rerun()
    
    # Seleção de template
    with st.container():
        st.markdown('<div class="expandable-section compact-section">', unsafe_allow_html=True)
        st.subheader("🎯 Selecione o Tipo de Planilha")
        
        template_options = {
            "Clientes": "Cadastro de clientes com dados financeiros",
            "Equipamentos": "Controle de equipamentos e garantias", 
            "Produtos": "Gestão de produtos e estoque",
            "Questionarios": "Formulários e pesquisas estruturadas"
        }
        
        cols = st.columns(4)
        selected_template = st.session_state.get('selected_template', None)
        
        for i, (template_name, description) in enumerate(template_options.items()):
            with cols[i]:
                is_selected = selected_template == template_name
                css_class = "template-card selected" if is_selected else "template-card"
                
                st.markdown(f"""
                <div class="{css_class}">
                    <h4 style="margin:0; font-size: 1.1rem;">{template_name}</h4>
                    <p style="margin:0.5rem 0 0 0; font-size: 0.8rem; color: #6c757d;">{description}</p>
                </div>
                """, unsafe_allow_html=True)
                
                if st.button("Selecionar", key=f"template_{i}", use_container_width=True):
                    st.session_state.selected_template = template_name
                    st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    if st.session_state.get('selected_template'):
        selected_template = st.session_state.selected_template
        
        # Upload de arquivos COM FUNDO ROXO E LETRAS BRANCAS
        with st.container():
            st.markdown('<div class="upload-section">', unsafe_allow_html=True)
            st.subheader(f"📁 Upload das Planilhas de {selected_template}")
            
            uploaded_files = st.file_uploader(
                f"Selecione as planilhas de {selected_template} para juntar",
                type=['xlsx', 'xls', 'csv', 'txt'],
                accept_multiple_files=True,
                help="Selecione duas ou mais planilhas do mesmo tipo"
            )
            
            if uploaded_files:
                st.write(f"**Arquivos selecionados:** {len(uploaded_files)}")
                
                for i, file in enumerate(uploaded_files):
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.write(f"**{i+1}. {file.name}**")
                    with col2:
                        st.write(f"`{file.size / 1024:.1f} KB`")
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Processamento
        if uploaded_files and len(uploaded_files) > 1:
            if st.button("🎯 Iniciar Junção", type="primary", use_container_width=True):
                with st.container():
                    st.markdown('<div class="expandable-section">', unsafe_allow_html=True)
                    st.subheader("🔄 Processamento de Junção")
                    
                    # Barra de progresso com porcentagem
                    progress_bar = st.progress(0)
                    progress_col1, progress_col2 = st.columns([3, 1])
                    with progress_col1:
                        status_text = st.empty()
                    with progress_col2:
                        progress_text = st.empty()
                        progress_text.text("0%")
                    
                    try:
                        merged_df, template_columns = merge_spreadsheets_with_template(
                            uploaded_files, progress_bar, progress_text, status_text, selected_template
                        )
                        
                        st.markdown("""
                        <div class="success-card">
                            <h3 style="margin:0; color:white; font-size: 1.3rem;">✅ Junção Concluída!</h3>
                            <p style="margin:0.5rem 0 0 0; color:white; font-size: 0.95rem;">
                            <strong>{}</strong> planilhas combinadas em um único arquivo com <strong>{}</strong> linhas e <strong>{}</strong> colunas.
                            </p>
                        </div>
                        """.format(len(uploaded_files), len(merged_df), len(merged_df.columns)), unsafe_allow_html=True)
                        
                        # Preview
                        st.subheader("👀 Preview do Resultado")
                        st.dataframe(merged_df.head(10), use_container_width=True)
                        
                        # Download
                        st.subheader("📥 Download do Arquivo Consolidado")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            excel_buffer = BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                merged_df.to_excel(writer, index=False, sheet_name=f'{selected_template}_Consolidado')
                            excel_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Excel Formatado",
                                data=excel_buffer,
                                file_name=f"{selected_template.lower()}_consolidado.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        with col2:
                            csv_buffer = BytesIO()
                            csv_buffer.write(merged_df.to_csv(index=False, encoding='utf-8').encode('utf-8'))
                            csv_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 CSV",
                                data=csv_buffer,
                                file_name=f"{selected_template.lower()}_consolidado.csv",
                                mime="text/csv",
                                use_container_width=True
                            )
                            
                    except Exception as e:
                        st.error(f"❌ Erro na junção: {str(e)}")
                    
                    st.markdown('</div>', unsafe_allow_html=True)
        elif uploaded_files and len(uploaded_files) == 1:
            st.markdown("""
            <div class="warning-card">
                <h4 style="margin:0; color:white; font-size: 1.1rem;">⚠️ Atenção</h4>
                <p style="margin:0.5rem 0 0 0; color:white; font-size: 0.9rem;">
                Selecione pelo menos <strong>2 planilhas</strong> para realizar a junção.
                </p>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

# Gerenciamento de estado
if 'page' not in st.session_state:
    st.session_state.page = 'main'

if 'selected_template' not in st.session_state:
    st.session_state.selected_template = None

# Navegação
if st.session_state.page == 'main':
    main_page()
elif st.session_state.page == 'split':
    split_page()
elif st.session_state.page == 'merge':
    merge_page()