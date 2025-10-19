import streamlit as st
import pandas as pd
import os
import tempfile
from math import ceil
import zipfile
from io import BytesIO
import time
import openpyxl
import numpy as np 
from datetime import datetime

# Configurar a página
st.set_page_config(
    page_title="Andrezin - Ferramentas de Planilhas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS personalizado PROFISSIONAL
st.markdown("""
<style>
    /* Fundo profissional */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        min-height: 100vh;
    }
    
    /* Container principal com vidro */
    .main-container {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(20px);
        border-radius: 24px;
        padding: 2.5rem;
        margin: 2rem auto;
        box-shadow: 0 20px 60px rgba(0,0,0,0.1);
        border: 1px solid rgba(255, 255, 255, 0.3);
        max-width: 1400px;
        position: relative;
        overflow: hidden;
    }
    
    .main-container::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(circle, rgba(255,255,255,0.1) 1px, transparent 1px);
        background-size: 20px 20px;
        animation: float 20s infinite linear;
        pointer-events: none;
    }
    
    @keyframes float {
        0% { transform: translate(0, 0) rotate(0deg); }
        100% { transform: translate(-20px, -20px) rotate(360deg); }
    }
    
    .main-header {
        font-size: 3rem;
        text-align: center;
        margin-bottom: 0.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    
    .subtitle {
        text-align: center;
        color: #6c757d;
        font-size: 1.2rem;
        margin-bottom: 3rem;
        font-weight: 300;
    }
    
    /* Cards de features modernos */
    .feature-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        gap: 2rem;
        margin: 2rem 0;
    }
    
    .feature-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border-radius: 20px;
        padding: 2.5rem 2rem;
        border: 2px solid rgba(255, 255, 255, 0.8);
        text-align: center;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        cursor: pointer;
        position: relative;
        overflow: hidden;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
    }
    
    .feature-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent);
        transition: left 0.6s;
    }
    
    .feature-card:hover::before {
        left: 100%;
    }
    
    .feature-card:hover {
        transform: translateY(-12px) scale(1.02);
        box-shadow: 0 25px 50px rgba(102, 126, 234, 0.15);
        border-color: #667eea;
    }
    
    .feature-icon {
        font-size: 4rem;
        margin-bottom: 1.5rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        filter: drop-shadow(0 4px 8px rgba(0,0,0,0.1));
    }
    
    .feature-title {
        font-size: 1.6rem;
        color: #2d3748;
        margin-bottom: 1rem;
        font-weight: 700;
    }
    
    .feature-description {
        font-size: 1rem;
        color: #4a5568;
        line-height: 1.6;
        font-weight: 400;
    }
    
    /* Upload section moderna */
    .upload-section {
        background: linear-gradient(135deg, rgba(102, 126, 234, 0.1) 0%, rgba(118, 75, 162, 0.1) 100%);
        border: 2px dashed rgba(102, 126, 234, 0.3);
        border-radius: 20px;
        padding: 3rem 2rem;
        margin: 2rem 0;
        text-align: center;
        transition: all 0.3s ease;
    }
    
    .upload-section:hover {
        border-color: rgba(102, 126, 234, 0.6);
        background: linear-gradient(135deg, rgba(102, 126, 234, 0.15) 0%, rgba(118, 75, 162, 0.15) 100%);
    }
    
    /* Botões modernos */
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        border: none !important;
        padding: 1rem 2rem !important;
        border-radius: 15px !important;
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        width: 100% !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.3) !important;
        position: relative;
        overflow: hidden;
    }
    
    .stButton>button::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
        transition: left 0.5s;
    }
    
    .stButton>button:hover::before {
        left: 100%;
    }
    
    .stButton>button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 35px rgba(102, 126, 234, 0.4) !important;
    }
    
    /* Seções expansíveis */
    .expandable-section {
        background: rgba(255, 255, 255, 0.8);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 2rem;
        margin: 1.5rem 0;
        border-left: 5px solid #667eea;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    
    .expandable-section:hover {
        box-shadow: 0 12px 40px rgba(0,0,0,0.15);
        transform: translateY(-2px);
    }
    
    /* Cards de status */
    .status-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 2rem;
        border-radius: 20px;
        margin: 1.5rem 0;
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.3);
        position: relative;
        overflow: hidden;
    }
    
    .status-card::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -50%;
        width: 100%;
        height: 100%;
        background: rgba(255,255,255,0.1);
        transform: rotate(30deg);
    }
    
    .success-card {
        background: linear-gradient(135deg, #00b09b 0%, #96c93d 100%);
        box-shadow: 0 10px 30px rgba(0, 176, 155, 0.3);
    }
    
    .warning-card {
        background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);
        box-shadow: 0 10px 30px rgba(255, 107, 107, 0.3);
    }
    
    /* Barra de progresso animada */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%) !important;
        border-radius: 10px;
        animation: pulse 2s infinite;
    }
    
    @keyframes pulse {
        0% { opacity: 1; }
        50% { opacity: 0.8; }
        100% { opacity: 1; }
    }
    
    /* Animações de entrada */
    @keyframes slideInUp {
        from {
            opacity: 0;
            transform: translateY(30px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .slide-in {
        animation: slideInUp 0.6s ease-out;
    }
    
    /* Melhorias para tabelas */
    .dataframe {
        border-radius: 15px !important;
        overflow: hidden !important;
        box-shadow: 0 5px 15px rgba(0,0,0,0.1) !important;
    }
    
    /* Ajustes para mobile */
    @media (max-width: 768px) {
        .main-container {
            margin: 1rem;
            padding: 1.5rem;
        }
        
        .main-header {
            font-size: 2rem;
        }
        
        .feature-grid {
            grid-template-columns: 1fr;
        }
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

# Função para juntar/formatar planilhas
def process_spreadsheets_with_template(uploaded_files, progress_bar, progress_text, status_text, template_type):
    try:
        status_text.text("🔍 Verificando templates...")
        progress_bar.progress(10)
        progress_text.text("10%")
        
        # Carregar template base se existir
        template_path = f"modelos/{template_type}ModeloExcel.xlsx"
        
        if os.path.exists(template_path):
            template_df = pd.read_excel(template_path, engine='openpyxl', nrows=0)
            template_columns = template_df.columns.tolist()
        else:
            # Usar estrutura da primeira planilha como template
            first_file = uploaded_files[0]
            file_ext = os.path.splitext(first_file.name)[1].lower()
            
            if file_ext in ['.csv', '.txt']:
                template_df = pd.read_csv(first_file, encoding='utf-8', nrows=0)
            else:
                template_df = pd.read_excel(first_file, engine='openpyxl', nrows=0)
            template_columns = template_df.columns.tolist()
        
        progress_bar.progress(30)
        progress_text.text("30%")
        status_text.text("📚 Processando planilhas...")
        
        dataframes = []
        total_files = len(uploaded_files)
        
        for i, uploaded_file in enumerate(uploaded_files):
            file_ext = os.path.splitext(uploaded_file.name)[1].lower()
            
            progress_percent = 30 + (i / total_files) * 50
            progress_bar.progress(int(progress_percent))
            progress_text.text(f"{int(progress_percent)}%")
            
            if total_files == 1:
                status_text.text(f"🎨 Formatando planilha...")
            else:
                status_text.text(f"📖 Processando arquivo {i+1} de {total_files}...")
            
            if file_ext in ['.csv', '.txt']:
                df = pd.read_csv(uploaded_file, encoding='utf-8')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            # Reordenar colunas conforme template e preencher colunas faltantes
            for col in template_columns:
                if col not in df.columns:
                    df[col] = ""
            
            df = df.reindex(columns=template_columns, fill_value="")
            dataframes.append(df)
            
            time.sleep(0.1)
        
        progress_bar.progress(85)
        progress_text.text("85%")
        
        if total_files == 1:
            status_text.text("🎨 Aplicando formatação...")
            merged_df = dataframes[0]  # Apenas formata a única planilha
        else:
            status_text.text("🔗 Combinando planilhas...")
            # Combinar verticalmente
            merged_df = pd.concat(dataframes, axis=0, ignore_index=True)
        
        progress_bar.progress(95)
        progress_text.text("95%")
        status_text.text("✨ Finalizando...")
        time.sleep(0.3)
        
        progress_bar.progress(100)
        progress_text.text("100%")
        
        if total_files == 1:
            status_text.text("✅ Formatação concluída!")
        else:
            status_text.text("✅ Junção concluída!")
        
        return merged_df, template_columns
        
    except Exception as e:
        progress_bar.progress(0)
        progress_text.text("0%")
        status_text.text("❌ Erro no processamento")
        raise Exception(f"Erro ao processar planilhas: {str(e)}")

# Página inicial
def main_page():
    st.markdown("""
    <div class="main-container slide-in">
        <div class="main-header">Andrezin - Ferramentas Auvo</div>
        <div class="subtitle">Soluções profissionais para gestão de planilhas</div>
    """, unsafe_allow_html=True)
    
    # Cards principais
    st.markdown('<div class="feature-grid">', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">📊</div>
            <div class="feature-title">Dividir Planilhas</div>
            <div class="feature-description">
                Transforme planilhas extensas em arquivos menores e gerenciáveis. 
                Ideal para processamento de grandes volumes de dados.
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚀 Iniciar Divisão", key="split_btn", use_container_width=True):
            st.session_state.page = "split"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="feature-card">
            <div class="feature-icon">🔗</div>
            <div class="feature-title">Formatar/Juntar Planilhas</div>
            <div class="feature-description">
                Formate planilhas individuais ou combine múltiplas planilhas seguindo templates padronizados.
                Perfeito para padronização e consolidação.
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚀 Iniciar Formatação", key="merge_btn", use_container_width=True):
            st.session_state.page = "merge"
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Estatísticas e informações
    with st.container():
        st.markdown('<div class="expandable-section">', unsafe_allow_html=True)
        st.subheader("📈 Estatísticas do Sistema")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Tipos de Template", "4", "Clientes, Equipamentos, Produtos, Questionários")
        with col2:
            st.metric("Formatos Suportados", "5", "XLSX, XLS, CSV, TXT")
        with col3:
            st.metric("Processamento", "Ilimitado", "Sem restrições de tamanho")
        with col4:
            st.metric("Atualização", "2024", "Versão mais recente")
        
        st.markdown('</div>', unsafe_allow_html=True)

# Página de divisão
def split_page():
    st.markdown("""
    <div class="main-container slide-in">
        <div class="main-header">Divisor de Planilhas</div>
        <div class="subtitle">Divida planilhas grandes em partes gerenciáveis</div>
    """, unsafe_allow_html=True)
    
    # Botão voltar
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Voltar", key="back_split", use_container_width=True, type="secondary"):
            st.session_state.page = "main"
            st.rerun()
    
    # Seção de upload
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
            st.markdown('<div class="expandable-section">', unsafe_allow_html=True)
            st.subheader("📋 Informações do Arquivo")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Nome", uploaded_file.name)
            with col2:
                st.metric("Tamanho", f"{uploaded_file.size / 1024:.1f} KB")
            with col3:
                st.metric("Tipo", uploaded_file.type.split('/')[-1].upper())
            with col4:
                st.metric("Status", "Pronto para processar", "✅")
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Configurações de divisão
        with st.container():
            st.markdown('<div class="expandable-section">', unsafe_allow_html=True)
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
                        <h3 style="margin:0; color:white; font-size: 1.5rem;">✅ Processamento Concluído!</h3>
                        <p style="margin:0.5rem 0 0 0; color:white; font-size: 1rem;">
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
                                    st.dataframe(df_preview.head(8), use_container_width=True)
                                    
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

# Página de junção/formatação
def merge_page():
    st.markdown("""
    <div class="main-container slide-in">
        <div class="main-header">Formatar e Juntar Planilhas</div>
        <div class="subtitle">Padronize planilhas individuais ou combine múltiplas planilhas</div>
    """, unsafe_allow_html=True)
    
    # Botão voltar
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Voltar", key="back_merge", use_container_width=True, type="secondary"):
            st.session_state.page = "main"
            st.rerun()
    
    # Seleção de template
    with st.container():
        st.markdown('<div class="expandable-section">', unsafe_allow_html=True)
        st.subheader("🎯 Selecione o Tipo de Planilha")
        
        template_options = {
            "Clientes": "Cadastro de clientes com dados financeiros",
            "Equipamentos": "Controle de equipamentos e garantias", 
            "Produtos": "Gestão de produtos e estoque",
            "Questionarios": "Formulários e pesquisas estruturadas"
        }
        
        # Usar st.radio para seleção única
        selected_template = st.radio(
            "Selecione o template:",
            options=list(template_options.keys()),
            format_func=lambda x: f"**{x}** - {template_options[x]}",
            key="template_radio"
        )
        
        # Atualizar session state
        st.session_state.selected_template = selected_template
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    if st.session_state.get('selected_template'):
        selected_template = st.session_state.selected_template
        
        # Upload de arquivos
        with st.container():
            st.markdown('<div class="upload-section">', unsafe_allow_html=True)
            st.subheader(f"📁 Upload das Planilhas de {selected_template}")
            
            uploaded_files = st.file_uploader(
                f"Selecione as planilhas de {selected_template} para formatar/juntar",
                type=['xlsx', 'xls', 'csv', 'txt'],
                accept_multiple_files=True,
                help="Selecione uma ou mais planilhas do mesmo tipo"
            )
            
            if uploaded_files:
                st.write(f"**Arquivos selecionados:** {len(uploaded_files)}")
                
                for i, file in enumerate(uploaded_files):
                    col1, col2, col3 = st.columns([3, 1, 1])
                    with col1:
                        st.write(f"**{i+1}. {file.name}**")
                    with col2:
                        st.write(f"`{file.size / 1024:.1f} KB`")
                    with col3:
                        st.write("✅ Pronto")
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Processamento
        if uploaded_files:
            process_label = "🎨 Formatando..." if len(uploaded_files) == 1 else "🔗 Juntando..."
            button_label = "🎯 Iniciar Formatação" if len(uploaded_files) == 1 else "🎯 Iniciar Junção"
            
            if st.button(button_label, type="primary", use_container_width=True, key="merge_process"):
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
                        merged_df, template_columns = process_spreadsheets_with_template(
                            uploaded_files, progress_bar, progress_text, status_text, selected_template
                        )
                        
                        success_message = "formatação" if len(uploaded_files) == 1 else "junção"
                        st.markdown(f"""
                        <div class="success-card">
                            <h3 style="margin:0; color:white; font-size: 1.5rem;">✅ {success_message.title()} Concluída!</h3>
                            <p style="margin:0.5rem 0 0 0; color:white; font-size: 1rem;">
                            <strong>{len(uploaded_files)}</strong> planilha(s) processada(s) com <strong>{len(merged_df)}</strong> linhas e <strong>{len(merged_df.columns)}</strong> colunas.
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Preview
                        st.subheader("👀 Preview do Resultado")
                        col1, col2 = st.columns([1, 3])
                        with col1:
                            st.metric("Total de Linhas", len(merged_df))
                            st.metric("Total de Colunas", len(merged_df.columns))
                            st.metric("Arquivos Processados", len(uploaded_files))
                        
                        with col2:
                            st.dataframe(merged_df.head(12), use_container_width=True)
                        
                        # Download
                        st.subheader("📥 Download do Arquivo Processado")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            excel_buffer = BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                merged_df.to_excel(writer, index=False, sheet_name=f'{selected_template}_Padronizado')
                            excel_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Excel Formatado",
                                data=excel_buffer,
                                file_name=f"{selected_template.lower()}_padronizado.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key="download_excel"
                            )
                        
                        with col2:
                            csv_buffer = BytesIO()
                            csv_buffer.write(merged_df.to_csv(index=False, encoding='utf-8').encode('utf-8'))
                            csv_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 CSV",
                                data=csv_buffer,
                                file_name=f"{selected_template.lower()}_padronizado.csv",
                                mime="text/csv",
                                use_container_width=True,
                                key="download_csv"
                            )
                            
                    except Exception as e:
                        st.error(f"❌ Erro no processamento: {str(e)}")
                    
                    st.markdown('</div>', unsafe_allow_html=True)
    
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
