import streamlit as st
import pandas as pd
import os
import tempfile
import shutil
import time
import traceback
from pdf_extractor_robust import processar_pdf_com_markup

# Configuração da página
st.set_page_config(
    page_title="Gerador de Catálogo com Preços",
    page_icon="📊",
    layout="centered"
)

# Título e descrição
st.title("Gerador de Catálogo com Preços")
st.write("Faça upload do seu catálogo PDF e do arquivo Excel com os preços para gerar um novo catálogo com preços.")

# Upload do arquivo PDF
pdf_file = st.file_uploader("Selecione o arquivo PDF do catálogo", type=["pdf"])

# Upload do arquivo Excel
excel_file = st.file_uploader("Selecione o arquivo Excel com os preços", type=["xlsx", "xls", "xlsm"])

# Definição do markup
markup = st.number_input("Defina o valor do markup", min_value=1.0, max_value=10.0, value=2.0, step=0.1)

# Seleção de cor para o rodapé
st.subheader("Personalização")
cor_option = st.selectbox(
    "Escolha a cor do rodapé:",
    ["Cinza", "Azul", "Verde", "Vermelho", "Preto", "Roxo"]
)

# Mapeamento de cores
cores = {
    "Cinza": (128, 128, 128),
    "Azul": (41, 98, 255),
    "Verde": (0, 128, 0),
    "Vermelho": (255, 0, 0),
    "Preto": (0, 0, 0),
    "Roxo": (128, 0, 128)
}

# Cor personalizada (opcional)
usar_cor_personalizada = st.checkbox("Usar cor personalizada")
if usar_cor_personalizada:
    col1, col2, col3 = st.columns(3)
    with col1:
        r = st.slider("R", 0, 255, 128)
    with col2:
        g = st.slider("G", 0, 255, 128)
    with col3:
        b = st.slider("B", 0, 255, 128)
    
    cor_tarja = (r, g, b)
    # Mostrar amostra da cor
    st.markdown(
        f"""
        <div style="background-color: rgb({r}, {g}, {b}); 
                    width: 100%; 
                    height: 50px; 
                    border-radius: 5px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    color: white;
                    font-weight: bold;">
            Amostra da cor selecionada
        </div>
        """, 
        unsafe_allow_html=True
    )
else:
    cor_tarja = cores[cor_option]
    # Mostrar amostra da cor
    r, g, b = cor_tarja
    st.markdown(
        f"""
        <div style="background-color: rgb({r}, {g}, {b}); 
                    width: 100%; 
                    height: 50px; 
                    border-radius: 5px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    color: white;
                    font-weight: bold;">
            Amostra da cor selecionada
        </div>
        """, 
        unsafe_allow_html=True
    )

# Botão para processar
if st.button("Gerar Catálogo com Preços"):
    if pdf_file is None or excel_file is None:
        st.error("Por favor, faça upload do arquivo PDF e do arquivo Excel.")
    else:
        with st.spinner("Processando... Isso pode levar alguns minutos."):
            try:
                # Criar diretório temporário dedicado para este processamento
                temp_dir = tempfile.mkdtemp(prefix="catalogo_")
                
                # Salvar os arquivos no diretório temporário
                pdf_path = os.path.join(temp_dir, "catalogo.pdf")
                excel_path = os.path.join(temp_dir, "precos.xlsx")
                output_pdf_path = os.path.join(temp_dir, "catalogo_com_precos.pdf")
                
                # Salvar os arquivos de entrada
                with open(pdf_path, "wb") as f:
                    f.write(pdf_file.getvalue())
                
                with open(excel_path, "wb") as f:
                    f.write(excel_file.getvalue())
                
                # Processar o PDF com tratamento de erros robusto
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                def update_progress(message, percent):
                    status_text.text(message)
                    progress_bar.progress(percent)
                
                # Processar o PDF
                num_produtos = processar_pdf_com_markup(
                    pdf_path, 
                    excel_path, 
                    output_pdf_path, 
                    markup,
                    cor_tarja=cor_tarja,
                    progress_callback=update_progress
                )
                
                # Verificar se o arquivo de saída foi criado
                if not os.path.exists(output_pdf_path):
                    raise FileNotFoundError("O arquivo de saída não foi criado. Verifique os logs para mais detalhes.")
                
                # Exibir resultado
                st.success(f"{num_produtos} produtos processados com sucesso!")
                
                # Botão para download
                with open(output_pdf_path, "rb") as file:
                    btn = st.download_button(
                        label="Baixar Catálogo com Preços",
                        data=file,
                        file_name="catalogo_com_precos.pdf",
                        mime="application/pdf"
                    )
                
            except Exception as e:
                st.error(f"Ocorreu um erro: {str(e)}")
                st.error("Detalhes do erro:")
                st.code(traceback.format_exc())
            
            finally:
                # Limpar arquivos temporários de forma segura
                try:
                    if 'temp_dir' in locals():
                        # Aguardar um momento para garantir que todos os processos terminaram
                        time.sleep(1)
                        shutil.rmtree(temp_dir, ignore_errors=True)
                except Exception as cleanup_error:
                    st.warning(f"Aviso: Não foi possível limpar todos os arquivos temporários. {str(cleanup_error)}")

# Adicionar informações sobre o aplicativo
st.markdown("---")
st.markdown("""
### Sobre o aplicativo
Este aplicativo permite adicionar preços aos catálogos PDF de forma automática. 
Ele extrai as imagens do PDF original, identifica os códigos dos produtos e adiciona os preços no rodapé de cada página.

#### Características:
- Extração de imagens em alta qualidade
- Detecção automática de códigos de produtos
- Aplicação de markup personalizado
- Arredondamento de preços para terminar em 7
- Personalização da cor do rodapé
""")
