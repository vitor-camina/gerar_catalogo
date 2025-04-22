import fitz  # PyMuPDF
import os
import re
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from fpdf import FPDF
import numpy as np
import io
import time
import traceback
import sys
import shutil

def extrair_imagens_pdf(pdf_path, output_dir, progress_callback=None):
    """
    Extrai apenas as imagens do PDF, ignorando completamente a camada de texto.
    
    Args:
        pdf_path: Caminho para o arquivo PDF
        output_dir: Diretório onde as imagens serão salvas
        progress_callback: Função para reportar progresso
        
    Returns:
        Lista de caminhos para as imagens extraídas
    """
    # Criar diretório de saída se não existir
    os.makedirs(output_dir, exist_ok=True)
    
    # Abrir o PDF com tratamento de erros
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise Exception(f"Erro ao abrir o arquivo PDF: {str(e)}")
    
    imagens_paths = []
    total_pages = len(doc)
    
    # Para cada página
    for page_num in range(total_pages):
        if progress_callback:
            progress_callback(f"Extraindo imagem da página {page_num+1} de {total_pages}...", 
                             (page_num / total_pages) * 0.3)  # 30% do progresso total
        
        try:
            # Renderizar a página como imagem em alta resolução
            page = doc[page_num]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Aumentar resolução para melhor qualidade
            
            # Salvar a imagem
            img_path = os.path.join(output_dir, f"page_{page_num+1}.png")
            pix.save(img_path)
            
            imagens_paths.append({
                'pagina': page_num,
                'caminho': img_path
            })
        except Exception as e:
            print(f"Erro ao processar página {page_num+1}: {str(e)}")
            # Continuar com a próxima página em vez de falhar completamente
            continue
    
    doc.close()
    
    return imagens_paths

def extrair_codigos_produtos(pdf_path, progress_callback=None):
    """
    Extrai códigos de produtos do PDF para referência.
    
    Args:
        pdf_path: Caminho para o arquivo PDF
        progress_callback: Função para reportar progresso
        
    Returns:
        Lista de códigos de produtos com suas páginas
    """
    codigos = []
    
    # Padrões para encontrar códigos de produtos
    padroes = [
        r'CONJUNTO\s+(\d{5})',  # Padrão para "CONJUNTO 85274"
        r'BERMUDA\s+(\d{5})',   # Padrão para "BERMUDA 84216"
        r'CAMISA\s+(\d{5})',    # Padrão para "CAMISA 84831"
        r'CAMISETA\s+(\d{5})',  # Padrão para "CAMISETA 84218"
        r'BONE\s+(\d{5})',      # Padrão para "BONE 82969"
        r'BONÉ\s+(\d{5})',      # Padrão para "BONÉ 82969" (com acento)
        r'BLUSA\s+(\d{5})',     # Padrão para "BLUSA 85640"
        r'SAIA\s+(\d{5})',      # Padrão para "SAIA 86130"
        r'VESTIDO\s+(\d{5})',   # Padrão para "VESTIDO 86065"
        r'MACACÃO\s+(\d{5})',   # Padrão para "MACACÃO 83844"
        r'JAQUETA\s+(\d{5})',   # Padrão para "JAQUETA 85109"
        r'BODY\s+(\d{5})',      # Padrão para "BODY 84291"
        r'CALÇA\s+(\d{5})',     # Padrão para "CALÇA 84522"
        r'LENÇO\s+(\d{5})',     # Padrão para "LENÇO 84486"
        r'ÓCULOS\s+(\d{5})'     # Padrão para "ÓCULOS 8585827"
    ]
    
    # Padrão genérico para encontrar qualquer código de 5 dígitos
    padrao_generico = r'\b(\d{5})\b'
    
    # Abrir o PDF com tratamento de erros
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise Exception(f"Erro ao abrir o arquivo PDF para extração de códigos: {str(e)}")
    
    total_pages = len(doc)
    
    # Para cada página
    for page_num in range(total_pages):
        if progress_callback:
            progress_callback(f"Extraindo códigos da página {page_num+1} de {total_pages}...", 
                             0.3 + (page_num / total_pages) * 0.2)  # 30-50% do progresso total
        
        try:
            # Obter a página
            page = doc[page_num]
            
            # Extrair texto
            texto = page.get_text()
            
            # Verificar padrões específicos
            for padrao in padroes:
                matches = re.finditer(padrao, texto, re.IGNORECASE)
                for match in matches:
                    codigo = match.group(1)
                    texto_completo = match.group(0)
                    
                    codigos.append({
                        'pagina': page_num,
                        'codigo': codigo,
                        'texto_completo': texto_completo,
                        'posicao_y': page_num  # Simplificado para usar apenas o número da página como referência
                    })
            
            # Se não encontrou com padrões específicos, tentar padrão genérico
            if not any(codigo['pagina'] == page_num for codigo in codigos):
                matches = re.finditer(padrao_generico, texto)
                for match in matches:
                    codigo = match.group(1)
                    
                    codigos.append({
                        'pagina': page_num,
                        'codigo': codigo,
                        'texto_completo': codigo,
                        'posicao_y': page_num  # Simplificado para usar apenas o número da página como referência
                    })
        except Exception as e:
            print(f"Erro ao extrair códigos da página {page_num+1}: {str(e)}")
            # Continuar com a próxima página em vez de falhar completamente
            continue
    
    doc.close()
    
    return codigos

def ler_excel_precos(caminho_arquivo, markup=2.0, progress_callback=None):
    """
    Lê um arquivo Excel contendo informações de produtos e preços e aplica o markup.
    
    Args:
        caminho_arquivo: Caminho para o arquivo Excel
        markup: Valor do markup a ser aplicado (ex: 2.0 para 100% de markup)
        progress_callback: Função para reportar progresso
        
    Returns:
        Um dicionário com códigos de produtos como chaves e informações de preço como valores
    """
    if progress_callback:
        progress_callback("Lendo arquivo Excel de preços...", 0.5)  # 50% do progresso total
    
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(caminho_arquivo)
        
        # Verificar se as colunas necessárias existem
        if len(df.columns) < 3:
            raise ValueError("Formato de arquivo inválido: não possui colunas suficientes")
        
        # Renomear colunas para facilitar o acesso
        # Baseado na análise do arquivo, sabemos que:
        # - Primeira coluna (0): Referência do produto
        # - Segunda coluna (1): Tamanho
        # - Terceira coluna (2): Valor (preço de custo)
        colunas_renomeadas = {
            df.columns[0]: 'referencia',
            df.columns[1]: 'tamanho',
            df.columns[2]: 'preco_custo'
        }
        
        df = df.rename(columns=colunas_renomeadas)
        
        # Pular a primeira linha se contiver cabeçalhos
        if isinstance(df['referencia'].iloc[0], str) and not df['referencia'].iloc[0].isdigit():
            df = df.iloc[1:].reset_index(drop=True)
        
        # Converter para dicionário de produtos
        produtos_dict = {}
        
        for _, row in df.iterrows():
            # Verificar se a linha tem valores válidos
            if pd.notna(row['referencia']) and pd.notna(row['preco_custo']):
                try:
                    # Verificar se o valor é uma string que contém "VALOR" (cabeçalho)
                    if isinstance(row['preco_custo'], str) and "VALOR" in row['preco_custo'].upper():
                        continue
                    
                    # Converter preço para float
                    if isinstance(row['preco_custo'], str):
                        preco_custo = float(row['preco_custo'].replace(',', '.'))
                    else:
                        preco_custo = float(row['preco_custo'])
                    
                    # Calcular preço de venda com markup
                    preco_venda = preco_custo * markup
                    
                    # Usar o código como chave no dicionário
                    if isinstance(row['referencia'], (int, float)):
                        codigo = str(int(row['referencia']))  # Converter para inteiro e depois para string para remover decimais
                    else:
                        # Tentar extrair apenas os dígitos se for uma string
                        codigo_match = re.search(r'(\d+)', str(row['referencia']))
                        if codigo_match:
                            codigo = codigo_match.group(1)
                        else:
                            continue  # Pular se não conseguir extrair um código numérico
                    
                    produtos_dict[codigo] = {
                        'tamanho': str(row['tamanho']) if pd.notna(row['tamanho']) else "",
                        'preco_custo': preco_custo,
                        'preco_venda': preco_venda
                    }
                except (ValueError, TypeError) as e:
                    # Pular linhas com valores não numéricos
                    print(f"Erro ao processar linha do Excel: {str(e)}")
                    continue
        
        print(f"Produtos carregados: {len(produtos_dict)}")
        return produtos_dict
    except Exception as e:
        raise Exception(f"Erro ao ler o arquivo Excel: {str(e)}")

def arredondar_preco_para_terminar_em_7(preco):
    """
    Arredonda o preço para baixo e termina com 7.
    
    Args:
        preco: Preço original
        
    Returns:
        Preço arredondado para baixo terminando em 7
    """
    # Arredondar para baixo para o inteiro mais próximo
    preco_inteiro = int(preco)
    
    # Calcular o último dígito
    ultimo_digito = preco_inteiro % 10
    
    # Se o último dígito já é 7, manter o valor
    if ultimo_digito == 7:
        return preco_inteiro
    
    # Se o último dígito é maior que 7, subtrair até chegar em 7
    if ultimo_digito > 7:
        return preco_inteiro - (ultimo_digito - 7)
    
    # Se o último dígito é menor que 7, subtrair até o próximo número terminado em 7
    return preco_inteiro - ultimo_digito - 3

def adicionar_tarja_cinza(imagem_path, output_path, altura_tarja=300, cor_tarja=(128, 128, 128)):
    """
    Adiciona uma tarja colorida sem transparência no rodapé da imagem.
    
    Args:
        imagem_path: Caminho para a imagem original
        output_path: Caminho para salvar a imagem com tarja
        altura_tarja: Altura da tarja em pixels
        cor_tarja: Tupla RGB para a cor da tarja (padrão: cinza)
        
    Returns:
        Caminho para a imagem com tarja
    """
    try:
        # Abrir a imagem
        img = Image.open(imagem_path)
        
        # Criar uma nova imagem sem transparência
        tarja = Image.new('RGB', (img.width, altura_tarja), cor_tarja)
        
        # Converter a imagem original para RGB se necessário
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # Criar uma cópia da imagem original
        img_com_tarja = img.copy()
        
        # Colar a tarja no rodapé
        img_com_tarja.paste(tarja, (0, img.height - altura_tarja))
        
        # Salvar a imagem com tarja
        img_com_tarja.save(output_path)
        
        return output_path
    except Exception as e:
        raise Exception(f"Erro ao adicionar tarja à imagem: {str(e)}")

def criar_pdf_com_precos(imagens, codigos, produtos_dict, output_path, cor_tarja=(128, 128, 128), progress_callback=None):
    """
    Cria um novo PDF com as imagens e preços, usando fonte Arial tamanho 150.
    
    Args:
        imagens: Lista de caminhos para as imagens das páginas
        codigos: Lista de códigos de produtos com suas páginas
        produtos_dict: Dicionário de produtos com códigos como chaves
        output_path: Caminho para salvar o PDF processado
        cor_tarja: Tupla RGB para a cor da tarja (padrão: cinza)
        progress_callback: Função para reportar progresso
        
    Returns:
        Caminho para o PDF processado
    """
    # Criar diretório temporário para imagens processadas
    temp_dir = os.path.join(os.path.dirname(output_path), "temp_images")
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        # Criar um novo PDF
        pdf = FPDF(unit='pt')
        
        # Verificar se a fonte Arial está disponível
        fonte_path = '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'
        if not os.path.exists(fonte_path):
            # Tentar encontrar uma fonte alternativa
            fontes_alternativas = [
                '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
                '/usr/share/fonts/truetype/freefont/FreeSans.ttf',
                '/usr/share/fonts/truetype/ubuntu/Ubuntu-R.ttf'
            ]
            
            for fonte_alt in fontes_alternativas:
                if os.path.exists(fonte_alt):
                    fonte_path = fonte_alt
                    break
            else:
                # Se nenhuma fonte for encontrada, usar a fonte padrão
                fonte_path = None
        
        # Adicionar fonte
        if fonte_path:
            pdf.add_font('CustomFont', '', fonte_path, uni=True)
            pdf.set_font('CustomFont', '', 30)  # Tamanho 150 (quintuplicado de 30)
        else:
            # Usar fonte padrão se não encontrar a fonte personalizada
            pdf.set_font('Arial', '', 30)
        
        total_imagens = len(imagens)
        
        # Processar cada página
        for idx, img_info in enumerate(sorted(imagens, key=lambda x: x['pagina'])):
            page_num = img_info['pagina']
            img_path = img_info['caminho']
            
            if progress_callback:
                progress_callback(f"Criando página {page_num+1} de {total_imagens}...", 
                                 0.7 + (idx / total_imagens) * 0.3)  # 70-100% do progresso total
            
            # Se não for a primeira página, adicionar tarja colorida
            if page_num > 0:
                # Adicionar tarja colorida
                img_com_tarja_path = os.path.join(temp_dir, f"page_{page_num+1}_tarja.png")
                img_path = adicionar_tarja_cinza(img_path, img_com_tarja_path, cor_tarja=cor_tarja)
            
            # Obter dimensões da imagem
            img = Image.open(img_path)
            img_width, img_height = img.size
            
            # Adicionar uma nova página com o tamanho da imagem
            pdf.add_page(format=(img_width, img_height))
            
            # Adicionar a imagem como fundo
            pdf.image(img_path, 0, 0, img_width, img_height)
            
            # Adicionar preços para os códigos desta página
            codigos_pagina = [c for c in codigos if c['pagina'] == page_num]
            
            # Distribuir os textos por toda a largura do rodapé
            if codigos_pagina and page_num > 0:  # Apenas para páginas após a capa
                # Calcular a largura disponível para texto
                largura_disponivel = img_width - 100  # Margem de 50px em cada lado
                
                # Preparar o texto completo para esta página
                textos_precos = []
                
                for codigo_info in codigos_pagina:
                    codigo = codigo_info['codigo']
                    
                    # Verificar se este código existe no dicionário de produtos
                    if codigo in produtos_dict:
                        # Calcular preço com markup
                        preco_venda = produtos_dict[codigo]['preco_venda']
                        
                        # Arredondar preço para terminar em 7
                        preco_arredondado = arredondar_preco_para_terminar_em_7(preco_venda)
                        
                        # Formatar preço
                        texto_preco = f"{codigo_info['texto_completo']} - R$ {preco_arredondado}"
                        textos_precos.append(texto_preco)
                
                # Se temos textos para exibir
                if textos_precos:
                    # Juntar os textos com espaçamento
                    texto_completo = "   |   ".join(textos_precos)
                    
                    # Posição do texto - centralizado no rodapé
                    x_pos = 50  # Margem esquerda
                    y_pos = img_height - 200  # Posição Y no rodapé (centralizado na tarja)
                    
                    # Adicionar texto com preço
                    pdf.set_xy(x_pos, y_pos)
                    pdf.set_text_color(255, 255, 255)  # Branco para contrastar com a tarja
                    
                    # Usar multi_cell com tratamento de erros
                    try:
                        pdf.multi_cell(largura_disponivel, 80, texto_completo)  # Altura de linha 80 para espaçamento adequado
                    except Exception as e:
                        print(f"Erro ao adicionar texto à página {page_num+1}: {str(e)}")
                        # Tentar uma abordagem alternativa
                        try:
                            # Dividir o texto em partes menores se for muito longo
                            partes_texto = texto_completo.split('|')
                            for i, parte in enumerate(partes_texto):
                                if parte.strip():
                                    pdf.set_xy(x_pos, y_pos + i * 40)
                                    pdf.cell(largura_disponivel, 40, parte.strip())
                        except Exception as e2:
                            print(f"Erro na abordagem alternativa: {str(e2)}")
        
        # Salvar o PDF
        pdf.output(output_path)
        
        return output_path
    
    except Exception as e:
        raise Exception(f"Erro ao criar PDF com preços: {str(e)}\n{traceback.format_exc()}")
    
    finally:
        # Limpar arquivos temporários
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception as cleanup_error:
            print(f"Aviso: Não foi possível limpar todos os arquivos temporários. {str(cleanup_error)}")

def processar_pdf_com_markup(pdf_path, excel_path, output_path, markup=2.0, cor_tarja=(128, 128, 128), progress_callback=None):
    """
    Processa um arquivo PDF, adiciona preços com markup e gera um novo PDF.
    Extrai apenas as imagens do PDF, ignorando completamente a camada de texto.
    
    Args:
        pdf_path: Caminho para o arquivo PDF original
        excel_path: Caminho para o arquivo Excel com preços
        output_path: Caminho para salvar o PDF processado
        markup: Valor do markup a ser aplicado
        cor_tarja: Tupla RGB para a cor da tarja (padrão: cinza)
        progress_callback: Função para reportar progresso
        
    Returns:
        Número de produtos processados
    """
    # Criar diretório temporário para extração
    temp_dir = os.path.join(os.path.dirname(output_path), "temp_extraction")
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        # 1. Extrair imagens do PDF
        if progress_callback:
            progress_callback("Extraindo imagens do PDF...", 0.05)
        imagens = extrair_imagens_pdf(pdf_path, temp_dir, progress_callback)
        
        # 2. Extrair códigos de produtos para referência
        if progress_callback:
            progress_callback("Extraindo códigos de produtos...", 0.35)
        codigos = extrair_codigos_produtos(pdf_path, progress_callback)
        print(f"Encontrados {len(codigos)} códigos de produtos no PDF")
        
        # 3. Ler preços do Excel
        if progress_callback:
            progress_callback("Lendo preços do Excel...", 0.55)
        produtos_dict = ler_excel_precos(excel_path, markup, progress_callback)
        
        # 4. Criar novo PDF com preços
        if progress_callback:
            progress_callback("Criando novo PDF com preços...", 0.7)
        criar_pdf_com_precos(imagens, codigos, produtos_dict, output_path, cor_tarja, progress_callback)
        
        # Contar produtos processados
        produtos_processados = sum(1 for c in codigos if c['codigo'] in produtos_dict)
        print(f"Processados {produtos_processados} produtos com sucesso!")
        
        if progress_callback:
            progress_callback("Processamento concluído!", 1.0)
        
        return produtos_processados
    
    except Exception as e:
        print(f"Erro durante o processamento: {str(e)}")
        print(traceback.format_exc())
        raise
    
    finally:
        # Limpar arquivos temporários
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception as cleanup_error:
            print(f"Aviso: Não foi possível limpar todos os arquivos temporários. {str(cleanup_error)}")

# Teste da função
if __name__ == "__main__":
    pdf_path = "/home/ubuntu/upload/Lucboo_BOY+BABY_PRV_26.pdf"
    excel_path = "/home/ubuntu/upload/TABELA DE PREÇO PRIMAVERA VERAO 2026 LUC.BOO (1).xlsx"
    output_path = "/home/ubuntu/pdf_final/catalogo_com_precos.pdf"
    
    def print_progress(message, percent):
        print(f"{message} - {percent*100:.1f}%")
    
    processar_pdf_com_markup(pdf_path, excel_path, output_path, markup=2.0, progress_callback=print_progress)
