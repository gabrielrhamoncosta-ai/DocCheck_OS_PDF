import os
import re
import fitz  # PyMuPDF
import pandas as pd  # Biblioteca para criar o Excel
import time  # Para pequenas pausas no sistema operacional

# --- CONFIGURAÇÃO GERAL ---
DIRETORIO_PDFS = '.'  # Diretório atual
NOME_ARQUIVO_EXCEL = 'Relatorio_Completo_OS.xlsx'

# --- CONFIGURAÇÃO DE VALIDAÇÃO ---
IGNORAR_VALIDACAO_ASSINATURA = False
IGNORAR_VALIDACAO_DESCRICAO = False

# Regex para encontrar a matrícula (6 dígitos)
REGEX_MATRICULA = r'(?<!\d)(\d{6})(?!\d)'
MATRICULA_MIN = 40000
MATRICULA_MAX = 76906

def verificar_assinatura_desenho(caminho_arquivo):
    """
    Verifica assinatura de forma abrangente (Tablet, DocuSign, Caneta).
    """
    try:
        with fitz.open(caminho_arquivo) as doc:
            # 1. Checagem de Texto
            for page in doc:
                texto_pagina = page.get_text("text").lower()
                termos_assinatura = ["docusign", "assinado digitalmente", "signed by", "assinatura digital", "icp-brasil"]
                if any(termo in texto_pagina for termo in termos_assinatura):
                    return True

            # 2. Checagem de Elementos Visuais
            for page in doc:
                for annot in page.annots():
                    if annot.type[1] in ['Ink', 'Stamp', 'Widget']: 
                        return True
                
                drawings = page.get_drawings()
                for draw in drawings:
                    for item in draw['items']:
                        if item[0] == 'c': # Curvas (escrita manual)
                            return True
        return False
    except Exception:
        return False

def limpar_string(texto):
    if not texto:
        return ""
    texto_limpo = texto.replace('"', '').replace('\n', ' ').strip().strip(',')
    return re.sub(r'\s+', ' ', texto_limpo)

def extrair_dados_texto(texto):
    """
    Extrai Matrícula, Nome, Função e valida Descrição.
    """
    dados = {
        "matricula": "N/A",
        "nome": "N/A",
        "funcao_cargo": "N/A",
        "status_descricao": "Ausente"
    }

    if not texto:
        return dados

    # 1. Segmentação
    segmento_funcionario = texto
    split_funcionario = re.split(r'DADOS DO FUNCIONÁRIO', texto, flags=re.IGNORECASE)
    if len(split_funcionario) > 1:
        segmento_funcionario = split_funcionario[1]

    # 2. Extrair Matrícula
    all_matches = re.findall(REGEX_MATRICULA, segmento_funcionario)
    if not all_matches:
        all_matches = re.findall(REGEX_MATRICULA, texto)
        
    for match_str in all_matches:
        try:
            numero = int(match_str)
            if MATRICULA_MIN <= numero <= MATRICULA_MAX:
                dados["matricula"] = match_str
                break
        except ValueError:
            continue

    # 3. Extrair Nome
    match_nome = re.search(
        r'Nome\W*[:]\s*(?:",")?\s*"?\s*(.*?)(?="|Centro|Matrícula|Endereço|Data|CNPJ)', 
        segmento_funcionario, 
        re.IGNORECASE | re.DOTALL
    )
    if match_nome:
        nome_extraido = match_nome.group(1)
        if "CONSTRUTORA" not in nome_extraido.upper():
             dados["nome"] = limpar_string(nome_extraido)

    # 4. Extrair Função
    match_funcao = re.search(
        r'Função\W*[:]\s*(?:",")?\s*"?\s*(.*?)(?="|CTPS|CNAE|Bairro|RG|CNPJ|Data)', 
        segmento_funcionario, 
        re.IGNORECASE | re.DOTALL
    )
    if match_funcao:
        dados["funcao_cargo"] = limpar_string(match_funcao.group(1))

    # 5. Verificar Descrição
    match_desc = re.search(r'Descrição Função\W*:(.*?)Agentes das Atividades', texto, re.DOTALL | re.IGNORECASE)
    
    status_real_descricao = "Ausente"
    if match_desc:
        conteudo_descricao = match_desc.group(1).strip()
        conteudo_limpo = re.sub(r'[.,;:"\s]', '', conteudo_descricao)
        if len(conteudo_limpo) > 2:
            status_real_descricao = "OK"
    
    dados["status_descricao"] = status_real_descricao
    return dados

def processar_renomeacao(diretorio, lista_dados):
    """
    Função interativa para renomear os arquivos com base nos dados já extraídos.
    """
    print("\n" + "="*50)
    print(">>> ETAPA DE RENOMEAÇÃO <<<")
    print("="*50)
    
    resposta = input("Deseja renomear os arquivos conforme as matrículas encontradas? (S/N): ").strip().lower()
    
    if resposta not in ['s', 'sim', 'y', 'yes']:
        print("Renomeação cancelada pelo usuário.")
        return

    print("\nIniciando renomeação...")
    cont_sucesso = 0
    cont_erro = 0

    for item in lista_dados:
        nome_atual = item["Nome do Arquivo .pdf"]
        matricula = item["Matrícula"]
        
        caminho_atual = os.path.join(diretorio, nome_atual)
        
        # Verifica se o arquivo ainda existe (pode ter sido movido ou deletado manualmente)
        if not os.path.exists(caminho_atual):
            continue

        novo_nome = ""
        
        if matricula != "N/A":
            # Verifica se já não está renomeado para evitar duplicação (Ex: 123456 - 123456 - arq.pdf)
            if nome_atual.startswith(f"{matricula} -"):
                print(f"Ignorado (Já renomeado): {nome_atual}")
                continue
                
            novo_nome = f"{matricula} - {nome_atual}"
        else:
            # Se não achou matrícula, opcionalmente renomeia para ERROR (como no seu script antigo)
            if not nome_atual.startswith("ERROR -"):
                novo_nome = f"ERROR - {nome_atual}"
            else:
                continue

        caminho_novo = os.path.join(diretorio, novo_nome)

        try:
            os.rename(caminho_atual, caminho_novo)
            if matricula != "N/A":
                print(f"[OK] Renomeado: '{nome_atual}' -> '{novo_nome}'")
                cont_sucesso += 1
            else:
                print(f"[AVISO] Matrícula não encontrada. Marcado como ERROR: '{novo_nome}'")
                cont_erro += 1
        except OSError as e:
            print(f"[ERRO] Falha ao renomear '{nome_atual}': {e}")

    print(f"\nResumo da Renomeação: {cont_sucesso} arquivos renomeados com sucesso, {cont_erro} marcados com erro.")


def gerar_relatorio_e_renomear(diretorio):
    print(f"--- Iniciando Análise em: {os.path.abspath(diretorio)} ---")
    
    lista_dados = []
    arquivos = [f for f in os.listdir(diretorio) if f.lower().endswith('.pdf')]
    total_arquivos = len(arquivos)
    
    # --- 1. LEITURA E EXTRAÇÃO (A parte "Perfeita") ---
    for i, nome_arquivo in enumerate(arquivos):
        caminho_completo = os.path.join(diretorio, nome_arquivo)
        print(f"[{i+1}/{total_arquivos}] Lendo: {nome_arquivo}...", end='\r')
        
        # Ignora arquivos temporários do Excel ou já processados como ERROR para leitura
        if nome_arquivo.startswith('~$') or nome_arquivo.startswith('ERROR -'):
             # Ainda adicionamos à lista para o Excel, mas pulamos processamento pesado se quiser
             pass 

        status_leitura = "OK"
        tem_assinatura_real = False
        dados_extraidos = {
            "matricula": "N/A", "nome": "N/A", 
            "funcao_cargo": "N/A", "status_descricao": "Ausente"
        }
        
        try:
            texto_completo = ""
            # O 'with' garante que o arquivo fecha logo após a leitura
            with fitz.open(caminho_completo) as doc:
                tem_assinatura_real = verificar_assinatura_desenho(caminho_completo)
                for p_index in range(min(3, len(doc))):
                    texto_completo += doc[p_index].get_text("text", sort=True) + "\n"
            
            dados_extraidos = extrair_dados_texto(texto_completo)
            
        except Exception as e:
            status_leitura = f"Erro: {e}"
        
        # Compilação dos resultados
        resultado_assinatura = "OK" if tem_assinatura_real else "Ausente"
        resultado_descricao = dados_extraidos["status_descricao"]

        if IGNORAR_VALIDACAO_ASSINATURA: resultado_assinatura = "OK"
        if IGNORAR_VALIDACAO_DESCRICAO: resultado_descricao = "OK"

        # Adiciona à lista de dados
        lista_dados.append({
            "Nome do Arquivo .pdf": nome_arquivo,
            "Matrícula": dados_extraidos["matricula"],
            "Nome": dados_extraidos["nome"],
            "Tem assinatura": resultado_assinatura,
            "Status de leitura": status_leitura,
            "Descrição da Função": resultado_descricao,
            "Função (Cargo)": dados_extraidos["funcao_cargo"]
        })

    print(f"\n\nLeitura finalizada! Gerando Excel...")

    # --- 2. GERAÇÃO DO EXCEL ---
    if lista_dados:
        colunas = [
            "Nome do Arquivo .pdf", "Matrícula", "Nome", 
            "Tem assinatura", "Status de leitura", 
            "Descrição da Função", "Função (Cargo)"
        ]
        
        df = pd.DataFrame(lista_dados)
        for col in colunas:
            if col not in df.columns: df[col] = "N/A"
        
        df = df[colunas]
        try:
            df.to_excel(NOME_ARQUIVO_EXCEL, index=False)
            print(f"SUCESSO! Relatório salvo como '{NOME_ARQUIVO_EXCEL}'")
        except PermissionError:
            print(f"ERRO: Não foi possível salvar o Excel. Verifique se '{NOME_ARQUIVO_EXCEL}' está aberto.")

        # Pausa para garantir que o OS liberou os handles dos arquivos PDF
        time.sleep(1)

        # --- 3. CHAMADA DA RENOMEAÇÃO ---
        processar_renomeacao(diretorio, lista_dados)

    else:
        print("Nenhum arquivo PDF encontrado.")

if __name__ == "__main__":
    gerar_relatorio_e_renomear(DIRETORIO_PDFS)
    input("\nPressione Enter para sair...")