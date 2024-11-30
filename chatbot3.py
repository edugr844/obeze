import streamlit as st
from openpyxl import Workbook, load_workbook
import io
import tempfile

import requests
import os


# Função para fazer o download do arquivo Fluxo_de_Caixa.xlsx do GitHub
def download_arquivo_fluxo():
    # URL do arquivo no seu repositório GitHub (ajuste conforme necessário)
    url = "https://github.com/edugr844/obeze/raw/main/Fluxo_de_Caixa.xlsx"
    
    # Caminho para salvar o arquivo temporariamente no Streamlit
    caminho_temp = os.path.join("temp", "Fluxo_de_Caixa.xlsx")

    # Baixando o arquivo
    response = requests.get(url)
    
    # Salvando o arquivo localmente
    with open(caminho_temp, "wb") as f:
        f.write(response.content)
    
    return caminho_temp

# Função para salvar dados no arquivo "Fluxo de Caixa"
def salvar_dados_excel(segmento, funcionarios, anos_operando, codigo):
    # Fazendo o download do arquivo "Fluxo_de_Caixa.xlsx" do GitHub
    caminho_fluxo_original = download_arquivo_fluxo()

    # Carrega o arquivo "Fluxo_de_Caixa" baixado
    wb = load_workbook(caminho_fluxo_original)
    ws = wb.active

    # Atualiza as células A1, A2 e A3 com os dados coletados
    ws['A1'] = f"Segmento: {segmento}"
    ws['A2'] = f"Funcionários: {funcionarios}"
    ws['A3'] = f"Anos Operando: {anos_operando}"

    # Cria um nome único para o arquivo com base no código
    file_name = f"fluxo_de_caixa_{codigo}.xlsx"

    # Salva o arquivo atualizado em um caminho temporário
    temp_file_path = os.path.join("temp", file_name)
    wb.save(temp_file_path)

    return temp_file_path


def salvar_dados_excel(segmento, funcionarios, anos_operando):
    # Cria um arquivo Excel temporário
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb = Workbook()
    ws = wb.active

    # Cria cabeçalho se for um novo arquivo
    ws.append(["Segmento de Mercado", "Número de Funcionários", "Anos Operando"])
    ws.append([segmento, funcionarios, anos_operando])

    # Salva o arquivo no caminho temporário
    wb.save(temp_file.name)
    return temp_file

def chatbot():
    st.title("Obezê")
    st.write("Por favor, responda às perguntas abaixo:")

    # Perguntas para o usuário
    segmento = st.text_input("Qual é o seu segmento de mercado?")
    funcionarios = st.number_input("Quantos funcionários você tem?", step=1, min_value=0)
    anos_operando = st.number_input("Há quantos anos sua empresa está operando?", step=1, min_value=0)

    # Botão para enviar os dados
    if st.button("Enviar"):
        if segmento and funcionarios and anos_operando:
            st.success("Dados enviados com sucesso!")
            st.write("Resumo das informações coletadas:")
            st.write(f"- Segmento de Mercado: {segmento}")
            st.write(f"- Número de Funcionários: {funcionarios}")
            st.write(f"- Anos Operando: {anos_operando}")
            
            # Salva os dados no Excel e cria um arquivo temporário
            temp_file = salvar_dados_excel(segmento, funcionarios, anos_operando)

            # Oferece o download do arquivo gerado
            with open(temp_file_path.name, "rb") as f:
                st.download_button(
                    label="Baixar Fluxo de Caixa Atualizado",
                    data=f,
                    file_name=f"fluxo_de_caixa_{segmento}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("Por favor, preencha todas as informações!")

# Executa o chatbot
if __name__ == "__main__":
    chatbot()
