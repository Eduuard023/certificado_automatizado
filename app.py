"""
#Pegar os dados da planilha
#Transferir os dados para o certificado
"""

#Pegar os dados da planilha
import openpyxl
from PIL import Image, ImageDraw, ImageFont

#Abrir a planilha
tabela = openpyxl.load_workbook('planilha_certificados.xlsx')
sheet_tabela = tabela['Planilha1']

for indice, linha in enumerate(sheet_tabela.iter_rows(min_row=2, max_row=2)):
    # cada célula que contém a info que precisamos
    nome_curso = linha[0].value # Nome do cruso
    nome_aluno = linha[1].value # Nome do Aluno
    carga_horaria = linha[2].value # Carga horária
    data_inicio = linha[3].value # Data de inicio
    data_final = linha[4].value # Data de conclusão
    data_emissao = linha[5].value # Data da emissao do certificado
    
    #Transferindo os dados da planilha para o certificado
    #Definindo as fontes
    fonte_nome = ImageFont.truetype('./tahomabd.ttf')
    fonte_geral = ImageFont.truetype('./tahoma.ttf')

    imagem = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((1020,827), nome_aluno, fill='black', font=fonte_nome)

    imagem.save(f'./{indice} {nome_aluno} certificado.png')
