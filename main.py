import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox
from fpdf import FPDF
import os
import re
import pdf2image


def arquivo():
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    return file_path

def gerar_pdf(tabela, pasta_destino):
    # Substituir NaN por "Não Informado"
    tabela = tabela.fillna("Não Informado");

    # Deixar a coluna de Contato Principal em captalize
    nomeContato = tabela.loc[0, 'Contato principal'].split(' ');
    for i in range(len(nomeContato)):
        nomeContato[i] = nomeContato[i].capitalize();
    
    tabela['Contato principal'] = ' '.join(nomeContato);

    if tabela.loc[0, 'CPF'] != "Não Informado":
        # Remover pontos e traços da Coluna de CPF
        tabela['CPF'] = tabela['CPF'].str.replace('[.-]', '', regex=True)
        # Formatar Coluna de CPF 
        cpf = tabela.loc[0, 'CPF']
        tabela['CPF'] = '{}.{}.{}-{}'.format(cpf[:3], cpf[3:6], cpf[6:9], cpf[9:]);
    
    pdf_file = FPDF();
    pdf_file.add_page();
    pdf_width = 210;
    pdf_height = 297;
    pdf_file.image('./base.png', x=0, y=0, w=pdf_width, h=pdf_height);
    
    campos = {
        'Data Criada': (20, 37),
        'Contato principal': (27, 61),
        'Data de Nascimento (texto)': (170, 61),
        'Endereço': (35, 68),
        'CEP': (163, 68),
        'Bairro': (28, 73),
        'Telefone comercial (contato)': (47, 80),
        'CPF': (163, 80),
        'E-mail': (28, 86),
        'Possui filhos?': (40, 94),
        'Se sim, quantos?': (129, 94),
        'Profissão': (145, 87),
        'Tem medo de dentista?': (61, 114),        
        'Qual a última vez que foi ao dentista?': (98, 122),
        'Alguma experiência negativa no último dentista?': (108, 130),
        'Está satisfeito com a estética do seu sorriso?': (98, 139),
        'Que tratamento está buscando?': (77, 147),
        'Possui alergia à medicamentos? Se sim, quais?': (114, 157),
        'Qual estilo musical mais gosta?': (85, 177),
        'Como nos conheceu?': (55, 185),
        'Tem Instagram?': (57, 195),
        'Autoriza o uso da foto?': (153, 243),        
    }
    
    for campo, posicao in campos.items():
        valor = tabela.loc[0, campo]
        pdf_file.set_font('helvetica', 'B', 10)
        pdf_file.set_xy(posicao[0], posicao[1])
        if campo == 'Autoriza o uso da foto?':
            if valor == 1:
                pdf_file.cell(0, 0, "Sim, autorizo.")
            else:
                pdf_file.cell(0, 0, "Não autorizo.")
        else:
            pdf_file.cell(0, 0, str(valor) if not pd.isna(valor) else "");
    
    # Criar a pasta com o nome da tabela
    Tabela_Nome = tabela.loc[0, 'Contato principal']
    pasta = Tabela_Nome
    diretorio = os.path.join(pasta_destino, pasta)
    os.makedirs(diretorio, exist_ok=True)

    nome_arquivo = os.path.join(diretorio, Tabela_Nome);
    pdf_file.output(f"{nome_arquivo}.pdf");
    
    image = pdf2image.convert_from_path(f"{nome_arquivo}.pdf", poppler_path=r'C:\Program Files (x86)\poppler-23.05.0\Library\bin')
    for page in image:
        page.save(f"{nome_arquivo}.jpg", "JPEG");

def click_butao():
    file_path = arquivo()
    if file_path:
        tabela = pd.read_excel(file_path, dtype={'CPF': str, 'Data de Nascimento (texto)': str})
        pasta_destino = filedialog.askdirectory()
        try:
            gerar_pdf(tabela, pasta_destino)
        except:
            messagebox.showerror("Erro", "Ocorreu um erro ao gerar o PDF e a imagem!");
            return;
        # Exibir mensagem de confirmação
        messagebox.showinfo("Sucesso", "O PDF e a imagem foram geradas com sucesso!");

janela = Tk()

janela.iconbitmap('./ico.ico')
janela.title('Automatizador de PDF')
janela.geometry('300x100+100+100')

botao_procurar = Button(janela, text="Procurar Arquivo", command=click_butao)
botao_procurar.place(x=100, y=25)

janela.mainloop()