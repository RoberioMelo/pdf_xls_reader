from tkinter import messagebox, filedialog, LEFT, FALSE, BOTTOM, TOP, RIGHT, StringVar, END
from customtkinter import CTk, CTkImage
from PIL import Image
import pandas as pd
import pdfplumber
import datetime
from pathlib import Path
import re
import os
import io
import threading
import customtkinter
from tqdm import tqdm


def remove_special_characters(str1):
    text = str1
    return re.sub(r"\D", '', text)


def salvar_resultado(df_final):
    folder_path = filedialog.askdirectory(title="Selecione a pasta para salvar os resultados")
    folder_path = Path(folder_path)
    current_date = datetime.datetime.now().strftime("%d-%m-%Y %H_%M")
    result_file = folder_path.joinpath(f'RESULTADO - {current_date}')
    df_final.to_excel(result_file.with_suffix('.xlsx'), index=False, header=True, sheet_name='EMPRESAS')


def converter_pdf_para_txt(lista_pdf):
    return ''.join([page.extract_text() for pdf_file in lista_pdf for page in pdfplumber.open(pdf_file).pages])


class Application:
    def __init__(self, master=None):
        self.df_final = None
        self.xlsx_sheet = None
        self.lista_pdf = None
        self.master = master
        master.resizable(0, 0)
        self.fontePadrao = customtkinter.CTkFont(family='Calibri', size=20)
        self.arquivo_pdf_selecionado = False
        self.arquivo_xlsx_selecionado = False

        # BORDAS DA TELA
        self.bordaContainer = customtkinter.CTkFrame(master=None, width=8, height=500, corner_radius=0)
        self.bordaContainer.pack(side=LEFT)
        self.bordaContainer.propagate(FALSE)

        self.bordaContainer = customtkinter.CTkFrame(master=None, width=500, height=8, corner_radius=0)
        self.bordaContainer.pack(side=BOTTOM)
        self.bordaContainer.propagate(FALSE)

        self.bordaContainer = customtkinter.CTkFrame(master=None, width=500, height=8, corner_radius=0)
        self.bordaContainer.pack(side=TOP)
        self.bordaContainer.propagate(FALSE)

        self.bordaContainer = customtkinter.CTkFrame(master=None, width=8, height=500, corner_radius=0)
        self.bordaContainer.pack(side=RIGHT)
        self.bordaContainer.propagate(FALSE)
        # BOTÃO SWITCH PARA TROCAR O TEMA
        self.appearance_mode = StringVar(value="System")
        self.buttonTema = customtkinter.CTkSwitch(master=None, text="☀", font=("Arial", 20), onvalue="Light",
                                                  offvalue="Dark", variable=self.appearance_mode,
                                                  command=self.change_appearance_mode_event)
        self.buttonTema.pack(padx=20, pady=10)
        self.buttonTema.pack(side=BOTTOM, anchor="w")
        self.buttonTema.propagate(FALSE)
        # FRAME DO CODIGO, USADO PARA SEPARAR OS ELEMENTOS
        self.primeiroContainer = customtkinter.CTkFrame(master=None, fg_color="transparent")
        self.primeiroContainer.pack(anchor="center")
        self.primeiroContainer.configure()

        self.segundoContainer = customtkinter.CTkFrame(master=None, fg_color="transparent")
        self.segundoContainer.pack(anchor="e", padx=10, pady=4)

        self.terceiroContainer = customtkinter.CTkFrame(master=None, fg_color="transparent")
        self.terceiroContainer.pack(anchor="e", padx=10, pady=4)
        # FRAME DA BARRA DE PROGRESSO
        self.sextoContainer = customtkinter.CTkProgressBar(master=None, orientation="horizontal",
                                                           progress_color="#00AB0B")
        self.sextoContainer.pack(anchor="e", padx=90, pady=25)
        self.progress = self.sextoContainer
        self.sextoContainer.destroy()

        self.quartoContainer = customtkinter.CTkFrame(master=None, fg_color="transparent")
        self.quartoContainer.pack(anchor="e", padx=120, pady=5)

        self.quintoContainer = customtkinter.CTkFrame(master=None, fg_color="transparent")
        self.quintoContainer.pack(anchor="e", padx=10, pady=4)
        # LOGO DO PROJETO
        image_path = "C:\\Users\\Analise\\Desktop\\Projetos\\exe_doc_diario\\img\\logo.png"
        pil_image = Image.open(image_path)
        ctk_image = CTkImage(pil_image, size=(80, 80))
        # LOGO
        self.titulo = customtkinter.CTkLabel(self.primeiroContainer, text="", image=ctk_image)
        self.titulo.pack()
        # LABEL "SELECIONAR PDFs"
        self.buttonPDFText = customtkinter.CTkLabel(self.segundoContainer, pady=8, text="Selecione o(s) PDF(s):")
        self.buttonPDFText.pack()
        # BOTÃO DE SELEÇÃO
        self.buttonPDFLabel = customtkinter.CTkButton(self.segundoContainer, text="Selecionar", width=70,
                                                      height=29, command=self.selecionar_pdf)
        self.buttonPDFLabel["font"] = self.fontePadrao
        self.buttonPDFLabel.pack(side=LEFT, padx=5)
        # CAIXA DE TEXTO
        self.PDF = customtkinter.CTkEntry(self.segundoContainer, placeholder_text="Adicione seus PDFs",
                                          width=280, height=30)
        self.PDF["font"] = self.fontePadrao
        self.PDF.pack()
        # LABEL "SELECIONAR EXCEL"
        self.buttonXLSXText = customtkinter.CTkLabel(self.terceiroContainer, pady=5,
                                                     text="Selecione a planilha de Empresas:")
        self.buttonXLSXText.pack()
        # BOTÃO DE SELEÇÃO
        self.buttonXLSXLabel = customtkinter.CTkButton(self.terceiroContainer, text="Selecionar", width=70,
                                                       height=29, command=self.selecionar_xlsx)
        self.buttonXLSXLabel["font"] = self.fontePadrao
        self.buttonXLSXLabel.pack(side=LEFT, padx=5)
        # CAIXA DE TEXTO
        self.XLSX = customtkinter.CTkEntry(self.terceiroContainer, placeholder_text="Adicione sua planilha de empresas",
                                           width=280, height=30)
        self.XLSX["font"] = self.fontePadrao
        self.XLSX.pack()
        # INICIAR
        self.executar = customtkinter.CTkButton(self.quartoContainer, text="Iniciar",
                                                command=lambda: ExecutarThread(self).start())
        self.executar["font"] = self.fontePadrao
        self.executar.pack()
        # Salvar
        self.salvarR = customtkinter.CTkButton(self.quartoContainer, text="Salvar",
                                               command=self.salvar)
        self.salvarR["font"] = self.fontePadrao
        self.salvarR.pack()

    # FUNÇÕES
    def change_appearance_mode_event(self):
        new_appearance_mode = self.appearance_mode.get()
        customtkinter.set_appearance_mode(new_appearance_mode)

    def selecionar_xlsx(self):
        self.xlsx_sheet = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo",
                                                     filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        if self.xlsx_sheet:
            self.arquivo_xlsx_selecionado = True

        self.XLSX.delete(0, END)
        self.XLSX.insert(END, self.xlsx_sheet)

    def selecionar_pdf(self):
        self.lista_pdf = filedialog.askopenfilenames(initialdir="/", title="Selecione um arquivo",
                                                     filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))
        if self.lista_pdf:
            self.arquivo_pdf_selecionado = True

        self.PDF.delete(0, END)
        for i in self.lista_pdf:
            self.PDF.insert(END, i + "/" + "\n")
        return self.lista_pdf

    @property
    def iniciar(self):
        if not (self.arquivo_pdf_selecionado and self.arquivo_xlsx_selecionado):  # Verifique se a variável é True
            messagebox.showerror("Erro", "Selecione os PDFs e a planilha antes de iniciar o processo.")

            return
        lista_pdf = self.lista_pdf
        pdf_text = converter_pdf_para_txt(lista_pdf)
        pdf_text = " ".join(pdf_text.split())
        pdf_lines = pdf_text.split("\n")

        xlsx_file = self.xlsx_sheet
        df = pd.read_excel(xlsx_file)
        cnpj_key = df.iloc[:, 2][3:]
        value = df.iloc[:, 0][3:]
        caceal_key = df.iloc[:, 25][3:]

        resultado_dict = {}
        for chave, valor in zip(cnpj_key, value):
            if isinstance(chave, str):
                for linha in pdf_lines:
                    if str(chave).lower() in str(linha).lower():
                        resultado_dict[chave] = valor
                        break
        print(f"Resultado dict: {resultado_dict}")  # Adiciona um print para verificar o dicionário resultado_dict

        resultado_dict1 = {}
        for chave, valor in zip(cnpj_key, value):
            chave = str(chave)
            stripped_key = remove_special_characters(chave)
            if isinstance(chave, str):
                if stripped_key.startswith(chave.replace(" ", "")):
                    for linha in pdf_lines:
                        if chave.lower() in linha.lower():
                            resultado_dict1[chave] = valor
                            break
        print(f"Resultado dict1: {resultado_dict1}")  # Adiciona um print para verificar o dicionário resultado_dict1

        resultado_dict3 = {}
        additional_character = '-'
        additional_character_position = 8
        for chave, valor in zip(caceal_key, value):
            chave = str(chave)
            stripped_key = remove_special_characters(chave[12:])
            if not stripped_key.strip():
                continue
            stripped_key = stripped_key[:additional_character_position] + additional_character + stripped_key[
                                                                                                 additional_character_position:]
            if isinstance(chave, str):
                for linha in pdf_lines:
                    if stripped_key in linha:
                        resultado_dict3[stripped_key] = valor
                        break

        print(f"Resultado dict3: {resultado_dict3}")  # Adiciona um print para verificar o dicionário resultado_dict3

        df_resultado = pd.DataFrame(list(resultado_dict.items()), columns=['CNPJ/CACEAL', 'EMPRESA'])
        df_resultado1 = pd.DataFrame(list(resultado_dict1.items()), columns=['CNPJ/CACEAL', 'EMPRESA'])
        df_resultado3 = pd.DataFrame(list(resultado_dict3.items()), columns=['CNPJ/CACEAL', 'EMPRESA'])

        df_final = pd.concat([df_resultado, df_resultado1, df_resultado3])
        # Criar lista para receber os valores de cada arquivo
        # Atribui o resultado da função à variável "df_final"
        self.df_final = df_final
        # Retorna None para indicar que a função não retorna nada

        return None

    def salvar(self):
        if not hasattr(self, 'df_final'):
            messagebox.showwarning("Aviso", "Por favor, execute a consulta primeiro.")
            return
        salvar_resultado(self.df_final)


class ExecutarThread(threading.Thread):
    def __init__(self, application):
        threading.Thread.__init__(self)
        self.application = application

    def run(self):
        self.application.iniciar


root = customtkinter.CTk()
root.geometry("400x450")
root.title("Consulta")
root.iconbitmap('C:\\Users\\Analise\\Desktop\\Projetos\\exe_doc_diario\\img\\logo.ico')
Application(root)
root.mainloop()
