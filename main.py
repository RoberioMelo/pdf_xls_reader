from tkinter import messagebox, filedialog, LEFT, BOTTOM, TOP, RIGHT, StringVar, END, PhotoImage
from tkinter import Tk, Label, Button, Entry, Scrollbar, Text, Frame, Canvas
from tkinter.ttk import Frame, Progressbar
from PIL import Image, ImageTk
import pandas as pd
import pdfplumber
import datetime
from pathlib import Path
import re
import os
import io
import threading
from tqdm import tqdm


class ScrollableFrame(Frame):
    def __init__(self, master, width, height, corner_radius):
        super().__init__(master)

        self.canvas = Canvas(self, width=width, height=height)  # Defina o atributo canvas aqui
        scrollbar = Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.config(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.interior = Frame(self.canvas)  # Use self.canvas como parent
        self.canvas.create_window((0, 0), window=self.interior, anchor="nw")

        self.interior.bind("<Configure>", self.set_scrollregion)

    def set_scrollregion(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))



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
    def __init__(self, master):
        self.df_final = None
        self.xlsx_sheet = None
        self.lista_pdf = None
        self.master = master
        master.resizable(0, 0)
        self.fontePadrao = ("Calibri", 20)
        self.arquivo_pdf_selecionado = False
        self.arquivo_xlsx_selecionado = False

        # # BORDAS DA TELA
        # self.bordaContainer = Frame(master, width=8, height=500)
        # self.bordaContainer.pack(side="left")
        # self.bordaContainer.propagate(False)
        #
        # self.bordaContainer = Frame(master=None, width=500, height=8)
        # self.bordaContainer.pack(side=BOTTOM)
        # self.bordaContainer.propagate(False)
        #
        # self.bordaContainer = Frame(master=None, width=500, height=8)
        # self.bordaContainer.pack(side=TOP)
        # self.bordaContainer.propagate(False)
        #
        # self.bordaContainer = Frame(master=None, width=500, height=8)
        # self.bordaContainer.pack(side=RIGHT)
        # self.bordaContainer.propagate(False)

        # # BOTÃO SWITCH PARA TROCAR O TEMA
        # self.appearance_mode = StringVar(value="System")
        # self.buttonTema = customtkinter.CTkSwitch(master=None, text="☀", font=("Arial", 20), onvalue="Light",
        #                                           offvalue="Dark", variable=self.appearance_mode,
        #                                           command=self.change_appearance_mode_event)
        # self.buttonTema.pack(padx=20, pady=10)
        # self.buttonTema.pack(side=BOTTOM, anchor="w")
        # self.buttonTema.propagate(False)

        # FRAME DA BARRA DE PROGRESSO
        # self.sextoContainer = customtkinter.CTkProgressBar(master=None, orientation="horizontal",
        #                                                    progress_color="#00AB0B")
        # self.sextoContainer.pack(anchor="e", padx=90, pady=25)
        # self.progress = self.sextoContainer
        # self.sextoContainer.destroy()

        # LOGO DO PROJETO
        image_path = "C:\\Users\\Analise\\Desktop\\Projetos\\pdf_xls_reader\\img\\logo.png"
        pil_image = Image.open(image_path)
        pil_image = pil_image.resize((80, 80), Image.BILINEAR)  # Redimensiona a imagem para o tamanho desejado
        tk_image = ImageTk.PhotoImage(pil_image)  # Use ImageTk.PhotoImage para criar a imagem

        # LOGO
        self.titulo = Label(text="", image=tk_image)
        self.titulo.image = tk_image  # Mantém uma referência à imagem
        self.titulo.grid(row=0, column=0,padx=(80, 0))

        # LABEL "SELECIONAR PDFs"
        self.buttonPDFText = Label(pady=8, text="Selecione o(s) PDF(s):")
        self.buttonPDFText.grid(row=1, column=0,padx=(80, 0))
        # BOTÃO DE SELEÇÃO
        self.buttonPDFLabel = Button(text="Selecionar", width=10, height=1, command=self.selecionar_pdf)
        self.buttonPDFLabel.grid(row=2, column=0,sticky="w",padx=(10, 5))

        # CAIXA DE TEXTO
        self.PDF = ScrollableFrame(master=None, width=255, height=100, corner_radius=0)
        self.PDF.grid(row=2, column=0, sticky="w", padx=(100, 0))

        # LABEL "SELECIONAR EXCEL"
        self.buttonXLSXText = Label(pady=5,
                                                     text="Selecione a planilha de Empresas:")
        self.buttonXLSXText.grid(row=3, column=0,padx=(80, 0))
        # BOTÃO DE SELEÇÃO
        self.buttonXLSXLabel = Button(text="Selecionar", width=10, height=1, command=self.selecionar_xlsx)
        self.buttonXLSXLabel.grid(row=4, column=0, sticky="w", padx=(10, 5))  # Ajuste o valor de padx aqui

        # CAIXA DE TEXTO
        self.XLSX = Entry(width=45)
        self.XLSX.grid(row=4, column=0, sticky="w", padx=(100, 0))

        # INICIAR
        self.botaoIniciar = Button(text="Iniciar Consulta", command=self.executar_thread)
        self.botaoIniciar.grid(row=6, column=0, pady=5)
        # Salvar
        self.salvarR = Button(text="Salvar",  command=self.salvar)
        self.salvarR.grid(row=7, column=0, pady = 5)
        # Salvar

    # FUNÇÕES
    def executar_thread(self):
        ExecutarThread(self).start()
    #
    # def change_appearance_mode_event(self):
    #     new_appearance_mode = self.appearance_mode.get()
    #     customtkinter.set_appearance_mode(new_appearance_mode)

    def selecionar_xlsx(self):
        self.xlsx_sheet = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo",
                                                     filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        if self.xlsx_sheet:
            self.arquivo_xlsx_selecionado = True

        self.XLSX.delete(0, END)
        self.XLSX.insert(END, self.xlsx_sheet)

    def selecionar_pdf(self):
        lista_nova = filedialog.askopenfilenames(initialdir="/", title="Selecione um arquivo",
                                                 filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))

        if lista_nova:
            self.lista_pdf = lista_nova
            self.arquivo_pdf_selecionado = True

            # Limpa o conteúdo do ScrollableFrame
            for widget in self.PDF.interior.winfo_children():
                widget.destroy()

            for i in self.lista_pdf:
                label = Label(self.PDF.interior, text=i)
                label.pack()

        return self.lista_pdf

    def iniciar(self):
        if not (self.arquivo_pdf_selecionado and self.arquivo_xlsx_selecionado):
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
        if self.df_final is None:
            messagebox.showwarning("Aviso", "Por favor, execute a consulta primeiro.")
            return
        salvar_resultado(self.df_final)

class ExecutarThread(threading.Thread):
    def __init__(self, application):
        threading.Thread.__init__(self)
        self.application = application

    def run(self):
        self.application.iniciar()


root = Tk()
root.geometry("380x390")
root.title("Consulta")
root.iconbitmap('C:\\Users\\Analise\\Desktop\\Projetos\\pdf_xls_reader\\img\\logo.ico')
app = Application(root)
root.mainloop()
