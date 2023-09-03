from tkinter import messagebox, filedialog, LEFT, BOTTOM, TOP, RIGHT, StringVar, END, PhotoImage,Tk, Label, Button, Entry,Listbox, Scrollbar, Text, Frame, Canvas
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


# class ScrollableFrame(Frame):
#     def __init__(self, master, width, height, corner_radius):
#         super().__init__(master)
#         try:
#             self.canvas = Canvas(self, width=width, height=height)  # Defina o atributo canvas aqui
#             scrollbar = Scrollbar(self, orient="vertical", command=self.canvas.yview)
#             self.canvas.config(yscrollcommand=scrollbar.set)
#
#             scrollbar.pack(side="right", fill="y")
#             self.canvas.pack(side="left", fill="both", expand=True)
#
#             self.interior = Frame(self.canvas)  # Use self.canvas como parent
#             self.canvas.create_window((0, 0), window=self.interior, anchor="nw")
#
#             self.interior.bind("<Configure>", self.set_scrollregion)
#         except Exception as e:
#             print(f'Erro no ScrollableFrame {e}')
#     def set_scrollregion(self, event=None):
#         try:
#             self.canvas.configure(scrollregion=self.canvas.bbox("all"))
#         except Exception as e:
#             print(f'Erro no set_scrollregion {e}')

def remove_special_characters(str1):
    try:
        text = str1
        return re.sub(r"\D", '', text)

    except Exception as e:
        print(f'Erro no remove_special_characters {e}')

def salvar_resultado(df_final):
    try:
        folder_path = filedialog.askdirectory(title="Selecione a pasta para salvar os resultados")
        folder_path = Path(folder_path)
        current_date = datetime.datetime.now().strftime('%d-%m-%Y %Hh%M')
        result_file = folder_path.joinpath(f'RESULTADO - {current_date}')
        df_final.to_excel(result_file.with_suffix('.xlsx'), index=False, header=True, sheet_name='EMPRESAS')
        print("salvo resultado")
    except Exception as e:
        print(f'Erro no salvar_resultado {e}')

def converter_pdf_para_txt(lista_pdf):
    try:
        pdf_text_dict = {}  # Dicionário para armazenar o texto de cada PDF
        for pdf_file in lista_pdf:
            pdf_name = os.path.basename(pdf_file)  # Obtém o nome do arquivo PDF
            with pdfplumber.open(pdf_file) as pdf:
                text = ''.join([page.extract_text() for page in pdf.pages])  # Extrai o texto de todas as páginas do PDF
                pdf_text_dict[pdf_name] = text  # Armazena o texto no dicionário usando o nome do PDF como chave

        return pdf_text_dict

    except Exception as e:
        print(f'Erro no converter_pdf_para_txt {e}')

class Application:
    def __init__(self, master):
        try:
            self.df_final = None
            self.xlsx_sheet = None
            self.lista_pdf = None
            self.master = master
            master.resizable(0, 0)
            self.fontePadrao = ("Calibri", 20)
            self.arquivo_pdf_selecionado = False
            self.arquivo_xlsx_selecionado = False

            # LOGO DO PROJETO
            image_path = "C:\\Users\\Analise\\Desktop\\Projetos\\pdf_xls_reader\\img\\logo.png"
            pil_image = Image.open(image_path)
            pil_image = pil_image.resize((80, 80), Image.BILINEAR)  # Redimensiona a imagem para o tamanho desejado
            tk_image = ImageTk.PhotoImage(pil_image)  # Use ImageTk.PhotoImage para criar a imagem

            # LOGO
            self.titulo = Label(text="", image=tk_image)
            self.titulo.image = tk_image  # Mantém uma referência à imagem
            self.titulo.grid(row=0, column=0,padx=(1, 0))

            self.imagem_concluido = Image.open("C:\\Users\\Analise\\Desktop\\Projetos\\pdf_xls_reader\\img\\check.png")
            self.imagem_concluido = self.imagem_concluido.resize((15, 15), Image.BILINEAR)
            self.imagem_concluido = ImageTk.PhotoImage(self.imagem_concluido)

            # # LABEL "Diario oficial"
            # self.buttonPDFText = Label(pady=8, text="Consulta Diario Oficial")
            # self.buttonPDFText.grid(row=0, column=0,padx=(0, 0), pady=(120,0))

            # LABEL "SELECIONAR PDFs"
            self.buttonPDFText = Label(pady=8, text="Selecione o(s) PDF(s):")
            self.buttonPDFText.grid(row=1, column=0,padx=(1, 0), pady=(20, 0))

            # BOTÃO DE SELEÇÃO
            self.buttonPDFLabel = Button(text="Selecionar", width=10, height=1, command=self.selecionar_pdf)
            self.buttonPDFLabel.grid(row=2, column=0,sticky="w",padx=(10, 5), pady=(0,50))

            # BOTÃO DE SELEÇÃO
            self.buttonPDFdelete = Button(text="Deletar", width=10, height=1, command=self.deletar_selecionados)
            self.buttonPDFdelete.grid(row=2, column=0,sticky="w",padx=(10, 5), pady=(50,0))

            # CAIXA DE TEXTO PDFs
            self.scroll_bar = Scrollbar(root)
            self.scroll_bar.grid(row=2, column=0, sticky="e", padx=(100, 0), ipady=60)

            self.PDF = Listbox(master=None,selectmode="multiple", yscrollcommand = self.scroll_bar.set)
            self.PDF.grid(row=2, column=0, sticky="e", padx=(0, 18), ipadx=65)

            # LABEL "SELECIONAR EXCEL"
            self.buttonXLSXText = Label(pady=5,text="Selecione a planilha de Empresas:")
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
            self.botaoIniciar.config(state="disabled")

            # Salvar
            self.salvarR = Button(text="Salvar",  command=self.salvar)
            self.salvarR.grid(row=7, column=0, pady=5,ipadx=25)
            self.salvarR.config(state="disabled")

        except Exception as e:
            print(f'Erro no Application {e}')


    # FUNÇÕES
    def executar_thread(self):
        try:
            ExecutarThread(self).start()

        except Exception as e:
            print(f'Erro no executar_thread {e}')
    def selecionar_xlsx(self):
        self.remover_imagem()
        try:
            self.xlsx_sheet = filedialog.askopenfilename(initialdir="/", title="Selecione um arquivo",
                                                     filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
            if self.xlsx_sheet:
                self.arquivo_xlsx_selecionado = True

            self.XLSX.delete(0, END)
            self.XLSX.insert(END, self.xlsx_sheet)
            self.verif_ati_bt()
            if not self.xlsx_sheet:
                self.xlsx_sheet = None
                self.verif_ati_bt()
                self.remover_imagem()
            print("Arquivo selecionado")
        except Exception as e:
            print(f'Erro no selecionar_xlsx {e}')

    def selecionar_pdf(self):
        self.remover_imagem()
        try:
            lista_nova = filedialog.askopenfilenames(initialdir="/", title="Selecione um arquivo",
                                                     filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))

            if lista_nova:
                if self.lista_pdf is None:
                    self.lista_pdf = []

                self.lista_pdf.extend(lista_nova)
                self.arquivo_pdf_selecionado = True

                for i in self.lista_pdf:
                    self.PDF.insert(END, os.path.basename(i))

            self.verif_ati_bt()
            print("PDFs selecionados")
            return self.lista_pdf

        except Exception as e:
            print(f'Erro no selecionar_pdf {e}')

    def deletar_selecionados(self):
        selecionados = self.PDF.curselection()  # Obtém os índices dos itens selecionados
        for indice in reversed(selecionados):  # Usamos reversed() para deletar de trás para frente
            self.lista_pdf.remove(self.lista_pdf[indice])
            if not self.lista_pdf:
                self.lista_pdf = None
                self.verif_ati_bt()
            self.PDF.delete(indice)  # Deleta o item no índice especificado
        self.remover_imagem()

    def verif_ati_bt(self):
        if self.xlsx_sheet is not None and self.lista_pdf is not None:
            self.botaoIniciar.config(state="normal")
        else:
            self.botaoIniciar.config(state="disabled")

    def verif_ati_bt_salvar(self):
        if self.df_final is not None:
            self.salvarR.config(state="normal")
        else:
            self.salvarR.config(state="disabled")

    def remover_imagem(self):
        if hasattr(self, 'imagem_concluido_label'):
            self.imagem_concluido_label.grid_forget()
            self.imagem_concluido_label1.grid_forget()

    def iniciar(self):
        global pdf_sources

        try:
            if not (self.arquivo_pdf_selecionado and self.arquivo_xlsx_selecionado):
                messagebox.showerror("Erro", "Selecione os PDFs e a planilha antes de iniciar o processo.")
                return
            lista_pdf = self.lista_pdf
            pdf_text_dict = converter_pdf_para_txt(lista_pdf)
            resultados_intermediarios = []  # Lista para armazenar os dataframes intermediários
            pdf_sources = []
            for chave_pdf, valor_text in pdf_text_dict.items():
                chave_pdf = chave_pdf
                pdf_text = " ".join(valor_text.split())
                pdf_lines = pdf_text.split("\n")


                xlsx_file = self.xlsx_sheet
                df = pd.read_excel(xlsx_file)
                cnpj_key = df.iloc[:, 2][3:]
                value = df.iloc[:, 0][3:]
                caceal_key = df.iloc[:, 25][3:]
                # Crie a lista para armazenar as fontes dos PDFs


                resultado_dict = {}
                for chave, valor in zip(cnpj_key, value):
                    if isinstance(chave, str):
                        for linha in pdf_lines:
                            if str(chave).lower() in str(linha).lower():
                                resultado_dict[chave] = valor
                                if not cnpj_key.empty and not value.empty: # Verifica se os valores estão presentes
                                    pdf_sources.append(
                                        os.path.basename(chave_pdf))  # Adicione o nome do PDF à lista de fontes
                                break
                print(
                    f"Resultado dict: {resultado_dict}")  # Adiciona um print para verificar o dicionário resultado_dict

                resultado_dict1 = {}
                for chave, valor in zip(cnpj_key, value):
                    chave = str(chave)
                    stripped_key = remove_special_characters(chave)
                    if isinstance(chave, str):
                        if stripped_key.startswith(chave.replace(" ", "")):
                            for linha in pdf_lines:
                                if str(chave).lower() in str(linha).lower():
                                    resultado_dict1[chave] = valor
                                    if not cnpj_key.empty and not value.empty:
                                        # Verifica se os valores estão presentes
                                        pdf_sources.append(
                                            os.path.basename(chave_pdf))  # Adicione o nome do PDF à lista de fontes
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
                            if str(chave).lower() in str(linha).lower():
                                resultado_dict3[chave] = valor
                                if not caceal_key and not value:  # Verifica se os valores estão presentes
                                    pdf_sources.append(
                                        os.path.basename(chave_pdf))# Adicione o nome do PDF à lista de fontes

                                break

                print(
                    f"Resultado dict3: {resultado_dict3}")  # Adiciona um print para verificar o dicionário resultado_dict3

                df_resultado = pd.DataFrame(list(resultado_dict.items()), columns=['CNPJ/CACEAL', 'EMPRESA'])
                df_resultado1 = pd.DataFrame(list(resultado_dict1.items()), columns=['CNPJ/CACEAL', 'EMPRESA'])
                df_resultado3 = pd.DataFrame(list(resultado_dict3.items()), columns=['CNPJ/CACEAL', 'EMPRESA'])

                df_final_intermediario = pd.concat([df_resultado, df_resultado1, df_resultado3], ignore_index=True)
                resultados_intermediarios.append(df_final_intermediario)


            # Concatena os dataframes intermediários para obter o dataframe final
            df_final = pd.concat(resultados_intermediarios, ignore_index=True)

            df_final["ARQUIVO"] = pdf_sources

            self.df_final = df_final
            self.imagem_concluido_label1 = Label(image=self.imagem_concluido)
            self.imagem_concluido_label1.grid(row=6, column=0, padx=(120, 0))
            self.verif_ati_bt_salvar()
            print("Finalizado")
            return None


        except Exception as e:
            print(f'Erro no iniciar {e}')


    def salvar(self):
        try:
            if self.df_final is None:
                messagebox.showwarning("Aviso", "Por favor, execute a consulta primeiro.")
                return
            salvar_resultado(self.df_final)
            self.imagem_concluido_label = Label(image=self.imagem_concluido)
            self.imagem_concluido_label.grid(row=7, column=0, padx=(120, 0))
        except Exception as e:
            print(f'Erro no salvar {e}')

class ExecutarThread(threading.Thread):
    def __init__(self, application):
        try:
            threading.Thread.__init__(self)
            self.application = application
        except Exception as e:
            print(f'Erro no ExecutarThread {e}')

    def run(self):
        try:
            self.application.iniciar()
        except Exception as e:
            print(f'Erro no run {e}')

root = Tk()
root.geometry("380x450")
root.title("Consulta")
root.iconbitmap('C:\\Users\\Analise\\Desktop\\Projetos\\pdf_xls_reader\\img\\logo.ico')
app = Application(root)
root.mainloop()