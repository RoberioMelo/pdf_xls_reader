import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from threading import Thread
from time import sleep
from PyPDF2 import PdfReader
import os


def converter_pdf_para_txt(caminho_pdf):
    with open(caminho_pdf, 'rb') as arquivo_pdf:
        leitor_pdf = PdfReader(arquivo_pdf)
        conteudo = ""
        for pagina in leitor_pdf.pages:
            conteudo += pagina.extract_text()
        return conteudo


def selecionar_pdf():
    caminhos_pdf = filedialog.askopenfilenames(title="Selecione os arquivos PDF", filetypes=(("Arquivos PDF", "*.pdf"),))
    for caminho_pdf in caminhos_pdf:
        lista_pdf.insert(tk.END, caminho_pdf)


def excluir_pdf():
    indices_selecionados = lista_pdf.curselection()
    for index in reversed(indices_selecionados):
        lista_pdf.delete(index)


def iniciar_conversao():
    caminhos_pdf = lista_pdf.get(0, tk.END)
    if not caminhos_pdf:
        messagebox.showwarning("Aviso", "Selecione pelo menos um arquivo PDF.")
        return

    pasta_destino = filedialog.askdirectory(title="Selecione a pasta de destino para os arquivos TXT")
    if not pasta_destino:
        return

    # Função para realizar a conversão em segundo plano
    def realizar_conversao():
        for caminho_pdf in caminhos_pdf:
            sleep(1)  # Simula o tempo de processamento
            nome_pdf = os.path.basename(caminho_pdf)
            nome_txt = os.path.splitext(nome_pdf)[0] + ".txt"
            conteudo_txt = converter_pdf_para_txt(caminho_pdf)
            caminho_txt = os.path.join(pasta_destino, nome_txt)
            with open(caminho_txt, 'w') as arquivo_txt:
                arquivo_txt.write(conteudo_txt)

            # Atualiza a janela principal para processar eventos
            janela.update()

        messagebox.showinfo("Conversão Concluída", "A conversão de PDF para TXT foi concluída com sucesso!")

    # Cria uma thread para executar a conversão
    t = Thread(target=realizar_conversao)
    t.start()


# Janela principal
janela = tk.Tk()
janela.title("Capturador de Chaves")
janela.geometry("300x300")

# Botão para selecionar os PDFs
btn_selecionar_pdf = tk.Button(janela, text="Selecionar PDFs", command=selecionar_pdf)
btn_selecionar_pdf.pack(pady=10)

# Lista para exibir os PDFs selecionados
lista_pdf = tk.Listbox(janela, selectmode=tk.MULTIPLE)
lista_pdf.pack(pady=10)

# Botão para excluir PDFs da lista
btn_excluir_pdf = tk.Button(janela, text="Excluir Selecionado", command=excluir_pdf)
btn_excluir_pdf.pack(pady=5)

# Botão para iniciar o processo de conversão
btn_iniciar_conversao = tk.Button(janela, text="Iniciar Conversão", command=iniciar_conversao)
btn_iniciar_conversao.pack(pady=10)

janela.mainloop()
