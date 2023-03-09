import tabula
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import os

class PDFToExcelConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.interaction_window = tk.Tk()
        self.interaction_window.title("Conversor PDF para Excel")
        self.interaction_window.geometry("500x250")
        self.progress_value = tk.DoubleVar()
    

        # Cria o botão para selecionar o arquivo PDF
        self.pdf_button = tk.Button(self.interaction_window, text="Selecionar PDF", command=self.select_pdf_file)
        self.pdf_button.pack(pady=10)

        # Cria o rótulo para mostrar o caminho do arquivo PDF selecionado
        self.pdf_path_label = tk.Label(self.interaction_window, text="Nenhum arquivo PDF selecionado")
        self.pdf_path_label.pack()

        # Cria o botão para selecionar o local para salvar o arquivo Excel
        self.excel_button = tk.Button(self.interaction_window, text="Salvar Excel", command=self.save_excel_file)
        self.excel_button.pack(pady=10)

         # Cria o rótulo para mostrar o caminho que o arquivo do excel vai ser salvo
        self.excel_path_label = tk.Label(self.interaction_window, text="Nenhum destino selecionado")
        self.excel_path_label.pack()


        # Cria o label para mostrar a porcentagem de progresso
        self.progress_label = tk.Label(self.interaction_window, text="0%")
        self.progress_label.pack()

        # Cria a barra de progresso para mostrar a % da conversão realizada
        self.progress_bar = tk.ttk.Progressbar(self.interaction_window, orient="horizontal", length=200, mode="determinate")
        self.progress_bar.pack(pady=10)


        # Cria o botão para iniciar a conversão
        self.convert_button = tk.Button(self.interaction_window, text="Converter", command=self.convert_pdf_to_excel)
        self.convert_button.pack(pady=10)

        self.interaction_window.mainloop()

    def select_pdf_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.pdf_path_label.configure(text=file_path)

    def save_excel_file(self):
        save_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if save_path:
            self.save_path = save_path
            self.excel_path_label.configure(text=save_path)

    def convert_pdf_to_excel(self):
       
        file_path = self.pdf_path_label.cget("text")
        save_path = self.excel_path_label.cget("text")
        try:
            if file_path:
                # Le o pdf na lista do dataframe
                dfs = tabula.read_pdf(file_path, pages='all')

                # convert PDF em CSV 
                tabula.convert_into(file_path, "output.csv", output_format="csv", pages='all')
                read_file = pd.read_csv ('output.csv',  encoding='latin-1')

                # Verificando se um local foi selecionado
                if save_path:
                    self.progress_bar["maximum"] = len(dfs)
                    for page_number, df in enumerate(dfs):
                        # Atualiza a barra de progresso a cada página convertida
                        self.progress_bar["value"] = page_number + 1
                        self.progress_label.configure(text="100%")
                        self.interaction_window.update()

                    read_file.to_excel (save_path, index = None, header=True)
                    os.remove('output.csv')
                    messagebox.showinfo("Conversão concluída", "O arquivo Excel foi salvo com sucesso.")
                else:
                    messagebox.showerror("Erro", "O local para salvar o arquivo Excel não foi selecionado.")
            else:
                messagebox.showerror("Erro", "O arquivo PDF não foi selecionado.")
        except:
            messagebox.showerror("Erro", "Não foi possivel realizar a conversão.")
        finally:
            # Reseta o valor da barra de progresso
            self.progress_bar["value"] = 0
            self.progress_label.configure(text="")



pdf_converter = PDFToExcelConverter()
pdf_converter.root.mainloop()

#criando um executavel pyinstaller --onefile --icon=meu_icone.ico nome_do_arquivo.py

