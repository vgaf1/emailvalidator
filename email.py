import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def validar_email(email):
    regex = r'^[a-z0-9.+_-]{1,64}@[a-z0-9.-]{3,64}\.[a-z]{2,}$'
    return re.match(regex, email) is not None

def validar_emails_planilha(caminho_planilha, coluna_emails):
    # leitura da planilha
    try:
        df = pd.read_excel(caminho_planilha) if caminho_planilha.endswith('.xlsx') else pd.read_csv(caminho_planilha)
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo não encontrado: {caminho_planilha}")
        return [], [], None

    # validacão dos emails
    df['Valido'] = df[coluna_emails].apply(validar_email)

    emails_validos = df[df['Valido'] == True][coluna_emails].tolist()
    emails_invalidos = df[df['Valido'] == False][coluna_emails].tolist()

    return emails_validos, emails_invalidos, df

#salvando resultado em arquivo excel
def salvar_resultado(df, caminho_saida):
    df.to_excel(caminho_saida, index=False)
    messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em: {caminho_saida}")

def escolher_arquivo():
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
    if caminho_arquivo:
        coluna_emails = coluna_entry.get()
        emails_validos, emails_invalidos, df = validar_emails_planilha(caminho_arquivo, coluna_emails)

        lista_validos.delete(0, tk.END)
        lista_invalidos.delete(0, tk.END)

        for email in emails_validos:
            lista_validos.insert(tk.END, email)

        for email in emails_invalidos:
            lista_invalidos.insert(tk.END, email)

        if df is not None:
            caminho_saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if caminho_saida:
                salvar_resultado(df, caminho_saida)

# criacão da interface grafica
root = tk.Tk()
root.title("Validador de Emails")

frame = tk.Frame(root)
frame.pack(pady=10)

coluna_label = tk.Label(frame, text="Nome da Coluna de Emails:")
coluna_label.grid(row=0, column=0, padx=5, pady=5)

coluna_entry = tk.Entry(frame)
coluna_entry.grid(row=0, column=1, padx=5, pady=5)

botao_arquivo = tk.Button(frame, text="Escolher Arquivo", command=escolher_arquivo)
botao_arquivo.grid(row=0, column=2, padx=5, pady=5)

frame_listas = tk.Frame(root)
frame_listas.pack(pady=10)

label_validos = tk.Label(frame_listas, text="Emails Válidos")
label_validos.grid(row=0, column=0, padx=10)

label_invalidos = tk.Label(frame_listas, text="Emails Inválidos")
label_invalidos.grid(row=0, column=1, padx=10)

lista_validos = tk.Listbox(frame_listas, width=40, height=20)
lista_validos.grid(row=1, column=0, padx=10, pady=10)

lista_invalidos = tk.Listbox(frame_listas, width=40, height=20)
lista_invalidos.grid(row=1, column=1, padx=10, pady=10)

root.mainloop()
