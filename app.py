import customtkinter as ctk
from tkinter import filedialog, messagebox
from gera_csv_do_pdf import extrair_nomes_de_pdf, gerar_csv, gerar_word_do_csv
import os

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def selecionar_pdf():
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not caminho_pdf:
        return

    try:
        nomes_com_rm = extrair_nomes_de_pdf(caminho_pdf)

        if not nomes_com_rm:
            messagebox.showwarning("Aviso", "Nenhum aluno encontrado no PDF.")
            return

        base = os.path.splitext(os.path.basename(caminho_pdf))[0]
        unidade = unidade_menu.get()

        caminho_csv = gerar_csv(nomes_com_rm, turma=base, unidade=unidade)
        caminho_docx = f"usuarios_{base}.docx"

        gerar_word_do_csv(
            caminho_csv,
            caminho_word=caminho_docx,
            caminho_imagem="informacao_acesso.jpg",
            nomes_com_rm=nomes_com_rm
        )

        messagebox.showinfo("Sucesso", "CSV e DOCX gerados com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


def processar_texto():
    texto = texto_box.get("1.0", "end").strip()

    if not texto:
        messagebox.showwarning("Atenção", "Digite pelo menos um nome.")
        return

    nomes = [(nome.strip(), "") for nome in texto.split("\n") if nome.strip()]

    unidade = unidade_menu.get()
    base = "manual"

    try:
        caminho_csv = gerar_csv(nomes, turma=base, unidade=unidade)
        caminho_docx = f"usuarios_{base}.docx"

        gerar_word_do_csv(
            caminho_csv,
            caminho_word=caminho_docx,
            caminho_imagem="informacao_acesso.jpg",
            nomes_com_rm=nomes
        )

        messagebox.showinfo("Sucesso", "Arquivos gerados com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


# ===============================
# INTERFACE
# ===============================
app = ctk.CTk()
app.title("Gerador de Acesso - CSV + DOCX")
app.geometry("500x600")

ctk.CTkLabel(app, text="Escolha como deseja inserir os nomes:", font=("Arial", 16)).pack(pady=15)

ctk.CTkLabel(app, text="Unidade:", anchor="w").pack(pady=(5, 0))

unidades_disponiveis = [
    "/Alunos/Alunos - centro",
    "/Alunos/Alunos - cohab",
    "/Alunos/Alunos - raposa",
    "/Unidade Bacabal/Alunos - bacabal"
]

unidade_menu = ctk.CTkOptionMenu(app, values=unidades_disponiveis)
unidade_menu.set(unidades_disponiveis[0])
unidade_menu.pack(pady=5)

btn_pdf = ctk.CTkButton(app, text="Importar PDF com nomes", command=selecionar_pdf)
btn_pdf.pack(pady=10)

ctk.CTkLabel(app, text="ou", font=("Arial", 12)).pack(pady=5)

ctk.CTkLabel(app, text="Digite um nome por linha:", anchor="w").pack(pady=(10, 0))
texto_box = ctk.CTkTextbox(app, height=200, width=400)
texto_box.pack(pady=10)

btn_manual = ctk.CTkButton(app, text="Gerar arquivos com nomes digitados", command=processar_texto)
btn_manual.pack(pady=20)

app.mainloop()