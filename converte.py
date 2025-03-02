import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pdfkit
import subprocess
import threading

# Função para converter DOC/DOCX para PDF
def convert_doc_to_pdf(input_file, output_file):
    try:
        # Caminho do executável do LibreOffice
        libreoffice_path = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
        
        # Comando para converter o arquivo
        command = [
            libreoffice_path,
            '--headless',  # Executa em modo headless (sem interface gráfica)
            '--convert-to', 'pdf',
            '--outdir', os.path.dirname(output_file),
            input_file
        ]
        
        # Executa o comando
        subprocess.run(command, check=True)
        
        # Verifica se o arquivo PDF foi criado
        return os.path.exists(output_file)
    except Exception as e:
        print(f"Erro ao converter DOC/DOCX para PDF: {e}")
        return False

# Função para converter HTML para PDF
def convert_html_to_pdf(input_file, output_file):
    try:
        # Configura o pdfkit para usar o wkhtmltopdf
        config = pdfkit.configuration(wkhtmltopdf='C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')  # Ajuste o caminho conforme necessário
        pdfkit.from_file(input_file, output_file, configuration=config)
        return os.path.exists(output_file)  # Verifica se o arquivo PDF foi criado
    except Exception as e:
        print(f"Erro ao converter HTML para PDF: {e}")
        return False

# Função para converter XLS/XLSX para PDF
def convert_xlsx_to_pdf(input_file, output_file):
    try:
        # Lê o arquivo XLS/XLSX
        df = pd.read_excel(input_file)
        
        # Cria um arquivo HTML temporário
        html_file = input_file.rsplit('.', 1)[0] + '.html'
        df.to_html(html_file)
        
        # Converte o HTML para PDF
        config = pdfkit.configuration(wkhtmltopdf='C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')  # Ajuste o caminho conforme necessário
        pdfkit.from_file(html_file, output_file, configuration=config)
        
        # Remove o arquivo HTML temporário
        os.remove(html_file)
        return os.path.exists(output_file)  # Verifica se o arquivo PDF foi criado
    except Exception as e:
        print(f"Erro ao converter XLS/XLSX para PDF: {e}")
        return False

# Função principal de conversão
def convert_to_pdf(input_file, output_file):
    file_extension = input_file.split('.')[-1].lower()
    
    if file_extension in ['doc', 'docx']:
        return convert_doc_to_pdf(input_file, output_file)
    elif file_extension in ['html', 'htm']:
        return convert_html_to_pdf(input_file, output_file)
    elif file_extension in ['xls', 'xlsx']:
        return convert_xlsx_to_pdf(input_file, output_file)
    else:
        print("Formato de arquivo não suportado.")
        return False

# Função para simular o progresso
def simulate_progress():
    for i in range(101):
        progress_bar['value'] = i
        root.update_idletasks()
        root.after(30)  # Ajuste o tempo para simular o progresso

# Função para enviar o arquivo
def upload_file():
    file_path = filedialog.askopenfilename(
        title="Selecione um arquivo",
        filetypes=(
            ("Arquivos de texto", "*.txt"),
            ("Documentos Word", "*.doc *.docx"),
            ("Planilhas Excel", "*.xls *.xlsx"),
            ("Arquivos HTML", "*.html *.htm"),
            ("Todos os arquivos", "*.*")
        )
    )
    
    if file_path:
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path) / 1024  # Tamanho em KB
        file_type = file_name.split('.')[-1].upper()
        
        # Atualiza as informações na interface
        file_info_label.config(text=f"Nome: {file_name}\nTamanho: {file_size:.2f} KB\nTipo: {file_type}")
        
        # Inicia a barra de progresso
        progress_bar['value'] = 0
        progress_bar.start()
        
        # Converte o arquivo para PDF em uma thread separada
        output_file = file_path.rsplit('.', 1)[0] + '.pdf'
        threading.Thread(target=convert_file, args=(file_path, output_file)).start()

# Função para converter o arquivo em uma thread separada
def convert_file(input_file, output_file):
    # Simula o progresso
    simulate_progress()
    
    # Converte o arquivo para PDF
    success = convert_to_pdf(input_file, output_file)
    
    # Para a barra de progresso
    progress_bar.stop()
    
    # Exibe o resultado
    if success:
        messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
        download_button.config(state=tk.NORMAL, command=lambda: open_file(output_file))
    else:
        messagebox.showerror("Erro", "Erro ao converter o arquivo. Tente novamente.")

# Função para abrir o arquivo PDF
def open_file(file_path):
    if os.name == 'nt':  # Windows
        os.startfile(file_path)
    else:  # Linux/Mac
        os.system(f'xdg-open {file_path}')

# Configuração da interface gráfica
root = tk.Tk()
root.title("Conversor de Arquivos para PDF")
root.geometry("400x300")

# Título
title_label = tk.Label(root, text="Converter arquivos para PDF", font=("Arial", 14))
title_label.pack(pady=10)

# Botão para enviar arquivo
upload_button = tk.Button(root, text="Enviar Arquivo", command=upload_file)
upload_button.pack(pady=10)

# Informações do arquivo
file_info_label = tk.Label(root, text="Nenhum arquivo selecionado", justify=tk.LEFT)
file_info_label.pack(pady=10)

# Barra de progresso
progress_bar = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=300, mode='determinate')
progress_bar.pack(pady=10)

# Botão para baixar PDF (inicialmente desabilitado)
download_button = tk.Button(root, text="Baixar PDF", state=tk.DISABLED)
download_button.pack(pady=10)

# Rodar a interface
root.mainloop()