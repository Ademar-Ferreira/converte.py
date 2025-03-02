# converte.py
O programa que desenvolvemos suporta a conversão dos seguintes **formatos de arquivos de texto** e outros formatos relacionados:

---

### Formatos suportados:

1. **Documentos de texto**:
   - `.txt`: Arquivos de texto simples.
   - `.doc`: Documentos do Microsoft Word (versões antigas).
   - `.docx`: Documentos do Microsoft Word (versões mais recentes).

2. **Planilhas**:
   - `.xls`: Planilhas do Microsoft Excel (versões antigas).
   - `.xlsx`: Planilhas do Microsoft Excel (versões mais recentes).

3. **Arquivos HTML**:
   - `.html`: Arquivos de marcação HTML.
   - `.htm`: Alternativa para arquivos HTML.

---

### Como o programa funciona para cada formato:

1. **Arquivos `.txt`**:
   - O programa lê o conteúdo do arquivo de texto e o converte diretamente em PDF.

2. **Arquivos `.doc` e `.docx`**:
   - Usa o **LibreOffice** (ou `unoconv`) para converter documentos do Word em PDF.

3. **Arquivos `.xls` e `.xlsx`**:
   - Converte a planilha em um arquivo HTML temporário e, em seguida, converte o HTML em PDF usando o **pdfkit**.

4. **Arquivos `.html` e `.htm`**:
   - Converte diretamente o arquivo HTML em PDF usando o **pdfkit**.

---

### Dependências necessárias:

Para que o programa funcione corretamente, as seguintes ferramentas precisam estar instaladas no sistema:

1. **LibreOffice**:
   - Necessário para converter arquivos `.doc` e `.docx`.
   - No Windows, o caminho padrão é:
     ```
     C:\Program Files\LibreOffice\program\soffice.exe
     ```
   - No Linux, instale com:
     ```bash
     sudo apt-get install libreoffice
     ```

2. **wkhtmltopdf**:
   - Necessário para converter arquivos `.html` e `.htm`.
   - Baixe e instale a partir do site oficial: [wkhtmltopdf](https://wkhtmltopdf.org/).
   - No Linux, instale com:
     ```bash
     sudo apt-get install wkhtmltopdf
     ```

3. **Bibliotecas Python**:
   - Certifique-se de que as bibliotecas necessárias estão instaladas:
     ```bash
     pip install pandas pdfkit python-docx
     ```

---

### Exemplo de uso:

1. Execute o programa.
2. Clique em **Enviar Arquivo**.
3. Selecione um arquivo de texto, documento do Word, planilha do Excel ou arquivo HTML.
4. Aguarde a conversão.
5. Clique em **Baixar PDF** para abrir o arquivo convertido.

---

### Observações:

- O programa **não suporta** formatos como `.pdf`, `.ppt`, `.pptx`, `.odt`, `.ods`, etc. Se precisar de suporte para esses formatos, é possível adicionar funcionalidades adicionais.
- Certifique-se de que as dependências (LibreOffice e wkhtmltopdf) estejam instaladas e configuradas corretamente.
