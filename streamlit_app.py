import fitz  # PyMuPDF
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

def combine_pdf_and_attachments(pdf_path, folder_path, output_path):
    if not os.path.isfile(pdf_path):
        messagebox.showerror("Feil", f"PDF-filen '{pdf_path}' ble ikke funnet.")
        return

    folder_path = Path(folder_path)
    combined_document = fitz.open()
    original_document = fitz.open(pdf_path)

    for page_num in range(len(original_document)):
        # Legg til en side av målebrevet
        page = original_document.load_page(page_num)
        combined_document.insert_pdf(original_document, from_page=page_num, to_page=page_num)
        
        # Finn vedleggene nevnt på denne siden
        text = page.get_text("text")
        links_text = text.split("Vedlagte dokumenter:")[1].strip().split("\n") if "Vedlagte dokumenter:" in text else []

        for link_text in links_text:
            link_text = link_text.strip()
            if "\\" in link_text and "." in link_text:
                file_name = link_text.split("\\")[-1].strip()
                file_path = folder_path / file_name

                print(f"Prøver å finne filen: {file_path}")

                if file_path.is_file():
                    print(f"Filen '{file_path}' funnet, legger til i PDF.")
                    attachment_document = fitz.open(file_path)
                    combined_document.insert_pdf(attachment_document)
                    attachment_document.close()
                else:
                    print(f"Advarsel: Filen '{file_name}' ble ikke funnet.")
                    continue

    original_document.close()

    combined_document.save(output_path)
    combined_document.close()
    messagebox.showinfo("Suksess", f"Kombinert PDF lagret som: {output_path}")
    print(f"Kombinert PDF lagret som: {output_path}")

def select_pdf_file():
    file_path = filedialog.askopenfilename(title="Velg PDF-fil med målebrev", filetypes=[("PDF filer", "*.pdf")])
    return file_path

def select_folder():
    folder_path = filedialog.askdirectory(title="Velg mappe med vedlegg")
    return folder_path

def select_output_file():
    output_path = filedialog.asksaveasfilename(title="Lagre kombinert PDF som", defaultextension=".pdf", filetypes=[("PDF filer", "*.pdf")])
    return output_path

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    pdf_path = select_pdf_file()
    if not pdf_path:
        messagebox.showerror("Feil", "Ingen PDF-fil valgt. Avslutter.")
    else:
        folder_path = select_folder()
        if not folder_path:
            messagebox.showerror("Feil", "Ingen mappe valgt. Avslutter.")
        else:
            output_path = select_output_file()
            if not output_path:
                messagebox.showerror("Feil", "Ingen lagringsplass valgt. Avslutter.")
            else:
                combine_pdf_and_attachments(pdf_path, folder_path, output_path)
