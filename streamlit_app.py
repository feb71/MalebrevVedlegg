import streamlit as st
import fitz  # PyMuPDF
from pathlib import Path
import re

def combine_pdf_and_attachments(pdf_file, folder_files):
    # Åpne hoved-PDF-filen fra opplastingen
    combined_document = fitz.open()
    original_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    
    # Opprett en dictionary for vedleggene, med bare filnavnet (uten sti) som nøkkel
    folder_dict = {Path(file.name).name: file for file in folder_files}

    # Iterer gjennom hver side i hoved-PDF-filen
    for page_num in range(len(original_document)):
        # Legg til en side av målebrevet
        page = original_document.load_page(page_num)
        combined_document.insert_pdf(original_document, from_page=page_num, to_page=page_num)
        
        # Finn vedleggene nevnt på denne siden
        text = page.get_text("text")
        links_text = text.split("Vedlagte dokumenter:")[1].strip().split("\n") if "Vedlagte dokumenter:" in text else []

        for link_text in links_text:
            # Fjern eventuelle ekstra mellomrom og unødvendige tegn
            link_text = link_text.strip().replace("\\", "/").split("/")[-1]
            
            # Sjekk om det ser ut som et gyldig filnavn (f.eks., må slutte med .pdf)
            if not re.match(r'.+\.pdf$', link_text):
                continue  # Hopp over hvis ikke en gyldig filnavnstruktur
            
            if link_text in folder_dict:
                attachment_file = folder_dict[link_text]
                st.info(f"Legger til vedlegget: {link_text} for side {page_num + 1}")
                
                # Gå til starten av filen før lesing
                attachment_file.seek(0)
                
                # Åpne vedleggs-PDF-en
                attachment_document = fitz.open(stream=attachment_file.read(), filetype="pdf")
                
                # Sett inn alle sidene i vedleggsdokumentet
                combined_document.insert_pdf(attachment_document)
                attachment_document.close()
            else:
                st.warning(f"Fant ikke vedlegget: {link_text} i opplastede filer")

    # Lagre det kombinerte dokumentet
    output_path = 'kombinert_dokument.pdf'
    combined_document.save(output_path)
    combined_document.close()

    return output_path

# Streamlit-grensesnittet
st.title("Kombiner målebrev med Vedlegg basert på forsidespesifikasjoner")

# Last opp hoved-PDF-filen
pdf_file = st.file_uploader("Last opp hoved PDF-filen", type="pdf")

# Last opp alle vedleggs-PDF-filene
folder_files = st.file_uploader("Last opp vedleggs-PDF-filer", type="pdf", accept_multiple_files=True)

if pdf_file is not None and folder_files:
    # Kombiner hoved-PDF og vedleggene
    st.write("Behandler filene, vennligst vent...")
    output_path = combine_pdf_and_attachments(pdf_file, folder_files)
    
    st.success("Kombinering fullført!")
    
    # Tilby den kombinerte filen for nedlasting
    with open(output_path, "rb") as f:
        st.download_button("Last ned kombinert PDF", f, file_name="kombinert_dokument.pdf")
