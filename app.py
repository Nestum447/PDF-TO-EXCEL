import streamlit as st
from img2table.document import Image, PDF
from img2table.ocr import TesseractOCR
import tempfile
import os

st.title("Convertidor de PDF o Imagen a Excel")

archivo = st.file_uploader(
    "Sube un PDF o imagen",
    type=["pdf", "png", "jpg", "jpeg"]
)

if archivo is not None:

    tmp = tempfile.NamedTemporaryFile(delete=False)
    tmp.write(archivo.read())

    ocr = TesseractOCR(lang="eng")

    if archivo.type == "application/pdf":
        doc = PDF(tmp.name)
    else:
        doc = Image(tmp.name)

    salida = "resultado.xlsx"

    doc.to_xlsx(dest=salida, ocr=ocr)

    st.success("Excel generado correctamente")

    with open(salida, "rb") as f:
        st.download_button(
            "Descargar Excel",
            f,
            file_name="resultado.xlsx"
        )

    os.remove(tmp.name)
