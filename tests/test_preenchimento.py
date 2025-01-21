import os
from src.preencher_certificados import preencher_certificados

def test_preenchimento_certificados():
    preencher_certificados(
        "data/alunos.xlsx",
        "data/modelo_certificado.docx",
        "output/"
    )
    assert len(os.listdir("output/")) > 0, "Certificados não foram gerados."
