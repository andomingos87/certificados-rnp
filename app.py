import csv
import io
import os
import smtplib
import ssl
import zipfile
from email.message import EmailMessage

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv
from email_validator import validate_email, EmailNotValidError
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject

# ───────────────────────── Constantes ─────────────────────────
CSV_COL_NAME = "Nome completo"
CSV_COL_MAIL = "E-mail"
PDF_FIELD = "{{ Nome do aluno }}"

# separadores disponíveis
SEP_LABELS = [
    "Auto (detectar)",
    "Vírgula ,",
    "Ponto e vírgula ;",
    "Tabulação \\t",
    "Pipe |",
]
SEP_MAP = {
    "Auto (detectar)": None,
    "Vírgula ,": ",",
    "Ponto e vírgula ;": ";",
    "Tabulação \\t": "\t",
    "Pipe |": "|",
}

# ──────────────────────── Credenciais SMTP ─────────────────────
load_dotenv()  # lê .env
SMTP_SERVER = os.getenv("SMTP_SERVER", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "0"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")
FROM_NAME = os.getenv("FROM_NAME", "")
FROM_ADDR = os.getenv("FROM_ADDR", SMTP_USER)

# ───────────────────── Configuração da página ──────────────────
st.set_page_config(page_title="Gerar & Enviar Certificados", layout="centered")
st.title("Gerador e Enviador de Certificados")

# ─────────────────── Uploads CSV e PDF ───────────────────
csv_file = st.file_uploader(
    f"CSV com colunas '{CSV_COL_NAME}' e '{CSV_COL_MAIL}'", type="csv"
)
pdf_template_file = st.file_uploader(
    f"PDF modelo com campo '{PDF_FIELD}'", type="pdf"
)

# ────────────── Escolha (ou auto) do separador ──────────────
st.subheader("Configuração do CSV")
sep_option = st.selectbox("Separador de colunas", SEP_LABELS, index=0)
sep_user = SEP_MAP[sep_option]

# ─────────────── Leitura do CSV com tratamento ───────────────
df = None
if csv_file:
    content_bytes = csv_file.getvalue()
    try:
        # detectar separador se necessário
        if sep_user is None:
            sample = content_bytes[:10000].decode(errors="ignore")
            try:
                sniffed = csv.Sniffer().sniff(sample, delimiters=";,|\t")
                sep_detected = sniffed.delimiter
            except csv.Error:
                sep_detected = ","
        else:
            sep_detected = sep_user

        df = pd.read_csv(io.BytesIO(content_bytes), sep=sep_detected)

        # verificar colunas obrigatórias
        missing_cols = [
            c for c in [CSV_COL_NAME, CSV_COL_MAIL] if c not in df.columns
        ]
        if missing_cols:
            st.error(
                f"Coluna(s) não encontrada(s): {', '.join(missing_cols)}.\n\n"
                f"Colunas disponíveis: {', '.join(df.columns)}"
            )
            df = None
        else:
            df = df.dropna(subset=[CSV_COL_NAME, CSV_COL_MAIL])
            st.subheader("Pré-visualização do CSV")
            st.dataframe(df.head())
    except Exception as e:
        st.error(f"Falha ao ler o CSV: {e}")
        df = None

# ───────────────── Modelo de e-mail ─────────────────
st.subheader("Modelo de e-mail")
subject_input = st.text_input(
    "Assunto (use {{Nome}} para personalizar)",
    value="Seu certificado de conclusão – {{Nome}}",
)
body_input = st.text_area(
    "Corpo em HTML (pode usar {{Nome}})", height=350, placeholder="Cole aqui seu HTML…"
)

# Pré-visualização em modal (se disponível) ou expander (fallback)
if st.button("Pré-visualizar e-mail"):
    if hasattr(st, "modal"):               # Streamlit >= 1.29
        container = st.modal("Pré-visualização do e-mail")
    else:                                  # versões mais antigas
        container = st.expander("Pré-visualização do e-mail", expanded=True)

    with container:
        st.markdown("#### Assunto")
        st.write(subject_input.replace("{{Nome}}", "Nome Exemplo"))

        st.markdown("#### Corpo")
        components.html(
            body_input.replace("{{Nome}}", "Nome Exemplo"),
            height=500,
            scrolling=True,
        )


# ──────────────── Validação das credenciais SMTP ───────────────
def smtp_ready() -> bool:
    required = [SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASS, FROM_ADDR]
    if not all(required):
        st.error(
            "Credenciais SMTP ausentes ou incompletas no `.env` "
            "(SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASS, FROM_ADDR)."
        )
        return False
    return True


# ────────────────────── Botão principal ──────────────────────
if (
    df is not None
    and pdf_template_file is not None
    and smtp_ready()
    and subject_input.strip()
    and body_input.strip()
):
    if st.button("Gerar certificados e enviar e-mails"):
        template_bytes = pdf_template_file.read()

        # verificar campo existente
        reader_check = PdfReader(io.BytesIO(template_bytes))
        if PDF_FIELD not in (reader_check.get_fields() or {}):
            st.error(f"O template não contém o campo '{PDF_FIELD}'.")
            st.stop()

        zip_buffer = io.BytesIO()
        ok, fail = 0, 0
        progress = st.progress(0.0, text="Iniciando…")

        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls(context=context)
            try:
                smtp.login(SMTP_USER, SMTP_PASS)
            except smtplib.SMTPAuthenticationError as e:
                st.error(f"Falha de autenticação SMTP: {e}")
                st.stop()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                total = len(df)
                for idx, row in df.iterrows():
                    nome = str(row[CSV_COL_NAME]).strip()
                    email_to = str(row[CSV_COL_MAIL]).strip()

                    # valida e-mail
                    try:
                        email_to = validate_email(email_to).email
                    except EmailNotValidError:
                        fail += 1
                        progress.progress(
                            (idx + 1) / total, text=f"E-mail inválido: {email_to}"
                        )
                        continue

                    # 1) gerar PDF
                    reader = PdfReader(io.BytesIO(template_bytes))
                    writer = PdfWriter()
                    writer.clone_document_from_reader(reader)
                    writer.update_page_form_field_values(
                        writer.pages[0], {PDF_FIELD: nome}
                    )
                    # marcar campo como read-only
                    for page in writer.pages:
                        if "/Annots" in page:
                            for annot in page["/Annots"]:
                                obj = annot.get_object()
                                if obj.get("/T") == PDF_FIELD:
                                    flags = obj.get("/Ff", 0)
                                    obj.update(
                                        {NameObject("/Ff"): NumberObject(flags | 1)}
                                    )

                    pdf_buf = io.BytesIO()
                    writer.write(pdf_buf)
                    pdf_buf.seek(0)
                    fname = f"certificado_{'_'.join(nome.lower().split())}.pdf"
                    zipf.writestr(fname, pdf_buf.read())

                    # 2) montar e-mail
                    msg = EmailMessage()
                    msg["Subject"] = subject_input.replace("{{Nome}}", nome)
                    msg["From"] = f"{FROM_NAME} <{FROM_ADDR}>"
                    msg["To"] = email_to
                    msg.set_content("Seu leitor não suporta HTML.")
                    msg.add_alternative(
                        body_input.replace("{{Nome}}", nome), subtype="html"
                    )
                    msg.add_attachment(
                        pdf_buf.getvalue(),
                        maintype="application",
                        subtype="pdf",
                        filename=fname,
                    )

                    # 3) enviar
                    try:
                        smtp.send_message(msg)
                        ok += 1
                        status = f"Enviado para {email_to}"
                    except Exception as e:
                        fail += 1
                        status = f"Falha ({email_to}): {e}"

                    progress.progress((idx + 1) / total, text=status)

        zip_buffer.seek(0)
        st.success(f"Concluído ✔️ {ok} enviados · ❌ {fail} falharam")
        st.download_button(
            "Baixar ZIP com todos os certificados",
            data=zip_buffer.getvalue(),
            file_name="certificados.zip",
            mime="application/zip",
        )
else:
    st.info(
        "Carregue CSV e PDF, defina o assunto/corpo, e garanta que as variáveis SMTP "
        "estejam no `.env` para habilitar o envio."
    )
