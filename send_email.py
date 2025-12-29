from email.mime.image import MIMEImage
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


EMAIL_REMETENTE = "leoprokopowiskii@gmail.com"
SENHA = "iquv lbcp fhry sntw"  # ideal depois jogar em variável de ambiente
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PORTA = 587

EMAIL_CC_PADRAO = "leonardo.lemes@sprogroup.com.br"

def enviar_email(destinatario, assunto, corpo_texto, corpo_html=None):

    msg = MIMEMultipart("related")
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = destinatario
    msg["Cc"] = EMAIL_CC_PADRAO
    msg["Subject"] = assunto

    # html (se existir)
    if corpo_html:
        msg.attach(MIMEText(corpo_html, "html", "utf-8"))

    with open("assets/assinatura.png", "rb") as f:
        img = MIMEImage(f.read())
        img.add_header("Content-ID", "<assinatura>")
        img.add_header("Content-Disposition", "inline", filename="assinatura.png")
        msg.attach(img)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA) as servidor:
        servidor.starttls()
        servidor.login(EMAIL_REMETENTE, SENHA)
        servidor.send_message(msg)

    print(f"📧 E-mail enviado para {destinatario}")