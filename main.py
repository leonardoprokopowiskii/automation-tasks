from excel_to_json import converter_excel_para_json
from group_tasks import agrupar_por_pessoa
from send_email import enviar_email

CAMINHO_EXCEL = "testes-automation.xlsx"

EMAIL_POR_PESSOA = {
    "João": "joao@email.com",
    "Maria": "maria@email.com",
    "Samuel Nascimento": "samuel.nascimento@sprogroup.com.br",
    "Daniel Pereira": ""
}


def montar_email_texto(pessoa, tasks):
    corpo = f"Olá {pessoa},\n\nSegue a lista das suas tasks:\n\n"

    for t in tasks:
        corpo += f"- #{t['ID']} | {t['Title']} | {t['State']}\n"

    corpo += "\nPor favor, verifique e ajuste conforme necessário.\n\nAtenciosamente."
    return corpo


def montar_email_html(pessoa, tasks):
    linhas = ""
    for t in tasks:
        linhas += f"""
            <tr>
                <td>{t['ID']}</td>
                <td>{t['Title']}</td>
                <td>{t['State']}</td>
            </tr>
        """

    return f"""
    <html>
      <body style="font-family: Arial, sans-serif; font-size: 14px;">
        <p>Olá <b>{pessoa}</b>,</p>

        <p>Segue a lista das suas tasks:</p>

        <table border="1" cellpadding="6" cellspacing="0"
               style="border-collapse: collapse;">
          <tr style="background-color:#f0f0f0;">
            <th>ID</th>
            <th>Título</th>
            <th>Status</th>
          </tr>
          {linhas}
        </table>

        <p style="margin-top:20px;">
          Por favor, verifique e ajuste se necessário.
        </p>

        <p>Atenciosamente,</p>

        <!-- assinatura depois -->
        <p><b>Automação</b></p>
      </body>
    </html>
    """


def main():
    tasks = converter_excel_para_json(CAMINHO_EXCEL)
    tasks_por_pessoa = agrupar_por_pessoa(tasks)

    for pessoa, lista_tasks in tasks_por_pessoa.items():

        if pessoa not in EMAIL_POR_PESSOA:
            print(f"⚠️ Sem email cadastrado para {pessoa}, pulando...")
            continue

        texto = montar_email_texto(pessoa, lista_tasks)
        html = montar_email_html(pessoa, lista_tasks)

        enviar_email(
            destinatario=EMAIL_POR_PESSOA[pessoa],
            assunto="Resumo das suas tasks",
            corpo_texto=texto,
            corpo_html=html
        )


if __name__ == "__main__":
    main()