from excel_to_json import converter_excel_para_json
from group_tasks import agrupar_por_pessoa
from send_email import enviar_email

CAMINHO_EXCEL = "testes-automation.xlsx"

EMAIL_POR_PESSOA = {
    "João": "joao@email.com",
    "Samuel Nascimento": "samuel.nascimento@sprogroup.com.br",
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
                <td>
                    <a href="https://dev.azure.com/sprogroup/SPRO_Fomento/_workitems/edit/{t['ID']}">
                        {t['ID']}
                    </a>
                </td>
                <td>{t['Title']}</td>
                <td>{t['State']}</td>
            </tr>
    """

    return f"""
        <html>
            <body style="
                font-family: Arial, Helvetica, sans-serif;
                font-size: 14px;
                color: #000;
            ">

                <p>Olá, <b>{pessoa}</b>,</p>

                <p>
                    Identificamos que você possui itens com irretgularidades (campos relacionados a horas e datas).
                    Contamos com seu apoio para revisar e ajustar essas atividades,
                    garantindo a padronização e rastreabilidade das informações.
                </p>

                <p style="font-style: italic;">
                    Obs: verifique se os campos estão corretamente preenchidos.
                </p>

                <p><b>Segue a relação de itens:</b></p>

                <table width="60%" cellpadding="6" cellspacing="0"
                    style="border-collapse: collapse; font-size: 13px;">
                    <tr style="background-color:#5B9BD5; color:#fff;">
                        <th align="left">ID</th>
                        <th align="left">Title</th>
                        <th align="left">State</th>
                    </tr>

                    {linhas}
                </table>

                <p style="margin-top:15px;">
                    Ao clicar no <b>ID</b>, você será direcionado à atividade correspondente.
                </p>

                <p>
                    Lembrando que a <b>atualização dos campos</b>, status real da atividade
                    e comentários de andamento são essenciais para garantir a visão
                    consolidada da Sprint.
                </p>

                <p><b>ATENÇÃO:</b> Ao finalizar uma atividade, incluir anexo/link das evidências de testes.</p>

                <p>
                    Caso as informações já estejam atualizadas, favor desconsiderar este e-mail.
                </p>

                <br>

                <p>Atenciosamente,</p>

                <!-- ASSINATURA -->
                <img src="assets/assinatura.png"
                    alt="SPRO Group"
                    style="width:420px;">

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