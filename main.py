from excel_to_json import converter_excel_para_json
from group_tasks import agrupar_por_pessoa
from send_email import enviar_email
from emails_list import EMAIL_POR_PESSOA    

CAMINHO_EXCEL = "testes-automation.xlsx"

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
<body style="margin:0; padding:0; width:100vw;">

<table width="100vw" cellpadding="0" cellspacing="0">
    <tr>
        <td align="center">

            <!-- CONTAINER LARGO / EMAIL NORMAL -->
            <table width="100vw" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="
                        font-family: Arial, Helvetica, sans-serif;
                        font-size:16px;
                        color:#000;
                    ">

                        <p>Olá, <b>{pessoa}</b>,</p>

                        <p>
                            Identificamos que você possui itens com irregularidades
                            (campos relacionados a horas e datas).
                            Contamos com seu apoio para revisar e ajustar essas atividades,
                            garantindo a padronização e rastreabilidade das informações.
                        </p>

                        <p style="font-style: italic; margin-bottom:56px">
                            Obs: verifique se os campos estão corretamente preenchidos.
                        </p>

                        <p><b>Segue a relação de itens:</b></p>

                        <!-- TABELA DE TASKS -->
                        <table width="60%" cellpadding="6" cellspacing="0"
                               style="border-collapse: collapse; font-size:15px;">
                            <tr style="background-color:#5B9BD5; color:#fff;">
                                <th align="left">ID</th>
                                <th align="left">Title</th>
                                <th align="left">State</th>
                            </tr>

                            {linhas}
                        </table>

                        <p style="margin-top: 56px">
                            Ao clicar no <b>ID</b>, você será direcionado à atividade correspondente.
                        </p>

                        <p>
                            Lembrando que a <b>atualização dos campos</b>, status real da atividade
                            e comentários de andamento são essenciais para garantir a visão
                            consolidada da Sprint.
                        </p>

                        <p>
                            Caso as informações já estejam atualizadas, favor desconsiderar este e-mail.
                        </p>

                        <br>

                        <p>Atenciosamente,</p>

                        <!-- ASSINATURA (50% VISUAL REAL) -->
                        <table width="100%" cellpadding="0" cellspacing="0">
                            <tr>
                                <td align="left">
                                    <img src="cid:assinatura"
                                         alt="SPRO Group"
                                         width="50%"
                                         style="display:block;">
                                </td>
                            </tr>
                        </table>

                    </td>
                </tr>
            </table>
            <!-- /CONTAINER -->

        </td>
    </tr>
</table>
<br>

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