import pandas as pd

def converter_excel_para_json(caminho_excel):
    df = pd.read_excel(caminho_excel)

    # remove linhas sem Assigned (se quiser)
    df = df.dropna(subset=["Assigned"])

    # garante que tudo vire string onde precisa
    df["Assigned"] = df["Assigned"].astype(str)

    # converte pra lista de dict (JSON)
    tasks = df.to_dict(orient="records")

    return tasks