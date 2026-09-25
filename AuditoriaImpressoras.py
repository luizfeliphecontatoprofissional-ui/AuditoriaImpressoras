import pandas as pd


def carregar_inventario(caminho):
    dados = pd.read_excel(
        caminho,
        sheet_name="Inventário & Volume",
        header=None
    )

    grupos = dados.iloc[8].ffill()
    campos = dados.iloc[9]

    datas = []

    for valor in dados.iloc[8]:
        data = pd.to_datetime(valor, errors="coerce")

        if pd.notna(data) and data not in datas:
            datas.append(data)

    if len(datas) < 2:
        raise ValueError(
            "Não foi possível encontrar duas datas no cabeçalho da planilha."
        )

    data_inicio = datas[0]
    data_fim = datas[1]

    print("=== PERÍODO ENCONTRADO ===")
    print("Início:", data_inicio.strftime("%d/%m/%Y"))
    print("Fim:", data_fim.strftime("%d/%m/%Y"))

    novas_colunas = []

    for grupo, campo in zip(grupos, campos):

        if pd.isna(campo):
            novas_colunas.append(None)
            continue

        grupo_data = pd.to_datetime(grupo, errors="coerce")

        if pd.notna(grupo_data):

            if grupo_data == data_inicio:
                novas_colunas.append(f"Start_{campo}")

            elif grupo_data == data_fim:
                novas_colunas.append(f"End_{campo}")

            else:
                novas_colunas.append(str(campo))

        elif str(grupo).strip().lower() == "volume":
            novas_colunas.append(f"Volume_{campo}")

        elif str(grupo).strip().lower() == "valor":
            novas_colunas.append(f"Valor_{campo}")

        else:
            novas_colunas.append(str(campo))

    dados.columns = novas_colunas

    return dados

def main():
    caminho = "dados/faturamento.xlsx"

    inventario = carregar_inventario(caminho)

    colunas_contadores = [
        ("Start_Mono", "End_Mono", "Volume_Mono"),
        ("Start_Mono A3", "End_Mono A3", "Volume_Mono A3"),
        ("Start_Color", "End_Color", "Volume_Color"),
        ("Start_Color A3", "End_Color A3", "Volume_Color A3"),
    ]

    print("=== TESTE DOS VOLUMES ===")

    for inicio, fim, volume in colunas_contadores:
        inicio_num = pd.to_numeric(inventario[inicio], errors="coerce")
        fim_num = pd.to_numeric(inventario[fim], errors="coerce")
        volume_num = pd.to_numeric(inventario[volume], errors="coerce")

        valido = inicio_num.notna() & fim_num.notna() & volume_num.notna()

        esperado = fim_num - inicio_num
        resultado = esperado == volume_num

        testes = resultado[valido]

        print(f"\n{volume}")
        print("Registros comparáveis:", len(testes))
        print("Corretos:", testes.sum())
        print("Divergentes:", (~testes).sum())

if __name__ == "__main__":
    main()