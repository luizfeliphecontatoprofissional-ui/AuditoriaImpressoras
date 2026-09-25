import pandas as pd


def main():
    caminho = "dados/faturamento.xlsx"

    dados = pd.read_excel(
        caminho,
        sheet_name="Inventário & Volume",
        header=None,
        nrows=10
    )

    print("=== CABEÇALHO SUPERIOR ===")

    for coluna in range(len(dados.columns)):
        valor_linha_8 = dados.iloc[8, coluna]
        valor_linha_9 = dados.iloc[9, coluna]

        print(
            f"Coluna {coluna}: "
            f"grupo = {valor_linha_8!r} | "
            f"campo = {valor_linha_9!r}"
        )


if __name__ == "__main__":
    main()