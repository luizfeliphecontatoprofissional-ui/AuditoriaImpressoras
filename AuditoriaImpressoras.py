import pandas as pd

def carregar_excel(caminho):
    return pd.read_excel(caminho)

def main():
    ndd = carregar_excel("dados/NDD.xlsx")
    faturamento = carregar_excel("dados/faturamento.xlsx")

    print("NDD:")
    print(ndd.columns.tolist())

    print("\nFaturamento:")
    print(faturamento.columns.tolist())

if __name__ == "__main__":
    main()