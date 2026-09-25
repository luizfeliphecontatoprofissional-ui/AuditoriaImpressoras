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

        if pd.notna(data):
            datas.append(data)

    datas = sorted(set(datas))

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

def verificar_counter_types(ndd, inventario):
    regras = {
        "1": ["A4"],
        "2": ["A4"],
        "3": ["A3", "A4"],
        "4": ["Print"],
        "5": ["A3", "A4"],
    }

    ndd["SerialNumber"] = (
        ndd["SerialNumber"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    ndd["CounterTypeDescription"] = (
        ndd["CounterTypeDescription"]
        .astype(str)
        .str.strip()
    )

    inventario["Numero de Série"] = (
        inventario["Numero de Série"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    inventario["Item"] = (
        inventario["Item"]
        .astype(str)
        .str.strip()
        .str.replace(r"^0+", "", regex=True)
    )

    tipos_por_serial = {}

    for _, linha in ndd.iterrows():

        serial = linha["SerialNumber"]
        tipo_contador = linha["CounterTypeDescription"]

        if serial == "NAN" or not serial:
            continue

        if serial not in tipos_por_serial:
            tipos_por_serial[serial] = set()

        tipos_por_serial[serial].add(tipo_contador)

    total_por_item = {}
    completos = 0
    faltando_serial = 0
    faltando_counter_type = 0

    print("\n=== VERIFICAÇÃO DOS COUNTER TYPES ===")

    for _, linha in inventario.iterrows():

        serial = linha["Numero de Série"]
        item = linha["Item"]

        if serial == "NAN" or not serial:
            continue

        if item not in regras:
            continue

        esperados = set(regras[item])

        encontrados = tipos_por_serial.get(serial, set())

        faltantes = esperados - encontrados

        if item not in total_por_item:
            total_por_item[item] = {
                "total": 0,
                "completos": 0,
                "faltando": 0
            }

        total_por_item[item]["total"] += 1

        if serial not in tipos_por_serial:
            faltando_serial += 1
            total_por_item[item]["faltando"] += 1

        elif faltantes:
            faltando_counter_type += 1
            total_por_item[item]["faltando"] += 1

        else:
            completos += 1
            total_por_item[item]["completos"] += 1

    # Resumo
    for item, dados in sorted(total_por_item.items()):
        print(
            f"Item {item}: "
            f"{dados['completos']} completos | "
            f"{dados['faltando']} incompletos | "
            f"Total: {dados['total']}"
        )

    print("\n=== RESUMO ===")
    print("Impressoras com CounterTypes esperados:", completos)
    print("Impressoras sem o serial no NDD:", faltando_serial)
    print("Impressoras com CounterType faltando:", faltando_counter_type)

def main():
    caminho_ndd = "dados/ndd.xlsx"
    caminho_faturamento = "dados/faturamento.xlsx"

    ndd = pd.read_excel(caminho_ndd)
    inventario = carregar_inventario(caminho_faturamento)

    verificar_counter_types(ndd, inventario)

    ndd["SerialNumber"] = (
        ndd["SerialNumber"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    inventario["Numero de Série"] = (
        inventario["Numero de Série"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    serials_ndd = {
        serial
        for serial in ndd["SerialNumber"]
        if serial and serial != "NAN"
    }

    serials_inventario = {
        serial
        for serial in inventario["Numero de Série"]
        if serial and serial != "NAN"
    }

    encontrados = serials_inventario & serials_ndd
    nao_encontrados = serials_inventario - serials_ndd

    print("=== CRUZAMENTO DE NÚMEROS DE SÉRIE ===")
    print("Seriais no Inventário:", len(serials_inventario))
    print("Seriais no NDD:", len(serials_ndd))
    print("Encontrados nos dois:", len(encontrados))
    print("Não encontrados no NDD:", len(nao_encontrados))

    print("\n=== TIPOS DE CONTADOR NO NDD ===")

    tipos_contador = (
        ndd["CounterTypeDescription"]
        .astype(str)
        .str.strip()
        .value_counts()
    )

    for tipo, quantidade in tipos_contador.items():
        print(f"{tipo}: {quantidade}")

if __name__ == "__main__":
    main()