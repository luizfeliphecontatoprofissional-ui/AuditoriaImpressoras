import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


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

    return dados, data_inicio, data_fim


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

    for item, dados in sorted(total_por_item.items()):

        print(
            f"Item {item}: "
            f"{dados['completos']} completos | "
            f"{dados['faltando']} incompletos | "
            f"Total: {dados['total']}"
        )

    print("\n=== RESUMO ===")
    print(
        "Impressoras com CounterTypes esperados:",
        completos
    )
    print(
        "Impressoras sem o serial no NDD:",
        faltando_serial
    )
    print(
        "Impressoras com CounterType faltando:",
        faltando_counter_type
    )


def criar_ndd_map(ndd):
    ndd_map = {}

    for _, linha in ndd.iterrows():

        serial = str(
            linha["SerialNumber"]
        ).strip().upper()

        tipo_contador = str(
            linha["CounterTypeDescription"]
        ).strip()

        if serial == "NAN" or not serial:
            continue

        if tipo_contador == "NAN" or not tipo_contador:
            continue

        chave = (
            serial,
            tipo_contador
        )

        ndd_map[chave] = linha

    return ndd_map


def para_numero(valor):
    numero = pd.to_numeric(
        valor,
        errors="coerce"
    )

    if pd.isna(numero):
        return None

    return float(numero)


def comparar_contador(
    divergencias,
    serial,
    item,
    nome_campo,
    valor_ndd,
    valor_faturamento
):
    ndd_num = para_numero(valor_ndd)
    fat_num = para_numero(valor_faturamento)

    if ndd_num is None and fat_num is None:
        return

    if ndd_num is None or fat_num is None:

        divergencias.append([
            serial,
            item,
            nome_campo,
            valor_ndd,
            valor_faturamento
        ])

        return

    if ndd_num != fat_num:

        divergencias.append([
            serial,
            item,
            nome_campo,
            valor_ndd,
            valor_faturamento
        ])


def comparar_tipo_1_2(ndd_map, inventario):
    divergencias = []

    for _, linha in inventario.iterrows():

        serial = str(
            linha["Numero de Série"]
        ).strip().upper()

        item = str(
            linha["Item"]
        ).strip().lstrip("0")

        if serial == "NAN" or not serial:
            continue

        if item not in ("1", "2"):
            continue

        registro_ndd = ndd_map.get(
            (serial, "A4")
        )

        if registro_ndd is None:
            continue

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start A4 vs Start Mono",
            registro_ndd["StartCounterTotal"],
            linha["Start_Mono"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End A4 vs End Mono",
            registro_ndd["EndCounterTotal"],
            linha["End_Mono"]
        )

    return divergencias


def comparar_tipo_3(ndd_map, inventario):
    divergencias = []

    for _, linha in inventario.iterrows():

        serial = str(
            linha["Numero de Série"]
        ).strip().upper()

        item = str(
            linha["Item"]
        ).strip().lstrip("0")

        if serial == "NAN" or not serial:
            continue

        if item != "3":
            continue

        registro_a3 = ndd_map.get(
            (serial, "A3")
        )

        registro_a4 = ndd_map.get(
            (serial, "A4")
        )

        if (
            registro_a3 is None
            or registro_a4 is None
        ):
            continue

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start A3 vs Start Mono A3",
            registro_a3["StartCounterTotal"],
            linha["Start_Mono A3"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start A4 vs Start Mono",
            registro_a4["StartCounterTotal"],
            linha["Start_Mono"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End A3 vs End Mono A3",
            registro_a3["EndCounterTotal"],
            linha["End_Mono A3"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End A4 vs End Mono",
            registro_a4["EndCounterTotal"],
            linha["End_Mono"]
        )

    return divergencias


def comparar_tipo_4(ndd_map, inventario):
    divergencias = []

    for _, linha in inventario.iterrows():

        serial = str(
            linha["Numero de Série"]
        ).strip().upper()

        item = str(
            linha["Item"]
        ).strip().lstrip("0")

        if serial == "NAN" or not serial:
            continue

        if item != "4":
            continue

        registro_print = ndd_map.get(
            (serial, "Print")
        )

        if registro_print is None:
            continue

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start Mono vs Start Mono",
            registro_print["StartCounterMono"],
            linha["Start_Mono"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start Color vs Start Color",
            registro_print["StartCounterColor"],
            linha["Start_Color"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End Mono vs End Mono",
            registro_print["EndCounterMono"],
            linha["End_Mono"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End Color vs End Color",
            registro_print["EndCounterColor"],
            linha["End_Color"]
        )

    return divergencias


def comparar_tipo_5(ndd_map, inventario):
    divergencias = []

    for _, linha in inventario.iterrows():

        serial = str(
            linha["Numero de Série"]
        ).strip().upper()

        item = str(
            linha["Item"]
        ).strip().lstrip("0")

        if serial == "NAN" or not serial:
            continue

        if item != "5":
            continue

        registro_a4 = ndd_map.get(
            (serial, "A4")
        )

        registro_a3 = ndd_map.get(
            (serial, "A3")
        )

        if (
            registro_a4 is None
            or registro_a3 is None
        ):
            continue

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start A4 Mono vs Start Mono",
            registro_a4["StartCounterMono"],
            linha["Start_Mono"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start A4 Color vs Start Color",
            registro_a4["StartCounterColor"],
            linha["Start_Color"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start A3 Mono vs Start Mono A3",
            registro_a3["StartCounterMono"],
            linha["Start_Mono A3"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "Start A3 Color vs Start Color A3",
            registro_a3["StartCounterColor"],
            linha["Start_Color A3"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End A4 Mono vs End Mono",
            registro_a4["EndCounterMono"],
            linha["End_Mono"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End A4 Color vs End Color",
            registro_a4["EndCounterColor"],
            linha["End_Color"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End A3 Mono vs End Mono A3",
            registro_a3["EndCounterMono"],
            linha["End_Mono A3"]
        )

        comparar_contador(
            divergencias,
            serial,
            item,
            "End A3 Color vs End Color A3",
            registro_a3["EndCounterColor"],
            linha["End_Color A3"]
        )

    return divergencias


def gerar_relatorio(
    divergencias,
    caminho_saida,
    data_inicio,
    data_fim
):
    colunas = [
        "SerialNumber",
        "Tipo/Item",
        "Campo Divergente",
        "Valor NDD",
        "Valor Tecprinters"
    ]

    df_relatorio = pd.DataFrame(
        divergencias,
        columns=colunas
    )

    os.makedirs(
        os.path.dirname(caminho_saida),
        exist_ok=True
    )

    with pd.ExcelWriter(
        caminho_saida,
        engine="openpyxl"
    ) as writer:

        df_relatorio.to_excel(
            writer,
            sheet_name="Divergências",
            index=False,
            startrow=4
        )

        planilha = writer.book["Divergências"]

        planilha.merge_cells("A1:E1")

        planilha["A1"] = (
            "Relatório de Divergências da Auditoria"
        )

        planilha["A1"].font = Font(
            bold=True,
            size=16
        )

        planilha["A1"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        planilha.merge_cells("A2:E2")

        planilha["A2"] = (
            f"Período: "
            f"{data_inicio.strftime('%d/%m/%Y')} "
            f"a "
            f"{data_fim.strftime('%d/%m/%Y')}"
        )

        planilha.merge_cells("A3:E3")

        planilha["A3"] = (
            f"Total de divergências: "
            f"{len(df_relatorio)}"
        )

        planilha["A2"].alignment = Alignment(
            horizontal="left"
        )

        planilha["A3"].alignment = Alignment(
            horizontal="left"
        )

        for celula in planilha[5]:

            celula.font = Font(
                bold=True
            )

            celula.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            celula.fill = PatternFill(
                fill_type="solid",
                fgColor="D9EAF7"
            )

        borda = Border(
            left=Side(
                style="thin",
                color="BFBFBF"
            ),
            right=Side(
                style="thin",
                color="BFBFBF"
            ),
            top=Side(
                style="thin",
                color="BFBFBF"
            ),
            bottom=Side(
                style="thin",
                color="BFBFBF"
            )
        )

        for linha in planilha.iter_rows(
            min_row=5,
            max_row=planilha.max_row,
            min_col=1,
            max_col=5
        ):

            for celula in linha:
                celula.border = borda

        larguras = {
            "A": 20,
            "B": 12,
            "C": 35,
            "D": 18,
            "E": 22
        }

        for coluna, largura in larguras.items():
            planilha.column_dimensions[coluna].width = largura

        for celula in planilha["C"]:

            celula.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

        for linha in range(
            6,
            planilha.max_row + 1
        ):

            planilha[f"D{linha}"].number_format = "0"
            planilha[f"E{linha}"].number_format = "0"

        planilha.auto_filter.ref = (
            f"A5:E{planilha.max_row}"
        )

        planilha.freeze_panes = "A6"

        planilha.row_dimensions[1].height = 25
        planilha.row_dimensions[5].height = 22

    print("\n=== RELATÓRIO ===")
    print("Relatório gerado com sucesso!")
    print("Local:", caminho_saida)
    print(
        "Quantidade de divergências:",
        len(df_relatorio)
    )


def executar_auditoria(
    caminho_ndd="dados/ndd.xlsx",
    caminho_faturamento="dados/faturamento.xlsx"
):
    try:

        ndd = pd.read_excel(
            caminho_ndd
        )

        inventario, data_inicio, data_fim = (
            carregar_inventario(
                caminho_faturamento
            )
        )

        ndd_map = criar_ndd_map(ndd)

        divergencias = []

        divergencias_tipo_1_2 = (
            comparar_tipo_1_2(
                ndd_map,
                inventario
            )
        )

        print("\n=== TIPO 1 E 2 ===")
        print(
            "Divergências encontradas:",
            len(divergencias_tipo_1_2)
        )

        divergencias.extend(
            divergencias_tipo_1_2
        )

        divergencias_tipo_3 = (
            comparar_tipo_3(
                ndd_map,
                inventario
            )
        )

        print("\n=== TIPO 3 ===")
        print(
            "Divergências encontradas:",
            len(divergencias_tipo_3)
        )

        divergencias.extend(
            divergencias_tipo_3
        )

        divergencias_tipo_4 = (
            comparar_tipo_4(
                ndd_map,
                inventario
            )
        )

        print("\n=== TIPO 4 ===")
        print(
            "Divergências encontradas:",
            len(divergencias_tipo_4)
        )

        divergencias.extend(
            divergencias_tipo_4
        )

        divergencias_tipo_5 = (
            comparar_tipo_5(
                ndd_map,
                inventario
            )
        )

        print("\n=== TIPO 5 ===")
        print(
            "Divergências encontradas:",
            len(divergencias_tipo_5)
        )

        divergencias.extend(
            divergencias_tipo_5
        )

        print("\n=== AUDITORIA ===")
        print(
            "Total de divergências:",
            len(divergencias)
        )

        caminho_relatorio = (
            "relatorios/divergencias.xlsx"
        )

        gerar_relatorio(
            divergencias,
            caminho_relatorio,
            data_inicio,
            data_fim
        )

        return True

    except FileNotFoundError as erro:

        print("\n=== ERRO ===")
        print("Arquivo não encontrado.")
        print(
            "Verifique se os arquivos estão "
            "na pasta 'dados'."
        )
        print("Detalhes:", erro)

        return False

    except PermissionError:

        print("\n=== ERRO ===")
        print(
            "Não foi possível acessar "
            "ou salvar um arquivo."
        )
        print(
            "Verifique se algum arquivo está "
            "aberto no Excel e tente novamente."
        )

        return False

    except KeyError as erro:

        print("\n=== ERRO ===")
        print(
            "Uma coluna esperada não foi "
            "encontrada na planilha."
        )
        print("Coluna:", erro)

        return False

    except ValueError as erro:

        print("\n=== ERRO ===")
        print(
            "Os dados da planilha não estão "
            "no formato esperado."
        )
        print("Detalhes:", erro)

        return False

    except Exception as erro:

        print("\n=== ERRO INESPERADO ===")
        print(
            "O programa encontrou um erro "
            "que não foi previsto."
        )
        print("Detalhes:", erro)

        return False


def selecionar_arquivo(entry):
    caminho = filedialog.askopenfilename(
        title="Selecionar arquivo",
        filetypes=[
            ("Arquivos Excel", "*.xlsx"),
            ("Todos os arquivos", "*.*")
        ]
    )

    if caminho:
        entry.delete(0, tk.END)
        entry.insert(0, caminho)


def abrir_relatorio():
    caminho_relatorio = os.path.abspath(
        "relatorios/divergencias.xlsx"
    )

    if not os.path.exists(caminho_relatorio):
        messagebox.showwarning(
            "Relatório não encontrado",
            "O relatório ainda não foi gerado."
        )
        return

    try:
        os.startfile(caminho_relatorio)
    except Exception as erro:
        messagebox.showerror(
            "Erro ao abrir relatório",
            f"Não foi possível abrir o relatório.\n\n{erro}"
        )


def executar_pela_interface():
    caminho_ndd = entrada_ndd.get().strip()
    caminho_faturamento = entrada_faturamento.get().strip()

    if not caminho_ndd or not caminho_faturamento:
        messagebox.showwarning(
            "Arquivos não selecionados",
            "Selecione os dois arquivos antes de executar a auditoria."
        )
        return

    if not os.path.isfile(caminho_ndd):
        messagebox.showwarning(
            "Arquivo NDD inválido",
            "O arquivo NDD selecionado não foi encontrado."
        )
        return

    if not os.path.isfile(caminho_faturamento):
        messagebox.showwarning(
            "Arquivo de faturamento inválido",
            "O arquivo de faturamento selecionado não foi encontrado."
        )
        return

    botao_executar.config(state="disabled")
    botao_abrir.config(state="disabled")
    status_var.set("Executando auditoria...")
    root.update_idletasks()

    sucesso = executar_auditoria(
        caminho_ndd,
        caminho_faturamento
    )

    if sucesso:
        status_var.set("✓ Auditoria concluída com sucesso.")
        botao_abrir.config(state="normal")

        messagebox.showinfo(
            "Auditoria concluída",
            "A auditoria foi executada com sucesso.\n\n"
            "O relatório foi salvo em:\n"
            "relatorios/divergencias.xlsx"
        )
    else:
        status_var.set("A auditoria não foi concluída.")
        messagebox.showerror(
            "Erro na auditoria",
            "A auditoria não foi concluída.\n\n"
            "Verifique os dados e os arquivos selecionados."
        )

    botao_executar.config(state="normal")


def main():
    global root
    global entrada_ndd
    global entrada_faturamento
    global botao_executar
    global botao_abrir
    global status_var

    root = tk.Tk()
    root.title("Auditoria de Impressoras")
    root.geometry("620x420")
    root.resizable(False, False)

    estilo = ttk.Style()
    try:
        estilo.theme_use("vista")
    except tk.TclError:
        pass

    frame_principal = ttk.Frame(
        root,
        padding=30
    )
    frame_principal.pack(fill="both", expand=True)

    titulo = ttk.Label(
        frame_principal,
        text="AUDITORIA DE IMPRESSORAS",
        font=("Segoe UI", 18, "bold")
    )
    titulo.pack(pady=(0, 5))

    subtitulo = ttk.Label(
        frame_principal,
        text="Comparação entre dados NDD e Faturamento"
    )
    subtitulo.pack(pady=(0, 25))

    frame_ndd = ttk.Frame(frame_principal)
    frame_ndd.pack(fill="x", pady=5)

    ttk.Label(
        frame_ndd,
        text="Arquivo NDD"
    ).pack(anchor="w")

    linha_ndd = ttk.Frame(frame_ndd)
    linha_ndd.pack(fill="x", pady=(5, 0))

    entrada_ndd = ttk.Entry(linha_ndd)
    entrada_ndd.pack(side="left", fill="x", expand=True)

    ttk.Button(
        linha_ndd,
        text="Procurar",
        command=lambda: selecionar_arquivo(entrada_ndd)
    ).pack(side="left", padx=(8, 0))

    frame_faturamento = ttk.Frame(frame_principal)
    frame_faturamento.pack(fill="x", pady=15)

    ttk.Label(
        frame_faturamento,
        text="Arquivo de Faturamento"
    ).pack(anchor="w")

    linha_faturamento = ttk.Frame(frame_faturamento)
    linha_faturamento.pack(fill="x", pady=(5, 0))

    entrada_faturamento = ttk.Entry(linha_faturamento)
    entrada_faturamento.pack(side="left", fill="x", expand=True)

    ttk.Button(
        linha_faturamento,
        text="Procurar",
        command=lambda: selecionar_arquivo(entrada_faturamento)
    ).pack(side="left", padx=(8, 0))

    botao_executar = ttk.Button(
        frame_principal,
        text="EXECUTAR AUDITORIA",
        command=executar_pela_interface
    )
    botao_executar.pack(pady=(20, 10), ipadx=20, ipady=5)

    botao_abrir = ttk.Button(
        frame_principal,
        text="ABRIR RELATÓRIO",
        command=abrir_relatorio,
        state="disabled"
    )
    botao_abrir.pack(pady=5, ipadx=20, ipady=3)

    status_var = tk.StringVar(
        value="Status: Aguardando arquivos..."
    )

    status = ttk.Label(
        frame_principal,
        textvariable=status_var,
        wraplength=540,
        anchor="center"
    )
    status.pack(pady=(20, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
