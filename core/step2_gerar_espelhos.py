import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font


def gerar_espelhos_motoristas(
    banco_consolidado_xlsx_path: str,
    modelo_xlsx_path: str,
    saida_espelhos_xlsx_path: str,
) -> str:
    """
    Gera um único arquivo Excel com uma aba por motorista (Espelhos_Motoristas.xlsx).

    Retorna o path do arquivo gerado.
    """
    if not os.path.exists(banco_consolidado_xlsx_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {banco_consolidado_xlsx_path}")
    if not os.path.exists(modelo_xlsx_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {modelo_xlsx_path}")

    df = pd.read_excel(banco_consolidado_xlsx_path)
    df.columns = df.columns.str.strip().str.lower()

    def achar_coluna(possiveis):
        for c in possiveis:
            if c in df.columns:
                return c
        return None

    col_motorista = achar_coluna(["nome do motorista", "motorista", "nome"])
    col_conta = achar_coluna(["conta", "conta corrente"])
    col_pix = achar_coluna(["pix", "chave pix"])
    col_data = achar_coluna(["data"])
    col_cidade = achar_coluna(["cidade"])
    col_status = achar_coluna(["status"])
    col_custo = achar_coluna(["custo", "valor", "valor unitario"])

    for obrigatoria in ["cpf", "banco", "agencia", "cliente", "romaneio"]:
        if obrigatoria not in df.columns:
            raise Exception(f"Coluna '{obrigatoria}' não encontrada no banco consolidado.")

    if not all([col_motorista, col_conta, col_pix, col_data, col_cidade, col_status, col_custo]):
        raise Exception("Coluna obrigatória não encontrada no banco consolidado (nome/cidade/data/status/custo etc).")

    # =========================
    # ESTILOS
    # =========================
    align_center = Alignment(horizontal="center", vertical="center")
    align_left_center = Alignment(horizontal="left", vertical="center")

    thin = Side(style="thin")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_lr = Border(left=thin, right=thin)

    fill_cliente = PatternFill(fill_type="solid", fgColor="D9D9D9")
    font_bold = Font(bold=True)
    font_bold_red = Font(bold=True, color="FF0000")

    formato_contabil = 'R$ #,##0.00_);R$ (#,##0.00)'

    def nome_aba_valido(nome):
        nome = re.sub(r'[\\/*?:\[\]]', '', str(nome))
        return nome[:31]

    def auto_largura_coluna_F(ws, start_row=2, end_row=7):
        max_len = 0
        for r in range(start_row, end_row + 1):
            v = ws[f"F{r}"].value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions["F"].width = max_len + 2

    # =========================
    # ABRIR MODELO
    # =========================
    wb = load_workbook(modelo_xlsx_path)
    aba_modelo = wb.active
    aba_modelo.title = "MODELO_BASE"

    # =========================
    # GERAR UMA ABA POR MOTORISTA
    # =========================
    for motorista in df[col_motorista].drop_duplicates():
        df_motorista = df[df[col_motorista] == motorista]
        linha_ref = df_motorista.iloc[0]

        ws = wb.copy_worksheet(aba_modelo)
        ws.title = nome_aba_valido(motorista)

        # PARTE 1 — DADOS FIXOS
        ws["C4"] = motorista
        ws["F2"] = f"Banco: {linha_ref['banco']}"
        ws["F3"] = f"Agência: {linha_ref['agencia']}"
        ws["F4"] = f"Conta: {linha_ref[col_conta]}"
        ws["F5"] = f"Favorecido: {motorista}"
        ws["F6"] = f"CPF/CNPJ do Favorecido: {linha_ref['cpf']}"
        ws["F7"] = f"PIX: {linha_ref[col_pix]}"

        for celula in ["F2", "F3", "F4", "F5", "F6", "F7"]:
            ws[celula].font = font_bold

        auto_largura_coluna_F(ws)

        # PARTE 2 — TABELA VARIÁVEL
        linha_atual = 12
        for cliente in df_motorista["cliente"].drop_duplicates():
            ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=6)
            cell = ws.cell(row=linha_atual, column=1)
            cell.value = cliente
            cell.alignment = align_center
            cell.fill = fill_cliente
            cell.font = font_bold

            for col in range(1, 7):
                ws.cell(row=linha_atual, column=col).border = border_all

            linha_atual += 1

            df_cliente = df_motorista[df_motorista["cliente"] == cliente]
            romaneios = df_cliente["romaneio"].dropna().drop_duplicates()

            for rom in romaneios:
                df_rom = df_cliente[df_cliente["romaneio"] == rom]

                ws.cell(row=linha_atual, column=1, value=rom).font = font_bold
                ws.cell(row=linha_atual, column=2, value=df_rom.shape[0]).font = font_bold
                ws.cell(row=linha_atual, column=4, value=df_rom.iloc[0][col_cidade]).font = font_bold
                ws.cell(row=linha_atual, column=5, value=df_rom.iloc[0][col_data]).font = font_bold
                ws.cell(row=linha_atual, column=6, value=df_rom.iloc[0][col_status]).font = font_bold

                for col in range(1, 7):
                    c = ws.cell(row=linha_atual, column=col)
                    c.alignment = align_center
                    c.border = border_all

                linha_atual += 1

        # PARTE 3 — MAPEAMENTO POR CIDADE
        linha_atual += 1

        ws.merge_cells(start_row=linha_atual, start_column=2, end_row=linha_atual, end_column=3)
        ws[f"B{linha_atual}"] = "CIDADE"
        ws[f"D{linha_atual}"] = "QUANTIDADE"
        ws[f"E{linha_atual}"] = "VALOR UNITÁRIO"
        ws[f"F{linha_atual}"] = "VALOR TOTAL"

        for col in ["B", "C", "D", "E", "F"]:
            ws[f"{col}{linha_atual}"].font = font_bold
            ws[f"{col}{linha_atual}"].alignment = align_center
            ws[f"{col}{linha_atual}"].border = border_all

        linha_atual += 1

        soma_geral_qtd = 0
        soma_geral_valor = 0

        for cliente in df_motorista["cliente"].drop_duplicates():
            ws.merge_cells(start_row=linha_atual, start_column=2, end_row=linha_atual, end_column=6)
            ws[f"B{linha_atual}"] = cliente
            ws[f"B{linha_atual}"].font = font_bold
            ws[f"B{linha_atual}"].alignment = align_center
            ws[f"B{linha_atual}"].fill = fill_cliente

            for col in ["B", "C", "D", "E", "F"]:
                ws[f"{col}{linha_atual}"].border = border_all

            linha_atual += 1

            df_cliente = df_motorista[df_motorista["cliente"] == cliente]
            cidades = df_cliente[col_cidade].dropna().unique()

            for cidade in cidades:
                quantidade = df_cliente[df_cliente[col_cidade] == cidade].shape[0]
                valor_unitario = df_cliente[df_cliente[col_cidade] == cidade].iloc[0][col_custo]
                valor_total = quantidade * valor_unitario

                soma_geral_qtd += quantidade
                soma_geral_valor += valor_total

                ws.merge_cells(start_row=linha_atual, start_column=2, end_row=linha_atual, end_column=3)
                ws[f"B{linha_atual}"] = cidade
                ws[f"D{linha_atual}"] = quantidade
                ws[f"E{linha_atual}"] = valor_unitario
                ws[f"F{linha_atual}"] = valor_total

                ws[f"E{linha_atual}"].number_format = formato_contabil
                ws[f"F{linha_atual}"].number_format = formato_contabil

                for col in ["B", "C", "D", "E", "F"]:
                    ws[f"{col}{linha_atual}"].alignment = align_center
                    ws[f"{col}{linha_atual}"].border = border_all

                linha_atual += 1

        ws.merge_cells(start_row=linha_atual, start_column=2, end_row=linha_atual, end_column=3)
        ws[f"B{linha_atual}"] = "TOTAL"
        ws[f"D{linha_atual}"] = soma_geral_qtd
        ws[f"E{linha_atual}"] = "-"
        ws[f"F{linha_atual}"] = soma_geral_valor
        ws[f"F{linha_atual}"].number_format = formato_contabil

        for col in ["B", "C", "D", "E", "F"]:
            ws[f"{col}{linha_atual}"].font = font_bold
            ws[f"{col}{linha_atual}"].alignment = align_center
            ws[f"{col}{linha_atual}"].border = border_all

        # VALOR TOTAL DA NOTA
        linha_atual += 2
        linha_valor_nota = linha_atual

        ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=5)
        ws[f"A{linha_atual}"] = "VALOR TOTAL DOS SERVIÇOS PRESTADOS NO PERÍODO (valor da nota fiscal)"
        ws[f"A{linha_atual}"].font = font_bold
        ws[f"A{linha_atual}"].fill = fill_cliente
        ws[f"A{linha_atual}"].alignment = align_left_center

        ws[f"F{linha_atual}"] = soma_geral_valor
        ws[f"F{linha_atual}"].font = font_bold
        ws[f"F{linha_atual}"].fill = fill_cliente
        ws[f"F{linha_atual}"].alignment = align_center
        ws[f"F{linha_atual}"].number_format = formato_contabil

        for col in ["A", "B", "C", "D", "E", "F"]:
            ws[f"{col}{linha_atual}"].border = border_all

        # DESCONTOS E VALOR LÍQUIDO
        linha_atual += 2
        linha_descontos = linha_atual

        ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=5)
        ws[f"A{linha_atual}"] = "(-) DESCONTOS"
        ws[f"A{linha_atual}"].font = font_bold
        ws[f"A{linha_atual}"].alignment = align_left_center

        ws[f"F{linha_atual}"].number_format = formato_contabil
        ws[f"F{linha_atual}"].font = font_bold_red
        ws[f"F{linha_atual}"].alignment = align_center

        for col in ["A", "B", "C", "D", "E", "F"]:
            ws[f"{col}{linha_atual}"].border = border_all

        for _ in range(5):
            linha_atual += 1
            ws[f"A{linha_atual}"].border = border_lr
            ws[f"F{linha_atual}"].border = border_lr

        linha_atual += 1

        ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=5)
        ws[f"A{linha_atual}"] = "VALOR LÍQUIDO A PAGAR AO PRESTADOR DE SERVIÇO"
        ws[f"A{linha_atual}"].font = font_bold
        ws[f"A{linha_atual}"].alignment = align_left_center

        ws[f"F{linha_atual}"] = f"=F{linha_valor_nota}-F{linha_descontos}"
        ws[f"F{linha_atual}"].number_format = formato_contabil
        ws[f"F{linha_atual}"].font = font_bold
        ws[f"F{linha_atual}"].alignment = align_center

        for col in ["A", "B", "C", "D", "E", "F"]:
            ws[f"{col}{linha_atual}"].border = border_all
            ws[f"{col}{linha_atual}"].fill = fill_cliente
            ws[f"{col}{linha_atual}"].font = font_bold

    # remover aba base e salvar
    del wb["MODELO_BASE"]

    os.makedirs(os.path.dirname(saida_espelhos_xlsx_path), exist_ok=True)
    wb.save(saida_espelhos_xlsx_path)
    return saida_espelhos_xlsx_path
