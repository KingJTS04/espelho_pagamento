import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment


def gerar_resumos(
    espelhos_xlsx_path: str,
    banco_consolidado_xlsx_path: str,
) -> str:
    """
    Cria as abas RESUMO e RESUMO TOTAL dentro do arquivo Espelhos_Motoristas.xlsx.

    Retorna o próprio path do espelhos (arquivo final).
    """
    if not os.path.exists(espelhos_xlsx_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {espelhos_xlsx_path}")
    if not os.path.exists(banco_consolidado_xlsx_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {banco_consolidado_xlsx_path}")

    # =========================
    # ESTILOS
    # =========================
    bold = Font(bold=True)
    bold_red = Font(bold=True, color="FF0000")
    red_font = Font(color="FF0000")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    formato_contabil = 'R$ #,##0.00_);R$ (#,##0.00)'

    # =========================
    # HELPERS
    # =========================
    def norm_text(v):
        return str(v).strip() if v is not None else ""

    def excel_sheet_ref(name: str) -> str:
        return name.replace("'", "''")

    def find_descontos_row(ws):
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                if norm_text(cell.value) == "(-) DESCONTOS":
                    return cell.row
        return None

    def find_mapeamento_header_row(ws):
        for r in range(1, ws.max_row + 1):
            b = norm_text(ws[f"B{r}"].value).upper()
            f = norm_text(ws[f"F{r}"].value).upper()
            if b == "CIDADE" and f == "VALOR TOTAL":
                return r
        return None

    def build_client_ranges_in_mapeamento(ws, clientes_set):
        header_row = find_mapeamento_header_row(ws)
        if not header_row:
            return {}

        ranges = {}
        current_client = None
        start_row = None

        r = header_row + 1
        while r <= ws.max_row:
            s = norm_text(ws[f"B{r}"].value)
            if s:
                sup = s.upper()
                if sup == "TOTAL":
                    break

                if s in clientes_set:
                    if current_client is not None and start_row is not None:
                        end_row = r - 1
                        ranges[current_client] = (start_row, end_row) if end_row >= start_row else None
                    current_client = s
                    start_row = r + 1
            r += 1

        if current_client is not None and start_row is not None:
            end_row = r - 1
            ranges[current_client] = (start_row, end_row) if end_row >= start_row else None

        return ranges

    def auto_ajuste(ws):
        for column_cells in ws.columns:
            max_length = 0
            col_letter = column_cells[0].column_letter
            for cell in column_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 3

    # =========================
    # ABRIR ARQUIVOS
    # =========================
    wb_espelhos = load_workbook(espelhos_xlsx_path, data_only=False)
    wb_banco = load_workbook(banco_consolidado_xlsx_path, data_only=True)
    ws_banco = wb_banco.active

    # remover abas se existirem
    for aba in ("RESUMO", "RESUMO TOTAL"):
        if aba in wb_espelhos.sheetnames:
            del wb_espelhos[aba]

    # =========================
    # CRIAR ABA RESUMO
    # =========================
    ws_resumo = wb_espelhos.create_sheet("RESUMO")
    ws_resumo["A1"] = "Relação dos Parceiros para Pagamento"
    ws_resumo["A1"].font = bold
    ws_resumo["A2"] = "Centro de Custo:"
    ws_resumo["A3"] = "Período:"
    ws_resumo["E3"] = "Vencimento:"
    ws_resumo["E3"].border = border_all

    # Cabeçalho RESUMO
    linha_inicio = 5
    cabecalhos = ["Nome do motorista", "Valor Bruto", "Desconto", "Valor Líquido", "Status NF"]
    for col, texto in enumerate(cabecalhos, start=1):
        cell = ws_resumo.cell(row=linha_inicio, column=col, value=texto)
        cell.font = bold_red if texto == "Desconto" else bold
        cell.alignment = center
        cell.border = border_all

    # Preencher motoristas
    linha_atual = linha_inicio + 1
    linha_primeiro_motorista = linha_atual
    motoristas_list = []
    motorista_to_sheet = {}

    for aba in wb_espelhos.sheetnames:
        if aba in ("RESUMO", "RESUMO TOTAL"):
            continue

        ws = wb_espelhos[aba]
        motorista = ws["C4"].value
        if not motorista:
            continue

        valor_bruto = 0
        desconto_linha = None

        for row in ws.iter_rows():
            for cell in row:
                if not cell.value:
                    continue
                texto = str(cell.value)

                if "VALOR TOTAL DOS SERVIÇOS PRESTADOS" in texto:
                    valor_bruto = ws[f"F{cell.row}"].value or 0

                if texto.strip() == "(-) DESCONTOS":
                    desconto_linha = cell.row

        ws_resumo.cell(row=linha_atual, column=1, value=motorista)
        motoristas_list.append(motorista)
        motorista_to_sheet[motorista] = aba

        c_bruto = ws_resumo.cell(row=linha_atual, column=2, value=valor_bruto)
        c_bruto.number_format = formato_contabil
        c_bruto.alignment = center

        if desconto_linha:
            c_desc = ws_resumo.cell(row=linha_atual, column=3, value=f"='{aba}'!F{desconto_linha}")
        else:
            c_desc = ws_resumo.cell(row=linha_atual, column=3, value=0)

        c_desc.font = red_font
        c_desc.number_format = formato_contabil
        c_desc.alignment = center

        c_liq = ws_resumo.cell(row=linha_atual, column=4, value=f"=B{linha_atual}-C{linha_atual}")
        c_liq.number_format = formato_contabil
        c_liq.alignment = center

        ws_resumo.cell(row=linha_atual, column=5, value="")

        for c in range(1, 6):
            ws_resumo.cell(row=linha_atual, column=c).border = border_all

        linha_atual += 1

    # Total RESUMO
    ws_resumo.cell(row=linha_atual, column=1, value="CUSTO TOTAL").font = bold

    ct_bruto = ws_resumo.cell(row=linha_atual, column=2, value=f"=SUM(B{linha_primeiro_motorista}:B{linha_atual-1})")
    ct_bruto.number_format = formato_contabil
    ct_bruto.font = bold
    ct_bruto.alignment = center

    ct_desc = ws_resumo.cell(row=linha_atual, column=3, value=f"=SUM(C{linha_primeiro_motorista}:C{linha_atual-1})")
    ct_desc.number_format = formato_contabil
    ct_desc.font = bold_red
    ct_desc.alignment = center

    ct_liq = ws_resumo.cell(row=linha_atual, column=4, value=f"=SUM(D{linha_primeiro_motorista}:D{linha_atual-1})")
    ct_liq.number_format = formato_contabil
    ct_liq.font = bold
    ct_liq.alignment = center

    for c in range(1, 6):
        cell = ws_resumo.cell(row=linha_atual, column=c)
        cell.border = border_all
        if c not in (2, 3, 4):
            cell.font = bold

    # =========================
    # CRIAR ABA RESUMO TOTAL
    # =========================
    ws_rt = wb_espelhos.create_sheet("RESUMO TOTAL")
    ws_rt["A1"] = "Relação dos Parceiros para Pagamento"
    ws_rt["A1"].font = bold
    ws_rt["A2"] = "Centro de Custo:"
    ws_rt["A3"] = "Período:"
    ws_rt["A5"] = "Nome do Motorista"
    ws_rt["A5"].font = bold
    ws_rt["A5"].border = border_all

    linha_cabecalho = 5
    linha_rt = 6

    for idx, motorista in enumerate(motoristas_list):
        cell = ws_rt.cell(row=linha_rt + idx, column=1, value=motorista)
        cell.border = border_all

    # Clientes únicos
    clientes_unicos = []
    col_cliente = None
    for col in range(1, ws_banco.max_column + 1):
        if norm_text(ws_banco.cell(row=1, column=col).value).lower() == "cliente":
            col_cliente = col
            break

    if col_cliente:
        for row in range(2, ws_banco.max_row + 1):
            cliente = ws_banco.cell(row=row, column=col_cliente).value
            if cliente and cliente not in clientes_unicos:
                clientes_unicos.append(cliente)

    clientes_set = set(clientes_unicos)

    col_inicio = 2
    for idx, cliente in enumerate(clientes_unicos):
        c = ws_rt.cell(row=linha_cabecalho, column=col_inicio + idx, value=cliente)
        c.font = bold
        c.alignment = center
        c.border = border_all

    # Colunas finais
    col_desconto = col_inicio + len(clientes_unicos)
    col_liquido = col_desconto + 1
    col_status = col_liquido + 1

    hd = ws_rt.cell(row=linha_cabecalho, column=col_desconto, value="Desconto")
    hd.font = bold_red
    hd.alignment = center
    hd.border = border_all

    hl = ws_rt.cell(row=linha_cabecalho, column=col_liquido, value="Valor Líquido")
    hl.font = bold
    hl.alignment = center
    hl.border = border_all

    hs = ws_rt.cell(row=linha_cabecalho, column=col_status, value="Status NF")
    hs.font = bold
    hs.alignment = center
    hs.border = border_all

    venc = ws_rt.cell(row=3, column=col_status, value="Vencimento:")
    venc.font = bold
    venc.alignment = center
    venc.border = border_all

    # pré-cálculos
    ranges_por_motorista = {}
    desconto_row_por_motorista = {}
    for motorista, sheet_name in motorista_to_sheet.items():
        ws_m = wb_espelhos[sheet_name]
        ranges_por_motorista[motorista] = build_client_ranges_in_mapeamento(ws_m, clientes_set)
        desconto_row_por_motorista[motorista] = find_descontos_row(ws_m)

    first_client_col = col_inicio
    last_client_col = col_inicio + len(clientes_unicos) - 1

    # preencher linhas
    for i, motorista in enumerate(motoristas_list):
        row_out = linha_rt + i
        sheet_name = motorista_to_sheet.get(motorista)
        sheet_ref = excel_sheet_ref(sheet_name) if sheet_name else None

        client_ranges = ranges_por_motorista.get(motorista, {})
        desc_row = desconto_row_por_motorista.get(motorista)

        # clientes
        for j, cliente in enumerate(clientes_unicos):
            out_cell = ws_rt.cell(row=row_out, column=col_inicio + j)
            rng = client_ranges.get(cliente)

            if sheet_ref and rng:
                s, e = rng
                out_cell.value = f"=SUM('{sheet_ref}'!F{s}:F{e})"
                out_cell.number_format = formato_contabil
            else:
                out_cell.value = ""

            out_cell.alignment = center
            out_cell.border = border_all

        # desconto
        cdesc = ws_rt.cell(row=row_out, column=col_desconto)
        if sheet_ref and desc_row:
            cdesc.value = f"='{sheet_ref}'!F{desc_row}"
            cdesc.number_format = formato_contabil
        else:
            cdesc.value = ""
        cdesc.font = red_font
        cdesc.alignment = center
        cdesc.border = border_all

        # líquido
        first_letter = ws_rt.cell(row=linha_cabecalho, column=first_client_col).column_letter
        last_letter = ws_rt.cell(row=linha_cabecalho, column=last_client_col).column_letter
        desc_letter = ws_rt.cell(row=linha_cabecalho, column=col_desconto).column_letter

        cliq = ws_rt.cell(
            row=row_out,
            column=col_liquido,
            value=f"=SUM({first_letter}{row_out}:{last_letter}{row_out})-N({desc_letter}{row_out})"
        )
        cliq.number_format = formato_contabil
        cliq.alignment = center
        cliq.border = border_all

        # status nf (vazio)
        cnf = ws_rt.cell(row=row_out, column=col_status, value="")
        cnf.alignment = center
        cnf.border = border_all

    # auto ajuste
    auto_ajuste(ws_resumo)
    auto_ajuste(ws_rt)

    wb_espelhos.save(espelhos_xlsx_path)
    return espelhos_xlsx_path
