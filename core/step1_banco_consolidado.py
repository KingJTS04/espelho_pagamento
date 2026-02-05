import os
import pandas as pd


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df


def gerar_banco_consolidado(
    motoristas_xlsx,
    fechamento_xlsx,
    saida_xlsx_path: str | None = None,
) -> pd.DataFrame:
    """
    Gera o banco consolidado juntando fechamento + motoristas.

    motoristas_xlsx / fechamento_xlsx:
      - pode ser caminho (str) para .xlsx
      - pode ser file-like (ex: request.files['motoristas'], BytesIO, etc)

    saida_xlsx_path:
      - se informado, salva o arquivo final nesse caminho
      - se None, apenas retorna o DataFrame
    """

    motoristas = normalizar_colunas(pd.read_excel(motoristas_xlsx))
    fechamento = normalizar_colunas(pd.read_excel(fechamento_xlsx))

    # ===============================
    # GARANTIR COLUNA "contrato" (VEM DE MOTORISTAS)
    # ===============================
    possiveis_contrato = [
        "contrato",
        "n contrato",
        "nº contrato",
        "numero do contrato",
        "número do contrato",
        "contrato nº",
        "contrato n",
        "contrato numero",
    ]

    col_contrato = next((c for c in possiveis_contrato if c in motoristas.columns), None)
    if not col_contrato:
        raise Exception("Coluna de contrato não encontrada em motoristas.xlsx (crie a coluna 'contrato').")
    if col_contrato != "contrato":
        motoristas = motoristas.rename(columns={col_contrato: "contrato"})

    # Merge
    if "nome do motorista" not in fechamento.columns:
        raise Exception("Coluna 'nome do motorista' não encontrada no fechamento.xlsx.")
    if "nome do motorista" not in motoristas.columns:
        raise Exception("Coluna 'nome do motorista' não encontrada no motoristas.xlsx.")

    banco_consolidado = fechamento.merge(motoristas, on="nome do motorista", how="left")

    # Formatar a coluna de data (você estava usando índice 2)
    if banco_consolidado.shape[1] >= 3:
        coluna_data = banco_consolidado.columns[2]
        banco_consolidado[coluna_data] = (
            pd.to_datetime(banco_consolidado[coluna_data], errors="coerce").dt.strftime("%d/%m/%Y")
        )

    # Salvar, se solicitado
    if saida_xlsx_path:
        os.makedirs(os.path.dirname(saida_xlsx_path), exist_ok=True)
        banco_consolidado.to_excel(saida_xlsx_path, index=False)

    return banco_consolidado