import os
import pandas as pd


def _normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()
    return df


def gerar_banco_consolidado(
    motoristas_xlsx_path: str,
    fechamento_xlsx_path: str,
    saida_xlsx_path: str,
    *,
    coluna_data_idx: int = 2,  # Coluna C = índice 2
    merge_key: str = "nome do motorista"
) -> str:
    """
    Junta motoristas + fechamento e salva banco_consolidado.xlsx.

    Retorna o path do arquivo gerado.
    """
    if not os.path.exists(motoristas_xlsx_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {motoristas_xlsx_path}")
    if not os.path.exists(fechamento_xlsx_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {fechamento_xlsx_path}")

    motoristas = _normalizar_colunas(pd.read_excel(motoristas_xlsx_path))
    fechamento = _normalizar_colunas(pd.read_excel(fechamento_xlsx_path))

    if merge_key not in motoristas.columns:
        raise ValueError(f"Coluna '{merge_key}' não encontrada no arquivo de motoristas.")
    if merge_key not in fechamento.columns:
        raise ValueError(f"Coluna '{merge_key}' não encontrada no arquivo de fechamento.")

    banco_consolidado = fechamento.merge(motoristas, on=merge_key, how="left")

    # Formatar data na coluna C (índice 2), se existir
    if len(banco_consolidado.columns) > coluna_data_idx:
        coluna_data = banco_consolidado.columns[coluna_data_idx]
        banco_consolidado[coluna_data] = (
            pd.to_datetime(banco_consolidado[coluna_data], errors="coerce")
            .dt.strftime("%d/%m/%Y")
        )

    os.makedirs(os.path.dirname(saida_xlsx_path), exist_ok=True)
    banco_consolidado.to_excel(saida_xlsx_path, index=False)
    return saida_xlsx_path
