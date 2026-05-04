import pandas as pd

def check_not_null(df, col):
    return df[col].notnull(), f"Campo {col} es obligatorio"

def check_range(df, col, min_val, max_val):
    return df[col].between(min_val, max_val), f"{col} fuera de rango ({min_val}-{max_val})"

def check_regex(df, col, pattern):
    return df[col].astype(str).str.match(pattern), f"Formato inválido en {col}"

def validate_range(df: pd.DataFrame, col: str, min_val: float, max_val: float):
    mask = df[col].between(min_val, max_val)
    return mask, f"Valor fuera de rango ({min_val} - {max_val})"

def validate_regex(df: pd.DataFrame, col: str, pattern: str):
    mask = df[col].astype(str).str.match(pattern, na=False)
    return mask, f"No cumple con el formato requerido"

def validate_not_null(df: pd.DataFrame, col: str):
    mask = df[col].notnull()
    return mask, "Este campo es obligatorio"

VALIDATION_CONFIG = {
    "Ventas": [
        {"func": check_not_null, "args": ["ID_Transaccion"]},
        {"func": check_range, "args": ["Monto", 0, 1000000]},
    ],
    "Inventario": [
        {"func": check_regex, "args": ["SKU", r"^[A-Z]{3}-\d+$"]},
    ]
}


