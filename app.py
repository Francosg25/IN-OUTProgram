import streamlit as st
import pandas as pd
import re
import difflib
import json
import os
import logging
from pathlib import Path
from typing import Dict, Any, Literal, List, Optional, Tuple
from io import BytesIO

def render_safe_table(df_obj):
    """
    Bypass de seguridad: Convierte DataFrames/Stylers a HTML puro para 
    evadir la carga de PyArrow en entornos con políticas de Application Control.
    """
    if df_obj is None or (hasattr(df_obj, 'empty') and df_obj.empty):
        return
    
    html_str = df_obj.to_html()
    
    st.markdown(
        f'<div style="overflow-x: auto; max-height: 400px; margin-bottom: 20px; border: 1px solid #e6e6e6; border-radius: 5px;">{html_str}</div>', 
        unsafe_allow_html=True
    )

try:
    from engine.allocation import CostAllocationEngine
except ImportError:
    # Fallback para demostración si el módulo no existe en el path
    class CostAllocationEngine:
        def calculate_outbound(self, df, costs): 
            df['Calc_Exp'] = 0
            df['%PCT'] = 0
            summary = df.groupby(['transport_type', 'bu'], as_index=False).sum(numeric_only=True)
            summary.rename(columns={'bu': 'BU', 'transport_type': 'Transport', 'Calc_Exp': 'Arg. Var $'}, inplace=True)
            summary.attrs['reconciliation'] = {'total_facturado': 0, 'total_asignado': 0, 'diferencia': 0, 'match_rate': 0}
            return summary

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

CACHE_FILENAME = "mapping_cache.json"


def smart_read_excel(file_obj) -> pd.DataFrame:
    """
    Escáner de densidad heurística Multi-Hoja con consolidación automática de columnas.
    """
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
            
        xls = pd.ExcelFile(file_obj, engine='openpyxl')
        
        # Diccionario expandido
        keywords = ['REF', 'REFERENCE', 'GUIA', 'WAYBILL', 'AWB', 'TRACKING', 
                    'BU', 'OU', 'BUSINESS UNIT', 'UNIDAD', 
                    'PESO', 'WEIGHT', 'KGS', 'LBS', 'GROSS', 
                    'PART', 'ITEM', 'QTY', 'PIECES', 'BULTOS', 
                    'CONTAINER', 'CONTENEDOR', 'CNTR', 'EQUIPO', 
                    'COST', 'PRICE', 'AMOUNT', 'USD', 'FIX COST']
        
        global_max_score = -1
        best_sheet = None
        best_idx = 0
        
        # 1. ESCÁNER MULTI-HOJA
        for sheet in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=30)
            
            for idx, row in df_temp.iterrows():
                row_list = row.tolist()
                score = 0
                non_null_count = 0
                for cell in row_list:
                    if pd.notna(cell) and str(cell).strip() != "":
                        non_null_count += 1 
                        cell_str = str(cell).upper().strip()
                        if any(kw in cell_str for kw in keywords):
                            score += 10 
                            
                total_score = score + non_null_count
                if total_score > global_max_score:
                    global_max_score = total_score
                    best_idx = idx
                    best_sheet = sheet
                    
        # 2. LECTURA DE LA PESTAÑA GANADORA
        df = pd.read_excel(xls, sheet_name=best_sheet, header=best_idx)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        # 3. CONSOLIDADOR INTELIGENTE DE PESOS MULTIPLES
        peso_cols = [col for col in df.columns if 'GROSS WEIGHT' in col or 'PESO BRUTO' in col]
        
        if len(peso_cols) > 0:
            df['CONSOLIDATED_GROSS_WEIGHT'] = pd.to_numeric(df[peso_cols[0]], errors='coerce').fillna(0)
            for col in peso_cols[1:]:
                df['CONSOLIDATED_GROSS_WEIGHT'] += pd.to_numeric(df[col], errors='coerce').fillna(0)
                
            df.drop(columns=peso_cols, inplace=True)
            df.rename(columns={'CONSOLIDATED_GROSS_WEIGHT': 'GROSS_WEIGHT'}, inplace=True)
        
        return df

    except Exception as e:
        st.error(f"Error técnico al leer el archivo: {e}")
        return pd.DataFrame()

def get_cache_path() -> str:
    return os.path.join(os.path.dirname(__file__), CACHE_FILENAME)


def load_mapping_cache() -> Dict[str, Dict[str, str]]:
    path = get_cache_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
            return {k: v for k, v in data.items() if isinstance(v, dict)}
    except Exception as e:
        logger.warning(f"No se pudo leer el caché de mapeos: {e}")
        return {}


def save_mapping_cache(cache: Dict[str, Dict[str, str]]) -> None:
    path = get_cache_path()
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(cache, handle, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning(f"No se pudo guardar el caché de mapeos: {e}")


TRAVEL_COLUMN_ALIASES = {
    'Sea': {
        'reference': ['reference', 'ref', 'guia', 'guía', 'awb', 'booking', 'shipment', 'documento'],
        'container_number': ['container', 'contenedor', 'cntr', 'box', 'equipo'],
        'bu': ['bu', 'BU', 'ou', 'business unit', 'unidad de negocio', 'unidad', 'area', 'division', 'empresa', 'departamento', 'department', 'depto', 'unit'],
        'gross_weight': ['gross_weight', 'peso', 'weight', 'kg', 'kgs', 'peso bruto'],
        'price': ['price', 'precio', 'cost', 'valor', 'amount', 'monto'],
        'part_number': ['part_number', 'part number', 'numero de parte', 'item', 'item code']
    },
    'Land': {
        'reference': ['reference', 'ref', 'doc', 'documento', 'orden', 'guia', 'guía'],
        'container_number': ['container', 'contenedor', 'truck', 'camion', 'vehiculo'],
        'bu': ['bu', 'BU', 'ou', 'business unit', 'unidad de negocio', 'unidad', 'area', 'division', 'empresa', 'departamento', 'department', 'depto', 'unit'],
        'gross_weight': ['gross_weight', 'peso', 'weight', 'kg', 'kgs', 'peso bruto'],
        'price': ['price', 'precio', 'cost', 'valor', 'amount', 'monto'],
        'part_number': ['part_number', 'part number', 'numero de parte', 'item', 'item code']
    },
    'Outbound': {
        'reference': ['reference', 'ref', 'shipment', 'export', 'tracking', 'documento', 'awb'],
        'container_number': ['container', 'contenedor', 'cntr', 'box', 'equipo'],
        'bu': ['bu', 'BU', 'ou', 'business unit', 'unidad de negocio', 'unidad', 'empresa', 'division'],
        'gross_weight': ['gross_weight', 'peso', 'weight', 'kg', 'kgs', 'peso bruto'],
        'price': ['price', 'precio', 'cost', 'valor', 'amount', 'monto'],
        'part_number': ['part_number', 'part number', 'numero de parte', 'item', 'item code']
    }
}

GENERIC_COLUMN_ALIASES = {
    'reference': ['reference', 'ref', 'guia', 'guía', 'documento', 'doc', 'tracking', 'awb', 'waybill'],
    'container_number': ['container', 'contenedor', 'cntr', 'box', 'equipo'],
    'bu': ['bu', 'BU', 'ou', 'business unit', 'unidad de negocio'],
    'gross_weight': ['gross_weight', 'peso', 'weight', 'kg', 'kgs', 'peso bruto'],
    'price': ['price', 'precio', 'cost', 'valor', 'amount', 'monto'],
    'part_number': ['part_number', 'part number', 'numero de parte', 'item', 'item code']
}


def normalize_col_name(name: str) -> str:
    return re.sub(r'[\W_]+', '', str(name).strip().lower())


def numeric_column_score(series: pd.Series) -> float:
    cleaned = series.astype(str).str.replace(r'[^0-9\-\.,]', '', regex=True).str.replace(',', '.')
    numeric = pd.to_numeric(cleaned, errors='coerce')
    valid = numeric.notna() & (numeric > 0)
    if len(series) == 0:
        return 0.0
    return valid.sum() / len(series)


def textual_column_score(series: pd.Series) -> float:
    text = series.astype(str).str.strip()
    non_empty = text[text != '']
    if len(non_empty) == 0:
        return 0.0
    numeric_like = non_empty.str.match(r'^[0-9\.,\-]+$').sum()
    if numeric_like / len(non_empty) > 0.5:
        return 0.0
    unique_ratio = non_empty.nunique() / len(non_empty)
    return min(1.0, max(0.3, unique_ratio))


def reference_column_score(series: pd.Series) -> float:
    text = series.astype(str).str.strip()
    non_empty = text[text != '']
    if len(non_empty) == 0:
        return 0.0
    has_letters = non_empty.str.contains(r'[A-Za-z]').mean()
    has_digits = non_empty.str.contains(r'\d').mean()
    unique_ratio = non_empty.nunique() / len(non_empty)
    return min(1.0, has_letters * 0.5 + has_digits * 0.3 + unique_ratio * 0.2)


def fuzzy_match_column(columns_list: list, keywords: list) -> Optional[str]:
    if not columns_list:
        return None

    normalized_keywords = [normalize_col_name(k) for k in keywords]
    best_match = None
    best_ratio = 0.0
    for col in columns_list:
        col_norm = normalize_col_name(col)
        for kw_norm in normalized_keywords:
            ratio = difflib.SequenceMatcher(None, col_norm, kw_norm).ratio()
            if ratio > best_ratio and ratio >= 0.70:
                best_ratio = ratio
                best_match = col
    return best_match


def find_best_column(columns_list: list, keywords: list) -> Optional[str]:
    """
    Escáner heurístico avanzado. Prioriza coincidencias exactas,
    luego palabras completas, luego coincidencias parciales y finalmente fuzzy.
    """
    if not columns_list:
        return None

    keywords_upper = [str(k).strip().upper() for k in keywords]

    # 1. MATCH EXACTO (Prioridad Máxima)
    for col in columns_list:
        col_clean = str(col).strip().upper()
        if col_clean in keywords_upper:
            return col

    # 2. MATCH DE PALABRA COMPLETA (Word Boundaries)
    for col in columns_list:
        col_clean = str(col).strip().upper()
        for kw_upper in keywords_upper:
            if re.search(rf'\b{re.escape(kw_upper)}\b', col_clean):
                return col

    # 3. MATCH PARCIAL (Fallback de rescate)
    for col in columns_list:
        col_clean = str(col).strip().upper()
        for kw_upper in keywords_upper:
            if kw_upper in col_clean:
                if kw_upper == 'BU' and 'BULTO' in col_clean:
                    continue
                return col

    # 3.5. MATCH NORMALIZADO SIN ESPACIOS/PUNTUACIÓN
    for col in columns_list:
        col_norm = normalize_col_name(col)
        for kw_upper in keywords_upper:
            if normalize_col_name(kw_upper) in col_norm:
                if kw_upper == 'BU' and 'BULTO' in col_norm:
                    continue
                return col

    # 4. FUZZY MATCHING
    return fuzzy_match_column(columns_list, keywords)


def detect_fallback_column(df: pd.DataFrame, field_name: str) -> Optional[str]:
    if field_name in ('gross_weight', 'price'):
        best = None
        best_score = 0.0
        for col in df.columns:
            score = numeric_column_score(df[col])
            if score > best_score:
                best_score = score
                best = col
        return best if best_score >= 0.65 else None

    if field_name == 'bu':
        # Buscamos primero columnas con tokens típicos de BU en el encabezado
        bu_tokens = ['bu', 'businessunit', 'unidaddenegocio', 'unidad', 'area', 'division', 'empresa', 'departamento', 'depto', 'unit', 'ou']
        for col in df.columns:
            col_norm = normalize_col_name(col)
            if any(token in col_norm for token in bu_tokens):
                return col

        best = None
        best_score = 0.0
        for col in df.columns:
            score = textual_column_score(df[col])
            if score > best_score:
                best_score = score
                best = col
        return best if best_score >= 0.5 else None

    if field_name == 'reference':
        best = None
        best_score = 0.0
        for col in df.columns:
            score = reference_column_score(df[col])
            if score > best_score:
                best_score = score
                best = col
        return best if best_score >= 0.4 else None

    return None


def suggest_mapping(df: pd.DataFrame, label: str) -> dict:
    cols = df.columns.tolist()
    aliases = TRAVEL_COLUMN_ALIASES.get(label, GENERIC_COLUMN_ALIASES)
    mapping = {}
    for field in ['reference', 'container_number', 'bu', 'gross_weight', 'price', 'part_number']:
        candidate = find_best_column(cols, aliases.get(field, GENERIC_COLUMN_ALIASES[field]))
        if not candidate:
            candidate = detect_fallback_column(df, field)
        mapping[field] = candidate
    return mapping


class LogisticsOrchestrator:
    """Orquestador de Datos: Soporta mapeo manual dinámico por archivo."""
    def __init__(self, allocation_engine):
        self.allocation_engine = allocation_engine
        self.canon_map = ['reference', 'bu', 'gross_weight']

    def _normalize_bu(self, raw_val: Any) -> str:
        """
        Sanitizador Estricto para Business Units (Ej: 1 -> M01, M-19 -> M19).
        """
        val = str(raw_val).strip().upper()
        if val in ['NAN', 'NONE', '', 'NULL']:
            return 'N/A'
            
        # Limpiamos caracteres raros, dejando solo letras y números
        clean_val = re.sub(r'[^A-Z0-9]', '', val)
        
        if clean_val.isdigit():
            num = int(clean_val)
            if 0 < num < 1000:
                return f"M{num:02d}" if num < 100 else f"M{num}"
            else:
                return "MISCELANEUS"  
            
        # CASO 2: Tiene la 'M' pero le falta el cero (Ej: "M1", "M19")
        match = re.match(r'^M(\d+)$', clean_val)
        if match:
            num = int(match.group(1))
            return f"M{num:02d}" if num < 100 else f"M{num}"
            
        # CASO 3: Es una cadena de texto (Ej: "CAPEX", "MISCELANEUS", o texto no mapeado)
        return clean_val
    
    def _clean_reference_key(self, raw_ref: str) -> str:
        """
        Limpia la llave primaria (Waybill) para asegurar un Match perfecto al 100%.
        Remueve sufijos como .M46, -2, etc.
        Ej: 'FG-R-2180LE25.M46-M45-2' -> 'FG-R-2180LE25'
        """
        val = str(raw_ref).strip().upper()
        if val in ['NAN', 'NONE', '']:
            return 'UNKNOWN'
        
        # 1. Cortar en el primer punto (Ej. quita .M46)
        if '.' in val:
            val = val.split('.')[0]
            
        # 2. Si la guía tiene el formato estándar de JE (ej. FG-R-XXXX), aseguramos mantener solo esa parte
        # Si tiene guiones extra al final que no son parte del prefijo, los cortamos.
        match = re.search(r'^(FG-R-\w+)', val)
        if match:
            return match.group(1)
            
        return val

    def standardize_source_manual(self, df: pd.DataFrame, label: str, mapping: Dict[str, str]) -> pd.DataFrame:
        """Estandariza un DF usando un mapeo manual proporcionado por el usuario."""
        try:
            # 1. Validación Estricta
            for canon, excel_col in mapping.items():
                if excel_col and excel_col not in df.columns:
                    raise ValueError(f"Columna '{excel_col}' no encontrada en el archivo {label}. Columnas disponibles: {list(df.columns)}")
            
            # 2. Extracción Robusta Anti-Duplicados
            df_filtered = pd.DataFrame()
            rename_dict = {v: k for k, v in mapping.items() if v}
            
            for excel_col, canon in rename_dict.items():
                col_data = df[excel_col]
                
                # Si el Excel tiene encabezados duplicados, Pandas devuelve un DataFrame.
                # Aislar la primera columna nos protege de crashes de dimensionalidad.
                if isinstance(col_data, pd.DataFrame):
                    col_data = col_data.iloc[:, 0]
                    
                df_filtered[canon] = col_data.copy()
            
            df_filtered['transport_type'] = label
            
            # Ahora es 100% seguro usar el accesor .str porque garantizamos que es una Series
            df_filtered['reference'] = df_filtered['reference'].astype(str).str.strip()
            
            # --- LÓGICA DE RESCATE DE BU ---
            if 'bu' not in df_filtered.columns:
                def extract_bu_from_reference(ref_str):
                    if pd.isna(ref_str) or str(ref_str).strip() == '':
                        return 'DEFAULT_BU'
                    match = re.search(r'M(\d{1,3})\b', str(ref_str).upper())
                    if match:
                        num = int(match.group(1))
                        return f"M{num:02d}"
                    return 'DEFAULT_BU'
                
                df_filtered['bu'] = df_filtered['reference'].apply(extract_bu_from_reference)
                
            # SANITIZADOR INTELIGENTE
            df_filtered['bu'] = df_filtered['bu'].apply(self._normalize_bu)

            # 3. Casteo de Pesos y Precios
            if 'gross_weight' not in df_filtered.columns:
                df_filtered['gross_weight'] = 1.0
            else:
                df_filtered['gross_weight'] = pd.to_numeric(df_filtered['gross_weight'], errors='coerce').fillna(1.0)
            
            if 'price' not in df_filtered.columns:
                df_filtered['price'] = 0.0
            else:
                df_filtered['price'] = pd.to_numeric(df_filtered['price'], errors='coerce').fillna(0.0)

            # 4. Auditoría de Datos
            df_filtered['note'] = ''
            missing_weight = df_filtered['gross_weight'] <= 0
            if missing_weight.any():
                df_filtered.loc[missing_weight, 'gross_weight'] = 1.0
                df_filtered.loc[missing_weight, 'note'] += 'Peso imputado; '

            missing_price = df_filtered['price'] <= 0
            if missing_price.any():
                df_filtered.loc[missing_price, 'note'] += 'Precio ausente; '

            if 'reference' not in df_filtered.columns or 'bu' not in df_filtered.columns:
                raise ValueError(f"Faltan columnas obligatorias (reference o bu) en {label}.")

            return df_filtered
            
        except ValueError as ve:
             raise ve
        except Exception as e:
            logger.error(f"Error interno al procesar {label}: {str(e)}", exc_info=True)
            import streamlit as st
            st.error(f"**Error técnico procesando {label}:** `{str(e)}`")
            raise RuntimeError(f"Fallo en {label}. Detalle técnico: {str(e)}") from e
        
    def run_pipeline(self, processed_dfs: List[pd.DataFrame], df_costs: pd.DataFrame) -> pd.DataFrame:
        """Ejecuta la unificación de DFs y el cálculo de prorrateo."""
        if not processed_dfs:
            raise ValueError("No hay datos procesados para ejecutar el prorrateo.")
        
        unified = pd.concat(processed_dfs, ignore_index=True)
        return self.allocation_engine.calculate_outbound(unified, df_costs)
    
    def standardize_outbound_robust(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Versión robusta para Outbound que resuelve la ambigüedad de 
        columnas duplicadas en los reportes de JE.
        """
        try:
            # 1. Identificar la columna de costo correcta (Fix Cost suele ser la maestra)
            # Buscamos la columna que sume ~126k, no la que sume ~32k
            potential_costs = [c for c in df.columns if 'FIX COST' in c.upper() or 'CALC_EXP' in c.upper()]
            cost_col = potential_costs[0] if potential_costs else df.columns[0]

            #  Identificar la BU correcta (Evitar las columnas con NaNs)
            # En tu Excel, la columna 'BU.2' es la que parece tener el mapeo final
            bu_candidates = [c for c in df.columns if 'BU' in c.upper()]
            # Elegimos la que tenga menos nulos
            bu_col = min(bu_candidates, key=lambda c: df[c].isnull().sum())

            # 3. Creación del DataFrame Limpio
            df_std = pd.DataFrame()
            df_std['reference'] = df['Reference'].astype(str).str.strip()
            df_std['bu'] = df[bu_col].astype(str).str.strip().str.upper()
            df_std['price'] = pd.to_numeric(df[cost_col], errors='coerce').fillna(0)
            df_std['transport_type'] = 'Outbound'

            # 4. Filtro de Integridad: Eliminar filas donde el costo es 0 o la BU es 'TOTAL'
            df_std = df_std[
                (df_std['price'] > 0) & 
                (~df_std['bu'].isin(['TOTAL', 'NAN', 'N/A']))
            ]

            return df_std

        except Exception as e:
            raise RuntimeError(f"Error en mapeo robusto Outbound: {str(e)}")


def build_executive_tables(flat_summary: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Transforma el resumen plano en vistas ejecutivas pivotadas (Formato Johnson Electric)."""
    if flat_summary.empty:
        return pd.DataFrame(), pd.DataFrame()

    df_c = flat_summary.copy()
    df_c['Transport'] = df_c['Transport'].fillna('UNKNOWN').astype(str)
    
    # --- RED DE SEGURIDAD (SANITIZACIÓN FINAL) ---
    # Forzamos mayúsculas y quitamos espacios residuales para evitar duplicados en el Pivot
    if 'BU' in df_c.columns:
        df_c['BU'] = df_c['BU'].astype(str).str.strip().str.upper()

    # Tabla de Porcentajes
    pct_pivot = df_c.pivot_table(
        index='Transport', 
        columns='BU', 
        values='%PCT', 
        aggfunc='sum', 
        fill_value=0
    )
    
    pct_pivot.index = pct_pivot.index + ' %PCT'
    pct_pivot.index.name = 'Type' # Coincide con la imagen

    # Tabla Monetaria
    mon_pivot = df_c.pivot_table(
        index='Transport', 
        columns='BU', 
        values='Arg. Var $', 
        aggfunc='sum', 
        fill_value=0
    )
    # Insertar total con el nombre exacto de la imagen
    mon_pivot.insert(0, 'Arg. Var $', mon_pivot.sum(axis=1))
    mon_pivot.index.name = 'Viewer' # Coincide con la imagen

    return pct_pivot, mon_pivot

def main():
    st.set_page_config(page_title="Sistema de cuentas", layout="wide", page_icon="")
    st.title("Sistema logistico ")
    st.markdown("""
    Esta herramienta permite cargar reportes logísticos con cualquier estructura y asignar sus costos 
    con detección automática de columnas según el tipo de viaje.
    """)
    st.divider()

    # 1. CARGA DE ARCHIVOS FINANCIEROS
    st.sidebar.header("Archivos Base")
    cost_file = st.sidebar.file_uploader("Consolidado Facturación ($)", type=["xlsx", "csv"])
    st.sidebar.info("Este archivo debe contener las Referencias y los montos a prorratear.")
    

    st.subheader("Carga de Bases Operativas")
    c_sea, c_land, c_out = st.columns(3)
    
    with c_sea: sea_file = st.file_uploader("Marítimo / China", type=["xlsx"], key="u_sea")
    with c_land: land_file = st.file_uploader("Terrestre / Land", type=["xlsx"], key="u_land")
    with c_out: out_file = st.file_uploader("Outbound / Export", type=["xlsx"], key="u_out")

    mapping_cache = load_mapping_cache()
    if "mapping_cache" not in st.session_state:
        st.session_state.mapping_cache = mapping_cache

    st.markdown("Mapeo de Columnas")
    st.caption("El sistema detecta automáticamente las columnas más probables según el tipo de viaje. Corrige sólo aquellas que no se identifiquen con precisión.")

   
    def create_mapping_ui(file, label):
        if file:
            df = smart_read_excel(file)
            suggested = suggest_mapping(df, label)
            saved = st.session_state.mapping_cache.get(label, {})
            cols = [""] + df.columns.tolist()
            with st.expander(f"Configurar columnas de {label}", expanded=True):
                st.write(f"**Columnas disponibles en {label}:** {', '.join(df.columns.tolist())}")
                st.markdown(
                    f"**Sugerido**: referencia = `{suggested.get('reference') or 'No detectada'}`, "
                    f"BU = `{suggested.get('bu') or 'No detectada'}`, "
                    f"peso = `{suggested.get('gross_weight') or 'No detectado'}`, "
                    f"precio = `{suggested.get('price') or 'No detectado'}`"
                )
                st.markdown(
                    f"**Último mapeo guardado**: referencia = `{saved.get('reference') or 'No guardado'}`, "
                    f"BU = `{saved.get('bu') or 'No guardado'}`, "
                    f"peso = `{saved.get('gross_weight') or 'No guardado'}`, "
                    f"precio = `{saved.get('price') or 'No guardado'}`"
                )
                col1, col2, col3, col4 = st.columns(4)
                def select_with_default(label_text, key, default):
                    index = cols.index(default) if default in cols else 0
                    return st.selectbox(label_text, cols, index=index, key=key)

                default_ref = saved.get('reference') or suggested.get('reference')
                default_container = saved.get('container_number') or suggested.get('container_number')
                default_bu = saved.get('bu') or suggested.get('bu')
                default_w = saved.get('gross_weight') or suggested.get('gross_weight')
                default_price = saved.get('price') or suggested.get('price')
                default_part = saved.get('part_number') or suggested.get('part_number')

                m_ref = select_with_default(f"Referencia / Guía ({label})", f"sel_ref_{label}", default_ref)
                m_container = select_with_default(f"Número de Contenedor ({label})", f"sel_container_{label}", default_container)
                m_bu = select_with_default(f"Unidad de Negocio ({label})", f"sel_bu_{label}", default_bu)
                m_w = select_with_default(f"Peso Bruto ({label}) - Opcional", f"sel_w_{label}", default_w)
                m_price = select_with_default(f"Precio / Valor ({label}) - Opcional", f"sel_price_{label}", default_price)
                m_part = select_with_default(f"Número de Parte ({label}) - Opcional", f"sel_part_{label}", default_part)

                if m_ref: 
                    mapping = {"reference": m_ref}
                    if m_bu: mapping["bu"] = m_bu
                    if m_container: mapping["container_number"] = m_container
                    if m_w: mapping["gross_weight"] = m_w
                    if m_price: mapping["price"] = m_price
                    if m_part: mapping["part_number"] = m_part 
                    st.session_state.mapping_cache[label] = mapping
                    return df, mapping
                else:
                    st.warning("Selecciona al menos Referencia para habilitar esta fuente.")
        return None, None

    sea_data, sea_map = create_mapping_ui(sea_file, "Sea")
    land_data, land_map = create_mapping_ui(land_file, "Land")
    out_data, out_map = create_mapping_ui(out_file, "Outbound")

    st.divider()

    if st.button("Ejecutar proceso", use_container_width=True):
        if not cost_file:
            st.error("Error: Debes cargar el archivo de Facturación ($) en la barra lateral.")
            return

        try:
            engine = CostAllocationEngine(allocation_type='weight')
            orchestrator = LogisticsOrchestrator(engine)
            
            to_process = []
            if sea_data is not None and sea_map: to_process.append(orchestrator.standardize_source_manual(sea_data, "Sea", sea_map))
            if land_data is not None and land_map: to_process.append(orchestrator.standardize_source_manual(land_data, "Land", land_map))
            if out_data is not None and out_map: to_process.append(orchestrator.standardize_source_manual(out_data, "Outbound", out_map))
            
            to_process = [d for d in to_process if not d.empty]
            
            if not to_process:
                st.warning("No hay fuentes operativas configuradas correctamente.")
                return

            if hasattr(cost_file, 'name') and cost_file.name.lower().endswith('.csv'):
                df_costs = pd.read_csv(cost_file, encoding='latin1')
            else:
                df_costs = smart_read_excel(cost_file)
            
            logger.info(f"Cost file loaded: {len(df_costs)} rows, columns: {df_costs.columns.tolist()}")
            if df_costs.empty:
                st.error("El archivo de costos está vacío o no se pudo leer correctamente.")
                return
            
            with st.spinner("Procesando datos y aplicando reglas de negocio..."):
                unified = pd.concat(to_process, ignore_index=True)
                final_summary = orchestrator.run_pipeline(to_process, df_costs)
                recon = final_summary.attrs.get('reconciliation', {})
                save_mapping_cache(st.session_state.mapping_cache)

                st.success("Procesamiento completado con éxito")
                
                # --- DASHBOARD DE RESULTADOS ---
                if recon:
                    st.markdown("#### Estado de Conciliación")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Total Facturado", f"${recon.get('total_facturado', 0):,.2f}")
                    m2.metric("Total Asignado", f"${recon.get('total_asignado', 0):,.2f}")
                    diff = recon.get('diferencia', 0)
                    m3.metric("Diferencia", f"${diff:,.2f}", delta=f"{diff:,.2f}", delta_color="inverse")
                    m4.metric("Match Rate", f"{recon.get('match_rate', 0):.1f}%")

                summary_report = final_summary[["BU", "Transport", "Arg. Var $", "%PCT"]].copy()
                summary_report = summary_report.sort_values(by=["Transport", "BU"]).reset_index(drop=True)

                pct_tab, mon_tab = build_executive_tables(final_summary)
                
                t1, t2 = st.tabs(["Reporte Ejecutivo", "Auditoría Detallada"])
                with t1:
                    st.write("### Summary final por BU y Tipo de Viaje")
                    # Usamos el Styler de Pandas y lo pasamos al bypass
                    styled_summary = summary_report.style.format({"%PCT": "{:.1%}", "Arg. Var $": "${:,.0f}"})
                    render_safe_table(styled_summary)
                    
                    st.write("**Asignación Porcentual (%PCT)**")
                    render_safe_table(pct_tab.style.format("{:.1%}"))
                    
                    st.write("**Distribución Monetaria ($)**")
                    render_safe_table(mon_tab.style.format("${:,.0f}"))

                
                    
                    try:
                        buffer = BytesIO()
                        # Si implementaste el generate_auditable_excel que definimos antes
                        if hasattr(engine, 'generate_auditable_excel'):
                            # El engine se encarga de crear el Excel con fórmulas y guardarlo en el buffer
                            engine.generate_auditable_excel(unified, df_costs, buffer)
                        else:
                            # Fallback de seguridad si aún no lo implementas
                            st.warning("Modo Auditable no encontrado en el Engine. Descargando datos planos.")
                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                pct_tab.to_excel(writer, sheet_name='Reporte', startrow=0)
                                mon_tab.to_excel(writer, sheet_name='Reporte', startrow=len(pct_tab) + 2)
                                final_summary.to_excel(writer, sheet_name='Base_Datos', index=False)
                        
                        st.download_button(
                            label="Descargar Excel Auditable (Con Fórmulas)",
                            data=buffer.getvalue(),
                            file_name="Consolidado_Logistico_IN_OUT_Auditoria.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                    except Exception as excel_err:
                        st.error(f"Fallo en la generación del Excel Auditable: {str(excel_err)}")

                with t2:
                    st.write("Datos procesados y cruzados:")
                    render_safe_table(final_summary)
                    
                    st.markdown("---")
                    notes = final_summary[final_summary['Note'].notna()] if 'Note' in final_summary.columns else pd.DataFrame()
                    if not notes.empty:
                        st.write("### Filas con imputaciones o reglas aplicadas")
                        render_safe_table(notes)
                    else:
                        st.info("No se detectaron imputaciones automáticas en el reporte final.")

        # --- CAPTURA DE ERRORES FAIL FAST ---
        except ValueError as ve:
            st.error("###Validación de Datos Fallida")
            st.warning(str(ve))
            st.info("Por favor, verifica el mapeo de columnas y los archivos subidos.")
        except Exception as e:
            st.error("###Error crítico en el sistema")
            st.exception(e)
            logger.exception("Traceback completo:")

if __name__ == "__main__":
    main()