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

# Importación del motor matemático
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



import pandas as pd
import streamlit as st

import pandas as pd
import streamlit as st

def smart_read_excel(file_obj) -> pd.DataFrame:
    """
    Escáner de densidad heurística con consolidación automática de columnas divididas (Ej. Múltiples Gross Weights).
    """
    try:
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
            
        xls = pd.ExcelFile(file_obj, engine='openpyxl')
        df_temp = pd.read_excel(xls, header=None, nrows=30)
        
        # Diccionario expandido para asegurar que atrapa la tabla real en cualquier formato
        keywords = ['REF', 'REFERENCE', 'GUIA', 'WAYBILL', 'AWB', 'TRACKING', 
                    'BU', 'BUSINESS UNIT', 'UNIDAD', 
                    'PESO', 'WEIGHT', 'KGS', 'LBS', 'GROSS', 
                    'PART', 'ITEM', 'QTY', 'PIECES', 'BULTOS', 
                    'CONTAINER', 'CONTENEDOR', 'CNTR', 'EQUIPO', 
                    'COST', 'PRICE', 'AMOUNT', 'USD', 'FIX COST']
        
        best_idx = 0
        max_score = -1
        
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
            if total_score > max_score:
                max_score = total_score
                best_idx = idx
                
        df = pd.read_excel(xls, header=best_idx)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        # Buscar todas las columnas que hablen de peso bruto
        peso_cols = [col for col in df.columns if 'GROSS WEIGHT' in col or 'PESO BRUTO' in col]
        
        if len(peso_cols) > 0:
            # Sumar horizontalmente todas las columnas de peso encontradas
            df['CONSOLIDATED_GROSS_WEIGHT'] = pd.to_numeric(df[peso_cols[0]], errors='coerce').fillna(0)
            for col in peso_cols[1:]:
                df['CONSOLIDATED_GROSS_WEIGHT'] += pd.to_numeric(df[col], errors='coerce').fillna(0)
                
            # Ocultamos las columnas sueltas para no confundir al usuario
            df.drop(columns=peso_cols, inplace=True)
        
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
        'bu': ['bu', 'business unit', 'unidad de negocio', 'unidad', 'area', 'division'],
        'gross_weight': ['gross_weight', 'peso', 'weight', 'kg', 'kgs', 'peso bruto'],
        'price': ['price', 'precio', 'cost', 'valor', 'amount', 'monto'],
        'part_number': ['part_number', 'part number', 'numero de parte', 'item', 'item code']
    },
    'Land': {
        'reference': ['reference', 'ref', 'doc', 'documento', 'orden', 'guia', 'guía'],
        'container_number': ['container', 'contenedor', 'truck', 'camion', 'vehiculo'],
        'bu': ['bu', 'business unit', 'unidad de negocio', 'unidad', 'area', 'division'],
        'gross_weight': ['gross_weight', 'peso', 'weight', 'kg', 'kgs', 'peso bruto'],
        'price': ['price', 'precio', 'cost', 'valor', 'amount', 'monto'],
        'part_number': ['part_number', 'part number', 'numero de parte', 'item', 'item code']
    },
    'Outbound': {
        'reference': ['reference', 'ref', 'shipment', 'export', 'tracking', 'documento', 'awb'],
        'container_number': ['container', 'contenedor', 'cntr', 'box', 'equipo'],
        'bu': ['bu', 'business unit', 'unidad de negocio', 'unidad', 'empresa', 'division'],
        'gross_weight': ['gross_weight', 'peso', 'weight', 'kg', 'kgs', 'peso bruto'],
        'price': ['price', 'precio', 'cost', 'valor', 'amount', 'monto'],
        'part_number': ['part_number', 'part number', 'numero de parte', 'item', 'item code']
    }
}

GENERIC_COLUMN_ALIASES = {
    'reference': ['reference', 'ref', 'guia', 'guía', 'documento', 'doc', 'tracking', 'awb', 'waybill'],
    'container_number': ['container', 'contenedor', 'cntr', 'box', 'equipo'],
    'bu': ['bu', 'business unit', 'unidad de negocio', 'unidad', 'area', 'division'],
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
    """
    Orquestador de Datos: Soporta mapeo manual dinámico por archivo.
    """
    def __init__(self, allocation_engine):
        self.allocation_engine = allocation_engine
        self.canon_map = ['reference', 'bu', 'gross_weight']

    def standardize_source_manual(self, df: pd.DataFrame, label: str, mapping: Dict[str, str]) -> pd.DataFrame:
        """Estandariza un DF usando un mapeo manual proporcionado por el usuario."""
        try:
            # Verificar que las columnas mapeadas existen en el DataFrame
            for canon, excel_col in mapping.items():
                if excel_col and excel_col not in df.columns:
                    st.error(f"Columna '{excel_col}' no encontrada en el archivo {label}. Columnas disponibles: {list(df.columns)}")
                    return pd.DataFrame()
            
            # Invertimos el mapeo para renombrar: {col_del_excel: nombre_canonico}
            rename_dict = {v: k for k, v in mapping.items() if v}
            
            # Solo tomamos las columnas mapeadas
            df_filtered = df[list(rename_dict.keys())].copy()
            df_filtered.rename(columns=rename_dict, inplace=True)
            
            df_filtered['transport_type'] = label
            df_filtered['reference'] = df_filtered['reference'].astype(str).str.strip()
            df_filtered['bu'] = df_filtered['bu'].astype(str).str.strip().str.upper()

            # Manejar gross_weight: imputar 1 si no está mapeada o si faltan valores.
            if 'gross_weight' not in df_filtered.columns:
                df_filtered['gross_weight'] = 1.0
            else:
                df_filtered['gross_weight'] = pd.to_numeric(df_filtered['gross_weight'], errors='coerce').fillna(1.0)
            
            # Manejar price opcional
            if 'price' not in df_filtered.columns:
                df_filtered['price'] = 0.0
            else:
                df_filtered['price'] = pd.to_numeric(df_filtered['price'], errors='coerce').fillna(0.0)

            df_filtered['note'] = ''
            missing_weight = df_filtered['gross_weight'] <= 0
            if missing_weight.any():
                df_filtered.loc[missing_weight, 'gross_weight'] = 1.0
                df_filtered.loc[missing_weight, 'note'] += 'Peso imputado; '

            missing_price = df_filtered['price'] <= 0
            if missing_price.any():
                df_filtered.loc[missing_price, 'note'] += 'Precio ausente; '

            # Asegurar que reference y bu existen
            if 'reference' not in df_filtered.columns or 'bu' not in df_filtered.columns:
                st.error(f"Faltan columnas obligatorias (reference o bu) en {label}.")
                return pd.DataFrame()

            return df_filtered
        except Exception as e:
            st.error(f"Error al procesar {label}: {e}")
            return pd.DataFrame()

    def run_pipeline(self, processed_dfs: List[pd.DataFrame], df_costs: pd.DataFrame) -> pd.DataFrame:
        """Ejecuta la unificación de DFs y el cálculo de prorrateo."""
        if not processed_dfs:
            raise ValueError("No hay datos procesados para ejecutar el prorrateo.")
        
        unified = pd.concat(processed_dfs, ignore_index=True)
        return self.allocation_engine.calculate_outbound(unified, df_costs)

def build_executive_tables(flat_summary: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Transforma el resumen plano en vistas ejecutivas pivotadas (Formato Johnson Electric)."""
    if flat_summary.empty:
        return pd.DataFrame(), pd.DataFrame()

    df_c = flat_summary.copy()
    df_c['Transport'] = df_c['Transport'].fillna('UNKNOWN').astype(str)

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
    
    # 2. CARGA DE BASES OPERATIVAS
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
    if "mapping_cache" not in st.session_state:
        st.session_state.mapping_cache = mapping_cache

   
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

                if m_ref and m_bu:
                    mapping = {"reference": m_ref, "bu": m_bu}
                    if m_container: mapping["container_number"] = m_container
                    if m_w: mapping["gross_weight"] = m_w
                    if m_price: mapping["price"] = m_price
                    if m_part: mapping["part_number"] = m_part # Guarda el mapeo del no. de parte
                    st.session_state.mapping_cache[label] = mapping
                    return df, mapping
                else:
                    st.warning("Selecciona al menos Referencia y BU para habilitar esta fuente.")
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
            engine = CostAllocationEngine()
            orchestrator = LogisticsOrchestrator(engine)
            
            to_process = []
            if sea_data is not None and sea_map: to_process.append(orchestrator.standardize_source_manual(sea_data, "Sea", sea_map))
            if land_data is not None and land_map: to_process.append(orchestrator.standardize_source_manual(land_data, "Land", land_map))
            if out_data is not None and out_map: to_process.append(orchestrator.standardize_source_manual(out_data, "Outbound", out_map))
            
            to_process = [d for d in to_process if not d.empty]
            
            if not to_process:
                st.warning("No hay fuentes operativas configuradas correctamente.")
                return

            if isinstance(cost_file, BytesIO) or hasattr(cost_file, 'name') and cost_file.name.lower().endswith('.csv'):
                df_costs = pd.read_csv(cost_file, encoding='latin1')
            else:
                df_costs = smart_read_excel(cost_file)
            
            with st.spinner("Procesando datos y aplicando reglas de negocio..."):
                unified = pd.concat(to_process, ignore_index=True)
                final_summary = orchestrator.run_pipeline(to_process, df_costs)
                recon = final_summary.attrs.get('reconciliation', {})
                save_mapping_cache(st.session_state.mapping_cache)

                st.success("Procesamiento completado con éxito")
                
                # --- DASHBOARD DE RESULTADOS ---
                if recon:
                    st.markdown("####Estado de Conciliación")
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
                    st.dataframe(summary_report.style.format({"%PCT": "{:.1%}", "Arg. Var $": "${:,.0f}"}), use_container_width=True)
                    st.write("**Asignación Porcentual (%PCT)**")
                    st.dataframe(pct_tab.style.format("{:.1%}"), use_container_width=True)
                    st.write("**Distribución Monetaria ($)**")
                    st.dataframe(mon_tab.style.format("${:,.0f}"), use_container_width=True)
                    
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        # Escribir tabla de porcentajes arriba
                        pct_tab.to_excel(writer, sheet_name='Reporte', startrow=0)
                        # Escribir tabla de dinero abajo (dejando un espacio de 2 filas)
                        mon_tab.to_excel(writer, sheet_name='Reporte', startrow=len(pct_tab) + 2)
                        # Pestaña de respaldo
                        final_summary.to_excel(writer, sheet_name='Base de Datos (Auditoría)', index=False)
                    
                    st.download_button(
                        label="📥 Descargar Reporte en Excel",
                        data=buffer.getvalue(),
                        file_name="Consolidado_Logistico_IN_OUT.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                
                with t2:
                    st.write("Datos procesados y cruzados:")
                    st.dataframe(final_summary, use_container_width=True)
                    st.markdown("---")
                    notes = final_summary[final_summary['Note'].notna()] if 'Note' in final_summary.columns else pd.DataFrame()
                    if not notes.empty:
                        st.write("### Filas con imputaciones o reglas aplicadas")
                        st.dataframe(notes, use_container_width=True)
                    else:
                        st.info("No se detectaron imputaciones automáticas en el reporte final.")
                    st.markdown("### Datos originales cargados con notas de auditoría")
                    audit_data = unified.copy()
                    if 'note' in audit_data.columns:
                        audit_data = audit_data[audit_data['note'].astype(bool)]
                    if not audit_data.empty:
                        st.dataframe(audit_data, use_container_width=True)
                    else:
                        st.info("No hay registros con notas de auditoría en las cargas originales.")

        except Exception as e:
            st.error(f"Fallo en el pipeline: {e}")
            logger.exception(e)

if __name__ == "__main__":
    main()
