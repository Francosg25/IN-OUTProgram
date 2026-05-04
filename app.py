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



def smart_read_excel(file_obj) -> pd.DataFrame:
    """
    Lee un Excel en memoria saltando títulos, evitando errores de puntero 
    y casteando tipos de celda de forma segura.
    """
    try:
        # 1. Regresar puntero
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
            
        xls = pd.ExcelFile(file_obj, engine='openpyxl')
        df_temp = pd.read_excel(xls, header=None, nrows=30)
        header_idx = 0
        
        keywords = ['REF', 'BU', 'PESO', 'WEIGHT', 'PART', 'ITEM', 'UNIT', 'GUIA', 'TRACKING', 'METHOD', 'CUSTOMER', 'WAYBILL']
        
        # 2. Escáner Seguro Celda por Celda
        for idx, row in df_temp.iterrows():
            row_list = row.tolist()
            contains_keyword = False
            
            for cell in row_list:
                # Solo evaluamos la celda si no es nula/vacía
                if pd.notna(cell):
                    # Forzamos la conversión a texto puro (Safe Cast)
                    cell_str = str(cell).upper().strip()
                    if any(kw in cell_str for kw in keywords):
                        contains_keyword = True
                        break # Encontramos la fila de encabezados
            
            if contains_keyword:
                header_idx = idx
                break
                
        # 3. Lectura final desde la fila detectada
        df = pd.read_excel(xls, header=header_idx)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
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


def suggest_mapping(df: pd.DataFrame, label: str) -> dict:
    cols = df.columns.tolist()
    aliases = TRAVEL_COLUMN_ALIASES.get(label, GENERIC_COLUMN_ALIASES)
    return {
        'reference': find_best_column(cols, aliases.get('reference', GENERIC_COLUMN_ALIASES['reference'])),
        'container_number': find_best_column(cols, aliases.get('container_number', GENERIC_COLUMN_ALIASES['container_number'])),
        'bu': find_best_column(cols, aliases.get('bu', GENERIC_COLUMN_ALIASES['bu'])),
        'gross_weight': find_best_column(cols, aliases.get('gross_weight', GENERIC_COLUMN_ALIASES['gross_weight'])),
        'price': find_best_column(cols, aliases.get('price', GENERIC_COLUMN_ALIASES['price'])),
        'part_number': find_best_column(cols, aliases.get('part_number', GENERIC_COLUMN_ALIASES['part_number']))
    }

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
    """Transforma el resumen plano en vistas ejecutivas pivotadas."""
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

    # Tabla Monetaria
    mon_pivot = df_c.pivot_table(
        index='Transport', 
        columns='BU', 
        values='Arg. Var $', 
        aggfunc='sum', 
        fill_value=0
    )
    mon_pivot.insert(0, 'Total Arg. Var $', mon_pivot.sum(axis=1))

    return pct_pivot, mon_pivot

def main():
    st.set_page_config(page_title="Logistics Auditor PRO", layout="wide", page_icon="🚢")
    st.title("Sistema logistico ")
    st.markdown("""
    Esta herramienta permite cargar reportes logísticos con cualquier estructura y asignar sus costos 
    mediante un mapeo manual de columnas.
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
                    
                    st.download_button(
                        "Descargar Resumen Final (CSV)",
                        summary_report.to_csv(index=False),
                        "auditoria_logistica_summary.csv",
                        "text/csv"
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
