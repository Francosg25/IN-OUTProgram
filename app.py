import streamlit as st
import pandas as pd
import re
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
    Lee un Excel escaneando las primeras 30 filas para saltar 
    títulos y celdas combinadas, encontrando el encabezado real.
    """
    try:
        # Leemos las primeras 30 filas sin asignar encabezados
        df_temp = smart_read_excel(file_obj)
        header_idx = 0
        
        # Palabras clave que suelen estar en un encabezado logístico
        keywords = ['REF', 'BU', 'PESO', 'WEIGHT', 'PART', 'ITEM', 'UNIT', 'GUIA', 'TRACKING', 'METHOD', 'CUSTOMER']
        
        for idx, row in df_temp.iterrows():
            row_str = row.astype(str).str.upper().tolist()
            # Si la fila contiene alguna de las palabras clave, esa es nuestra fila de encabezados
            if any(any(kw in cell for kw in keywords) for cell in row_str if cell != 'NAN'):
                header_idx = idx
                break
                
        # Ahora leemos el archivo completo empezando desde la fila correcta
        df = smart_read_excel(file_obj)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        return df
    except Exception as e:
        import streamlit as st
        st.error(f"Error al leer el archivo: {e}")
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


def normalize_col_name(name: str) -> str:
    return re.sub(r'[\W_]+', '', str(name).strip().lower())



def find_best_column(columns_list: list, keywords: list) -> str:
    """
    Escáner heurístico avanzado. Prioriza coincidencias exactas, 
    luego palabras completas (Regex) y finalmente coincidencias parciales.
    """
    if not columns_list:
        return None

    # 1. MATCH EXACTO (Prioridad Máxima)
    # Ejemplo: Si la columna se llama exactamente "BU"
    for col in columns_list:
        col_clean = str(col).strip().upper()
        if col_clean in [k.upper() for k in keywords]:
            return col

    # 2. MATCH DE PALABRA COMPLETA (Word Boundaries)
    # Ejemplo: "BU DESTINO" hace match con "BU", pero "BULTOS" es ignorado.
    for col in columns_list:
        col_clean = str(col).strip().upper()
        for kw in keywords:
            kw_upper = kw.upper()
            # \b indica un límite de palabra (espacios, guiones, inicio/fin)
            if re.search(rf'\b{kw_upper}\b', col_clean):
                return col

    # 3. MATCH PARCIAL (Fallback de rescate)
    # Ejemplo: "PESO_BRUTO" hace match con "PESO"
    for col in columns_list:
        col_clean = str(col).strip().upper()
        for kw in keywords:
            kw_upper = kw.upper()
            if kw_upper in col_clean:
                # Regla de Excepción Crítica para Logística
                if kw_upper == 'BU' and 'BULTO' in col_clean:
                    continue # Ignoramos "BULTOS" para que no se asigne a "BU"
                return col
                
    return None

def suggest_mapping(df: pd.DataFrame) -> dict:
    """
    Genera sugerencias de mapeo usando el escáner heurístico.
    """
    cols = df.columns.tolist()
    return {
        'reference': find_best_column(cols, ['reference', 'ref', 'guia', 'guía', 'documento', 'doc', 'tracking', 'awb', 'waybill']),
        'bu': find_best_column(cols, ['bu', 'unidad de negocio', 'unidad', 'business unit', 'businessunit', 'area', 'division']),
        'gross_weight': find_best_column(cols, ['gross_weight', 'peso', 'weight', 'kg', 'kgs']),
        'price': find_best_column(cols, ['price', 'precio', 'cost', 'valor', 'amount', 'monto']),
        # ¡Añadimos la columna del número de parte para que funcione el motor de reglas (Capex/Misc)!
        'part_number': find_best_column(cols, ['part_number', 'part number', 'numero de parte', 'no. parte', 'item', 'item code'])
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
            suggested = suggest_mapping(df)
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
                default_bu = saved.get('bu') or suggested.get('bu')
                default_w = saved.get('gross_weight') or suggested.get('gross_weight')
                default_price = saved.get('price') or suggested.get('price')
                default_part = saved.get('part_number') or suggested.get('part_number')

                m_ref = select_with_default(f"Referencia / Guía ({label})", f"sel_ref_{label}", default_ref)
                m_bu = select_with_default(f"Unidad de Negocio ({label})", f"sel_bu_{label}", default_bu)
                m_w = select_with_default(f"Peso Bruto ({label}) - Opcional", f"sel_w_{label}", default_w)
                m_price = select_with_default(f"Precio / Valor ({label}) - Opcional", f"sel_price_{label}", default_price)
                m_part = select_with_default(f"Número de Parte ({label}) - Opcional", f"sel_part_{label}", default_part)

                if m_ref and m_bu:
                    mapping = {"reference": m_ref, "bu": m_bu}
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
                df_costs = pd.read_csv(cost_file)
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
