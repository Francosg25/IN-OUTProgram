import streamlit as st
import pandas as pd
import re
import logging
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
            # Invertimos el mapeo para renombrar: {col_del_excel: nombre_canonico}
            rename_dict = {v: k for k, v in mapping.items() if v}
            
            # Solo tomamos las columnas mapeadas
            df_filtered = df[list(rename_dict.keys())].copy()
            df_filtered.rename(columns=rename_dict, inplace=True)
            
            df_filtered['transport_type'] = label
            df_filtered['gross_weight'] = pd.to_numeric(df_filtered.get('gross_weight', 0), errors='coerce').fillna(0)
            
            if 'reference' not in df_filtered.columns:
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

    # 3. PANEL DE MAPEO DINÁMICO
    st.markdown("Mapeo de Columnas")
    st.caption("Selecciona qué columna de tu Excel corresponde a cada dato necesario.")
    
    def create_mapping_ui(file, label):
        if file:
            df = pd.read_excel(file)
            cols = [""] + df.columns.tolist()
            with st.expander(f"Configurar columnas de {label}", expanded=True):
                col1, col2, col3 = st.columns(3)
                m_ref = col1.selectbox(f"Referencia / Guía ({label})", cols, key=f"sel_ref_{label}")
                m_bu = col2.selectbox(f"Unidad de Negocio ({label})", cols, key=f"sel_bu_{label}")
                m_w = col3.selectbox(f"Peso Bruto ({label})", cols, key=f"sel_w_{label}")
                
                if m_ref and m_bu:
                    mapping = {"reference": m_ref, "bu": m_bu, "gross_weight": m_w}
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

            df_costs = pd.read_excel(cost_file)
            
            with st.spinner("Procesando datos y aplicando reglas de negocio..."):
                final_summary = orchestrator.run_pipeline(to_process, df_costs)
                recon = final_summary.attrs.get('reconciliation', {})

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

                pct_tab, mon_tab = build_executive_tables(final_summary)
                
                t1, t2 = st.tabs(["📊 Reporte Ejecutivo", "🔍 Auditoría Detallada"])
                with t1:
                    st.write("**Asignación Porcentual (%PCT)**")
                    st.dataframe(pct_tab.style.format("{:.1%}"), use_container_width=True)
                    st.write("**Distribución Monetaria ($)**")
                    st.dataframe(mon_tab.style.format("${:,.0f}"), use_container_width=True)
                    
                    st.download_button(
                        "📥 Descargar Reporte Completo (CSV)",
                        final_summary.to_csv(index=False),
                        "auditoria_logistica.csv",
                        "text/csv"
                    )
                
                with t2:
                    st.write("Datos procesados y cruzados:")
                    st.dataframe(final_summary, use_container_width=True)

        except Exception as e:
            st.error(f"Fallo en el pipeline: {e}")
            logger.exception(e)

if __name__ == "__main__":
    main()
