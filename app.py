import streamlit as st
import pandas as pd
import re
import logging
from typing import Dict, Any, Literal
from io import BytesIO


# --- CONFIGURACIÓN DE LOGS ---
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- MOTOR DE ASIGNACIÓN (Engine) ---
class CostAllocationEngine:
    def __init__(self, allocation_type: Literal['weight', 'full_container'] = 'weight'):
        self.allocation_type = allocation_type

    def _clean_ref(self, ref: Any) -> str:
        if pd.isna(ref): return None
        pattern = r"\w?\w-J-\d{4}LI\d{2}"
        match = re.search(pattern, str(ref).upper())
        return match.group(0) if match else None

    def calculate_outbound(self, df_transactions: pd.DataFrame, df_costs: pd.DataFrame) -> pd.DataFrame:
        df = df_transactions.copy()
        costs = df_costs.copy()

        df['clean_key'] = df['reference'].apply(self._clean_ref)
        df['clean_key'] = df['clean_key'].fillna(df['bu'])

        if self.allocation_type == 'weight':
            total_weight_per_ref = df.groupby('clean_key')['gross_weight'].transform('sum')
            df['Proportion'] = df['gross_weight'] / total_weight_per_ref.replace(0, 1)
        else:
            items_per_ref = df.groupby('clean_key')['clean_key'].transform('count')
            df['Proportion'] = 1.0 / items_per_ref

        ref_col_costs = [c for c in costs.columns if 'ref' in c.lower() or 'bu' in c.lower()][0]
        cost_col = [c for c in costs.columns if 'cost' in c.lower() or 'amount' in c.lower() or 'usd' in c.lower()][0]

        costs_subset = costs[[ref_col_costs, cost_col]].rename(columns={ref_col_costs: 'clean_key', cost_col: 'Total Cost'})
        costs_subset['clean_key'] = costs_subset['clean_key'].apply(lambda x: self._clean_ref(x) if self._clean_ref(x) else str(x).strip().upper())
        costs_subset = costs_subset.groupby('clean_key', as_index=False)['Total Cost'].sum()

        df = df.merge(costs_subset, on='clean_key', how='left')
        df['Total Cost'] = pd.to_numeric(df['Total Cost'], errors='coerce').fillna(0)
        df['Calc_Exp'] = df['Total Cost'] * df['Proportion']

        group_cols = ['transport_type', 'bu']
        if 'method' in df.columns:
            group_cols.insert(1, 'method')
            df['method'] = df['method'].fillna('N/A')

        summary = df.groupby(group_cols, as_index=False)['Calc_Exp'].sum()
        rename_map = {'bu': 'BU', 'transport_type': 'Transport', 'Calc_Exp': 'Arg. Var $'}
        if 'method' in df.columns: rename_map['method'] = 'Method'
        summary.rename(columns=rename_map, inplace=True)
        
        total_exp_per_transport = summary.groupby('Transport')['Arg. Var $'].transform('sum')
        summary['%PCT'] = summary['Arg. Var $'] / total_exp_per_transport.replace(0, 1)
        
        summary.attrs['reconciliation'] = {
            'total_facturado': costs_subset['Total Cost'].sum(),
            'total_asignado': summary['Arg. Var $'].sum(),
            'diferencia': costs_subset['Total Cost'].sum() - summary['Arg. Var $'].sum(),
            'match_rate': (len(df[df['Total Cost'] > 0]) / len(df)) * 100 if len(df) > 0 else 0
        }
        return summary

# --- ORQUESTADOR (Service) ---
class LogisticsPipelineOrchestrator:
    def __init__(self, allocation_engine):
        self.allocation_engine = allocation_engine
        self.aliases = {
            'reference': ['REFERENCE', 'CONTAINER NUMBER', 'WAYBILL NUMBER', 'REFERENCIA', 'CONTAINER'],
            'bu': ['BU', 'OU', 'BUSINESS UNIT', 'UNIDAD DE NEGOCIO'],
            'gross_weight': ['GROSS WEIGHT (KGS)', 'TOTAL GROSS WEIGHT', 'PESO BRUTO (KGS)', 'WEIGHT'],
            'inbound': ['INBOUND'],
            'method': ['METHOD'],
            'part_number': ['NO DE PARTE', 'PART NUMBER', 'ITEM CODE']
        }

    def clean_bu_code(self, raw_bu: Any) -> str:
        if pd.isna(raw_bu): return 'DEFAULT_BU'
        clean_str = str(raw_bu).strip().upper()
        clean_str = re.sub(r'[^A-Z0-9]', '', clean_str)
        match = re.match(r'^M(\d{1,4})$', clean_str)
        if match and 0 <= int(match.group(1)) <= 1000: return clean_str
        return clean_str if clean_str else 'DEFAULT_BU'

    def _extract_and_standardize(self, file_obj, transport_label: str) -> pd.DataFrame:
        try:
            df_raw = pd.read_excel(file_obj)
            df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()
            df_std = pd.DataFrame()
            mandatory_cols = ['reference', 'bu', 'gross_weight']
            
            for canon_col, possible_names in self.aliases.items():
                for name in possible_names:
                    matched_cols = [c for c in df_raw.columns if name in c]
                    if matched_cols:
                        df_std[canon_col] = df_raw[matched_cols[0]]
                        break
                if canon_col not in df_std.columns:
                    if canon_col in mandatory_cols:
                        if canon_col == 'bu': df_std['bu'] = 'DEFAULT_BU'
                        else: return pd.DataFrame()
                    else: df_std[canon_col] = None

            if 'bu' in df_std.columns:
                df_std['bu'] = df_std['bu'].apply(self.clean_bu_code)

            def apply_business_rules(row):
                part_val = str(row['part_number']).upper() if row['part_number'] else ""
                if "CAPEX" in part_val or (part_val and part_val != 'NONE' and not any(char.isdigit() for char in part_val)):
                    return "Capex"
                num_spaces = part_val.count(' ')
                if num_spaces > 3 or (len(part_val) > 25 and num_spaces >= 2):
                    return 'Miscelaneus'
                for kw in ['TAPA PLASTICA', 'CHAROLA', 'BASE PLASTICA', 'TAPA', 'BASE']:
                    if kw in part_val: return 'Miscelaneus'
                return row['bu']
            
            if 'part_number' in df_std.columns and 'bu' in df_std.columns:
                df_std['bu'] = df_std.apply(apply_business_rules, axis=1)

            df_std['transport_type'] = transport_label
            df_std['gross_weight'] = pd.to_numeric(df_std['gross_weight'], errors='coerce').fillna(0)
            return df_std
        except Exception as e:
            logger.error(f"Error: {e}")
            return pd.DataFrame()

    def run_pipeline(self, files: Dict[Any, str], df_costs: pd.DataFrame) -> pd.DataFrame:
        standardized_dfs = []
        for file_obj, label in files.items():
            df = self._extract_and_standardize(file_obj, label)
            if not df.empty: standardized_dfs.append(df)
        if not standardized_dfs: raise ValueError("Archivos inválidos.")
        return self.allocation_engine.calculate_outbound(pd.concat(standardized_dfs, ignore_index=True), df_costs)



def build_executive_tables(flat_summary: pd.DataFrame):
    """
    Transforma el summary plano (Normalizado) en tablas pivote (Wide Format) 
    para la vista ejecutiva.
    """
    if flat_summary.empty:
        return pd.DataFrame(), pd.DataFrame()

    # Orden lógico de los transportes en logística
    transport_order = ['Sea', 'Land', 'Outbound']

    #TABLA DE PORCENTAJES (%PCT) ---
    pct_table = flat_summary.pivot_table(
        index='Transport', 
        columns='BU', 
        values='%PCT', 
        aggfunc='sum',
        fill_value=0
    )
    pct_table.index.name = 'Type'
    
    pct_table.index = pct_table.index + ' %PCT'
    
    pct_order = [t + ' %PCT' for t in transport_order if t + ' %PCT' in pct_table.index]
    pct_table = pct_table.reindex(pct_order)


    #TABLA DE COSTOS (Arg. Var $) ---
    money_table = flat_summary.pivot_table(
        index='Transport', 
        columns='BU', 
        values='Arg. Var $', 
        aggfunc='sum',
        fill_value=0
    )
    money_table.index.name = 'Viewer'
    
    money_table.insert(0, 'Arg. Var $', money_table.sum(axis=1))

    mon_order = [t for t in transport_order if t in money_table.index]
    money_table = money_table.reindex(mon_order)

    return pct_table, money_table


def main():
    st.set_page_config(page_title="In-Out Logistics", layout="wide")
    st.title("Logistics Data Orchestrator")
    st.markdown("---")

    with st.sidebar:
        st.header("Facturacion")
        cost_file = st.file_uploader("Consolidado de Facturación (Costos)", type=["xlsx"])
        
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.subheader("Sea / China")
        sea_file = st.file_uploader("Reporte Marítimo", type=["xlsx"], key="sea")
    with col_b:
        st.subheader("Land / Impos")
        land_file = st.file_uploader("Impos Land", type=["xlsx"], key="land")
    with col_c:
        st.subheader("Outbound / Expos")
        outbound_file = st.file_uploader("Reporte Expos", type=["xlsx", "xls"], key="out")

    st.markdown("---")

    if st.button("Ejecutar Procesamiento", use_container_width=True):
        if not cost_file:
            st.warning("Por favor, sube el archivo de Consolidado de Facturación.")
            return

        try:
            with st.spinner("🔄 Procesando datos y reglas de negocio..."):
                engine = CostAllocationEngine(allocation_type='weight')
                orchestrator = LogisticsPipelineOrchestrator(engine)
                
                df_costs = pd.read_excel(cost_file)
                files_to_process = {}
                if sea_file: files_to_process[sea_file] = "Sea"
                if land_file: files_to_process[land_file] = "Land"
                if outbound_file: files_to_process[outbound_file] = "Outbound"

                if not files_to_process:
                    st.error("Sube al menos un archivo de operación.")
                    return

                # Única ejecución del pipeline
                final_summary = orchestrator.run_pipeline(files_to_process, df_costs)
                recon = final_summary.attrs.get('reconciliation', {})
                
                st.success("Procesamiento completado con éxito.")
                
                if recon:
                    st.markdown("###Estado de Conciliación")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Total Facturado", f"${recon['total_facturado']:,.2f}")
                    c2.metric("Total Asignado", f"${recon['total_asignado']:,.2f}")
                    diff = recon['diferencia']
                    c3.metric("Diferencia", f"${diff:,.2f}", delta=f"{diff:,.2f}", delta_color="normal" if abs(diff) < 1 else "inverse")
                    c4.metric("Match Rate", f"{recon['match_rate']:.1f}%")

                pct_table, money_table = build_executive_tables(final_summary)
                
                tab_summary, tab_details = st.tabs(["📊 Resumen Ejecutivo", "📝 Detalles y Auditoría"])
                
                with tab_summary:
                    st.markdown("### Tabla de Porcentajes de Asignación")
                    st.dataframe(pct_table.style.format("{:.0%}"))
                    
                    st.markdown("### Tabla de Cargos y Costos")
                    st.dataframe(money_table.style.format("${:,.0f}"))
                    
                    st.download_button(
                        "Descargar Base Normalizada (CSV)",
                        data=final_summary.to_csv(index=False),
                        file_name="summary_logistica_master.csv",
                        mime="text/csv"
                    )
                
                with tab_details:
                    st.dataframe(final_summary, use_container_width=True)

        except Exception as e:
            st.error(f"Error durante la ejecución: {str(e)}")
            logger.exception("Detalle del error:")

if __name__ == "__main__":
    main()