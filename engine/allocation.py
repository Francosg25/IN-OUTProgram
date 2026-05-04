import pandas as pd
import logging
import re
from typing import Literal, Any

logger = logging.getLogger(__name__)

class CostAllocationEngine:
    def __init__(self, allocation_type: Literal['weight', 'full_container'] = 'weight'):
        self.allocation_type = allocation_type

    def _clean_ref(self, ref: Any) -> str:
        """
        Extrae la referencia usando el patrón: \w?\w-J-\d{4}LI\d{2}
        """
        if pd.isna(ref): return None
        
        # Patrón flexible: letras/números opcionales + -J- + 4 dígitos + LI + 2 dígitos
        pattern = r"\w?\w-J-\d{4}LI\d{2}"
        match = re.search(pattern, str(ref).upper())
        
        if match:
            return match.group(0)
        return None

    def calculate_outbound(self, df_transactions: pd.DataFrame, df_costs: pd.DataFrame) -> pd.DataFrame:
        df = df_transactions.copy()
        costs = df_costs.copy()

        # Price is optional but gross_weight must exist for the engine logic.
        if 'price' in df.columns:
            df['price'] = pd.to_numeric(df['price'], errors='coerce').fillna(0.0)
        else:
            df['price'] = 0.0

        df['gross_weight'] = pd.to_numeric(df['gross_weight'], errors='coerce').fillna(0.0)

        req_cols = ['reference', 'bu', 'gross_weight', 'transport_type']
        
        missing_cols = [col for col in req_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Faltan columnas canónicas: {missing_cols}")

        try:
            # 1. Limpieza de Referencias con Regex y Fallback a BU
            df['clean_key'] = df['reference'].apply(self._clean_ref)
            
            # Si la limpieza falló (None), usamos la BU como llave de cruce
            df['clean_key'] = df['clean_key'].fillna(df['bu'])

            # 2. Aplicar validadores de costos fijos y notas para cada transporte
            df['fixed_cost'] = 0.0
            df['note'] = ''
            df.loc[df['gross_weight'] <= 0, 'note'] += 'Peso imputado; '
            df.loc[df['price'] <= 0, 'note'] += 'Precio ausente; '

            transport_norm = df['transport_type'].astype(str).str.strip().str.lower()
            sea_missing_cost = (transport_norm == 'sea') & (df['gross_weight'] <= 0) & (df['price'] <= 0)
            land_missing_cost = (transport_norm == 'land') & (df['gross_weight'] <= 0) & (df['price'] <= 0)
            outbound_missing_cost = (transport_norm == 'outbound') & (df['gross_weight'] <= 0) & (df['price'] <= 0)

            df.loc[sea_missing_cost, 'fixed_cost'] = 2500.0
            df.loc[sea_missing_cost, 'note'] += 'Costo fijo SEA aplicado; '
            df.loc[land_missing_cost, 'fixed_cost'] = 1800.0
            df.loc[land_missing_cost, 'note'] += 'Costo fijo LAND aplicado; '
            df.loc[outbound_missing_cost, 'fixed_cost'] = 2000.0
            df.loc[outbound_missing_cost, 'note'] += 'Costo fijo OUTBOUND aplicado; '

            # 3. Cálculo de Proporción (Usamos la llave limpia para agrupar)
            if self.allocation_type == 'weight':
                allocation_weight = df['gross_weight'].where(~df['fixed_cost'].gt(0), 0).replace(0, 1)
                total_weight_per_ref = allocation_weight.groupby(df['clean_key']).transform('sum')
                df['Proportion'] = allocation_weight / total_weight_per_ref.replace(0, 1)
            elif self.allocation_type == 'full_container':
                items_per_ref = df.groupby('clean_key')['clean_key'].transform('count')
                df['Proportion'] = 1.0 / items_per_ref

            # 4. Detección Heurística de Costos en el archivo financiero
            ref_col_costs = [c for c in costs.columns if 'ref' in c.lower() or 'bu' in c.lower()][0]
            cost_cols = [c for c in costs.columns if 'cost' in c.lower() or 'amount' in c.lower() or 'usd' in c.lower()]
            cost_col = cost_cols[0] if cost_cols else costs.columns[-1]

            costs_subset = costs[[ref_col_costs, cost_col]].rename(columns={ref_col_costs: 'clean_key', cost_col: 'Total Cost'})
            
            # Limpiamos también las llaves en el archivo de costos para que coincidan
            costs_subset['clean_key'] = costs_subset['clean_key'].apply(lambda x: self._clean_ref(x) if self._clean_ref(x) else str(x).strip().upper())
            
            costs_subset = costs_subset.groupby('clean_key', as_index=False)['Total Cost'].sum()

            # JOIN Relacional usando la llave limpia (Referencia o BU)
            df = df.merge(costs_subset, on='clean_key', how='left')
            df['Total Cost'] = pd.to_numeric(df['Total Cost'], errors='coerce').fillna(0)

            # Gasto Calculado
            df['Calc_Exp'] = df['Total Cost'] * df['Proportion']
            df.loc[df['fixed_cost'] > 0, 'Calc_Exp'] = df.loc[df['fixed_cost'] > 0, 'fixed_cost']

            #Agrupación Particionada
            group_cols = ['transport_type', 'bu']
            if 'method' in df.columns:
                group_cols.insert(1, 'method')
                # Llenamos nulos en method para evitar que groupby los ignore
                df['method'] = df['method'].fillna('N/A')

            summary = df.groupby(group_cols, as_index=False).agg({
                'Calc_Exp': 'sum',
                'fixed_cost': 'sum',
                'note': lambda texts: '; '.join(sorted({t for t in texts if t}))
            })
            summary['Note'] = summary['note'].replace('', None)
            summary.drop(columns=['note', 'fixed_cost'], inplace=True)
            
            # Renombrado de columnas para reporte
            rename_map = {'bu': 'BU', 'transport_type': 'Transport', 'Calc_Exp': 'Arg. Var $'}
            if 'method' in df.columns:
                rename_map['method'] = 'Method'
            
            summary.rename(columns=rename_map, inplace=True)
            
            # %PCT Particionado (basado en el transporte)
            total_exp_per_transport = summary.groupby('Transport')['Arg. Var $'].transform('sum')
            summary['%PCT'] = summary['Arg. Var $'] / total_exp_per_transport.replace(0, 1)
            
            # --- NUEVA VALIDACIÓN DE CONCILIACIÓN ---
            total_default_cost = df['fixed_cost'].sum()
            total_input_cost = costs_subset['Total Cost'].sum() + total_default_cost
            total_allocated_cost = summary['Arg. Var $'].sum()
            diff = total_input_cost - total_allocated_cost
            match_rate = 0.0
            if total_input_cost:
                match_rate = max(0.0, 100.0 * (1 - abs(diff) / total_input_cost))
            
            logger.info(f"AUDITORÍA: Total Facturado: {total_input_cost} | Total Asignado: {total_allocated_cost} | Diff: {diff}")
            
            # Guardamos la métrica en el dataframe para que la UI pueda leerla
            summary.attrs['reconciliation'] = {
                'total_facturado': total_input_cost,
                'total_asignado': total_allocated_cost,
                'diferencia': diff,
                'match_rate': match_rate
            }
            
            sort_cols = ['Transport', 'BU']
            if 'Method' in summary.columns:
                sort_cols.insert(1, 'Method')
                
            summary = summary.sort_values(by=sort_cols).reset_index(drop=True)
            
            return summary

        except Exception as e:
            logger.error(f"Fallo en motor de prorrateo: {str(e)}")
            raise