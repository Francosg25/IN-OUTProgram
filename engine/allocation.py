import pandas as pd
import numpy as np
import logging
import re
from typing import Literal, Any

logger = logging.getLogger(__name__)

class BusinessRulesEngine:
    """Clasificador Semántico de Números de Parte"""
    @staticmethod
    def apply_classification_rules(df: pd.DataFrame) -> pd.DataFrame:
        if 'part_number' not in df.columns or 'bu' not in df.columns:
            return df
            
        df_rules = df.copy()
        
        def classify_part(row):
            part = str(row['part_number']).upper().strip()
            current_bu = str(row['bu']).strip()
            
            if 'CAPEX' in part or (part.isalpha() and len(part) > 3 and 'TAPA' not in part and 'CAJA' not in part):
                return 'Capex'
                
            misc_keywords = ['TAPA', 'CAJA', 'BASE', 'CHAROLA', 'PALLET', 'CARTON', 'PLASTICA', 'PLASTICO', 'WOOD']
            if any(kw in part for kw in misc_keywords) or part.count(' ') >= 3:
                return 'Miscelaneus'
                
            return current_bu
            
        df_rules['bu'] = df_rules.apply(classify_part, axis=1)
        return df_rules

class CostAllocationEngine:
    def __init__(self, allocation_type: Literal['weight', 'full_container'] = 'weight'):
        self.allocation_type = allocation_type
        # Tarifas maestras de Johnson Electric
        self.fallback_costs = {
            'sea': 2500.0,
            'land': 1200.0,  # Corregido al valor BZ1 de tu Excel
            'outbound': 2000.0
        }

    def _clean_ref(self, ref: Any) -> str:
        if pd.isna(ref): return None
        pattern = r"\w?\w-J-\d{4}LI\d{2}"
        match = re.search(pattern, str(ref).upper())
        if match: return match.group(0)
        return re.sub(r'[^A-Z0-9]', '', str(ref).upper())

    def calculate_outbound(self, df_transactions: pd.DataFrame, df_costs: pd.DataFrame) -> pd.DataFrame:
        df = df_transactions.copy()
        costs = df_costs.copy()

        if 'price' in df.columns:
            df['price'] = pd.to_numeric(df['price'], errors='coerce').fillna(0.0)
        else:
            df['price'] = 0.0

        df['gross_weight'] = pd.to_numeric(df.get('gross_weight', 0), errors='coerce').fillna(0.0)

        df = BusinessRulesEngine.apply_classification_rules(df)

        # --- NUEVA REGLA: CUSTOMER COMO BU PARA OUTBOUND ---
        # Buscamos si existe alguna columna de Customer en el archivo
        customer_col = next((c for c in df.columns if 'customer' in c.lower()), None)
        
        if 'bu' not in df.columns and customer_col:
            # Si no hay BU, el Customer asume su lugar
            df['bu'] = df[customer_col]
        elif 'bu' in df.columns and customer_col:
            # Si existe BU pero viene vacía en algunas filas, la rellenamos con el Customer
            df['bu'] = df['bu'].fillna(df[customer_col])
            
        # Salvavidas: si una fila queda huérfana de BU y Customer, le ponemos 'N/A'
        if 'bu' in df.columns:
            df['bu'] = df['bu'].fillna('N/A')
        else:
            df['bu'] = 'N/A'
        

        req_cols = ['reference', 'bu', 'gross_weight', 'transport_type']
        missing_cols = [col for col in req_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Faltan columnas canónicas: {missing_cols}")

        try:
            # 2. Agrupación Dinámica (Contenedor vs Referencia)
            if 'container_number' in df.columns:
                df['group_key'] = df['container_number'].fillna(df['reference']).apply(self._clean_ref)
            else:
                df['group_key'] = df['reference'].apply(self._clean_ref)
                
            df['group_key'] = df['group_key'].fillna(df['bu'])

            # 3. Ponderación de Peso (Recreación de fórmula Excel =BO9/SUMIFS)
            df['note'] = ''
            if self.allocation_type == 'weight':
                total_weight_per_group = df.groupby('group_key')['gross_weight'].transform('sum')
                df['Proportion'] = df['gross_weight'] / total_weight_per_group.replace(0, np.nan)
                df['Proportion'] = df['Proportion'].fillna(1.0)
            else:
                items_per_group = df.groupby('group_key')['group_key'].transform('count')
                df['Proportion'] = 1.0 / items_per_group

            # 4. Cruce Financiero
            ref_col_costs = [c for c in costs.columns if 'ref' in c.lower() or 'bu' in c.lower()][0]
            cost_cols = [c for c in costs.columns if 'cost' in c.lower() or 'amount' in c.lower() or 'usd' in c.lower()]
            cost_col = cost_cols[0] if cost_cols else costs.columns[-1]

            costs_subset = costs[[ref_col_costs, cost_col]].rename(columns={ref_col_costs: 'financial_key', cost_col: 'Total Cost'})
            costs_subset['financial_key'] = costs_subset['financial_key'].apply(lambda x: self._clean_ref(x) if self._clean_ref(x) else str(x).strip().upper())
            costs_subset = costs_subset.groupby('financial_key', as_index=False)['Total Cost'].sum()

            df = df.merge(costs_subset, left_on='group_key', right_on='financial_key', how='left')
            
            # 5. Sistema Fallback Tarifario
            df['fixed_cost'] = 0.0
            
            def apply_fallback(row):
                if pd.isna(row['Total Cost']) or row['Total Cost'] == 0:
                    trans_type = str(row['transport_type']).strip().lower()
                    fallback = self.fallback_costs.get(trans_type, 0.0)
                    if fallback > 0:
                        return fallback, fallback
                return row['Total Cost'], 0.0

            df[['Total Cost', 'fixed_cost']] = df.apply(apply_fallback, axis=1, result_type='expand')
            df['Total Cost'] = pd.to_numeric(df['Total Cost'], errors='coerce').fillna(0)
            
            df.loc[df['fixed_cost'] > 0, 'note'] += f'Costo estándar aplicado; '
            df.loc[df['gross_weight'] <= 0, 'note'] += 'Peso imputado; '

            # 6. Cálculo Final
            df['Calc_Exp'] = df['Total Cost'] * df['Proportion']

            # 7. Resumen Ejecutivo
            group_cols = ['transport_type', 'bu']
            if 'method' in df.columns:
                group_cols.insert(1, 'method')
                df['method'] = df['method'].fillna('N/A')

            summary = df.groupby(group_cols, as_index=False).agg({
                'Calc_Exp': 'sum',
                'fixed_cost': 'sum',
                'note': lambda texts: '; '.join(sorted({t for t in texts if t}))
            })
            summary['Note'] = summary['note'].replace('', None)
            summary.drop(columns=['note', 'fixed_cost'], inplace=True)
            
            rename_map = {'bu': 'BU', 'transport_type': 'Transport', 'Calc_Exp': 'Arg. Var $'}
            if 'method' in df.columns: rename_map['method'] = 'Method'
            summary.rename(columns=rename_map, inplace=True)
            
            total_exp_per_transport = summary.groupby('Transport')['Arg. Var $'].transform('sum')
            summary['%PCT'] = summary['Arg. Var $'] / total_exp_per_transport.replace(0, 1)
            
            # Métrica UI
            total_default_cost = df['fixed_cost'].drop_duplicates().sum()
            total_input_cost = costs_subset['Total Cost'].sum() + total_default_cost
            total_allocated_cost = summary['Arg. Var $'].sum()
            diff = total_input_cost - total_allocated_cost
            match_rate = max(0.0, 100.0 * (1 - abs(diff) / total_input_cost)) if total_input_cost else 0.0
            
            summary.attrs['reconciliation'] = {
                'total_facturado': total_input_cost,
                'total_asignado': total_allocated_cost,
                'diferencia': diff,
                'match_rate': match_rate
            }
            
            sort_cols = ['Transport', 'BU']
            if 'Method' in summary.columns: sort_cols.insert(1, 'Method')
            return summary.sort_values(by=sort_cols).reset_index(drop=True)

        except Exception as e:
            logger.error(f"Fallo en motor de prorrateo: {str(e)}")
            raise
        
    
        

