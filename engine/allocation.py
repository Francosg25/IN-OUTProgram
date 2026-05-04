import pandas as pd
import numpy as np
import logging
import re
from typing import Literal, Any

# 1. IMPORTAR EL MOTOR DE REGLAS
from engine.rules import BusinessRulesEngine

logger = logging.getLogger(__name__)

class CostAllocationEngine:
    def __init__(self, allocation_type: Literal['weight', 'full_container'] = 'weight'):
        self.allocation_type = allocation_type
        # Tarifas de respaldo en caso de que Finanzas no reporte el costo
        self.fallback_costs = {
            'sea': 2500.0,
            'land': 1200.0,
            'outbound': 2000.0
        }

    def _clean_ref(self, ref: Any) -> str:
        """
        Limpia referencias y contenedores usando Regex o limpieza alfanumérica estándar.
        """
        if pd.isna(ref): return None
        
        # Patrón específico si aplica
        pattern = r"\w?\w-J-\d{4}LI\d{2}"
        match = re.search(pattern, str(ref).upper())
        
        if match:
            return match.group(0)
        # Fallback genérico: quitar espacios y caracteres raros
        return re.sub(r'[^A-Z0-9]', '', str(ref).upper())

    def calculate_outbound(self, df_transactions: pd.DataFrame, df_costs: pd.DataFrame) -> pd.DataFrame:
        df = df_transactions.copy()
        costs = df_costs.copy()

        # Limpieza inicial
        if 'price' in df.columns:
            df['price'] = pd.to_numeric(df['price'], errors='coerce').fillna(0.0)
        else:
            df['price'] = 0.0

        df['gross_weight'] = pd.to_numeric(df.get('gross_weight', 0), errors='coerce').fillna(0.0)

        # 2. APLICAR REGLAS DE NEGOCIO (Capex y Misceláneos)
        df = BusinessRulesEngine.apply_classification_rules(df)

        # Validación de columnas obligatorias
        req_cols = ['reference', 'bu', 'gross_weight', 'transport_type']
        missing_cols = [col for col in req_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Faltan columnas canónicas: {missing_cols}")

        try:
            # 3. DEFINICIÓN DE AGRUPACIÓN LOGÍSTICA (Contenedor vs Guía Terrestre)
            if 'container_number' in df.columns:
                # Si hay contenedor (SEA), limpiamos y agrupamos por contenedor. 
                # Si una fila no trae contenedor, cae a la referencia (guía).
                df['group_key'] = df['container_number'].fillna(df['reference']).apply(self._clean_ref)
            else:
                # Logística terrestre o aérea (Land/Outbound)
                df['group_key'] = df['reference'].apply(self._clean_ref)
                
            # Si tanto contenedor como referencia fallan, usamos la BU como último recurso
            df['group_key'] = df['group_key'].fillna(df['bu'])

            # 4. PONDERACIÓN MATEMÁTICA
            df['note'] = ''
            if self.allocation_type == 'weight':
                # Suma total del peso por Contenedor/Guía
                total_weight_per_group = df.groupby('group_key')['gross_weight'].transform('sum')
                
                # Proporción = Peso Individual / Peso Total
                df['Proportion'] = df['gross_weight'] / total_weight_per_group.replace(0, np.nan)
                
                # Manejo de nulos (Equivalente al IFERROR de tu Excel)
                df['Proportion'] = df['Proportion'].fillna(1.0)
            else:
                items_per_group = df.groupby('group_key')['group_key'].transform('count')
                df['Proportion'] = 1.0 / items_per_group

            # 5. DETECCIÓN FINANCIERA (Archivo de Costos)
            ref_col_costs = [c for c in costs.columns if 'ref' in c.lower() or 'bu' in c.lower()][0]
            cost_cols = [c for c in costs.columns if 'cost' in c.lower() or 'amount' in c.lower() or 'usd' in c.lower()]
            cost_col = cost_cols[0] if cost_cols else costs.columns[-1]

            costs_subset = costs[[ref_col_costs, cost_col]].rename(columns={ref_col_costs: 'financial_key', cost_col: 'Total Cost'})
            
            # Limpiamos las llaves financieras
            costs_subset['financial_key'] = costs_subset['financial_key'].apply(lambda x: self._clean_ref(x) if self._clean_ref(x) else str(x).strip().upper())
            costs_subset = costs_subset.groupby('financial_key', as_index=False)['Total Cost'].sum()

            # 6. JOIN Relacional Operaciones <-> Finanzas
            df = df.merge(costs_subset, left_on='group_key', right_on='financial_key', how='left')
            
            # 7. SISTEMA DE RESPALDO (Fallback Costs)
            df['fixed_cost'] = 0.0 # Columna de rastreo para auditoría
            
            def apply_fallback(row):
                if pd.isna(row['Total Cost']) or row['Total Cost'] == 0:
                    trans_type = str(row['transport_type']).strip().lower()
                    # Si no hay costo financiero, inyectamos la tarifa pactada (Ej. 2500 para Sea)
                    fallback = self.fallback_costs.get(trans_type, 0.0)
                    if fallback > 0:
                        return fallback, fallback # Retornamos Costo Final, y Registro de Fallback
                return row['Total Cost'], 0.0

            # Aplicamos la función a dos columnas (Total Cost y fixed_cost)
            df[['Total Cost', 'fixed_cost']] = df.apply(apply_fallback, axis=1, result_type='expand')
            df['Total Cost'] = pd.to_numeric(df['Total Cost'], errors='coerce').fillna(0)
            
            # Anotaciones para auditoría
            df.loc[df['fixed_cost'] > 0, 'note'] += f'Costo estándar aplicado; '
            df.loc[df['gross_weight'] <= 0, 'note'] += 'Peso imputado; '

            # 8. CÁLCULO FINAL DE EXPENSAS
            df['Calc_Exp'] = df['Total Cost'] * df['Proportion']

            # 9. AGRUPACIÓN PARTICIONADA (Resumen Ejecutivo)
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
            
            # Renombrado de columnas para reporte
            rename_map = {'bu': 'BU', 'transport_type': 'Transport', 'Calc_Exp': 'Arg. Var $'}
            if 'method' in df.columns:
                rename_map['method'] = 'Method'
            
            summary.rename(columns=rename_map, inplace=True)
            
            # %PCT Particionado (basado en el transporte)
            total_exp_per_transport = summary.groupby('Transport')['Arg. Var $'].transform('sum')
            summary['%PCT'] = summary['Arg. Var $'] / total_exp_per_transport.replace(0, 1)
            
            # --- 10. MÉTRICAS DE CONCILIACIÓN ---
            total_default_cost = df['fixed_cost'].drop_duplicates().sum() # Sumamos los respaldos aplicados
            total_input_cost = costs_subset['Total Cost'].sum() + total_default_cost
            total_allocated_cost = summary['Arg. Var $'].sum()
            diff = total_input_cost - total_allocated_cost
            
            match_rate = 0.0
            if total_input_cost:
                match_rate = max(0.0, 100.0 * (1 - abs(diff) / total_input_cost))
            
            logger.info(f"AUDITORÍA: Facturado (Real+Fallback): {total_input_cost} | Asignado: {total_allocated_cost} | Diff: {diff}")
            
            # Diccionario para UI
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