import pandas as pd
import logging
from typing import Literal

logger = logging.getLogger(__name__)

class CostAllocationEngine:
    """
    Motor central para transformar transacciones logísticas y datos de facturación
    en el Summary final por BU (Business Unit).
    """
    
    def __init__(self, allocation_type: Literal['weight', 'full_container'] = 'weight'):
        # Permite alternar entre prorratear por peso o cobrar el contenedor entero
        self.allocation_type = allocation_type

    def calculate_outbound(self, df_transactions: pd.DataFrame, df_costs: pd.DataFrame) -> pd.DataFrame:
        """
        Reemplaza la lógica de SUMIFS y XLOOKUP del Excel original.
        """
        # Trabajamos con copias para mantener inmutabilidad
        df = df_transactions.copy()
        costs = df_costs.copy()

        # 1. Asegurar nombres de columnas (Asumiendo que el Extractor previo ya las normalizó)
        req_cols = ['Reference', 'BU', 'Gross Weight']
        if not all(col in df.columns for col in req_cols):
            raise ValueError(f"Faltan columnas requeridas en transacciones: {req_cols}")

        try:
            # 2. Cálculo de Proporción (Reemplaza =BO9/SUMIFS($BO:$BO,$BI:$BI,BI9))
            if self.allocation_type == 'weight':
                # .transform('sum') calcula el total por Referencia sin colapsar el DataFrame
                total_weight_per_ref = df.groupby('Reference')['Gross Weight'].transform('sum')
                # Evitar división por cero
                df['Proportion'] = df['Gross Weight'] / total_weight_per_ref.replace(0, 1)
            
            elif self.allocation_type == 'full_container':
                # Regla de negocio: Si usaron el contenedor, absorben el costo completo de su porción de la factura
                # independientemente de si el peso fue de 1kg o 10,000kg.
                items_per_ref = df.groupby('Reference')['Reference'].transform('count')
                df['Proportion'] = 1.0 / items_per_ref

            # 3. Cruce de Costos (Reemplaza =XLOOKUP(...))
            # Hacemos un LEFT JOIN usando la 'Reference' (ej. Waybill o Container Number)
            df = df.merge(costs[['Reference', 'Total Cost']], on='Reference', how='left')
            
            # Llenar facturas no encontradas con 0 para evitar propagación de NaNs
            df['Total Cost'] = df['Total Cost'].fillna(0)

            # 4. Gasto Calculado (Calc_Exp)
            df['Calc_Exp'] = df['Total Cost'] * df['Proportion']

            # 5. Generación del Summary por BU
            summary = df.groupby('BU', as_index=False)['Calc_Exp'].sum()
            total_exp = summary['Calc_Exp'].sum()
            
            # Calcular porcentaje final por BU (Reemplaza la tabla final de D3 a H3)
            summary['%PCT'] = summary['Calc_Exp'] / total_exp if total_exp > 0 else 0
            
            return summary

        except Exception as e:
            logger.error(f"Fallo en motor de prorrateo: {str(e)}")
            raise