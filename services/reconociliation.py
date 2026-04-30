import pandas as pd
import logging
import re

logger = logging.getLogger(__name__)

class ReconciliationService:
    """
    Servicio encargado de cruzar facturas/pagos con las operaciones 
    de transporte (Sea/Land/Outbound).
    """

    def clean_reference(self, ref: str) -> str:
        """Limpia referencias para mejorar el cruce (Ej: 'Cont. 123' -> '123')"""
        if pd.isna(ref): return ""
        # Elimina espacios, guiones y convierte a mayúsculas
        return re.sub(r'[^a-zA-Z0-9]', '', str(ref)).upper()

    def merge_costs(self, df_ops: pd.DataFrame, df_finance: pd.DataFrame) -> pd.DataFrame:
        """
        Cruza las operaciones logísticas con el estado de cuenta/consolidado.
        """
        try:
            # 1. Normalización de llaves de cruce
            # En el consolidado la columna suele ser 'REFERENCIA' o 'REFERENCIA '
            df_ops['ref_clean'] = df_ops['reference'].apply(self.clean_reference)
            
            # Buscamos la columna de referencia en el archivo financiero
            finance_ref_col = [c for c in df_finance.columns if 'REFERENCIA' in c.upper()][0]
            df_finance['ref_clean'] = df_finance[finance_ref_col].apply(self.clean_reference)

            # 2. Identificar columna de costo (USD)
            cost_col = [c for c in df_finance.columns if 'COST' in c.upper() or 'AMOUNT' in c.upper()][0]

            # 3. Join (Izquierda: mantenemos todas las operaciones)
            # Esto permite ver qué operaciones NO tienen factura asociada aún.
            merged = pd.merge(
                df_ops, 
                df_finance[['ref_clean', cost_col]], 
                on='ref_clean', 
                how='left'
            )

            # Renombrar para el motor de prorrateo
            merged.rename(columns={cost_col: 'Total Cost'}, inplace=True)
            merged['Total Cost'] = merged['Total Cost'].fillna(0)

            # 4. Reporte de 'Missing Matches' para el usuario de Streamlit
            missing = merged[merged['Total Cost'] == 0]['reference'].unique()
            if len(missing) > 0:
                logger.warning(f"Referencias sin costo encontrado: {len(missing)}")

            return merged

        except Exception as e:
            logger.error(f"Error en reconciliación: {e}")
            raise