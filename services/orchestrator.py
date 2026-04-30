import pandas as pd
import logging
import re
from typing import Dict, Any

logger = logging.getLogger(__name__)

class LogisticsPipelineOrchestrator:
    def __init__(self, allocation_engine):
        self.allocation_engine = allocation_engine
        
        self.aliases = {
            'reference': ['REFERENCE', 'CONTAINER NUMBER', 'WAYBILL NUMBER', 'REFERENCIA'],
            'bu': ['BU', 'OU', 'BUSINESS UNIT', 'UNIDAD DE NEGOCIO'],
            'gross_weight': ['GROSS WEIGHT (KGS)', 'TOTAL GROSS WEIGHT', 'PESO BRUTO (KGS)', 'WEIGHT'],
            'inbound': ['INBOUND'],
            'method': ['METHOD'],
            'part_number': ['NO DE PARTE', 'PART NUMBER']
        }
        
        # Lista Blanca (Whitelist) de BUs permitidas.
        # Ajusta esta lista con los 5 o6 códigos reales que operan en la planta.
        self.valid_bus = {'M01', 'M02', 'M19', 'M23', 'M45'} 
    
    def clean_bu_code(self, raw_bu: Any) -> str:
        """
        Sanitizador puro: Elimina basura, normaliza y valida contra la lista blanca.
        """
        if pd.isna(raw_bu):
            return 'DEFAULT_BU'
            
        # 1. Cast a string, mayúsculas, y eliminación de espacios en los extremos
        clean_str = str(raw_bu).strip().upper()
        
        # 2. Regex: Eliminar cualquier caracter que no sea alfanumérico (ej. \xa0, guiones)
        clean_str = re.sub(r'[^A-Z0-9]', '', clean_str)
        
        return clean_str if clean_str else 'DEFAULT_BU'

    def _extract_and_standardize(self, file_path: Any, source_type: str, transport_label: str) -> pd.DataFrame:
        try:
            df_raw = pd.read_excel(file_path) 
            df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()

            df_std = pd.DataFrame()
            mandatory_cols = ['reference', 'bu', 'gross_weight']
            
            for canon_col, possible_names in self.aliases.items():
                found = False
                for name in possible_names:
                    matched_cols = [c for c in df_raw.columns if name in c]
                    if matched_cols:
                        df_std[canon_col] = df_raw[matched_cols[0]]
                        found = True
                        break
                
                if not found:
                    if canon_col in mandatory_cols:
                        if canon_col == 'bu':
                            df_std['bu'] = 'DEFAULT_BU'
                        else:
                            logger.error(f"Falta columna crítica '{canon_col}' en {transport_label}")
                            return pd.DataFrame()
                    else:
                        # Columnas opcionales (como inbound, method, part_number)
                        df_std[canon_col] = None

            # --- NUEVA CAPA DE SANITIZACIÓN ---
            # Aplicamos la función de limpieza vectorizada a toda la columna
            if 'bu' in df_std.columns:
                df_std['bu'] = df_std['bu'].apply(self.clean_bu_code)

            df_std['transport_type'] = transport_label
            df_std['gross_weight'] = pd.to_numeric(df_std['gross_weight'], errors='coerce').fillna(0)

            return df_std

        except Exception as e:
            logger.error(f"Error estandarizando fuente {transport_label}: {e}")
            return pd.DataFrame()

    def run_pipeline(self, files: Dict[str, Dict[str, str]], df_costs: pd.DataFrame) -> pd.DataFrame:
        standardized_dfs = []
        
        for file_obj, config in files.items():
            df = self._extract_and_standardize(
                file_path=file_obj, 
                source_type=config['type'], 
                transport_label=config['label']
            )
            if not df.empty:
                standardized_dfs.append(df)

        if not standardized_dfs:
            raise ValueError("No se pudieron extraer datos válidos de ningún archivo. Verifica los nombres de las columnas.")

        # Unificación
        df_unified = pd.concat(standardized_dfs, ignore_index=True)
        
        # Procesamiento final (Prorrateo)
        summary_df = self.allocation_engine.calculate_outbound(df_unified, df_costs)
        
        return summary_df