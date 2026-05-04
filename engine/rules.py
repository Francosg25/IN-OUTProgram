import pandas as pd

class BusinessRulesEngine:
    """
    Motor de reglas de negocio para clasificar Unidades de Negocio (BUs)
    basado en la semántica de los Números de Parte.
    """
    @staticmethod
    def apply_classification_rules(df: pd.DataFrame) -> pd.DataFrame:
        if 'part_number' not in df.columns or 'bu' not in df.columns:
            return df
            
        df_rules = df.copy()
        
        def classify_part(row):
            part = str(row['part_number']).upper().strip()
            current_bu = str(row['bu']).strip()
            
            # REGLA CAPEX
            if 'CAPEX' in part or (part.isalpha() and len(part) > 3 and 'TAPA' not in part and 'CAJA' not in part):
                return 'Capex'
                
            # REGLA MISCELÁNEOS
            misc_keywords = ['TAPA', 'CAJA', 'BASE', 'CHAROLA', 'PALLET', 'CARTON', 'PLASTICA', 'PLASTICO', 'WOOD']
            if any(kw in part for kw in misc_keywords) or part.count(' ') >= 3:
                return 'Miscelaneus'
                
            return current_bu
            
        df_rules['bu'] = df_rules.apply(classify_part, axis=1)
        return df_rules