import pandas as pd
import numpy as np
import logging
import re
from typing import Literal, Any
from io import BytesIO

logger = logging.getLogger(__name__)

class BusinessRulesEngine:
    """Clasificador Semántico de Números de Parte y Descripciones"""
    
    @staticmethod
    def apply_classification_rules(df: pd.DataFrame) -> pd.DataFrame:
        if 'part_number' not in df.columns or 'bu' not in df.columns:
            return df
            
        df_rules = df.copy()
        
        # Diccionarios de detección (Johnson Electric Standards)
        misc_keywords = [
            'TAPA', 'CAJA', 'BASE', 'CHAROLA', 'PALLET', 'CARTON', 'PLASTICA', 'PLASTICO', 'WOOD', 
            'CABLE', 'LENS', 'KIT', 'MODULE', 'ADAPTER', 'DISPLAY', 'MONITOR', 'SCREEN'
        ]
        
        capex_keywords = [
            'CAPEX', 'TOOLING', 'MACHINE', 'TESTER', 'POWER SUPPLY', 'ROBOT', 'STATION', 
            'CONVEYOR', 'MOTOR', 'PUMP'
        ]
        
        def classify_part(row):
            part = str(row['part_number']).upper().strip()
            current_bu = str(row['bu']).strip().upper()
            
            if part in ['NAN', 'NONE', '', 'NULL']:
                return current_bu

            if any(kw in part for kw in capex_keywords):
                return 'CAPEX' 
                
            part_only_letters = part.isalpha()
            if part_only_letters and len(part) > 3 and not any(kw in part for kw in misc_keywords):
                return 'CAPEX' 

            if any(kw in part for kw in misc_keywords):
                return 'MISCELANEUS' 
                
            num_spaces = part.count(' ')
            if num_spaces >= 3 or len(part) > 30:
                return 'MISCELANEUS' 
                
            return current_bu
            
        df_rules['bu'] = df_rules.apply(classify_part, axis=1)
        return df_rules

class CostAllocationEngine:
    def __init__(self, allocation_type: str = 'weight', fallback_costs: dict = None, default_cost: float = 0.0):
        self.output_sheet = "Audit_Report"
        self.allocation_type = allocation_type if allocation_type in ('weight', 'equal') else 'weight'
        self.fallback_costs = fallback_costs or {}
        self.default_cost = default_cost

    def _find_cost_reference_column(self, df: pd.DataFrame) -> str:
        """Identifica la columna de referencia en el archivo de costos."""
        candidates = [
            c for c in df.columns
            if any(tok in c.lower() for tok in ['ref', 'reference', 'waybill', 'awb', 'tracking', 'guia', 'shipment', 'documento'])
        ]
        if candidates:
            return candidates[0]
        candidates = [c for c in df.columns if 'bu' in c.lower()]
        if candidates:
            return candidates[0]
        raise ValueError("No se pudo identificar la columna de referencia en el archivo de costos.")

    def _find_cost_value_column(self, df: pd.DataFrame) -> str:
        """Identifica la columna de valor/costo en el archivo de costos."""
        candidates = [
            c for c in df.columns
            if any(tok in c.lower() for tok in ['fix cost', 'fix', 'amount', 'cost', 'usd', 'valor', 'importe', 'price', 'monto'])
        ]
        if candidates:
            return candidates[0]
        return df.columns[-1]

    def _clean_ref(self, raw_ref: Any) -> str:
        """
        LIMPIEZA CRÍTICA: Remueve sufijos para asegurar match de costos.
        Ej: 'FG-R-2180LE25.M46-M45-2' -> 'FG-R-2180LE25'
        """
        val = str(raw_ref).strip().upper()
        if val in ['NAN', 'NONE', '', 'NULL']:
            return 'UNKNOWN'
        
        # 1. Cortar en el primer punto (Quita .M46, .PFA, etc)
        if '.' in val:
            val = val.split('.')[0]
            
        # 2. Extraer solo el patrón de guía estándar si existe
        match = re.search(r'^(FG-R-\w+)', val)
        if match:
            return match.group(1)
            
        return val
    
    def calculate_in_memory(self, df_transactions: pd.DataFrame, df_costs: pd.DataFrame = None) -> pd.DataFrame:
        """
        Motor de cálculo en Pandas para poblar el Dashboard interactivo de Streamlit.
        Replica la lógica de Excel (SUMIFS y XLOOKUP) en memoria.
        """
        df = df_transactions.copy()
        
        # 1. Limpieza de Llave (Garantiza el match)
        df['group_key'] = df['reference'].apply(self._clean_ref)
        
        # 2. Equivalente a SUMIFS de Excel: Proporción por Peso
        total_weight = df.groupby('group_key')['gross_weight'].transform('sum')
        # Evitamos división por cero asignando 0 si no hay peso
        df['%PCT'] = np.where(total_weight > 0, df['gross_weight'] / total_weight, 0.0)
        
        # Inicializamos la columna monetaria
        df['Arg. Var $'] = 0.0
        
        # 3. Equivalente a XLOOKUP de Excel: Cruce de Costos (Outbound)
        if df_costs is not None and not df_costs.empty:
            try:
                # Detectar columnas de la tabla de costos
                ref_col = [c for c in df_costs.columns if 'REF' in c.upper()][0]
                val_col = [c for c in df_costs.columns if any(k in c.upper() for k in ['FIX', 'COST', 'USD'])][0]
                
                # Consolidar Costos Facturados
                costs_clean = df_costs[[ref_col, val_col]].copy()
                costs_clean.columns = ['Cost_Ref', 'Amount']
                costs_clean['Cost_Ref'] = costs_clean['Cost_Ref'].apply(self._clean_ref)
                costs_clean = costs_clean.groupby('Cost_Ref', as_index=False)['Amount'].sum()
                
                # Diccionario de cruce rápido (Hash Map)
                cost_map = dict(zip(costs_clean['Cost_Ref'], costs_clean['Amount']))
                
                # Filtrar transacciones Outbound
                mask_outbound = df['transport_type'].str.upper() == 'OUTBOUND'
                
                # Mapear costo total facturado a cada transacción
                df.loc[mask_outbound, 'Total_Invoice_Cost'] = df.loc[mask_outbound, 'group_key'].map(cost_map).fillna(0)
                
                # Aplicar prorrateo financiero: Costo * Porcentaje
                df.loc[mask_outbound, 'Arg. Var $'] = df.loc[mask_outbound, '%PCT'] * df.loc[mask_outbound, 'Total_Invoice_Cost']
            
            except Exception as e:
                logger.warning(f"Advertencia en cálculo en memoria: {e}")
                
        return df

    def generate_auditable_excel(self, df_transactions: pd.DataFrame, df_costs: pd.DataFrame, buffer: BytesIO):
        """
        Enrutador que genera el reporte con inyección de fórmulas dinámicas.
        """
        try:
            writer = pd.ExcelWriter(buffer, engine='xlsxwriter')
            
            # Separar flujos
            df_inbound = df_transactions[df_transactions['transport_type'].isin(['Land', 'Sea'])].copy()
            df_outbound = df_transactions[df_transactions['transport_type'].str.upper() == 'OUTBOUND'].copy()
            
            if not df_inbound.empty:
                self._inject_inbound_formulas(writer, df_inbound, df_costs)
                
            if not df_outbound.empty:
                self._inject_outbound_formulas(writer, df_outbound, df_costs)
                
            writer.close()
            return buffer
        except Exception as e:
            logger.error(f"Error en generación de Excel: {e}")
            raise

    def _inject_outbound_formulas(self, writer: pd.ExcelWriter, df: pd.DataFrame, df_costs: pd.DataFrame):
        """
        Estrategia Outbound: Inyecta XLOOKUP contra matriz de costos limpiada.
        """
        # 1. Aplicar Reglas de Negocio
        df = BusinessRulesEngine.apply_classification_rules(df)
        df['group_key'] = df['reference'].apply(self._clean_ref)
        
        # 2. Preparar Matriz de Costos (Deduplicada por Llave Limpia)
        ref_col_cost = [c for c in df_costs.columns if 'REF' in c.upper()][0]
        val_col_cost = [c for c in df_costs.columns if any(k in c.upper() for k in ['FIX', 'COST', 'USD'])][0]
        
        costs_clean = df_costs[[ref_col_cost, val_col_cost]].copy()
        costs_clean.columns = ['Cost_Ref', 'Amount']
        costs_clean['Cost_Ref'] = costs_clean['Cost_Ref'].apply(self._clean_ref)
        # Sumamos costos si la misma guía aparece varias veces en factura
        costs_clean = costs_clean.groupby('Cost_Ref', as_index=False)['Amount'].sum()

        # 3. Escribir tablas en Excel
        # Col A-B: Diccionario de Costos | Col D-I: Datos Operativos
        costs_clean.to_excel(writer, sheet_name='Outbound_Auditable', index=False, startcol=0)
        
        cols_to_exp = ['reference', 'bu', 'part_number', 'gross_weight', 'transport_type', 'group_key']
        df[cols_to_exp].to_excel(writer, sheet_name='Outbound_Auditable', index=False, startcol=3)
        
        workbook = writer.book
        worksheet = writer.sheets['Outbound_Auditable']
        
        # Formatos
        money = workbook.add_format({'num_format': '$#,##0.00'})
        pct = workbook.add_format({'num_format': '0.000%'})
        header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        
        col_weight = 'G' 
        col_key = 'I'   
        
        worksheet.write('K1', '% Proportion', header)
        worksheet.write('L1', 'Calc_Amount', header)
        
        for i in range(2, len(df) + 2):
            # Prorrateo por peso: Peso_Fila / Suma_Pesos_Misma_Guia
            f_prop = f'=IFERROR({col_weight}{i}/SUMIFS({col_weight}:{col_weight}, {col_key}:{col_key}, {col_key}{i}), 0)'
            worksheet.write_formula(f'K{i}', f_prop, pct)
            
            # Cruce de Costo: Buscar GroupKey en Col A y traer Monto de Col B
            f_val = f'=XLOOKUP({col_key}{i}, A:A, B:B, 0) * K{i}'
            worksheet.write_formula(f'L{i}', f_val, money)

        worksheet.set_column('A:L', 15)

    def _inject_inbound_formulas(self, writer: pd.ExcelWriter, df: pd.DataFrame, df_costs: pd.DataFrame):
        """Estrategia para Land/Sea (Costo Global Fijo)."""
        df = BusinessRulesEngine.apply_classification_rules(df)
        df['group_key'] = df['reference'].apply(self._clean_ref).fillna(df['bu'])
        
        cols_to_export = ['reference', 'bu', 'gross_weight', 'transport_type', 'group_key']
        df[cols_to_export].to_excel(writer, sheet_name='Inbound_Auditable', index=False)
        
        workbook, worksheet = writer.book, writer.sheets['Inbound_Auditable']
        money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
        pct_fmt = workbook.add_format({'num_format': '0.000%'})
        
        # Inyección de costo base (Fallback temporal, asumiendo primer costo)
        cost_col = [c for c in df_costs.columns if 'cost' in c.lower() or 'usd' in c.lower()][0]
        costo_total = pd.to_numeric(df_costs[cost_col].iloc[0], errors='coerce') if not df_costs.empty else 0.0
        
        worksheet.write('Z1', costo_total, money_fmt)
        worksheet.write('Y1', 'Total Cost:')
        
        for i in range(2, len(df) + 2):
            worksheet.write_formula(f'F{i}', f'=IFERROR(C{i}*1/SUMIFS(C:C, E:E, E{i}), 0)', pct_fmt)
            worksheet.write_formula(f'G{i}', f'=F{i}*$Z$1', money_fmt)
            
        worksheet.write(0, 5, '% Propot'); worksheet.write(0, 6, 'Amount')
        worksheet.set_column('A:G', 15)

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
            
            # Debug: Log matches
            matched = df['Total Cost'].notna().sum()
            total = len(df)
            logger.info(f"Cost matching: {matched}/{total} rows have cost data")
            if matched == 0:
                logger.warning("No cost matches found. Check reference formats in cost file.")
                logger.info(f"Sample group_keys: {df['group_key'].head().tolist()}")
                logger.info(f"Sample financial_keys: {costs_subset['financial_key'].head().tolist() if not costs_subset.empty else 'No cost data'}")
            df['fixed_cost'] = 0.0
            
            def apply_fallback(row):
                if pd.isna(row['Total Cost']) or row['Total Cost'] == 0:
                    trans_type = str(row['transport_type']).strip().lower()
                    fallback = self.fallback_costs.get(trans_type, 0.0)
                    if fallback > 0:
                        return fallback, fallback
                return row['Total Cost'], 0.0

            df[['Total Cost', 'fixed_cost']] = df.apply(apply_fallback, axis=1, result_type='expand')
            df['Total Cost'] = pd.to_numeric(df['Total Cost'], errors='coerce')
            
            # If no matches or partial matches, use average cost for missing ones
            if df['Total Cost'].isna().any() and not costs_subset.empty:
                avg_cost = costs_subset['Total Cost'].mean()
                df['Total Cost'] = df['Total Cost'].fillna(avg_cost)
                logger.info(f"Using average cost {avg_cost} for unmatched references")
            else:
                df['Total Cost'] = df['Total Cost'].fillna(0)
            
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

    def generate_outbound_auditable_excel(self, df_transactions: pd.DataFrame, df_costs: pd.DataFrame, output_path_or_buffer):
     
        try:
            # 1. Limpieza y preparación de datos base
            df = df_transactions.copy()
            df['gross_weight'] = pd.to_numeric(df.get('gross_weight', 0), errors='coerce').fillna(0.0)
            df = BusinessRulesEngine.apply_classification_rules(df)
            
            # Limpieza de Llave Primaria (Reference)
            df['group_key'] = df['reference'].apply(self._clean_ref).fillna('UNKNOWN')
            
            # Preparar tabla de costos para el XLOOKUP
            ref_col_costs = [c for c in df_costs.columns if 'ref' in c.lower() or 'bu' in c.lower()][0]
            cost_col = [c for c in df_costs.columns if 'cost' in c.lower() or 'amount' in c.lower() or 'usd' in c.lower()][0]
            
            costs_subset = df_costs[[ref_col_costs, cost_col]].copy()
            costs_subset.columns = ['Cost_Reference', 'Fix_Cost']
            costs_subset['Cost_Reference'] = costs_subset['Cost_Reference'].apply(self._clean_ref)
            costs_subset = costs_subset.groupby('Cost_Reference', as_index=False)['Fix_Cost'].sum()

            # 2. Inicializar XlsxWriter
            writer = pd.ExcelWriter(output_path_or_buffer, engine='xlsxwriter')
            
            # Escribir primero la tabla de Costos (Serán nuestras columnas A y B)
            costs_subset.to_excel(writer, sheet_name='Outbound_Auditable', index=False, startcol=0, startrow=0)
            
            # Escribir las transacciones al lado (Empezando en la columna D, que es índice 3)
            cols_to_export = ['reference', 'bu', 'gross_weight', 'transport_type', 'group_key']
            df_export = df[cols_to_export]
            df_export.to_excel(writer, sheet_name='Outbound_Auditable', index=False, startcol=3, startrow=0)
            
            workbook = writer.book
            worksheet = writer.sheets['Outbound_Auditable']
            
            # Formatos de Alta Fidelidad
            money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
            pct_fmt = workbook.add_format({'num_format': '0.00000%'})
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
            
            # 3. Mapeo de Columnas para Inyección
            # Tabla de Costos: A = Cost_Reference, B = Fix_Cost
            # Tabla Transaccional: D=Ref, E=BU, F=Weight, G=Type, H=Group_Key (Referencia Limpia)
            col_cost_ref = 'A'
            col_cost_val = 'B'
            
            col_trans_weight = 'F'
            col_trans_key = 'H'
            
            col_idx_propot = 3 + len(cols_to_export) # Índice 8 -> Columna I
            col_idx_amount = col_idx_propot + 1      # Índice 9 -> Columna J
            
            col_propot_letter = 'I'
            col_amount_letter = 'J'
            
            worksheet.write(0, col_idx_propot, '% Proportion', header_fmt)
            worksheet.write(0, col_idx_amount, 'Calc_Exp', header_fmt)
            
            total_rows = len(df_export)
            
            # 4. Inyección del Motor (XLOOKUP + SUMIFS)
            for i in range(2, total_rows + 2):
                # =F2 / SUMIFS(F:F, H:H, H2)
                f_propot = f'=IFERROR({col_trans_weight}{i}/SUMIFS({col_trans_weight}:{col_trans_weight}, {col_trans_key}:{col_trans_key}, {col_trans_key}{i}), 0)'
                worksheet.write_formula(f'{col_propot_letter}{i}', f_propot, pct_fmt)
                
                # =XLOOKUP(H2, A:A, B:B, 0) * I2
                # Nota: Si el XLOOKUP no encuentra la ref, devuelve 0 para no romper el Excel
                f_amount = f'=XLOOKUP({col_trans_key}{i}, {col_cost_ref}:{col_cost_ref}, {col_cost_val}:{col_cost_val}, 0) * {col_propot_letter}{i}'
                worksheet.write_formula(f'{col_amount_letter}{i}', f_amount, money_fmt)
            
            worksheet.set_column('A:J', 16)
            writer.close()
            
            logger.info("Excel Outbound Auditable con XLOOKUP inyectado correctamente.")
            return output_path_or_buffer
            
        except Exception as e:
            logger.error(f"Fallo crítico en motor Outbound: {e}")
            raise
        
    
        

