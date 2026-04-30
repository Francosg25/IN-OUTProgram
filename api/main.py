from fastapi import FastAPI, UploadFile, File, Form, HTTPException
import pandas as pd
import io
from services.orchestrator import LogisticsPipelineOrchestrator
from engine.allocation import CostAllocationEngine
import uvicorn

app = FastAPI(title="Logistics Backend API")

# Inicializamos el motor
engine = CostAllocationEngine(allocation_type='weight')
orchestrator = LogisticsPipelineOrchestrator(allocation_engine=engine)

@app.post("/process-logistics")
async def process_logistics(
    cost_file: UploadFile = File(...),
    sea_file: UploadFile = File(None),
    land_file: UploadFile = File(None),
    outbound_file: UploadFile = File(None)
):
    try:
        # 1. Leer archivo de costos
        cost_content = await cost_file.read()
        df_costs = pd.read_excel(io.BytesIO(cost_content))

        # 2. Preparar archivos de operación
        files_to_process = {}
        
        if sea_file:
            content = await sea_file.read()
            files_to_process[io.BytesIO(content)] = {'type': 'china_sea', 'label': 'Sea'}
            
        if land_file:
            content = await land_file.read()
            files_to_process[io.BytesIO(content)] = {'type': 'impos_land', 'label': 'Land'}
            
        if outbound_file:
            content = await outbound_file.read()
            files_to_process[io.BytesIO(content)] = {'type': 'expos', 'label': 'Outbound'}

        if not files_to_process:
            raise HTTPException(status_code=400, detail="No se enviaron archivos de operación.")

        # 3. Ejecutar Pipeline
        final_summary = orchestrator.run_pipeline(files_to_process, df_costs)
        
        # 4. Preparar respuesta (incluyendo metadatos de conciliación)
        response = {
            "summary": final_summary.to_dict(orient="records"),
            "reconciliation": final_summary.attrs.get('reconciliation', {})
        }
        
        return response

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
