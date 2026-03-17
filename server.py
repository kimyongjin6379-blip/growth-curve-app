"""
Growth Curve Data Automation Tool — FastAPI 백엔드
"""

from __future__ import annotations

import json
import os
import uuid
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from processor import process_file

app = FastAPI(title="Growth Curve Data Automation Tool", version="1.1.0")

# 임시 파일 저장소
TEMP_DIR = Path(tempfile.gettempdir()) / "growth_curve_outputs"
TEMP_DIR.mkdir(exist_ok=True)


@app.post("/api/process")
async def process_upload(
    file: UploadFile = File(...),
    experiment_date: str = Form(""),
    goal: str = Form(""),
    strain: str = Form(""),
    sample_map_json: str = Form("[]"),
):
    """업로드된 원본 엑셀 파일을 처리하고, 가공 결과를 반환."""
    # 파일 확장자 검증
    if file.filename:
        ext = Path(file.filename).suffix.lower()
        if ext not in (".xlsx", ".xls", ".csv"):
            raise HTTPException(
                status_code=400,
                detail="지원되지 않는 파일 형식입니다. .xlsx 또는 .csv 파일을 업로드해 주세요.",
            )

    # 파일 바이트 읽기
    file_bytes = await file.read()
    if not file_bytes:
        raise HTTPException(status_code=400, detail="업로드된 파일이 비어 있습니다.")

    metadata = {
        "experiment_date": experiment_date,
        "goal": goal,
        "strain": strain,
    }

    # 샘플 매핑 JSON 파싱
    try:
        sample_map_list = json.loads(sample_map_json) if sample_map_json else []
    except json.JSONDecodeError:
        sample_map_list = []

    try:
        excel_bytes, chart_data = process_file(file_bytes, metadata, sample_map_list)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"데이터 처리 오류: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"서버 처리 오류: {str(e)}")

    # 결과 파일 임시 저장
    file_id = str(uuid.uuid4())
    original_stem = Path(file.filename or "uploaded").stem
    output_filename = f"{original_stem}_processed.xlsx"
    output_path = TEMP_DIR / f"{file_id}.xlsx"
    output_path.write_bytes(excel_bytes)

    return JSONResponse(
        content={
            "success": True,
            "file_id": file_id,
            "filename": output_filename,
            "chart_data": chart_data,
            "message": "데이터 가공이 완료되었습니다.",
        }
    )


@app.get("/api/download/{file_id}")
async def download_file(file_id: str):
    """가공된 엑셀 파일 다운로드."""
    output_path = TEMP_DIR / f"{file_id}.xlsx"
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다. 다시 가공해 주세요.")

    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"processed_{file_id[:8]}.xlsx",
    )


# 정적 파일 서빙 (프론트엔드)
static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
