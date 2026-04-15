"""
Growth Curve Data Automation Tool — FastAPI 백엔드
"""

from __future__ import annotations

import json
import logging
import os
import uuid
import tempfile
from pathlib import Path

import httpx
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

logger = logging.getLogger(__name__)

from processor import process_file, process_media_optimization

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
    base_media: str = Form(""),
    media_type: str = Form("peptone_screening"),
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
        "base_media": base_media,
        "media_type": media_type,
    }

    # 페이로드 파싱: 구버전(list) 또는 신버전(dict with experiment_type)
    try:
        parsed_payload = json.loads(sample_map_json) if sample_map_json else []
    except json.JSONDecodeError:
        parsed_payload = []

    # payload 형태 판별
    is_media_opt = (
        isinstance(parsed_payload, dict)
        and parsed_payload.get("experiment_type") == "media_optimization"
    )

    if is_media_opt:
        sample_map_list = []  # 배지 최적화에선 사용 X
        base_medium = parsed_payload.get("base_medium", {}) or {}
        # 신 포맷: composition_groups (composition_groups → per-SM variations 변환)
        # 구 포맷: variations (legacy, 직접 사용)
        composition_groups = parsed_payload.get("composition_groups", []) or []
        if composition_groups:
            variations = []
            for cg in composition_groups:
                applied = cg.get("applied_samples") or []
                cg_name = (cg.get("name") or "").strip()
                cg_strain = (cg.get("strain") or "").strip()
                cg_desc = (cg.get("description") or "").strip()
                cg_comp = cg.get("composition") or []
                for sm_code in applied:
                    variations.append({
                        "code": sm_code,
                        "strain": cg_strain,
                        "condition_name": cg_name,
                        "description": cg_desc,
                        "composition": cg_comp,   # full per-SM composition
                        "group_id": cg.get("id"),
                    })
        else:
            variations = parsed_payload.get("variations", []) or []
    else:
        # 펩톤 스크리닝 / 기타: 구버전 호환 (list) 또는 sample_map 래핑
        if isinstance(parsed_payload, list):
            sample_map_list = parsed_payload
        else:
            sample_map_list = parsed_payload.get("sample_map", [])
        base_medium = None
        variations = []
        composition_groups = []

    try:
        if is_media_opt:
            excel_bytes, chart_data = process_media_optimization(
                file_bytes, metadata, base_medium, variations,
                composition_groups=composition_groups,
            )
        else:
            excel_bytes, chart_data = process_file(
                file_bytes, metadata, sample_map_list
            )
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

    # PeptoMatch DB에 성장 데이터 자동 전송 (fire-and-forget)
    # 의미있는 매핑이 있을 때만 전송 (첫 번째 가공은 매핑 없이 진행되므로 skip)
    if is_media_opt:
        # 배지 최적화: composition_group 또는 variation 중 하나라도 의미 있는 정보가 있으면 전송
        has_meaningful_info = (
            bool(composition_groups) and any(
                cg.get("strain") or cg.get("name") or cg.get("description")
                or cg.get("composition") or cg.get("applied_samples")
                for cg in composition_groups
            )
        ) or any(
            v.get("strain") or v.get("condition_name") or v.get("description")
            or v.get("composition") or v.get("overrides")
            for v in variations
        )
    else:
        has_meaningful_info = any(
            entry.get("peptone_1") or entry.get("name")
            for entry in sample_map_list
        ) if sample_map_list else False

    peptomatch_url = os.getenv("PEPTOMATCH_INGEST_URL", "https://web-production-02f4.up.railway.app/api/ingest")
    if peptomatch_url and chart_data and has_meaningful_info:
        try:
            ingest_payload = {
                "metadata": metadata,
                "experiment_type": "media_optimization" if is_media_opt else "peptone_screening",
                "sample_map": sample_map_list,
                "base_medium": base_medium,
                "variations": variations,
                "composition_groups": composition_groups if is_media_opt else [],
                "chart_data": chart_data,
                "source_filename": file.filename or "unknown",
            }
            async with httpx.AsyncClient(timeout=10.0) as client:
                resp = await client.post(peptomatch_url, json=ingest_payload)
                logger.info(f"PeptoMatch ingest response: {resp.status_code}")
        except Exception as e:
            # 전송 실패해도 가공 결과는 정상 반환 (fire-and-forget)
            logger.warning(f"PeptoMatch ingest failed (non-blocking): {e}")

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
