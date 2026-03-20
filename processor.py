"""
Magellan Microplate Reader 원본 데이터 → 가공 엑셀 변환 모듈
업데이트된 process_growth_curve.py 로직을 웹 서비스에서 사용 가능하도록 모듈화.

주요 변경사항:
- SAMPLE_MAP 지원 (SM 그룹코드 → 샘플명/펩톤 농도 매핑)
- 종합 시트에 샘플 매핑 테이블 추가
- raw 시트에 그룹코드, 샘플명, 펩톤 농도 컬럼 추가
- data 가공 시트에 균주명, 그룹코드, 샘플명, 펩톤 농도 컬럼 추가
"""

import io
import re
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# ---------------------------------------------------------------------------
# 샘플 매핑 헬퍼
# ---------------------------------------------------------------------------
def get_display_name(group_code: str, sample_map: dict) -> str:
    """SM 그룹코드 → 표시용 샘플명 변환."""
    if group_code in sample_map:
        return sample_map[group_code][0]
    return group_code


def get_peptone_pct(group_code: str, sample_map: dict):
    """SM 그룹코드 → 펩톤 농도(%) 반환. 매핑 없으면 None."""
    if group_code in sample_map:
        return sample_map[group_code][1]
    return None


# ---------------------------------------------------------------------------
# 1. 원본 데이터 읽기 (첫 번째 Raw OD 블록만 추출)
# ---------------------------------------------------------------------------
def read_raw_block(filepath_or_bytes):
    """첫 번째 시트에서 Raw OD 데이터 블록만 추출.

    Parameters
    ----------
    filepath_or_bytes : str | bytes | io.BytesIO
        파일 경로 또는 파일 바이트

    Returns
    -------
    df : DataFrame with columns: Well, Sample, T0, T1, T2, ...
    time_seconds : list[int] — 각 타임포인트의 초 단위 값
    """
    if isinstance(filepath_or_bytes, bytes):
        filepath_or_bytes = io.BytesIO(filepath_or_bytes)

    wb = pd.ExcelFile(filepath_or_bytes, engine="openpyxl")
    sheet_name = wb.sheet_names[0]
    raw = pd.read_excel(
        filepath_or_bytes, sheet_name=sheet_name, header=None, engine="openpyxl"
    )

    # 첫 번째 행: 시간 헤더 탐색 (숫자 + 's' 또는 '숫자 s' 패턴)
    header_row_idx = None
    for i, row in raw.iterrows():
        vals = row.dropna().astype(str).tolist()
        if any(re.match(r"^\d+\s*s$", v) for v in vals):
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("시간 헤더 행(예: '0s', '3593s' ...)을 찾을 수 없습니다.")

    # 데이터 행에서 실제 OD가 있는 범위 결정
    first_data_row = header_row_idx + 1
    if first_data_row >= len(raw):
        raise ValueError("데이터 행을 찾을 수 없습니다.")

    # 첫 데이터 행의 non-null 개수로 실제 시간 포인트 수 결정
    first_od = raw.iloc[first_data_row, 2:]
    n_times = first_od.notna().sum()

    # 시간(초) 추출 — 처음 n_times개만 사용
    raw_headers = raw.iloc[header_row_idx, 2 : 2 + n_times].tolist()
    time_seconds = []
    for h in raw_headers:
        m = re.match(r"(\d+)\s*s", str(h))
        time_seconds.append(int(m.group(1)) if m else 0)

    # 고유 컬럼명 생성 (T0, T1, T2, ...)
    time_cols = [f"T{i}" for i in range(n_times)]

    # 데이터 행 추출
    rows = []
    for i in range(first_data_row, len(raw)):
        well = raw.iloc[i, 0]
        sample = raw.iloc[i, 1]
        if pd.isna(well) or str(well).strip() == "":
            break
        od_values = raw.iloc[i, 2 : 2 + n_times].tolist()
        rows.append([str(well).strip(), str(sample).strip()] + od_values)

    if not rows:
        raise ValueError("데이터 행을 찾을 수 없습니다.")

    df = pd.DataFrame(rows, columns=["Well", "Sample"] + time_cols)
    for col in time_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, time_seconds


# ---------------------------------------------------------------------------
# 2. Blank 보정
# ---------------------------------------------------------------------------
def blank_correct(df: pd.DataFrame) -> pd.DataFrame:
    """BL 샘플 평균을 SM 샘플에서 차감한 보정 DataFrame 반환."""
    time_cols = [c for c in df.columns if c.startswith("T")]

    bl_mask = df["Sample"].str.contains("BL", case=False, na=False)
    sm_mask = df["Sample"].str.contains("SM", case=False, na=False)

    if bl_mask.sum() == 0:
        raise ValueError("Blank(BL) 샘플을 찾을 수 없습니다.")
    if sm_mask.sum() == 0:
        raise ValueError("Sample(SM) 샘플을 찾을 수 없습니다.")

    bl_mean = df.loc[bl_mask, time_cols].mean(axis=0)
    corrected = df.loc[sm_mask].copy()
    corrected[time_cols] = corrected[time_cols].subtract(bl_mean, axis=1)

    return corrected


# ---------------------------------------------------------------------------
# 3. 그룹별 통계 (Mean / SD)
# ---------------------------------------------------------------------------
def extract_group_name(sample: str) -> str:
    """SM1_1 → SM1, SM12_3 → SM12 (마지막 _숫자 접미사 제거)."""
    m = re.match(r"^(.+?)_\d+$", sample)
    return m.group(1) if m else sample


def compute_group_stats(df: pd.DataFrame):
    """그룹별 평균(Mean)과 표준편차(SD)를 반환."""
    time_cols = [c for c in df.columns if c.startswith("T")]
    df = df.copy()
    df["Group"] = df["Sample"].apply(extract_group_name)

    mean_df = df.groupby("Group")[time_cols].mean()
    sd_df = df.groupby("Group")[time_cols].std(ddof=1)  # 표본 표준편차

    # 단일 반복인 경우 SD = 0 처리
    sd_df = sd_df.fillna(0)

    return mean_df, sd_df


# ---------------------------------------------------------------------------
# 4. 엑셀 출력
# ---------------------------------------------------------------------------
def write_output_bytes(
    corrected_df: pd.DataFrame,
    mean_df: pd.DataFrame,
    sd_df: pd.DataFrame,
    time_cols: list,
    time_seconds: list,
    metadata: Optional[Dict[str, str]] = None,
    sample_map: Optional[Dict[str, Tuple[str, float]]] = None,
) -> bytes:
    """3개 시트를 가진 결과 엑셀 파일을 바이트로 생성하여 반환."""
    wb = Workbook()
    meta = metadata or {}
    smap = sample_map or {}
    strain_name = meta.get("strain", "")

    # ── 스타일 정의 ──
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    center_align = Alignment(horizontal="center", vertical="center")

    groups = mean_df.index.tolist()

    # =================================================================
    # Sheet 1: 종합 (메타데이터 양식 + 샘플 매핑 테이블)
    # =================================================================
    ws1 = wb.active
    ws1.title = "종합"

    meta_fields = [
        ("일자", meta.get("experiment_date", "")),
        ("배경", ""),
        ("목표", meta.get("goal", "")),
        ("방법", ""),
        ("균주", strain_name),
        ("배지", ""),
        ("배양조건", ""),
        ("장비", "Magellan Microplate Reader"),
        ("비고", ""),
    ]
    for i, (label, default) in enumerate(meta_fields, start=1):
        cell_label = ws1.cell(row=i, column=1, value=label)
        cell_label.font = header_font
        cell_label.fill = header_fill
        cell_label.border = thin_border
        cell_label.alignment = center_align
        cell_val = ws1.cell(row=i, column=2, value=default)
        cell_val.border = thin_border
    ws1.column_dimensions["A"].width = 15
    ws1.column_dimensions["B"].width = 60

    # ── 샘플 매핑 테이블 ──
    map_start_row = len(meta_fields) + 3  # 빈 행 하나 띄움

    map_title = ws1.cell(row=map_start_row, column=1, value="실험 샘플 구성")
    map_title.font = Font(bold=True, size=12)

    map_headers = ["그룹코드", "샘플명 (배지/펩톤)", "펩톤 농도 (%)"]
    map_header_fill = PatternFill("solid", fgColor="B4C6E7")
    for j, h in enumerate(map_headers, start=1):
        cell = ws1.cell(row=map_start_row + 1, column=j, value=h)
        cell.font = header_font
        cell.fill = map_header_fill
        cell.border = thin_border
        cell.alignment = center_align

    for i, grp in enumerate(groups, start=map_start_row + 2):
        # 그룹코드
        cell_code = ws1.cell(row=i, column=1, value=grp)
        cell_code.border = thin_border
        cell_code.alignment = center_align

        # 샘플명 (매핑 있으면 채움, 없으면 빈칸)
        display = get_display_name(grp, smap)
        cell_name = ws1.cell(
            row=i, column=2, value=display if display != grp else ""
        )
        cell_name.border = thin_border

        # 펩톤 농도 (매핑 있으면 채움, 없으면 빈칸)
        pct = get_peptone_pct(grp, smap)
        cell_pct = ws1.cell(row=i, column=3, value=pct if pct is not None else "")
        cell_pct.border = thin_border
        cell_pct.alignment = center_align
        if pct is not None:
            cell_pct.number_format = "0.0"

    ws1.column_dimensions["C"].width = 18

    # =================================================================
    # Sheet 2: raw (Blank 보정 Well별 데이터)
    # =================================================================
    ws2 = wb.create_sheet(title="raw")

    # 시간 인덱스 (시간 단위) 계산
    time_hours = [s / 3600 for s in time_seconds]

    # 헤더
    raw_headers = ["Well", "그룹코드", "샘플명", "펩톤 (%)"]
    for j, h in enumerate(raw_headers, start=1):
        cell = ws2.cell(row=1, column=j, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_align

    time_start_col = len(raw_headers) + 1
    for j, t in enumerate(time_cols):
        h = time_hours[j]
        cell = ws2.cell(row=1, column=time_start_col + j, value=f"{h:.2f}h")
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # 데이터
    for i, (_, row) in enumerate(corrected_df.iterrows(), start=2):
        grp = extract_group_name(row["Sample"])
        display = get_display_name(grp, smap)
        pct = get_peptone_pct(grp, smap)

        ws2.cell(row=i, column=1, value=row["Well"]).border = thin_border
        ws2.cell(row=i, column=2, value=row["Sample"]).border = thin_border

        cell_name = ws2.cell(
            row=i, column=3, value=display if display != grp else ""
        )
        cell_name.border = thin_border

        cell_pct = ws2.cell(
            row=i, column=4, value=pct if pct is not None else ""
        )
        cell_pct.border = thin_border
        cell_pct.alignment = center_align

        for j, t in enumerate(time_cols):
            cell = ws2.cell(
                row=i, column=time_start_col + j, value=round(row[t], 5)
            )
            cell.border = thin_border
            cell.number_format = "0.00000"

    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 25
    ws2.column_dimensions["D"].width = 12

    # =================================================================
    # Sheet 3: data 가공 (Wide-form: Mean | SD)
    # =================================================================
    ws3 = wb.create_sheet(title="data 가공")

    n_times = len(time_cols)

    # 행 1: 구분 라벨
    info_headers = ["균주명", "그룹코드", "샘플명", "펩톤 (%)"]
    for j, h in enumerate(info_headers, start=1):
        cell = ws3.cell(row=1, column=j, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_align

    mean_start_col = len(info_headers) + 1

    # Mean 영역 헤더
    mean_label_cell = ws3.cell(row=1, column=mean_start_col, value="Mean")
    mean_label_cell.font = Font(bold=True, color="FFFFFF")
    mean_label_cell.fill = PatternFill("solid", fgColor="4472C4")
    mean_label_cell.alignment = center_align
    if n_times > 1:
        ws3.merge_cells(
            start_row=1,
            start_column=mean_start_col,
            end_row=1,
            end_column=mean_start_col + n_times - 1,
        )

    # SD 영역 헤더
    sd_start_col = mean_start_col + n_times
    sd_label_cell = ws3.cell(row=1, column=sd_start_col, value="SD")
    sd_label_cell.font = Font(bold=True, color="FFFFFF")
    sd_label_cell.fill = PatternFill("solid", fgColor="ED7D31")
    sd_label_cell.alignment = center_align
    if n_times > 1:
        ws3.merge_cells(
            start_row=1,
            start_column=sd_start_col,
            end_row=1,
            end_column=sd_start_col + n_times - 1,
        )

    # 행 2: 시간 인덱스
    for j, h in enumerate(info_headers, start=1):
        cell = ws3.cell(row=2, column=j, value="")
        cell.border = thin_border
    ws3.cell(row=2, column=len(info_headers), value="Time Index").font = header_font
    ws3.cell(row=2, column=len(info_headers)).fill = header_fill
    ws3.cell(row=2, column=len(info_headers)).border = thin_border

    for j in range(n_times):
        # Mean 시간 인덱스
        cell_m = ws3.cell(row=2, column=mean_start_col + j, value=j)
        cell_m.font = header_font
        cell_m.fill = PatternFill("solid", fgColor="D6E4F0")
        cell_m.alignment = center_align
        cell_m.border = thin_border
        # SD 시간 인덱스
        cell_s = ws3.cell(row=2, column=sd_start_col + j, value=j)
        cell_s.font = header_font
        cell_s.fill = PatternFill("solid", fgColor="FCE4D6")
        cell_s.alignment = center_align
        cell_s.border = thin_border

    # 데이터 행
    for i, grp in enumerate(groups, start=3):
        display = get_display_name(grp, smap)
        pct = get_peptone_pct(grp, smap)

        # 균주명
        ws3.cell(
            row=i, column=1, value=strain_name if strain_name else ""
        ).border = thin_border

        # 그룹코드
        ws3.cell(row=i, column=2, value=grp).border = thin_border

        # 샘플명
        cell_name = ws3.cell(
            row=i, column=3, value=display if display != grp else ""
        )
        cell_name.border = thin_border

        # 펩톤 농도
        cell_pct = ws3.cell(
            row=i, column=4, value=pct if pct is not None else ""
        )
        cell_pct.border = thin_border
        cell_pct.alignment = center_align

        for j, t in enumerate(time_cols):
            # Mean
            cell_m = ws3.cell(
                row=i,
                column=mean_start_col + j,
                value=round(mean_df.loc[grp, t], 5),
            )
            cell_m.border = thin_border
            cell_m.number_format = "0.00000"
            # SD
            cell_s = ws3.cell(
                row=i,
                column=sd_start_col + j,
                value=round(sd_df.loc[grp, t], 5),
            )
            cell_s.border = thin_border
            cell_s.number_format = "0.00000"

    ws3.column_dimensions["A"].width = 20
    ws3.column_dimensions["B"].width = 12
    ws3.column_dimensions["C"].width = 25
    ws3.column_dimensions["D"].width = 12

    # =================================================================
    # Sheet 4: raw (세로) — Blank 보정 Well별 데이터 (전치: 시간=행, Well=열)
    # =================================================================
    ws4 = wb.create_sheet(title="raw (세로)")

    # 보정 데이터에서 Well/Sample 정보 추출
    wells = corrected_df["Well"].tolist()
    samples = corrected_df["Sample"].tolist()

    # ── 행 1: 헤더 — "Time (h)" + 각 Well의 Sample 이름
    cell_h = ws4.cell(row=1, column=1, value="Time (h)")
    cell_h.font = header_font
    cell_h.fill = header_fill
    cell_h.border = thin_border
    cell_h.alignment = center_align

    for ci, sample_name in enumerate(samples, start=2):
        cell = ws4.cell(row=1, column=ci, value=sample_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = center_align

    # ── 행 2: 균주명
    cell_s = ws4.cell(row=2, column=1, value="균주명")
    cell_s.font = header_font
    cell_s.fill = PatternFill("solid", fgColor="E2EFDA")
    cell_s.border = thin_border
    cell_s.alignment = center_align

    for ci, sample_name in enumerate(samples, start=2):
        grp = extract_group_name(sample_name)
        cell = ws4.cell(row=2, column=ci, value=strain_name if strain_name else "")
        cell.border = thin_border
        cell.alignment = center_align

    # ── 행 3: 샘플명
    cell_n = ws4.cell(row=3, column=1, value="샘플명")
    cell_n.font = header_font
    cell_n.fill = PatternFill("solid", fgColor="E2EFDA")
    cell_n.border = thin_border
    cell_n.alignment = center_align

    for ci, sample_name in enumerate(samples, start=2):
        grp = extract_group_name(sample_name)
        display = get_display_name(grp, smap)
        cell = ws4.cell(row=3, column=ci, value=display if display != grp else grp)
        cell.border = thin_border
        cell.alignment = center_align

    # ── 데이터 행: 시간별 OD 값
    data_start_row = 4
    for ti, t in enumerate(time_cols):
        row_idx = data_start_row + ti
        h = time_hours[ti]

        cell_time = ws4.cell(row=row_idx, column=1, value=round(h, 4))
        cell_time.border = thin_border
        cell_time.number_format = "0.00"
        cell_time.alignment = center_align

        for ci, (_, row_data) in enumerate(corrected_df.iterrows(), start=2):
            cell = ws4.cell(row=row_idx, column=ci, value=round(row_data[t], 5))
            cell.border = thin_border
            cell.number_format = "0.00000"

    ws4.column_dimensions["A"].width = 12

    # =================================================================
    # Sheet 5: data 가공 (세로) — 그룹 통계 전치 (시간=행, 그룹=열)
    # =================================================================
    ws5 = wb.create_sheet(title="data 가공 (세로)")

    n_groups = len(groups)

    # ── 행 1: 구분 라벨 — Mean 영역 + SD 영역
    cell_h5 = ws5.cell(row=1, column=1, value="Time (h)")
    cell_h5.font = header_font
    cell_h5.fill = header_fill
    cell_h5.border = thin_border
    cell_h5.alignment = center_align

    mean_col_start = 2
    sd_col_start = 2 + n_groups

    # Mean 구분 라벨 (merge)
    mean_label = ws5.cell(row=1, column=mean_col_start, value="Mean")
    mean_label.font = Font(bold=True, color="FFFFFF")
    mean_label.fill = PatternFill("solid", fgColor="4472C4")
    mean_label.alignment = center_align
    mean_label.border = thin_border
    if n_groups > 1:
        ws5.merge_cells(
            start_row=1,
            start_column=mean_col_start,
            end_row=1,
            end_column=mean_col_start + n_groups - 1,
        )

    # SD 구분 라벨 (merge)
    sd_label = ws5.cell(row=1, column=sd_col_start, value="SD")
    sd_label.font = Font(bold=True, color="FFFFFF")
    sd_label.fill = PatternFill("solid", fgColor="ED7D31")
    sd_label.alignment = center_align
    sd_label.border = thin_border
    if n_groups > 1:
        ws5.merge_cells(
            start_row=1,
            start_column=sd_col_start,
            end_row=1,
            end_column=sd_col_start + n_groups - 1,
        )

    # ── 행 2: 그룹코드
    cell_gc = ws5.cell(row=2, column=1, value="그룹코드")
    cell_gc.font = header_font
    cell_gc.fill = header_fill
    cell_gc.border = thin_border
    cell_gc.alignment = center_align

    for gi, grp in enumerate(groups):
        # Mean 열
        cell_m = ws5.cell(row=2, column=mean_col_start + gi, value=grp)
        cell_m.font = header_font
        cell_m.fill = PatternFill("solid", fgColor="D6E4F0")
        cell_m.border = thin_border
        cell_m.alignment = center_align
        # SD 열
        cell_s = ws5.cell(row=2, column=sd_col_start + gi, value=grp)
        cell_s.font = header_font
        cell_s.fill = PatternFill("solid", fgColor="FCE4D6")
        cell_s.border = thin_border
        cell_s.alignment = center_align

    # ── 행 3: 균주명
    cell_strain = ws5.cell(row=3, column=1, value="균주명")
    cell_strain.font = header_font
    cell_strain.fill = PatternFill("solid", fgColor="E2EFDA")
    cell_strain.border = thin_border
    cell_strain.alignment = center_align

    for gi in range(n_groups):
        cell_m = ws5.cell(
            row=3, column=mean_col_start + gi,
            value=strain_name if strain_name else "",
        )
        cell_m.border = thin_border
        cell_m.alignment = center_align
        cell_s = ws5.cell(
            row=3, column=sd_col_start + gi,
            value=strain_name if strain_name else "",
        )
        cell_s.border = thin_border
        cell_s.alignment = center_align

    # ── 행 4: 샘플명
    cell_sn = ws5.cell(row=4, column=1, value="샘플명")
    cell_sn.font = header_font
    cell_sn.fill = PatternFill("solid", fgColor="E2EFDA")
    cell_sn.border = thin_border
    cell_sn.alignment = center_align

    for gi, grp in enumerate(groups):
        display = get_display_name(grp, smap)
        disp_val = display if display != grp else grp

        cell_m = ws5.cell(row=4, column=mean_col_start + gi, value=disp_val)
        cell_m.border = thin_border
        cell_m.alignment = center_align
        cell_s = ws5.cell(row=4, column=sd_col_start + gi, value=disp_val)
        cell_s.border = thin_border
        cell_s.alignment = center_align

    # ── 데이터 행: 시간별 Mean/SD 값
    v_data_start = 5
    for ti, t in enumerate(time_cols):
        row_idx = v_data_start + ti
        h = time_hours[ti]

        cell_time = ws5.cell(row=row_idx, column=1, value=round(h, 4))
        cell_time.border = thin_border
        cell_time.number_format = "0.00"
        cell_time.alignment = center_align

        for gi, grp in enumerate(groups):
            # Mean
            cell_m = ws5.cell(
                row=row_idx,
                column=mean_col_start + gi,
                value=round(mean_df.loc[grp, t], 5),
            )
            cell_m.border = thin_border
            cell_m.number_format = "0.00000"
            # SD
            cell_s = ws5.cell(
                row=row_idx,
                column=sd_col_start + gi,
                value=round(sd_df.loc[grp, t], 5),
            )
            cell_s.border = thin_border
            cell_s.number_format = "0.00000"

    ws5.column_dimensions["A"].width = 12

    # 바이트로 저장
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()


# ---------------------------------------------------------------------------
# 5. 차트 데이터 추출 (JSON 용)
# ---------------------------------------------------------------------------
def extract_chart_data(
    mean_df: pd.DataFrame,
    sd_df: pd.DataFrame,
    time_seconds: list,
    sample_map: Optional[Dict[str, Tuple[str, float]]] = None,
) -> Dict:
    """Plotly 차트에 필요한 데이터를 딕셔너리로 반환."""
    smap = sample_map or {}
    time_cols = [c for c in mean_df.columns if c.startswith("T")]
    time_hours = [round(s / 3600, 2) for s in time_seconds]

    groups = mean_df.index.tolist()
    series = []
    for grp in groups:
        display = get_display_name(grp, smap)
        label = f"{display} ({grp})" if display != grp else grp

        mean_vals = mean_df.loc[grp, time_cols].tolist()
        sd_vals = sd_df.loc[grp, time_cols].tolist()
        # NaN → None for JSON serialization
        mean_vals = [None if (isinstance(v, float) and np.isnan(v)) else round(v, 5) for v in mean_vals]
        sd_vals = [None if (isinstance(v, float) and np.isnan(v)) else round(v, 5) for v in sd_vals]
        series.append({
            "name": label,
            "group_code": grp,
            "mean": mean_vals,
            "sd": sd_vals,
        })

    return {
        "time_hours": time_hours,
        "series": series,
    }


# ---------------------------------------------------------------------------
# 6. 샘플 매핑 파싱 (프론트엔드 JSON → dict)
# ---------------------------------------------------------------------------
def parse_sample_map(sample_map_list: Optional[List[Dict]] = None) -> Dict[str, Tuple[str, float]]:
    """프론트엔드에서 전달된 샘플 매핑 리스트를 dict로 변환.

    Parameters
    ----------
    sample_map_list : list of dict
        [{"code": "SM1", "name": "MRS (Control)", "peptone_pct": 0.0}, ...]

    Returns
    -------
    dict: {"SM1": ("MRS (Control)", 0.0), ...}
    """
    if not sample_map_list:
        return {}

    result = {}
    for item in sample_map_list:
        code = item.get("code", "").strip()
        name = item.get("name", "").strip()
        pct = item.get("peptone_pct")
        if code and name:
            try:
                pct_val = float(pct) if pct is not None and pct != "" else 0.0
            except (ValueError, TypeError):
                pct_val = 0.0
            result[code] = (name, pct_val)

    return result


# ---------------------------------------------------------------------------
# 통합 처리 함수
# ---------------------------------------------------------------------------
def process_file(
    file_bytes: bytes,
    metadata: Optional[Dict[str, str]] = None,
    sample_map_list: Optional[List[Dict]] = None,
) -> Tuple[bytes, Dict]:
    """
    원본 엑셀 바이트 → (가공 Excel 바이트, 차트 JSON 데이터) 반환.

    Parameters
    ----------
    file_bytes : bytes
        업로드된 원본 엑셀 파일 바이트
    metadata : dict, optional
        {"experiment_date": "...", "goal": "...", "strain": "..."}
    sample_map_list : list of dict, optional
        [{"code": "SM1", "name": "MRS (Control)", "peptone_pct": 0.0}, ...]

    Returns
    -------
    (excel_bytes, chart_data)
    """
    sample_map = parse_sample_map(sample_map_list)

    # 1) Raw 블록 추출
    df, time_seconds = read_raw_block(file_bytes)
    time_cols = [c for c in df.columns if c.startswith("T")]

    # 2) Blank 보정
    corrected = blank_correct(df)

    # 3) 그룹 통계
    mean_df, sd_df = compute_group_stats(corrected)

    # 4) 엑셀 출력 (바이트)
    excel_bytes = write_output_bytes(
        corrected, mean_df, sd_df, time_cols, time_seconds, metadata, sample_map
    )

    # 5) 차트 데이터
    chart_data = extract_chart_data(mean_df, sd_df, time_seconds, sample_map)

    # 6) 그룹 목록 (프론트엔드 샘플 매핑 UI 용)
    chart_data["groups"] = mean_df.index.tolist()

    return excel_bytes, chart_data
