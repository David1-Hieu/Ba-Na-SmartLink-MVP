from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

import pandas as pd

from schema import CODE_TO_FIELD, CORE_REQUIRED_FIELDS, INDICATOR_FIELDS, REPORT_COLUMNS


def normalize_text(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.strip().lower().replace("đ", "d").replace("Đ", "d")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_")


def to_int(value: Any) -> Optional[int]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, str):
        value = value.strip().replace(".", "").replace(",", "")
        if value == "":
            return None
    try:
        if pd.isna(value):
            return None
        return int(float(value))
    except Exception:
        return None


def to_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def parse_vietnamese_period(text: str) -> str:
    text = to_text(text)
    year_match = re.search(r"(20\d{2})", text)
    year = year_match.group(1) if year_match else "2026"
    lowered = text.lower()
    # Check longer Roman numerals first so "Quý II" is not caught by "Quý I".
    quarter_patterns = [
        (r"qu[ýy]\s*iv\b|qu[ýy]\s*4\b", "Q4"),
        (r"qu[ýy]\s*iii\b|qu[ýy]\s*3\b", "Q3"),
        (r"qu[ýy]\s*ii\b|qu[ýy]\s*2\b", "Q2"),
        (r"qu[ýy]\s*i\b|qu[ýy]\s*1\b", "Q1"),
    ]
    for pattern, q in quarter_patterns:
        if re.search(pattern, lowered):
            return f"{year}-{q}"
    month_match = re.search(r"tháng\s*(\d{1,2})", lowered)
    if month_match:
        return f"{year}-{int(month_match.group(1)):02d}"
    return text or f"{year}-Q2"


def parse_datetime_text(value: Any) -> str:
    text = to_text(value)
    if not text:
        return ""
    for fmt in ["%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"]:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%Y-%m-%d %H:%M") if "%H" in fmt else dt.strftime("%Y-%m-%d")
        except Exception:
            pass
    return text


def _empty_report_row() -> Dict[str, Any]:
    return {col: None for col in REPORT_COLUMNS}


def _find_label_value(raw: pd.DataFrame, label_keyword: str) -> str:
    keyword = normalize_text(label_keyword)
    for r in range(raw.shape[0]):
        for c in range(raw.shape[1]):
            if keyword in normalize_text(raw.iat[r, c]):
                # Common template has label in column 0 and value in column 1.
                if c + 1 < raw.shape[1]:
                    return to_text(raw.iat[r, c + 1])
    return ""


def _find_period(raw: pd.DataFrame) -> str:
    for r in range(raw.shape[0]):
        for c in range(raw.shape[1]):
            text = to_text(raw.iat[r, c])
            if "kỳ báo cáo" in text.lower() or "ky bao cao" in normalize_text(text):
                return parse_vietnamese_period(text)
    return "2026-Q2"


def parse_vertical_report_excel(path_or_buffer: Any, source_file: str = "") -> pd.DataFrame:
    raw = pd.read_excel(path_or_buffer, sheet_name=0, header=None)
    row = _empty_report_row()
    row.update(
        {
            "commune_name": "Xã Bà Nà",
            "village_name": _find_label_value(raw, "Đơn vị báo cáo"),
            "period": _find_period(raw),
            "report_date": "2026-06-19",
            "reporter_name": _find_label_value(raw, "Người lập báo cáo"),
            "reporter_title": _find_label_value(raw, "Chức danh"),
            "phone": _find_label_value(raw, "Số điện thoại"),
            "due_at": parse_datetime_text(_find_label_value(raw, "Hạn nộp")),
            "submitted_at": "",
            "submission_status": "Đã nộp",
            "days_late": 0,
            "note": "Nhập từ phiếu báo cáo từng thôn",
            "source_file": source_file,
        }
    )

    header_row = None
    for idx in range(raw.shape[0]):
        values = [normalize_text(x) for x in raw.iloc[idx].tolist()]
        if "ma_ct" in values and "so_lieu" in values:
            header_row = idx
            break
    if header_row is None:
        raise ValueError("Không tìm thấy bảng Mã CT/Số liệu trong file Excel.")

    header_values = [normalize_text(x) for x in raw.iloc[header_row].tolist()]
    code_col = header_values.index("ma_ct")
    value_col = header_values.index("so_lieu")

    for idx in range(header_row + 1, raw.shape[0]):
        code = to_text(raw.iat[idx, code_col]).upper()
        if code in CODE_TO_FIELD:
            row[CODE_TO_FIELD[code]] = raw.iat[idx, value_col]

    return pd.DataFrame([row], columns=REPORT_COLUMNS)


def _clean_summary_header(value: Any) -> str:
    text = to_text(value)
    code_match = re.search(r"CT\d{2}", text.upper())
    if code_match:
        return CODE_TO_FIELD.get(code_match.group(0), normalize_text(text))
    norm = normalize_text(text)
    aliases = {
        "thon": "village_name",
        "trang_thai_nop": "submission_status",
        "stt": "stt",
    }
    return aliases.get(norm, norm)


def parse_summary_workbook(path_or_buffer: Any, source_file: str = "") -> Tuple[pd.DataFrame, pd.DataFrame]:
    xls = pd.ExcelFile(path_or_buffer)
    summary_sheet = "Tong hop" if "Tong hop" in xls.sheet_names else xls.sheet_names[0]
    raw = pd.read_excel(path_or_buffer, sheet_name=summary_sheet, header=None)

    period = "2026-Q2"
    commune = "Xã Bà Nà"
    for c in range(raw.shape[1]):
        text = to_text(raw.iat[1, c]) if raw.shape[0] > 1 else ""
        if "Kỳ báo cáo" in text or "ky bao cao" in normalize_text(text):
            period = parse_vietnamese_period(text)
            m = re.search(r"(Xã|Phường)\s+[^—-]+", text, flags=re.IGNORECASE)
            if m:
                commune = m.group(0).strip()
            break

    header_idx = None
    for r in range(raw.shape[0]):
        row_norm = [normalize_text(x) for x in raw.iloc[r].tolist()]
        if "stt" in row_norm and "thon" in row_norm:
            header_idx = r
            break
    if header_idx is None:
        raise ValueError("Không tìm thấy bảng tổng hợp theo thôn trong workbook.")

    headers = [_clean_summary_header(x) for x in raw.iloc[header_idx].tolist()]
    data = raw.iloc[header_idx + 1 :].copy()
    data.columns = headers
    data = data[data.get("village_name").notna()]
    data = data[~data["village_name"].astype(str).str.contains("TỔNG|TONG", case=False, na=False)]

    rows: List[Dict[str, Any]] = []
    for _, src in data.iterrows():
        village = to_text(src.get("village_name"))
        if not village:
            continue
        status = to_text(src.get("submission_status")) or "Chưa nộp"
        row = _empty_report_row()
        row.update(
            {
                "commune_name": commune,
                "village_name": village,
                "period": period,
                "report_date": "2026-06-19",
                "submission_status": status,
                "source_file": source_file,
                "note": "Seed từ file tổng hợp Drive",
            }
        )
        for field in INDICATOR_FIELDS:
            row[field] = src.get(field)
        rows.append(row)

    reports = pd.DataFrame(rows, columns=REPORT_COLUMNS)

    progress = pd.DataFrame()
    if "Theo doi tien do" in xls.sheet_names:
        prog_raw = pd.read_excel(path_or_buffer, sheet_name="Theo doi tien do", header=None)
        header_idx = None
        for r in range(prog_raw.shape[0]):
            row_norm = [normalize_text(x) for x in prog_raw.iloc[r].tolist()]
            if "thon" in row_norm and "trang_thai" in row_norm:
                header_idx = r
                break
        if header_idx is not None:
            prog_headers = [normalize_text(x) for x in prog_raw.iloc[header_idx].tolist()]
            prog = prog_raw.iloc[header_idx + 1 :].copy()
            prog.columns = prog_headers
            if "thon" in prog.columns:
                prog = prog[prog["thon"].notna()]
                progress = pd.DataFrame(
                    {
                        "village_name": prog["thon"].map(to_text),
                        "reporter_name": prog.get("nguoi_lap_bia", ""),
                        "phone": prog.get("sdt_bia", ""),
                        "submitted_at": prog.get("thoi_diem_nop", ""),
                        "submission_status_progress": prog.get("trang_thai", ""),
                        "days_late": prog.get("so_ngay_tre", ""),
                    }
                )
                progress["reporter_name"] = progress["reporter_name"].map(to_text)
                progress["phone"] = progress["phone"].map(to_text)
                progress["submitted_at"] = progress["submitted_at"].map(parse_datetime_text)
                progress["submission_status_progress"] = progress["submission_status_progress"].map(to_text)
                progress["days_late"] = progress["days_late"].map(to_int)
                reports = reports.merge(progress, on="village_name", how="left", suffixes=("", "_p"))
                for col in ["reporter_name", "phone", "submitted_at", "days_late"]:
                    alt = f"{col}_p"
                    if alt in reports.columns:
                        current = reports[col] if col in reports.columns else pd.Series([None] * len(reports))
                        mask_missing = current.isna() | (current.astype(str).str.strip() == "")
                        reports[col] = current.where(~mask_missing, reports[alt])
                        reports = reports.drop(columns=[alt])
                if "submission_status_progress" in reports.columns:
                    reports["submission_status"] = reports["submission_status_progress"].where(
                        reports["submission_status_progress"].astype(str).str.len() > 0,
                        reports["submission_status"],
                    )
                    reports = reports.drop(columns=["submission_status_progress"])

    reports["period"] = period
    reports["commune_name"] = commune
    reports["due_at"] = reports["due_at"].fillna("2026-06-15 17:00")
    reports["source_file"] = source_file
    return reports[REPORT_COLUMNS], progress


def parse_flat_table_excel(path_or_buffer: Any, source_file: str = "") -> pd.DataFrame:
    df = pd.read_excel(path_or_buffer, header=0)
    rename: Dict[str, str] = {}
    aliases = {
        "xa_phuong": "commune_name",
        "xa": "commune_name",
        "thon": "village_name",
        "ten_thon": "village_name",
        "ky_bao_cao": "period",
        "ngay_bao_cao": "report_date",
        "nguoi_lap": "reporter_name",
        "so_dien_thoai": "phone",
        "trang_thai_nop": "submission_status",
        "so_ho": "ct01_households",
        "tong_so_ho_dan": "ct01_households",
        "so_nhan_khau": "ct02_population",
        "tong_so_nhan_khau": "ct02_population",
        "ho_ngheo": "ct03_poor_households",
        "so_ho_ngheo": "ct03_poor_households",
        "ho_can_ngheo": "ct04_near_poor_households",
        "so_ho_can_ngheo": "ct04_near_poor_households",
        "ghi_chu": "note",
    }
    for col in df.columns:
        norm = normalize_text(col)
        code_match = re.search(r"ct\d{2}", norm)
        if code_match:
            rename[col] = CODE_TO_FIELD.get(code_match.group(0).upper(), aliases.get(norm, norm))
        else:
            rename[col] = aliases.get(norm, norm)
    out = df.rename(columns=rename)
    for col in REPORT_COLUMNS:
        if col not in out.columns:
            out[col] = None
    out["source_file"] = source_file
    out["commune_name"] = out["commune_name"].fillna("Xã Bà Nà")
    out["period"] = out["period"].fillna("2026-Q2")
    out["submission_status"] = out["submission_status"].fillna("Đã nộp")
    return out[REPORT_COLUMNS]


def parse_excel_to_reports(path_or_buffer: Any, source_file: str = "") -> pd.DataFrame:
    """Parse either the official Drive summary workbook, one village report, or a flat table."""
    # Use a fresh reader each time; Streamlit UploadedFile supports seek.
    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)
    xls = pd.ExcelFile(path_or_buffer)
    sheet_names = [normalize_text(s) for s in xls.sheet_names]

    if any(s == "tong_hop" for s in sheet_names):
        if hasattr(path_or_buffer, "seek"):
            path_or_buffer.seek(0)
        reports, _ = parse_summary_workbook(path_or_buffer, source_file)
        return reports

    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)
    first = pd.read_excel(path_or_buffer, sheet_name=0, header=None)
    flat_text = " ".join(to_text(x) for x in first.fillna("").to_numpy().flatten()[:80])
    if "Mã CT" in flat_text and "Số liệu" in flat_text:
        if hasattr(path_or_buffer, "seek"):
            path_or_buffer.seek(0)
        return parse_vertical_report_excel(path_or_buffer, source_file)

    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)
    return parse_flat_table_excel(path_or_buffer, source_file)


def get_raw_preview(path_or_buffer: Any, max_rows: int = 30) -> pd.DataFrame:
    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)
    return pd.read_excel(path_or_buffer, sheet_name=0, header=None).head(max_rows)


def _issue(row: Dict[str, Any], severity: str, error_type: str, field_name: str, message: str) -> Dict[str, Any]:
    return {
        "village_name": to_text(row.get("village_name")) or "(chưa rõ)",
        "period": to_text(row.get("period")) or "(chưa rõ)",
        "severity": severity,
        "error_type": error_type,
        "field_name": field_name,
        "message": message,
    }


def standardize_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    clean = df.copy()
    for field in INDICATOR_FIELDS + ["days_late"]:
        if field in clean.columns:
            clean[field] = clean[field].map(to_int)
    for col in ["commune_name", "village_name", "period", "reporter_name", "reporter_title", "phone", "due_at", "submitted_at", "submission_status", "note", "source_file"]:
        if col in clean.columns:
            clean[col] = clean[col].map(to_text)
    return clean


def validate_reports(df: pd.DataFrame, allow_not_submitted_rows: bool = True) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
    clean = standardize_numeric_columns(df)
    issues: List[Dict[str, Any]] = []

    for idx, row_obj in clean.iterrows():
        row = row_obj.to_dict()
        status_norm = normalize_text(row.get("submission_status"))
        not_submitted = "chua_nop" in status_norm

        for field in ["commune_name", "village_name", "period"]:
            metadata_required_fields = {
                    "reporter_name": "Thiếu người lập báo cáo.",
                    "reporter_title": "Thiếu chức danh người lập báo cáo.",
                    "phone": "Thiếu số điện thoại người lập báo cáo.",
                    "due_at": "Thiếu hạn nộp báo cáo.",
                }

            for field, message in metadata_required_fields.items():
                    if not to_text(row.get(field)):
                        severity = "BLOCKER" if field in ["reporter_name", "phone"] else "WARNING"
                        issues.append(
                            _issue(
                                row,
                                severity,
                                "Thiếu thông tin biểu mẫu",
                                field,
                                message,
                            )
                        )
            if not to_text(row.get(field)):
                issues.append(_issue(row, "BLOCKER", "Thiếu dữ liệu", field, f"Thiếu trường bắt buộc: {field}."))

        if not_submitted and allow_not_submitted_rows:
            continue

        for field in CORE_REQUIRED_FIELDS:
            if field in ["commune_name", "village_name", "period"]:
                continue
            value = row.get(field)
            if value is None:
                issues.append(_issue(row, "BLOCKER", "Thiếu số liệu", field, f"Thiếu số liệu bắt buộc: {field}."))

        for field in INDICATOR_FIELDS:
            value = row.get(field)
            if value is not None and value < 0:
                issues.append(_issue(row, "BLOCKER", "Sai định dạng", field, "Số liệu phải là số nguyên không âm."))

        hh = row.get("ct01_households")
        pop = row.get("ct02_population")
        poor = row.get("ct03_poor_households")
        near_poor = row.get("ct04_near_poor_households")
        children = row.get("ct07_children_under_16")
        special_children = row.get("ct08_special_children")
        cultural = row.get("ct09_cultural_households")
        working = row.get("ct10_working_age")
        bhi = row.get("ct11_health_insurance")
        phone = to_text(row.get("phone"))

        if phone and not re.fullmatch(r"0\d{9}", phone):
            issues.append(
                _issue(
                    row,
                    "WARNING",
                    "Sai định dạng",
                    "phone",
                    "Số điện thoại không đúng định dạng 10 chữ số bắt đầu bằng 0.",
                )
            )
        if hh is not None and pop is not None:
            if pop < hh:
                issues.append(_issue(row, "BLOCKER", "Sai logic", "ct02_population", "Tổng nhân khẩu nhỏ hơn tổng số hộ dân."))
            ratio = pop / hh if hh else None
            if ratio is not None and (ratio < 2.0 or ratio > 5.5):
                issues.append(_issue(row, "WARNING", "Bất thường", "ct02_population", f"Tỷ lệ nhân khẩu/hộ = {ratio:.2f}, cần kiểm tra lại."))

        if hh is not None and poor is not None and poor > hh:
            issues.append(_issue(row, "BLOCKER", "Sai logic", "ct03_poor_households", "Số hộ nghèo lớn hơn tổng số hộ dân."))
        if hh is not None and near_poor is not None and poor is not None and poor + near_poor > hh:
            issues.append(_issue(row, "BLOCKER", "Sai logic", "ct03_ct04", "Hộ nghèo + hộ cận nghèo lớn hơn tổng số hộ dân."))
        if pop is not None and children is not None and children > pop:
            issues.append(_issue(row, "BLOCKER", "Sai logic", "ct07_children_under_16", "Số trẻ em dưới 16 tuổi lớn hơn tổng nhân khẩu."))
        if children is not None and special_children is not None and special_children > children:
            issues.append(_issue(row, "BLOCKER", "Sai logic", "ct08_special_children", "Trẻ em hoàn cảnh đặc biệt lớn hơn tổng số trẻ em."))
        if hh is not None and cultural is not None and cultural > hh:
            issues.append(_issue(row, "BLOCKER", "Sai logic", "ct09_cultural_households", "Số hộ gia đình văn hóa lớn hơn tổng số hộ dân."))
        if pop is not None and working is not None and working > pop:
            issues.append(_issue(row, "BLOCKER", "Sai logic", "ct10_working_age", "Số người trong độ tuổi lao động lớn hơn tổng nhân khẩu."))
        if pop is not None and bhi is not None and bhi > pop:
            issues.append(_issue(row, "BLOCKER", "Sai logic", "ct11_health_insurance", "Số người tham gia BHYT lớn hơn tổng nhân khẩu."))
        if phone and not re.fullmatch(r"0\d{9}", phone):
            issues.append(_issue(row, "WARNING", "Sai định dạng", "phone", "Số điện thoại không đúng định dạng 10 chữ số bắt đầu bằng 0."))

    return clean[REPORT_COLUMNS], issues


def has_blockers(issues: Iterable[Dict[str, Any]]) -> bool:
    return any(item.get("severity") == "BLOCKER" for item in issues)
