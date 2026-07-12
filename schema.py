from __future__ import annotations

INDICATORS = [
    ("ct01_households", "CT01", "Tổng số hộ dân", "Hộ"),
    ("ct02_population", "CT02", "Tổng số nhân khẩu", "Người"),
    ("ct03_poor_households", "CT03", "Số hộ nghèo", "Hộ"),
    ("ct04_near_poor_households", "CT04", "Số hộ cận nghèo", "Hộ"),
    ("ct05_revolution_contributors", "CT05", "Số người có công với cách mạng", "Người"),
    ("ct06_social_protection", "CT06", "Số đối tượng bảo trợ xã hội đang hưởng trợ cấp", "Người"),
    ("ct07_children_under_16", "CT07", "Số trẻ em dưới 16 tuổi", "Người"),
    ("ct08_special_children", "CT08", "Số trẻ em có hoàn cảnh đặc biệt", "Người"),
    ("ct09_cultural_households", "CT09", "Số hộ đạt Gia đình văn hóa", "Hộ"),
    ("ct10_working_age", "CT10", "Số người trong độ tuổi lao động", "Người"),
    ("ct11_health_insurance", "CT11", "Số người tham gia BHYT", "Người"),
    ("ct12_digital_team_members", "CT12", "Số thành viên Tổ công nghệ số cộng đồng", "Người"),
    ("ct13_online_public_service_guided", "CT13", "Số người dân được hướng dẫn dùng DVC trực tuyến trong kỳ", "Người"),
    ("ct14_domestic_violence_cases", "CT14", "Số vụ bạo lực gia đình ghi nhận trong kỳ", "Vụ"),
]

CODE_TO_FIELD = {code: field for field, code, _, _ in INDICATORS}
FIELD_TO_CODE = {field: code for field, code, _, _ in INDICATORS}
FIELD_LABELS = {field: label for field, _, label, _ in INDICATORS}
INDICATOR_FIELDS = [field for field, _, _, _ in INDICATORS]

CORE_REQUIRED_FIELDS = [
    "commune_name",
    "village_name",
    "period",
    "ct01_households",
    "ct02_population",
    "ct03_poor_households",
    "ct04_near_poor_households",
]

REPORT_COLUMNS = [
    "commune_name",
    "village_name",
    "period",
    "report_date",
    "reporter_name",
    "reporter_title",
    "phone",
    "due_at",
    "submitted_at",
    "submission_status",
    "days_late",
    *INDICATOR_FIELDS,
    "note",
    "source_file",
]
