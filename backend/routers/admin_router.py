# =============================================================================
# routers/admin_router.py  –  Admin-only endpoints
# =============================================================================
import io
import uuid
from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, status
from pydantic import BaseModel

import sheets
from auth import hash_password
from dependencies import require_role

router = APIRouter(prefix="/admin", tags=["Admin"])

_admin_dep = Depends(require_role("ADMIN"))


# ─────────────────────────────────────────────────────────────────────────────
# CLASSES
# ─────────────────────────────────────────────────────────────────────────────
class AddClassRequest(BaseModel):
    class_id: str
    class_name: str
    year: Optional[str] = ""
    section: Optional[str] = ""
    semester: Optional[str] = ""


@router.get("/classes")
async def list_all_classes(_=_admin_dep):
    return sheets.get_all_classes()


@router.post("/classes")
async def add_class(body: AddClassRequest, _=_admin_dep):
    sheets.add_class(body.model_dump())
    return {"message": f"Class '{body.class_name}' added."}


class RenameClassRequest(BaseModel):
    new_name: str


@router.patch("/classes/{class_id}/rename")
async def rename_class(class_id: str, body: RenameClassRequest, _=_admin_dep):
    sheets.update_class_name(class_id, body.new_name)
    return {"message": "Class renamed."}


@router.delete("/classes/{class_id}")
async def deactivate_class(class_id: str, _=_admin_dep):
    sheets.deactivate_class(class_id)
    return {"message": "Class deactivated."}


# ─────────────────────────────────────────────────────────────────────────────
# SUBJECTS
# ─────────────────────────────────────────────────────────────────────────────
class AddSubjectRequest(BaseModel):
    subject_id: str
    class_id: str
    subject_name: str
    faculty_id: Optional[str] = ""


@router.get("/classes/{class_id}/subjects")
async def list_subjects(class_id: str, _=_admin_dep):
    return sheets.get_subjects_for_class(class_id)


@router.post("/subjects")
async def add_subject(body: AddSubjectRequest, _=_admin_dep):
    sheets.add_subject(body.model_dump())
    return {"message": f"Subject '{body.subject_name}' added."}


@router.delete("/subjects/{subject_id}")
async def delete_subject(subject_id: str, _=_admin_dep):
    sheets.delete_subject(subject_id)
    return {"message": "Subject deactivated."}


# ─────────────────────────────────────────────────────────────────────────────
# STUDENTS
# ─────────────────────────────────────────────────────────────────────────────
class AddStudentRequest(BaseModel):
    reg_no: str
    name: str
    class_id: str
    class_name: Optional[str] = ""
    # Student login credentials
    username: Optional[str] = None   # defaults to reg_no
    password: Optional[str] = None   # defaults to reg_no


@router.get("/classes/{class_id}/students")
async def list_students(class_id: str, _=_admin_dep):
    return sheets.get_students_by_class(class_id)


@router.post("/students")
async def add_student(body: AddStudentRequest, _=_admin_dep):
    cls = sheets.get_class_by_id(body.class_id)
    class_name = body.class_name or (cls["ClassName"] if cls else body.class_id)

    # Add to Students master sheet
    sheets.add_student({
        "reg_no": body.reg_no,
        "name": body.name,
        "class_id": body.class_id,
        "class_name": class_name,
    })

    # Create login account
    username = body.username or body.reg_no
    password = body.password or body.reg_no
    sheets.create_user({
        "user_id": body.reg_no,
        "name": body.name,
        "username": username,
        "password_hash": hash_password(password),
        "role": "STUDENT",
        "class_id": body.class_id,
    })
    return {"message": f"Student '{body.name}' added with login '{username}'."}


@router.delete("/students/{reg_no}")
async def deactivate_student(reg_no: str, _=_admin_dep):
    sheets.deactivate_student(reg_no)
    return {"message": "Student deactivated."}


# ─────────────────────────────────────────────────────────────────────────────
# FACULTY ACCOUNTS
# ─────────────────────────────────────────────────────────────────────────────
class AddFacultyRequest(BaseModel):
    faculty_id: str
    name: str
    username: str
    password: str


@router.post("/faculty")
async def add_faculty(body: AddFacultyRequest, _=_admin_dep):
    sheets.create_user({
        "user_id": body.faculty_id,
        "name": body.name,
        "username": body.username,
        "password_hash": hash_password(body.password),
        "role": "FACULTY",
        "class_id": "",
    })
    return {"message": f"Faculty '{body.name}' added."}


@router.delete("/users/{username}")
async def delete_user(username: str, _=_admin_dep):
    sheets.delete_user(username)
    return {"message": f"User '{username}' deleted."}


# ─────────────────────────────────────────────────────────────────────────────
# REPORTS
# ─────────────────────────────────────────────────────────────────────────────
@router.get("/reports/class/{class_id}")
async def class_report(class_id: str, _=_admin_dep):
    return sheets.get_class_attendance_report(class_id)


@router.get("/reports/student/{reg_no}")
async def student_report(reg_no: str, _=_admin_dep):
    return sheets.get_student_attendance_summary(reg_no)


# ─────────────────────────────────────────────────────────────────────────────
# PAST ATTENDANCE IMPORT (Excel upload)
# ─────────────────────────────────────────────────────────────────────────────
@router.post("/import/attendance")
async def import_attendance(file: UploadFile = File(...), _=Depends(require_role("ADMIN"))):
    """
    Upload an .xlsx file with columns:
    Date(DD-MM-YYYY) | ClassID | Hour | SubjectID | RegNo | Status | FacultyID(optional)

    All rows are validated then bulk-inserted into Attendance_Log (skipping duplicates).
    Also writes to individual class sheets.
    """
    import openpyxl

    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported.")

    contents = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
    ws_xl = wb.active

    headers = [str(c.value).strip() if c.value else "" for c in next(ws_xl.iter_rows(max_row=1))]
    required = {"Date", "ClassID", "Hour", "SubjectID", "RegNo", "Status"}
    missing = required - set(headers)
    if missing:
        raise HTTPException(
            status_code=400,
            detail=f"Missing required columns: {', '.join(missing)}. "
                   f"Expected headers: Date, ClassID, Hour, SubjectID, RegNo, Status, FacultyID",
        )

    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    errors = []
    valid_statuses = {"P", "A", "-"}

    for row_num, row in enumerate(ws_xl.iter_rows(min_row=2, values_only=True), start=2):
        row_data = {h: (str(row[i]).strip() if row[i] is not None else "") for h, i in idx.items()}
        if not any(row_data.values()):
            continue   # skip empty rows

        # Basic validation
        if row_data.get("Status", "").upper() not in valid_statuses:
            errors.append(f"Row {row_num}: Invalid status '{row_data.get('Status')}'.")
            continue
        if row_data.get("Hour", "").upper() not in sheets.ALL_HOURS:
            errors.append(f"Row {row_num}: Invalid hour '{row_data.get('Hour')}'.")
            continue

        rows.append({
            "date": row_data["Date"],
            "class_id": row_data["ClassID"],
            "hour": row_data["Hour"].upper(),
            "subject_id": row_data["SubjectID"],
            "reg_no": row_data["RegNo"],
            "status": row_data["Status"].upper(),
            "faculty_id": row_data.get("FacultyID", "IMPORT"),
        })

    if errors:
        raise HTTPException(
            status_code=422,
            detail={"validation_errors": errors[:20], "message": "Fix errors and re-upload."},
        )

    inserted = sheets.import_past_attendance_rows(rows)
    return {
        "message": f"Import complete. {inserted} rows inserted (duplicates skipped).",
        "total_rows": len(rows),
        "inserted": inserted,
        "skipped": len(rows) - inserted,
    }
