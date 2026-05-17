# backend/routes/students.py
"""
Student import and listing routes.
POST /students/import/{course_id}  — upload the SIT xlsx roster
GET  /students/{course_id}         — list students for a course
"""
import io
from typing import List
from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from sqlalchemy.ext.asyncio import AsyncSession
from backend.database.connection import get_db
from backend.core.logger import get_logger
from backend.core.auth import get_current_user

logger = get_logger(__name__)
router = APIRouter(prefix="/students", tags=["Students"])

def _parse_sit_roster(file_bytes: bytes) -> List[dict]:
    """
    Parse SIT student roster xlsx.
    Structure:
      Row 1-3: institute/batch/branch headers — skip
      Row 4: 'SR. No.' | 'PRN' | 'AIML' | ...  (column headers)
      Row 5: 'Section A' — section marker
      Student rows: col B = PRN (24xxxxxxx), col C = name
      Section B/C markers appear mid-sheet — tracked for section label
    """
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    students = []
    current_section = "A"

    for row in rows:
        if not row or len(row) < 3:
            continue
        col_a = str(row[0]).strip() if row[0] else ""
        col_b = row[1]
        col_c = str(row[2]).strip() if row[2] else ""

        # Detect section marker rows
        if "Section A" in col_a or col_a == "Section A ":
            current_section = "A"; continue
        if "Section B" in col_a or col_a == "Section B ":
            current_section = "B"; continue
        if "Section C" in col_a or col_a == "Section C ":
            current_section = "C"; continue

        # Student row: PRN is a number starting with 24
        if col_b and str(col_b).startswith("24") and col_c and col_c not in ("AIML","PRN",""):
            try:
                prn = int(col_b)
            except (ValueError, TypeError):
                continue
            students.append({
                "prn":     str(prn),
                "name":    col_c,
                "section": current_section,
            })

    return students


@router.post("/import/{course_id}")
async def import_students(
    course_id: int,
    file: UploadFile = File(...),
    db: AsyncSession = Depends(get_db),
    current_user = Depends(get_current_user),
):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only .xlsx files are supported")

    file_bytes = await file.read()
    try:
        students = _parse_sit_roster(file_bytes)
    except Exception as e:
        logger.error(f"Student parse error: {e}")
        raise HTTPException(400, f"Failed to parse roster: {str(e)}")

    if not students:
        raise HTTPException(400, "No students found in the file. Check the format.")

    # Persist to DB — upsert pattern
    from sqlalchemy import text
    # Ensure table exists (created by init_db normally, but guard here)
    await db.execute(text("""
        CREATE TABLE IF NOT EXISTS students (
            id SERIAL PRIMARY KEY,
            course_id INTEGER NOT NULL,
            prn TEXT NOT NULL,
            name TEXT NOT NULL,
            section TEXT DEFAULT 'A',
            UNIQUE(course_id, prn)
        )
    """))

    inserted = 0
    for s in students:
        try:
            await db.execute(text("""
                INSERT INTO students (course_id, prn, name, section)
                VALUES (:course_id, :prn, :name, :section)
                ON CONFLICT (course_id, prn) DO UPDATE
                  SET name=EXCLUDED.name, section=EXCLUDED.section
            """), {"course_id": course_id, "prn": s["prn"], "name": s["name"], "section": s["section"]})
            inserted += 1
        except Exception:
            pass

    await db.commit()
    logger.info(f"Imported {inserted} students for course_id={course_id}")
    return {
        "status": "success",
        "data": {
            "imported": inserted,
            "total_parsed": len(students),
            "sections": list({s["section"] for s in students}),
        }
    }


@router.get("/{course_id}")
async def get_students(
    course_id: int,
    db: AsyncSession = Depends(get_db),
    current_user = Depends(get_current_user),
):
    from sqlalchemy import text
    try:
        result = await db.execute(
            text("SELECT prn, name, section FROM students WHERE course_id=:cid ORDER BY section, name"),
            {"cid": course_id}
        )
        rows = result.fetchall()
        students = [{"prn": r[0], "name": r[1], "section": r[2]} for r in rows]
        return {"status": "success", "data": students, "total": len(students)}
    except Exception as e:
        logger.error(f"Get students error: {e}")
        return {"status": "success", "data": [], "total": 0}


from pydantic import BaseModel

class StudentRosterUpdate(BaseModel):
    students: list  # [{prn, name, section}, ...]


@router.put("/update/{course_id}")
async def update_students(
    course_id: int,
    body: StudentRosterUpdate,
    db: AsyncSession = Depends(get_db),
    current_user = Depends(get_current_user),
):
    """
    Replace the entire student roster for a course with the provided list.
    Deletes all existing students for the course, then inserts the new list.
    """
    from sqlalchemy import text
    students = body.students
    if not isinstance(students, list):
        raise HTTPException(400, "students must be a list")

    try:
        # Ensure table exists
        await db.execute(text("""
            CREATE TABLE IF NOT EXISTS students (
                id SERIAL PRIMARY KEY,
                course_id INTEGER NOT NULL,
                prn TEXT NOT NULL,
                name TEXT NOT NULL,
                section TEXT DEFAULT 'A',
                UNIQUE(course_id, prn)
            )
        """))

        # Delete existing roster for this course
        await db.execute(text("DELETE FROM students WHERE course_id=:cid"), {"cid": course_id})

        inserted = 0
        for s in students:
            prn = str(s.get("prn", "")).strip()
            name = str(s.get("name", "")).strip()
            section = str(s.get("section", "A")).strip() or "A"
            if not prn or not name:
                continue
            await db.execute(text("""
                INSERT INTO students (course_id, prn, name, section)
                VALUES (:course_id, :prn, :name, :section)
                ON CONFLICT (course_id, prn) DO UPDATE
                  SET name=EXCLUDED.name, section=EXCLUDED.section
            """), {"course_id": course_id, "prn": prn, "name": name, "section": section})
            inserted += 1

        await db.commit()
        logger.info(f"Updated roster for course_id={course_id}: {inserted} students saved")
        return {
            "status": "success",
            "data": {"saved": inserted, "total": len(students)}
        }
    except Exception as e:
        logger.error(f"Student roster update error: {e}")
        raise HTTPException(500, f"Failed to update roster: {str(e)}")
