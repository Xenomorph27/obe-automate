"""
backend/services/course_file_service.py

Generates the complete OBE Course File (.docx) with all 13 sections:
  1.  Vision & Mission of the Department
  2.  Program Outcomes (POs), PEOs, PSOs
  3.  Syllabus + Personal Timetable  (placeholder if not stored)
  4.  CO Statements + CO-PO-PSO Mapping with justification
  5.  Previous year CO Attainment + Action Plan
  6.  Session Plan with CO mapping
  7.  Evaluation Plan with CO mapping + Marksheets
  8.  Slow & Advanced Learners + Action Plans
  9.  CO Attainment of Internal Evaluation
  10. Activity Reports
  11. Learning Material
  12. Question Bank
  13. Attendance links
"""

import json
import tempfile
from pathlib import Path
from typing import Optional

from sqlalchemy import text, select
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.logger import get_logger
from backend.core.storage import get_storage
from backend.services.course_service import CourseService

logger = get_logger(__name__)

_CATEGORY = "course_files"

# ── Colours & Fonts (shared with other services) ──────────────────────────────
_NAVY = "1F3864"
_LIGHT_BLUE = "D6DCE4"
_YELLOW = "FFFF00"
_GREEN = "E2EFDA"
_ORANGE = "FCE4D6"


# ─────────────────────────────────────────────────────────────────────────────
# Helper: build docx via Node docx library
# ─────────────────────────────────────────────────────────────────────────────

_JS_RUNNER = "/tmp/course_file_gen.js"


def _build_js(data: dict) -> str:
    """Return the Node.js script that builds the docx from the data dict."""
    data_json = json.dumps(data, ensure_ascii=False, default=str)

    return r"""
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, Header, Footer, PageBreak, ExternalHyperlink,
  LevelFormat,
} = require('docx');
const fs = require('fs');

const DATA = """ + data_json + r""";

// ── Style helpers ──────────────────────────────────────────────────────────
const NAVY = "1F3864";
const LIGHT = "D9E1F2";
const GREEN = "E2EFDA";
const ORANGE = "FCE4D6";

function run(text, opts={}) {
  return new TextRun({ text: String(text||''), font:'Calibri', size: opts.size||20,
    bold:opts.bold||false, color:opts.color||'000000', ...opts });
}

function para(children, opts={}) {
  if (typeof children === 'string') children = [run(children, opts)];
  return new Paragraph({
    alignment: opts.align||AlignmentType.LEFT,
    spacing: { before: opts.before||60, after: opts.after||60 },
    heading: opts.heading||undefined,
    pageBreakBefore: opts.pageBreak||false,
    numbering: opts.numbering||undefined,
    children,
  });
}

function heading1(text) {
  return para([run(text, {bold:true,size:24,color:NAVY})], {
    before:240, after:120, align:AlignmentType.LEFT
  });
}

function heading2(text) {
  return para([run(text, {bold:true,size:22})], { before:180, after:80 });
}

function sectionTitle(num, title) {
  return para([run(`${num}. ${title}`, {bold:true,size:24,color:NAVY})], {
    pageBreak: num>1, before:200, after:120
  });
}

function cell(text, opts={}) {
  const fills = { navy: NAVY, light: LIGHT, green: GREEN, orange: ORANGE };
  const fillColor = fills[opts.fill] || (opts.fill || 'FFFFFF');
  const textColor = opts.fill==='navy' ? 'FFFFFF' : '000000';
  const b = { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' };
  return new TableCell({
    borders: { top:b, bottom:b, left:b, right:b },
    shading: { fill: fillColor, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top:80, bottom:80, left:120, right:120 },
    columnSpan: opts.span||1,
    children: [new Paragraph({
      alignment: opts.align||AlignmentType.LEFT,
      children: [run(text, { bold:opts.bold||opts.fill==='navy',
        size:opts.size||18, color:textColor })],
    })],
  });
}

function makeTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a,b)=>a+b,0);
  const hRow = new TableRow({
    tableHeader: true,
    children: headers.map((h,i)=>cell(h,{fill:'navy',size:18,align:AlignmentType.CENTER,
      bold:true, width:colWidths[i]})),
  });
  const dataRows = rows.map((r,ri)=>new TableRow({
    children: r.map((v,ci)=>cell(String(v||''),{
      fill: ri%2===0 ? 'FFFFFF':'light', align:AlignmentType.LEFT,
    })),
  }));
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [hRow, ...dataRows],
  });
}

// ── Section builders ───────────────────────────────────────────────────────
const children = [];

// Cover info block
children.push(
  para([run(`Department of: ${DATA.department}`, {bold:true,size:22})], {before:0,after:60,align:AlignmentType.CENTER}),
  para([run(`${DATA.course_name} (${DATA.course_code}) Course File`, {bold:true,size:28,color:NAVY})],
    {before:60,after:60,align:AlignmentType.CENTER}),
  para([run(`A.Y ${DATA.academic_year} (${DATA.semester} Semester)`, {size:22})],
    {before:40,after:40,align:AlignmentType.CENTER}),
  para([run(DATA.batch ? `Batch ${DATA.batch}` : '', {size:22})],
    {before:20,after:20,align:AlignmentType.CENTER}),
  para(''),
);

// Table of contents (simple list)
children.push(
  para([run('Course File Contents', {bold:true,size:24,color:NAVY})],
    {before:120,after:80,align:AlignmentType.CENTER}),
);
const tocEntries = [
  'Vision & Mission of the Department',
  'Program Outcomes (POs), Program Educational Objectives (PEOs) and Program Specific Outcomes (PSOs)',
  'Syllabus, Personal Timetable',
  'CO Statements, CO-PO-PSO Mapping with justification',
  'CO Attainment of the course from the previous academic year and the action plan',
  'Session Plan',
  'Evaluation plan with CO Mapping',
  'List of Slow and Advanced learners and the action plans',
  'CO Attainment of internal evaluation',
  'The reports of the activities planned and conducted',
  'Learning Material',
  'Question Bank',
  'Compiled Attendance',
];
const tocRows = tocEntries.map((t,i)=>[[`${i+1}`,t]]).flat().map(r=>new TableRow({
  children: [
    cell(r[0], {fill:'navy',align:AlignmentType.CENTER,bold:true}),
    cell(r[1], {fill:'FFFFFF'}),
  ],
}));
// Re-build as proper rows
const tocTableRows = tocEntries.map((t,i)=>new TableRow({
  children:[cell(`${i+1}`,{fill:'navy',align:AlignmentType.CENTER,bold:true,size:18}),
             cell(t,{fill:'FFFFFF',size:18})],
}));
children.push(new Table({
  width:{size:9000,type:WidthType.DXA},
  columnWidths:[800,8200],
  rows:tocTableRows,
}));

// ── 1. Vision & Mission ────────────────────────────────────────────────────
children.push(sectionTitle(1,'Vision & Mission of the Department'));

if(DATA.vision_text){
  children.push(heading2('VISION OF THE DEPARTMENT'));
  children.push(para(DATA.vision_text||''));
}
if(DATA.mission_text){
  children.push(heading2('MISSION OF THE DEPARTMENT'));
  const missions = (DATA.mission_text||'').split('\n').filter(Boolean);
  missions.forEach(m=>children.push(para(m,{before:40,after:40})));
}
if(!DATA.vision_text && !DATA.mission_text){
  children.push(para('[Vision & Mission not yet filled. Edit in Course File section of the app.]',
    {color:'888888'}));
}

// ── 2. POs, PEOs, PSOs ────────────────────────────────────────────────────
children.push(sectionTitle(2,'Program Outcomes (POs), Program Educational Objectives (PEOs) and Program Specific Outcomes (PSOs)'));
children.push(heading2('Program Outcomes (POs)'));

if(DATA.pos && DATA.pos.length){
  const poRows = DATA.pos.map(p=>new TableRow({
    children:[
      cell(p.po_id,{fill:'navy',align:AlignmentType.CENTER,bold:true,size:18}),
      cell(p.statement||p.description||'',{fill:'FFFFFF',size:18}),
    ],
  }));
  children.push(new Table({
    width:{size:9000,type:WidthType.DXA},columnWidths:[1000,8000],rows:poRows,
  }));
} else {
  children.push(para('[POs not configured. Add POs during course setup.]',{color:'888888'}));
}

children.push(heading2('Program Educational Objectives (PEOs)'));
children.push(para('[PEOs to be added by faculty — standard institutional PEOs apply.]',{color:'888888'}));

children.push(heading2('Program Specific Outcomes (PSOs)'));
children.push(para('[PSOs to be added by faculty — standard institutional PSOs apply.]',{color:'888888'}));

// ── 3. Syllabus & Timetable ───────────────────────────────────────────────
children.push(sectionTitle(3,'Syllabus, Personal Timetable'));
children.push(heading2('Syllabus'));
if(DATA.syllabus_units && DATA.syllabus_units.length){
  DATA.syllabus_units.forEach(u=>{
    children.push(para([run(`Unit ${u.unit_number}: ${u.unit_title||''}`,{bold:true,size:20})],
      {before:120,after:40}));
    if(u.topics && u.topics.length){
      u.topics.forEach(t=>children.push(para(`  • ${t}`,{before:20,after:20})));
    }
  });
} else {
  children.push(para('[Syllabus will be extracted from session plan. Generate session plan first.]',{color:'888888'}));
}
children.push(heading2('Personal Timetable'));
children.push(para('[Faculty timetable — to be attached separately.]',{color:'888888'}));

// ── 4. CO Statements + CO-PO Mapping ─────────────────────────────────────
children.push(sectionTitle(4,'CO Statements, CO-PO-PSO Mapping with justification'));

if(DATA.cos && DATA.cos.length){
  const coRows = DATA.cos.map(c=>new TableRow({
    children:[
      cell(c.co_id,{fill:'navy',align:AlignmentType.CENTER,bold:true,size:18}),
      cell(c.statement||'',{fill:'FFFFFF',size:18}),
      cell(c.bloom_level||'',{fill:'light',align:AlignmentType.CENTER,size:18}),
    ],
  }));
  const coHeaderRow = new TableRow({
    children:[cell('CO',{fill:'navy',bold:true,align:AlignmentType.CENTER,size:18}),
               cell('Statement',{fill:'navy',bold:true,size:18}),
               cell("Bloom's Level",{fill:'navy',bold:true,align:AlignmentType.CENTER,size:18})],
  });
  children.push(new Table({
    width:{size:9000,type:WidthType.DXA},columnWidths:[900,6500,1600],
    rows:[coHeaderRow,...coRows],
  }));
}

// CO-PO Matrix
children.push(heading2('CO-PO Mapping'));
if(DATA.co_po_matrix && DATA.pos && DATA.pos.length){
  const poIds = DATA.pos.map(p=>p.po_id);
  const matrixHeaderCells = [cell('CO',{fill:'navy',bold:true,align:AlignmentType.CENTER,size:16}),
    ...poIds.map(p=>cell(p,{fill:'navy',bold:true,align:AlignmentType.CENTER,size:14}))];
  const matrixRows = [new TableRow({children:matrixHeaderCells})];
  (DATA.cos||[]).forEach(co=>{
    const mapping = DATA.co_po_matrix[co.co_id]||{};
    matrixRows.push(new TableRow({
      children:[cell(co.co_id,{fill:'light',bold:true,align:AlignmentType.CENTER,size:16}),
        ...poIds.map(p=>cell(mapping[p]||'-',{align:AlignmentType.CENTER,
          fill:mapping[p]?'green':'FFFFFF',size:16}))],
    }));
  });
  const colW = Math.floor(9000/(1+poIds.length));
  children.push(new Table({
    width:{size:9000,type:WidthType.DXA},
    columnWidths:[colW,...poIds.map(()=>colW)],
    rows:matrixRows,
  }));
}

// ── 5. Previous CO Attainment ─────────────────────────────────────────────
children.push(sectionTitle(5,'CO Attainment of the course from the previous academic year and the action plan'));
if(DATA.prev_co_attainment){
  children.push(para(DATA.prev_co_attainment,{before:60,after:60}));
} else {
  children.push(para('[Previous year CO attainment data not yet entered.]',{color:'888888'}));
}
children.push(heading2('Action Plan'));
if(DATA.action_plan){
  children.push(para(DATA.action_plan));
} else {
  children.push(para('[Action plan not yet entered.]',{color:'888888'}));
}

// ── 6. Session Plan ───────────────────────────────────────────────────────
children.push(sectionTitle(6,'Session Plan with CO mapping to each lecture'));
if(DATA.session_rows && DATA.session_rows.length){
  const sessionHeaders = ['Lect. No','Unit No','Points to Cover','Methodology','Type','CO Mapped'];
  const sessionColW = [800,800,4000,1200,1000,1200];
  const sessionDataRows = DATA.session_rows.map(r=>[
    r.lect_no||r.lectNo||'',
    r.unit_no||r.unitNo||'',
    r.points_to_cover||r.pointsToCover||r.topic||'',
    r.methodology||'',
    r.lecture_exp_eval||r.type||'Lecture',
    r.co||r.co_mapped||'',
  ]);
  children.push(makeTable(sessionHeaders, sessionDataRows, sessionColW));
} else {
  children.push(para('[Session plan not yet generated. Use the Session Plan page to generate and save.]',{color:'888888'}));
}

// Textbooks / references from session plan
if(DATA.textbooks && DATA.textbooks.length){
  children.push(heading2('Text Books & References'));
  DATA.textbooks.forEach(t=>children.push(para(`• ${t.title||t}`,{before:20,after:20})));
}
if(DATA.web_links && DATA.web_links.length){
  children.push(heading2('Web Links / Online Resources'));
  DATA.web_links.forEach(w=>children.push(para(`• ${w.title||w.url||w}`,{before:20,after:20})));
}

// Tutorial questions
if(DATA.tutorial_questions && DATA.tutorial_questions.length){
  children.push(heading2('Tutorial Questions with CO Mapping'));
  const tqHeaders=['Q No','Question','CO'];
  const tqW=[600,7200,1200];
  children.push(makeTable(tqHeaders, DATA.tutorial_questions.map((q,i)=>[i+1,q.question_text||q,q.co_id||'']), tqW));
}

// ── 7. Evaluation Plan & Marksheets ──────────────────────────────────────
children.push(sectionTitle(7,'Evaluation plan with CO Mapping'));
if(DATA.eval_rows && DATA.eval_rows.length){
  const evHeaders=['Sr.No','Component','Units/Syllabus','CO Mapped','Marks','Weightage','Tentative Date'];
  const evW=[500,2000,2500,1200,700,900,1200];
  const evData = DATA.eval_rows.map(r=>[
    r.sr_no||r.srNo||'',
    r.component||r.comp||r.name||'',
    r.unit_syllabus||r.units||'',
    r.co||r.co_mapped||'',
    r.marks||r.total_marks||'',
    r.weightage||'',
    r.date||r.tentative_date||'',
  ]);
  children.push(makeTable(evHeaders, evData, evW));
} else {
  children.push(para('[Evaluation plan not yet generated. Use the Evaluation Plan page first.]',{color:'888888'}));
}

// CA marksheets
(DATA.ca_sheets||[]).forEach(ca=>{
  if(!ca.qp || !ca.qp.length) return;
  children.push(heading2(`${ca.ca_label} — Question Paper`));
  const qpH=['Q.No','Question','Marks','CO','BL'];
  const qpW=[500,5800,700,1000,1000];
  children.push(makeTable(qpH, ca.qp.map(q=>[q.q_no||'',q.question_text||'',q.marks||'',q.co_id||'',q.bloom_level||'']),qpW));

  if(ca.marks && Object.keys(ca.marks).length){
    children.push(heading2(`${ca.ca_label} — Marks`));
    const qNos=(ca.qp||[]).map(q=>q.q_no||q.question_text?.substring(0,20));
    const mkH=['PRN','Name',...qNos,'Total'];
    const mkW=[1400,2000,...qNos.map(()=>Math.floor(4000/Math.max(qNos.length,1))),800];
    const mkData=Object.entries(ca.marks).map(([prn,mks])=>{
      const row=[prn, DATA.student_map?.[prn]||''];
      let tot=0;
      (ca.qp||[]).forEach(q=>{const v=parseFloat(mks[q.q_no]||0);row.push(v||'');tot+=v;});
      row.push(tot||'');
      return row;
    });
    children.push(makeTable(mkH,mkData,mkW));
  }
});

// ── 8. Slow & Advanced Learners ───────────────────────────────────────────
children.push(sectionTitle(8,'List of Slow and Advanced learners and the action plans'));
children.push(heading2('Slow Learners'));
if(DATA.slow_learners){
  if(Array.isArray(DATA.slow_learners_parsed)){
    const slHeaders=['Sr.No','PRN','Name','Marks Obtained'];
    const slW=[600,1600,4000,2800];
    children.push(makeTable(slHeaders,
      DATA.slow_learners_parsed.map((s,i)=>[i+1,s.prn||'',s.name||'',s.marks||'']), slW));
  } else {
    children.push(para(DATA.slow_learners));
  }
} else {
  children.push(para('[Slow learner list not yet entered. Complete CA marks to auto-generate.]',{color:'888888'}));
}
children.push(heading2('Advanced Learners'));
if(DATA.advanced_learners){
  children.push(para(DATA.advanced_learners));
} else {
  children.push(para('[Advanced learner list not yet entered.]',{color:'888888'}));
}

// ── 9. CO Attainment (internal) ───────────────────────────────────────────
children.push(sectionTitle(9,'CO Attainment of internal evaluation'));
if(DATA.co_attainment && Object.keys(DATA.co_attainment).length){
  const caHeaders=['CO','Attainment (%)','Level'];
  const caW=[1200,3000,4800];
  const caData=Object.entries(DATA.co_attainment).map(([co,val])=>{
    const pct=typeof val==='number'?val:0;
    const level=pct>=70?3:pct>=40?2:1;
    return [co,`${pct.toFixed(1)}%`,String(level)];
  });
  children.push(makeTable(caHeaders,caData,caW));
} else {
  children.push(para('[CO attainment will be shown here once marks are entered in the Master Attainment File.]',
    {color:'888888'}));
}

// ── 10. Activity Reports ──────────────────────────────────────────────────
children.push(sectionTitle(10,'The reports of the activities planned and conducted'));
if(DATA.activity_reports){
  DATA.activity_reports.split('\n').filter(Boolean).forEach(line=>
    children.push(para(line,{before:40,after:40})));
} else {
  children.push(para('[Activity reports not yet entered. Add them in the Course File section.]',{color:'888888'}));
}

// ── 11. Learning Material ─────────────────────────────────────────────────
children.push(sectionTitle(11,'Learning Material'));
if(DATA.learning_material_links){
  DATA.learning_material_links.split('\n').filter(Boolean).forEach(link=>{
    children.push(new Paragraph({
      spacing:{before:40,after:40},
      children:[new ExternalHyperlink({
        children:[new TextRun({text:link,font:'Calibri',size:20,color:'0563C1',underline:{}})],
        link:link,
      })],
    }));
  });
} else {
  children.push(para('[Learning material links not yet entered. Add them in the Course File section.]',{color:'888888'}));
}

// ── 12. Question Bank ─────────────────────────────────────────────────────
children.push(sectionTitle(12,'Question Bank'));
if(DATA.questions && DATA.questions.length){
  (DATA.cos||[]).forEach(co=>{
    const coQs=DATA.questions.filter(q=>q.co_id===co.co_id);
    if(!coQs.length) return;
    children.push(heading2(`${co.co_id}`));
    coQs.forEach((q,i)=>children.push(para(`${i+1}. ${q.question_text}`,{before:40,after:20})));
  });
  // Questions with no CO
  const unmapped=DATA.questions.filter(q=>!q.co_id);
  if(unmapped.length){
    children.push(heading2('General'));
    unmapped.forEach((q,i)=>children.push(para(`${i+1}. ${q.question_text}`,{before:40,after:20})));
  }
} else {
  children.push(para('[Question bank is empty. Use the Question Bank page to generate questions.]',{color:'888888'}));
}

// ── 13. Attendance ────────────────────────────────────────────────────────
children.push(sectionTitle(13,'Compiled Attendance'));
if(DATA.attendance_links){
  DATA.attendance_links.split('\n').filter(Boolean).forEach(link=>{
    children.push(new Paragraph({
      spacing:{before:40,after:40},
      children:[new ExternalHyperlink({
        children:[new TextRun({text:link,font:'Calibri',size:20,color:'0563C1',underline:{}})],
        link:link,
      })],
    }));
  });
} else {
  children.push(para('[Attendance links not yet entered. Add them in the Course File section.]',{color:'888888'}));
}

// ── Build document ────────────────────────────────────────────────────────
const doc = new Document({
  styles:{
    default:{
      document:{ run:{ font:'Calibri', size:20 } },
    },
  },
  numbering:{
    config:[
      { reference:'bullets',
        levels:[{ level:0, format:LevelFormat.BULLET, text:'•', alignment:AlignmentType.LEFT,
          style:{ paragraph:{ indent:{ left:720, hanging:360 } } } }] },
    ],
  },
  sections:[{
    properties:{
      page:{
        size:{ width:11906, height:16838 },
        margin:{ top:1080, right:1080, bottom:1080, left:1080 },
      },
    },
    headers:{
      default: new Header({ children:[
        new Paragraph({ alignment:AlignmentType.RIGHT,
          children:[run(`${DATA.course_name} (${DATA.course_code}) — Course File`,{size:16,color:'888888'})],
        }),
      ]}),
    },
    footers:{
      default: new Footer({ children:[
        new Paragraph({ alignment:AlignmentType.CENTER,
          children:[run('Page ',{size:16,color:'888888'}),
            new TextRun({children:[PageNumber.CURRENT],font:'Calibri',size:16,color:'888888'}),
            run(' | Generated by OBE Automate',{size:16,color:'888888'})],
        }),
      ]}),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync(process.argv[2], buf);
  console.log('OK');
}).catch(e=>{ console.error(e); process.exit(1); });
"""


class CourseFileService:
    def __init__(self, db: AsyncSession):
        self.db = db

    @staticmethod
    def get_filepath(course_id: int) -> str:
        storage = get_storage()
        p = storage.get_path(_CATEGORY, f"course_file_{course_id}.docx")
        return str(p) if p else str(
            get_storage()._dir(_CATEGORY) / f"course_file_{course_id}.docx"
        )

    async def _get_students(self, course_id: int):
        try:
            result = await self.db.execute(
                text("SELECT prn, name, section FROM students WHERE course_id=:cid ORDER BY section, name"),
                {"cid": course_id}
            )
            return [{"prn": r[0], "name": r[1], "section": r[2]} for r in result.fetchall()]
        except Exception:
            return []

    async def _get_session_rows(self, course_id: int):
        from backend.database.models import SessionPlanRow
        result = await self.db.execute(
            select(SessionPlanRow).where(SessionPlanRow.course_id == course_id)
        )
        row = result.scalar_one_or_none()
        return row.rows if row else []

    async def _get_eval_rows(self, course_id: int):
        from backend.database.models import EvalPlanRow
        result = await self.db.execute(
            select(EvalPlanRow).where(EvalPlanRow.course_id == course_id)
        )
        row = result.scalar_one_or_none()
        return row.rows if row else []

    async def _get_ca_sheets(self, course_id: int):
        from backend.database.models import CASheet
        result = await self.db.execute(
            select(CASheet).where(CASheet.course_id == course_id)
        )
        sheets = result.scalars().all()
        return [{"ca_label": s.ca_label, "qp": s.qp, "marks": s.marks} for s in sheets]

    async def _get_questions(self, course_id: int):
        from backend.database.models import Question
        result = await self.db.execute(
            select(Question).where(Question.course_id == course_id)
        )
        qs = result.scalars().all()
        return [{"question_text": q.question_text, "co_id": q.co_id,
                 "bloom_level": q.bloom_level, "marks": q.marks} for q in qs]

    async def _get_co_attainment(self, course_id: int, students, ca_sheets, cos):
        """Calculate CO attainment % from CA sheets marks."""
        attainment = {}
        for co in cos:
            cid = co["co_id"]
            total_pct = 0
            count = 0
            for sheet in ca_sheets:
                qp = [q for q in (sheet.get("qp") or []) if q.get("co_id") == cid]
                max_marks = sum(float(q.get("marks", 0)) for q in qp)
                if not max_marks or not students:
                    continue
                marks_data = sheet.get("marks") or {}
                passed = sum(
                    1 for s in students
                    if sum(float((marks_data.get(s["prn"]) or {}).get(q.get("q_no"), 0))
                           for q in qp) / max_marks * 100 >= 60
                )
                total_pct += (passed / len(students)) * 100
                count += 1
            attainment[cid] = round(total_pct / count, 1) if count else 0.0
        return attainment

    async def _get_slow_advanced(self, course_id: int, students, ca_sheets, cos):
        """Derive slow/advanced learners from marks."""
        if not students or not ca_sheets:
            return [], []

        student_totals = {}
        student_max = {}
        for s in students:
            student_totals[s["prn"]] = 0.0
            student_max[s["prn"]] = 0.0

        for sheet in ca_sheets:
            qp = sheet.get("qp") or []
            marks_data = sheet.get("marks") or {}
            total_marks = sum(float(q.get("marks", 0)) for q in qp)
            if not total_marks:
                continue
            for s in students:
                obtained = sum(float((marks_data.get(s["prn"]) or {}).get(q.get("q_no"), 0))
                               for q in qp)
                student_totals[s["prn"]] += obtained
                student_max[s["prn"]] += total_marks

        scored = []
        for s in students:
            mx = student_max.get(s["prn"], 0)
            tot = student_totals.get(s["prn"], 0)
            pct = (tot / mx * 100) if mx else 0
            scored.append({"prn": s["prn"], "name": s["name"],
                           "marks": f"{tot:.1f}/{mx:.0f}", "pct": pct})

        scored.sort(key=lambda x: x["pct"])
        slow = [s for s in scored if s["pct"] < 40]
        advanced = [s for s in scored if s["pct"] >= 75]
        return slow, advanced

    async def _get_extra(self, course_id: int):
        from backend.database.models import CourseFileExtra
        result = await self.db.execute(
            select(CourseFileExtra).where(CourseFileExtra.course_id == course_id)
        )
        extra = result.scalar_one_or_none()
        return extra.to_dict() if extra else {}

    def _extract_syllabus_from_session(self, session_rows):
        """Build unit/topic structure from saved session plan rows."""
        units = {}
        for row in session_rows:
            unit_no = row.get("unit_no", row.get("unitNo", ""))
            topic = row.get("points_to_cover", row.get("pointsToCover", row.get("topic", "")))
            if not unit_no:
                continue
            key = str(unit_no)
            if key not in units:
                units[key] = {"unit_number": unit_no,
                               "unit_title": row.get("unit_title", f"Unit {unit_no}"),
                               "topics": []}
            if topic and topic not in units[key]["topics"]:
                units[key]["topics"].append(topic)
        return list(units.values())

    def _extract_tutorial_questions(self, questions, max_per_co=5):
        """Pick tutorial questions per CO."""
        if not questions:
            return []
        by_co = {}
        for q in questions:
            co = q.get("co_id") or "General"
            by_co.setdefault(co, []).append(q)
        result = []
        for co, qs in by_co.items():
            result.extend(qs[:max_per_co])
        return result

    async def generate(self, course_id: int) -> dict:
        import subprocess

        course_svc = CourseService(self.db)
        course = await course_svc.get_course(course_id)
        students = await self._get_students(course_id)
        session_rows = await self._get_session_rows(course_id)
        eval_rows = await self._get_eval_rows(course_id)
        ca_sheets = await self._get_ca_sheets(course_id)
        questions = await self._get_questions(course_id)
        extra = await self._get_extra(course_id)
        cos = course.cos

        co_attainment = await self._get_co_attainment(course_id, students, ca_sheets, cos)
        slow_list, advanced_list = await self._get_slow_advanced(course_id, students, ca_sheets, cos)
        syllabus_units = self._extract_syllabus_from_session(session_rows)
        tutorial_qs = self._extract_tutorial_questions(questions)

        student_map = {s["prn"]: s["name"] for s in students}

        data = {
            "course_name": course.course_name,
            "course_code": course.course_code,
            "department": course.department,
            "faculty_name": course.faculty_name,
            "semester": course.semester,
            "academic_year": course.academic_year,
            "credits": course.credits,
            "batch": extra.get("batch", ""),
            "cos": cos,
            "pos": course.pos,
            "co_po_matrix": course.co_po_matrix,
            # Section 1
            "vision_text": extra.get("vision_text", ""),
            "mission_text": extra.get("mission_text", ""),
            # Section 3
            "syllabus_units": syllabus_units,
            # Section 5
            "prev_co_attainment": extra.get("prev_co_attainment", ""),
            "action_plan": extra.get("action_plan", ""),
            # Section 6
            "session_rows": session_rows,
            "tutorial_questions": tutorial_qs,
            # Section 7
            "eval_rows": eval_rows,
            "ca_sheets": ca_sheets,
            "student_map": student_map,
            # Section 8
            "slow_learners": extra.get("slow_learners", "") or
                             ("\n".join(f"{s['prn']} — {s['name']} ({s['marks']})" for s in slow_list)),
            "advanced_learners": extra.get("advanced_learners", "") or
                                  ("\n".join(f"{s['prn']} — {s['name']} ({s['marks']})" for s in advanced_list)),
            "slow_learners_parsed": slow_list,
            # Section 9
            "co_attainment": co_attainment,
            # Section 10
            "activity_reports": extra.get("activity_reports", ""),
            # Section 11
            "learning_material_links": extra.get("learning_material_links", ""),
            # Section 12
            "questions": questions,
            # Section 13
            "attendance_links": extra.get("attendance_links", ""),
        }

        js_code = _build_js(data)

        _storage = get_storage()
        _filename = f"course_file_{course_id}.docx"

        with tempfile.TemporaryDirectory() as tmp:
            js_path = Path(tmp) / "build.js"
            out_path = Path(tmp) / _filename
            js_path.write_text(js_code, encoding="utf-8")

            result = subprocess.run(
                ["node", str(js_path), str(out_path)],
                capture_output=True, text=True, timeout=60
            )
            if result.returncode != 0:
                raise RuntimeError(f"docx generation failed:\n{result.stderr[:500]}")

            _storage.save_from_path(_CATEGORY, _filename, out_path)

        filepath = str(_storage.get_path(_CATEGORY, _filename))
        logger.info(f"Course file saved -> {filepath}")

        return {
            "course_id": course_id,
            "course_name": course.course_name,
            "filename": _filename,
            "download_url": f"/course-file/download/{course_id}",
            "sections_with_data": {
                "session_plan": len(session_rows) > 0,
                "evaluation_plan": len(eval_rows) > 0,
                "ca_marks": len(ca_sheets) > 0,
                "question_bank": len(questions) > 0,
                "vision_mission": bool(extra.get("vision_text")),
            },
        }
