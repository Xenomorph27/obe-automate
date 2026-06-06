"""
backend/services/co_po_engine.py

CO-PO Mapping Engine — Clear, Explainable, Accurate
=====================================================

PHILOSOPHY
----------
Each PO has a GATE: a minimum evidence threshold that must be met by the CO
statement itself before any mapping is assigned. If the gate is not met → 0 (blank).

This means:
  - PO6, PO7, PO8, PO9, PO10, PO11  →  0 by default, only assigned if CO
    explicitly mentions the relevant domain keyword(s).
  - PO1, PO2, PO3, PO4, PO5, PO12   →  can be inferred from Bloom's verb +
    subject-matter keywords, but only if genuinely relevant.

STRENGTH SCALE
--------------
  3 = High  — CO directly and substantially develops this PO skill
  2 = Medium — CO moderately develops this PO skill (partial or indirect)
  1 = Low   — CO peripherally touches this PO (minor contribution)
  0 = None  — CO has no meaningful relationship to this PO → shown as blank

MAPPING PIPELINE
----------------
  1. Bloom's verb  → identifies cognitive level and base PO relevance
  2. Subject keywords → adds/boosts specific POs based on content domain
  3. Gate check    → enforces that "soft" POs (PO6–PO11) only activate on
                     explicit evidence. Never assumed. Never defaulted.
  4. AI layer      → LLM reads the actual CO text and independently assigns
                     strengths; used as a check on rule output.
  5. Merge         → AI is trusted for POs where rules give 0 (AI may catch
                     what rules miss); rules are trusted for strength values
                     when both agree a PO is relevant.

Works for ANY engineering course. No course-specific hardcoding.
"""

import re
import logging

logger = logging.getLogger(__name__)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1: BLOOM'S VERB → BASE PO MAPPING
#
# Each Bloom's level activates specific POs at defined strengths.
# POs NOT listed here get 0 from the verb step (may be added by keywords).
# ═══════════════════════════════════════════════════════════════════════════════

_VERB_MAP = [
    # L1 — Remember / Recall
    # Cognitive skill: retrieving knowledge from memory.
    # → Needs engineering knowledge (PO1). Minimal analysis. Lifelong learning enabled.
    (
        r"\b(list|recall|define|state|identify|name|label|recognize|memorize|enumerate|outline)\b",
        {"PO1": 2, "PO2": 1, "PO12": 1}
    ),

    # L2 — Understand / Explain
    # Cognitive skill: constructing meaning from information.
    # → Needs knowledge (PO1) and some analytical thinking (PO2).
    (
        r"\b(explain|describe|summarize|interpret|classify|paraphrase|illustrate|understand|discuss|express|translate)\b",
        {"PO1": 3, "PO2": 2, "PO12": 1}
    ),

    # L3 — Apply
    # Cognitive skill: carrying out a procedure in a given situation.
    # → Needs knowledge (PO1), analysis to select method (PO2), solution-building (PO3), tools (PO5).
    (
        r"\b(apply|use|utilize|demonstrate|compute|solve|execute|implement|employ|perform|operate|calculate)\b",
        {"PO1": 3, "PO2": 2, "PO3": 2, "PO5": 1, "PO12": 1}
    ),

    # L4 — Analyze / Compare
    # Cognitive skill: breaking material into parts and determining how they relate.
    # → Strong on knowledge (PO1), strong analysis (PO2), investigation methods (PO4).
    (
        r"\b(analyze|analyse|differentiate|compare|contrast|examine|categorize|dissect|distinguish|break.?down|deconstruct)\b",
        {"PO1": 3, "PO2": 3, "PO4": 2, "PO12": 1}
    ),

    # L5 — Evaluate / Justify
    # Cognitive skill: making judgements based on criteria and standards.
    # → Needs strong knowledge (PO1), deep analysis (PO2), research/evidence (PO4).
    (
        r"\b(evaluate|justify|assess|critique|judge|defend|recommend|argue|validate|verify|appraise|rank)\b",
        {"PO1": 3, "PO2": 3, "PO4": 3, "PO12": 1}
    ),

    # L6 — Create / Design / Develop / Build
    # Cognitive skill: producing something new by combining elements.
    # → Needs all three core engineering POs: knowledge (PO1), analysis (PO2),
    #   design/solution (PO3), modern tools (PO5).
    (
        r"\b(design|develop|build|create|construct|formulate|synthesize|architect|engineer|model|generate|produce|invent|prototype)\b",
        {"PO1": 3, "PO2": 2, "PO3": 3, "PO5": 2, "PO12": 1}
    ),

    # L4/5 — Investigate / Experiment / Test
    # Cognitive skill: systematic inquiry using research methods.
    # → Needs knowledge (PO1), analysis (PO2), experimental methods (PO4), tools (PO5).
    (
        r"\b(investigate|experiment|test|measure|simulate|benchmark|survey|observe|probe|explore|research)\b",
        {"PO1": 3, "PO2": 2, "PO4": 3, "PO5": 2, "PO12": 1}
    ),
]


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2: SUBJECT-MATTER KEYWORD → PO BOOST
#
# Applied AFTER the verb step. Boosts are ADDED to existing values (capped at 3).
# A PO that is 0 after the verb step CAN be raised here — but only for
# PO1–PO5 and PO12. PO6–PO11 are gated separately (Section 3).
# ═══════════════════════════════════════════════════════════════════════════════

_KEYWORD_BOOSTS = [

    # ── CORE ENGINEERING KNOWLEDGE (PO1) ──────────────────────────────────────
    # Technical domain terms that reinforce engineering knowledge requirement.
    (r"\b(algorithm|theorem|formula|principle|concept|theory|technique|method|mechanism|"
     r"architecture|protocol|standard|specification|model|paradigm)\b",
     {"PO1": 1}),

    # ── ANALYTICAL REASONING (PO2) ─────────────────────────────────────────────
    # Keywords implying reasoning, deduction, or inference tasks.
    (r"\b(reason|infer|deduc\w+|interpret|diagnos\w+|troubleshoot|root.?cause|trade.?off|"
     r"constraint|limitation|assumption|hypothesis|criterion|criteria)\b",
     {"PO2": 1}),

    # ── SYSTEM / SOLUTION DESIGN (PO3) ────────────────────────────────────────
    # Keywords implying building, structuring, or designing a system.
    (r"\b(system|pipeline|workflow|solution|structure|component|module|subsystem|"
     r"interface|integration|deployment|infrastructure|circuit|network)\b",
     {"PO3": 1}),

    # ── DATA / EXPERIMENTAL METHODS (PO4) ─────────────────────────────────────
    # Keywords implying structured data collection, analysis, or experiments.
    (r"\b(data|dataset|experiment|measurement|statistic|observation|metric|"
     r"performance|accuracy|precision|recall|f1|loss|error|residual|"
     r"sample|population|distribution|hypothesis.test|significance)\b",
     {"PO4": 1}),

    # ── MODERN TOOLS / SOFTWARE (PO5) ─────────────────────────────────────────
    # Specific tools, platforms, libraries, or computational methods.
    (r"\b(tool|software|platform|library|framework|tensorflow|pytorch|sklearn|"
     r"keras|opencv|matlab|simulink|arduino|raspberry|cloud|aws|azure|gcp|"
     r"docker|kubernetes|git|api|sdk|ide|compiler|interpreter|debugger|"
     r"simulation|visualization|jupyter|colab|pandas|numpy|scipy)\b",
     {"PO5": 1}),

    # Deep learning / neural network architectures → tools (PO5) + design (PO3)
    (r"\b(deep.learning|neural.network|cnn|rnn|lstm|gru|transformer|gan|"
     r"autoencoder|bert|gpt|llm|attention|backpropagation|convolution)\b",
     {"PO5": 1, "PO3": 1}),

    # Machine learning methods → analysis (PO2) + tools (PO5)
    (r"\b(machine.learning|supervised|unsupervised|reinforcement|regression|"
     r"classification|clustering|svm|random.forest|decision.tree|naive.bayes|"
     r"knn|k-means|dimensionality|pca|feature.extract|feature.select)\b",
     {"PO2": 1, "PO5": 1}),

    # Optimization methods → analysis (PO2) + design (PO3)
    (r"\b(optim\w+|loss.function|gradient|convergence|hyperparameter|tuning|"
     r"regulariz\w+|overfitting|underfitting|cross.valid\w+|grid.search)\b",
     {"PO2": 1, "PO3": 1}),

    # IoT / embedded / hardware → tools (PO5)
    (r"\b(iot|sensor|embedded|microcontroller|fpga|hardware|firmware|"
     r"edge.computing|real.time|interrupt|actuator|transducer)\b",
     {"PO5": 1}),

    # Lifelong learning reinforcement — self-directed study topics
    (r"\b(current|emerging|recent|state.of.the.art|latest|trend|advancement|"
     r"future|evolving|research.area|open.problem|frontier)\b",
     {"PO12": 1}),
]


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3: GATED POs — PO6 to PO11
#
# These POs are NEVER assigned unless the CO statement contains an explicit
# trigger keyword from the gate list below. This prevents "inflation" where
# every CO gets mapped to ethics, environment, teamwork, etc.
#
# Each gate entry: (PO_id, regex_pattern, strength_if_triggered)
# ═══════════════════════════════════════════════════════════════════════════════

_GATED_POS = [
    # PO6 — Engineer & Society
    # Only if CO explicitly addresses societal, legal, health, or safety impact.
    (
        "PO6",
        r"\b(society|social|community|public|human.impact|welfare|health|safety|"
        r"legal|cultural|civic|policy|regulation|standard|compliance|"
        r"accessib\w+|inclusion|diversity|equity)\b",
        2
    ),

    # PO7 — Environment & Sustainability
    # Only if CO explicitly mentions environmental or sustainability concepts.
    (
        "PO7",
        r"\b(environment\w*|sustain\w+|green|eco.?friendly|carbon|energy.effici\w+|"
        r"renewable|emission|waste|lifecycle|footprint|conservation|"
        r"biodiversity|climate|pollution)\b",
        2
    ),

    # PO8 — Ethics
    # Only if CO explicitly mentions ethical reasoning, responsibility, or bias.
    (
        "PO8",
        r"\b(ethic\w*|moral\w*|responsib\w+|bias|fairness|accountab\w+|"
        r"transparen\w+|integrity|privacy|consent|intellectual.property|"
        r"professional.conduct|code.of.conduct|security|cryptograph\w+|"
        r"encrypt\w+|authentication|vulnerability|threat|attack|malware)\b",
        2
    ),

    # PO9 — Individual & Team Work
    # Only if CO explicitly mentions collaborative or team activities.
    (
        "PO9",
        r"\b(team\w*|collaborat\w+|group|peer|cooperat\w+|collective|together|"
        r"co.?operat\w+|interdisciplin\w+|multidisciplin\w+|joint|"
        r"coordinat\w+|partner\w*)\b",
        2
    ),

    # PO10 — Communication
    # Only if CO explicitly mentions communication, reporting, or presentation.
    (
        "PO10",
        r"\b(report\w*|document\w*|present\w+|communicat\w+|write|writing|"
        r"publish|disseminat\w+|articul\w+|verbal|written|oral|"
        r"technical.writing|documentation|diagram|chart|visualiz\w+)\b",
        2
    ),

    # PO11 — Project Management & Finance
    # Only if CO explicitly mentions project planning, management, or resources.
    (
        "PO11",
        r"\b(project.manag\w*|manag\w+.project|schedule|budget|resource.plan\w*|deliverable|"
        r"milestone|agile|scrum|sprint|stakeholder|cost.estimat\w*|timeline|"
        r"feasibility|procurement|leadership.in.project)\b",
        2
    ),
]


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4: RULE ENGINE — applies Sections 1, 2, 3 to a single CO
# ═══════════════════════════════════════════════════════════════════════════════

_ALL_POS = [f"PO{i}" for i in range(1, 13)]


def _map_single_co(co_statement: str) -> dict:
    """
    Map one CO statement to PO strengths (0–3).
    Returns a dict {PO1: int, PO2: int, ..., PO12: int}.

    Reasoning is fully traceable:
      Step 1 → Bloom's verb sets base strengths for PO1–PO5, PO12.
      Step 2 → Subject keywords boost specific POs.
      Step 3 → Gated POs (PO6–PO11) are set ONLY if explicit triggers found.
               PO6–PO11 remain 0 if no gate is triggered.
    """
    text = co_statement.lower()
    mapping = {po: 0 for po in _ALL_POS}

    # ── Step 1: Bloom's verb ──────────────────────────────────────────────────
    verb_hit = False
    for pattern, scores in _VERB_MAP:
        if re.search(pattern, text):
            for po, val in scores.items():
                mapping[po] = val
            verb_hit = True
            break

    # No verb matched → treat as L2 (understand) as safe fallback
    if not verb_hit:
        mapping["PO1"] = 3
        mapping["PO2"] = 1
        mapping["PO12"] = 1

    # ── Step 2: Subject keyword boosts (PO1–PO5, PO12 only) ──────────────────
    for pattern, boosts in _KEYWORD_BOOSTS:
        if re.search(pattern, text):
            for po, delta in boosts.items():
                mapping[po] = min(3, mapping[po] + delta)

    # ── Step 3: Gated POs — PO6 to PO11 ─────────────────────────────────────
    # Always start at 0. Only assign if gate keyword explicitly found.
    for po_id, pattern, strength in _GATED_POS:
        mapping[po_id] = 0  # enforce: start clean
        if re.search(pattern, text):
            mapping[po_id] = strength

    return mapping


def rule_engine(cos: list, all_po_ids: list) -> dict:
    """
    Run the rule engine for all COs.

    Args:
        cos:         list of {co_id, co_statement}
        all_po_ids:  list of PO/PSO ids the frontend sent

    Returns:
        {co_id: {po_id: strength (0–3)}}
    """
    result = {}
    for co in cos:
        co_id = co["co_id"]
        stmt  = co.get("co_statement") or co.get("statement") or co.get("description") or ""
        scores = _map_single_co(stmt)
        # Map to the exact PO ids requested; PSOs default 0 (AI fills them)
        result[co_id] = {po: scores.get(po, 0) for po in all_po_ids}
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5: MERGE — Rule engine + AI
#
# Rules establish the baseline; AI acts as a second opinion.
#
# Decision table (per CO-PO cell):
#
#   Rule | AI  | Decision
#   -----|-----|----------------------------------------------------------
#    0   |  0  | 0  — both say no relationship → blank
#    0   | >0  | AI value  — AI found something rules missed → accept
#   >0   |  0  | Rule value — rules found relationship, AI missed it → keep
#   same | same| That value — full agreement
#   diff | diff| Weighted average (rules 60%, AI 40%), rounded to nearest int
#
# PO6–PO11 special rule:
#   If rule says 0 (gate not triggered) AND AI says >0:
#   → Accept AI's value BUT cap at 1 (low). AI may be pattern-matching loosely;
#     a cap of 1 prevents inflation while still recording a weak link.
# ═══════════════════════════════════════════════════════════════════════════════

# PO6–PO11 are GATED — rule engine keyword gate is the ONLY authority.
# AI output is completely ignored for these POs.
# They stay 0 unless the CO text explicitly contains a gate keyword.
_GATED_HARD = {"PO6", "PO7", "PO8", "PO9", "PO10", "PO11"}

# PO1–PO5, PO12 — core technical POs where AI is a valid second opinion.
_CORE_POS   = {"PO1", "PO2", "PO3", "PO4", "PO5", "PO12"}


def merge_rule_ai(rule_map: dict, ai_map: dict, co_ids: list, all_po_ids: list) -> dict:
    """
    Merge rule-based and AI mappings into a final CO-PO matrix.

    For PO6–PO11 (gated):
        Rule engine result is FINAL. AI is ignored entirely.
        This prevents the AI from inflating soft POs it has no evidence for.

    For PO1–PO5, PO12 (core technical POs):
        Both rule engine and AI are considered:
        - Both 0          → 0 (no relationship)
        - Rule 0, AI > 0  → take AI (AI found something rules missed)
        - Rule > 0, AI 0  → take rule (rule anchor holds)
        - Both > 0        → weighted blend: rules 60%, AI 40%, rounded
    """
    merged = {}
    for co_id in co_ids:
        r = rule_map.get(co_id, {})
        a = ai_map.get(co_id, {})
        merged[co_id] = {}
        for po in all_po_ids:
            rv = int(r.get(po, 0))
            av = int(a.get(po, 0))

            if po in _GATED_HARD:
                # Hard gate: rule engine is sole authority. AI ignored.
                val = rv

            elif rv == 0 and av == 0:
                val = 0

            elif rv == 0 and av > 0:
                # AI found something rules missed — trust AI for core POs
                val = av

            elif rv > 0 and av == 0:
                # Rule anchor holds
                val = rv

            else:
                # Both agree there is a relationship — weighted blend
                val = round(rv * 0.6 + av * 0.4)

            merged[co_id][po] = max(0, min(3, val))

    return merged
