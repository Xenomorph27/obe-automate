"""
backend/services/co_po_engine.py

Rule-based CO-PO mapping engine with AI validation layer.

Architecture:
  1. Rule engine  — verb extraction + keyword scanning → deterministic base mapping
  2. AI layer     — validates and fills gaps the rules couldn't cover
  3. Merge        — weighted blend: rules anchor, AI adjusts within ±1
  4. Universal    — PO1=3, PO12=1 always enforced at the end

Works for ANY engineering course (ML, IoT, Cybersecurity, Civil, etc.)
No course-specific hardcoding anywhere.
"""

import re
import json
import logging

logger = logging.getLogger(__name__)

# ── Bloom's verb → PO strength table ─────────────────────────────────────────
# Keys are regex patterns matched against the CO action verb.
# Values are {PO: strength} partial mappings.
# Order matters — first match wins for verb detection.

_VERB_MAP = [
    # Remember / Recall
    (r"\b(list|recall|define|state|identify|name|label|recognize|memorize)\b",
     {"PO1": 3, "PO2": 1, "PO12": 1}),

    # Understand / Explain
    (r"\b(explain|describe|summarize|interpret|classify|paraphrase|illustrate|understand)\b",
     {"PO1": 3, "PO2": 2, "PO12": 1}),

    # Apply
    (r"\b(apply|use|utilize|demonstrate|compute|solve|execute|implement|employ)\b",
     {"PO1": 3, "PO2": 2, "PO3": 2, "PO5": 2, "PO12": 1}),

    # Analyze / Discuss
    (r"\b(analyze|analyse|discuss|differentiate|compare|contrast|examine|categorize|dissect)\b",
     {"PO1": 3, "PO2": 3, "PO4": 2, "PO12": 1}),

    # Evaluate / Justify
    (r"\b(evaluate|justify|assess|critique|judge|defend|recommend|argue)\b",
     {"PO1": 3, "PO2": 3, "PO4": 3, "PO12": 1}),

    # Create / Design / Develop / Build
    (r"\b(design|develop|build|create|construct|formulate|synthesize|architect|engineer|model)\b",
     {"PO1": 3, "PO2": 2, "PO3": 3, "PO5": 2, "PO12": 1}),

    # Investigate / Experiment
    (r"\b(investigate|experiment|test|measure|simulate|verify|validate|benchmark)\b",
     {"PO1": 3, "PO2": 2, "PO4": 3, "PO5": 1, "PO12": 1}),
]

# ── Keyword → PO boost table ──────────────────────────────────────────────────
# Each entry: (regex pattern, {PO: delta})
# Deltas are ADDED to the verb base. Values are capped at 3.

_KEYWORD_BOOSTS = [
    # Technical domains → PO1 reinforcement
    (r"\b(algorithm|theorem|formula|principle|concept|theory|technique|method)\b",
     {"PO1": 1}),

    # Analysis / reasoning → PO2
    (r"\b(analyz|reason|infer|deduc|evaluat|assess|interpret|diagnos)\w*\b",
     {"PO2": 1}),

    # System / architecture / model design → PO3
    (r"\b(system|architecture|framework|pipeline|model|network|infrastructure|circuit|protocol)\b",
     {"PO3": 1}),

    # Data / experiments / measurement → PO4
    (r"\b(data|dataset|experiment|measurement|statistic|survey|observation|benchmark|metric)\b",
     {"PO4": 1}),

    # Tools / software / platforms → PO5
    (r"\b(tool|software|platform|library|framework|tensorflow|pytorch|sklearn|matlab|arduino|"
     r"raspberry|cloud|api|docker|kubernetes|simulation|ide|compiler)\b",
     {"PO5": 1}),

    # Deep learning / neural nets → PO5 extra
    (r"\b(deep.learning|neural.network|cnn|rnn|lstm|transformer|gan|autoencoder|bert|llm)\b",
     {"PO5": 1}),

    # IoT / embedded / hardware → PO5
    (r"\b(iot|sensor|embedded|microcontroller|fpga|hardware|firmware|edge.computing)\b",
     {"PO5": 1}),

    # Security / cryptography → PO8 + PO5
    (r"\b(security|cryptograph|encrypt|decrypt|authentication|authorization|firewall|"
     r"vulnerability|threat|attack|penetration|malware|forensic|privacy)\b",
     {"PO8": 1, "PO5": 1}),

    # Ethics explicitly mentioned → PO8
    (r"\b(ethic|moral|responsib|bias|fairness|accountab|transparen)\w*\b",
     {"PO8": 2}),

    # Societal / human impact → PO6
    (r"\b(society|social|community|human|impact|welfare|accessib|inclusion|diversity)\b",
     {"PO6": 2}),

    # Environment / sustainability → PO7
    (r"\b(environment|sustain|green|eco|carbon|energy.efficient|renewable|emission)\w*\b",
     {"PO7": 2}),

    # Team / collaboration → PO9
    (r"\b(team|collaborat|group|peer|cooperat|collective|together)\w*\b",
     {"PO9": 2}),

    # Communication / documentation → PO10
    (r"\b(report|document|present|communicat|write|publish|disseminat|articul)\w*\b",
     {"PO10": 2}),

    # Project management → PO11
    (r"\b(project|manage|schedule|budget|resource|plan|deliverable|milestone|agile|scrum)\b",
     {"PO11": 2}),

    # Clustering / unsupervised / dimensionality → PO2 + PO4
    (r"\b(cluster|unsupervised|dimensionality|reduction|anomaly|pattern.recognition)\b",
     {"PO2": 1, "PO4": 1}),

    # Optimization → PO2 + PO3
    (r"\b(optimiz|loss.function|gradient|convergence|hyperparameter|tuning|regulariz)\w*\b",
     {"PO2": 1, "PO3": 1}),
]

_ALL_POS = [f"PO{i}" for i in range(1, 13)]


def _rule_map_single(co_statement: str) -> dict:
    """Apply verb + keyword rules to a single CO statement. Returns {PO: 0-3}."""
    text = co_statement.lower()
    mapping = {po: 0 for po in _ALL_POS}

    # Step 1: verb detection — first match wins
    verb_scores = {}
    for pattern, scores in _VERB_MAP:
        if re.search(pattern, text):
            verb_scores = scores
            break

    # If no verb matched, default to "explain" level
    if not verb_scores:
        verb_scores = {"PO1": 3, "PO2": 1, "PO12": 1}

    for po, val in verb_scores.items():
        mapping[po] = val

    # Step 2: keyword boosts
    for pattern, boosts in _KEYWORD_BOOSTS:
        if re.search(pattern, text):
            for po, delta in boosts.items():
                mapping[po] = min(3, mapping[po] + delta)

    # Step 3: universal rules
    mapping["PO1"] = 3   # Engineering knowledge always 3 for technical COs
    mapping["PO12"] = max(1, mapping["PO12"])  # Lifelong learning always ≥ 1

    return mapping


def rule_engine(cos: list, all_po_ids: list) -> dict:
    """
    Run rule engine for all COs.
    cos: list of {co_id, co_statement}
    all_po_ids: list of PO ids to include (e.g. ["PO1"..."PO12","PSO1","PSO2"])
    Returns {co_id: {po_id: strength}}
    """
    result = {}
    for co in cos:
        co_id   = co["co_id"]
        stmt    = co.get("co_statement", co.get("description", ""))
        scores  = _rule_map_single(stmt)
        # Include only requested PO ids; PSOs default to 0 (AI will fill)
        result[co_id] = {po: scores.get(po, 0) for po in all_po_ids}
    return result


def merge_rule_ai(rule_map: dict, ai_map: dict, co_ids: list, all_po_ids: list) -> dict:
    """
    Merge rule-based and AI mappings.

    Strategy:
    - PO1 and PO12: rule engine always wins (universal NBA rules)
    - For other POs:
        * If rule says 0 and AI says >0: take AI (AI found something rules missed)
        * If rule says >0 and AI says 0: keep rule (rules are anchored)
        * If both agree: use that value
        * If they disagree by 1: average (round up)
        * If they disagree by >1: trust rule, nudge by 1 toward AI
    """
    merged = {}
    for co_id in co_ids:
        r = rule_map.get(co_id, {})
        a = ai_map.get(co_id, {})
        merged[co_id] = {}
        for po in all_po_ids:
            rv = int(r.get(po, 0))
            av = int(a.get(po, 0))

            if po == "PO1":
                merged[co_id][po] = 3  # Always
            elif po == "PO12":
                merged[co_id][po] = 1  # Always
            elif rv == 0 and av > 0:
                merged[co_id][po] = av  # AI found something
            elif rv > 0 and av == 0:
                merged[co_id][po] = rv  # Rule anchor holds
            elif rv == av:
                merged[co_id][po] = rv  # Agreement
            elif abs(rv - av) == 1:
                merged[co_id][po] = max(rv, av)  # Small diff → round up
            else:
                # Large disagreement → trust rule, nudge 1 toward AI
                merged[co_id][po] = rv + (1 if av > rv else -1)

            # Clamp
            merged[co_id][po] = max(0, min(3, merged[co_id][po]))

    return merged
