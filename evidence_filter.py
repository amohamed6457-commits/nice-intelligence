"""
Filter appraisals by WHICH endpoint categories had no usable data.

Two different things get called "insufficient", and a filter that mixes them
is worse than no filter at all:

  * G-BA "keine bewertbaren Daten" -- the committee could NOT assess the
    evidence. An evidence-quality problem, and the thing worth filtering on.
  * HAS "SMR insuffisant" -- the committee DID assess, and judged the
    clinical benefit inadequate for reimbursement. That is a decision, not a
    gap, and 2,026 appraisals carry it. Filtering those out would throw away
    real verdicts.

So the G-BA filter asks which categories were unassessable, and the HAS
filter asks whether a grade was recorded at all. Different questions, and
the labels say so.

G-BA, across 1,051 appraisals with verdicts, no usable data in:

    Mortality         498   47%
    Morbidity         539   51%
    Quality of life   711   68%
    Adverse events    519   49%
    all four          468   45%

Ticking several means ALL of them were missing, so the filter narrows as you
add categories rather than widening. Ticking all four gives the 468
appraisals where nothing at all could be assessed -- the evidence-gap
population, which is a selection worth having in its own right and not only
something to hide.

A category counts as missing for an appraisal when NO subgroup produced
usable data for it. An appraisal where one subgroup had quality-of-life data
and another did not is not counted as missing: something usable exists.

    import evidence_filter
    evidence_filter.attach_gba(subgroups_frame)
    evidence_filter.attach_has(ratings_frame)
    ids = evidence_filter.ids_for("G-BA", ["Quality of life"])
"""

import endpoint_summary

# One chip per endpoint category, in the order G-BA reports them.
GBA_OPTIONS = [name for _field, name in endpoint_summary.CATEGORIES]

HAS_OPTIONS = [
    "Has a recorded SMR grade",
    "No grade recorded",
]

# Ratings that record no judgement at all.
_NO_GRADE = {"non précisé", "non précisée", "commentaires", "sans objet",
             "non applicable au dossier",
             "commentaires sans chiffrage de l'asmr", "nan", ""}

_GBA = None     # appraisal_id -> {category: True when nothing usable}
_HAS = None     # appraisal_id -> True when a real SMR grade exists


def attach_gba(subgroups):
    global _GBA
    _GBA = None
    if subgroups is None or not len(subgroups):
        return
    try:
        index = {}
        ids = subgroups["appraisal_id"].astype(str).str.strip()
        for key, group in subgroups.groupby(ids):
            missing = {}
            for field, name in endpoint_summary.CATEGORIES:
                verdicts = [endpoint_summary.verdict(getattr(row, field, None))
                            for row in group.itertuples()]
                verdicts = [v for v in verdicts if v]
                if verdicts:
                    # Missing only when NO subgroup produced usable data.
                    missing[name] = all(v == "NO DATA" for v in verdicts)
            if missing:
                index[key] = missing
        _GBA = index
    except Exception:                                      # noqa: BLE001
        _GBA = None


def attach_has(ratings):
    global _HAS
    _HAS = None
    if ratings is None or not len(ratings):
        return
    try:
        frame = ratings[ratings["rating_type"].astype(str).str.upper() == "SMR"]
        ids = frame["appraisal_id"].astype(str).str.strip()
        grade = frame["rating"].astype(str).str.strip().str.lower()
        real = ~grade.isin(_NO_GRADE)
        index = {}
        for key, ok in zip(ids, real):
            index[key] = index.get(key, False) or bool(ok)
        _HAS = index
    except Exception:                                      # noqa: BLE001
        _HAS = None


def options_for(agency):
    if agency == "G-BA" and _GBA:
        return list(GBA_OPTIONS)
    if agency == "HAS" and _HAS:
        return list(HAS_OPTIONS)
    return []


def counts_for(agency, direction=None):
    """How many appraisals each option matches ON ITS OWN, for the chip
    labels. Ticking two does not give the sum: the filter is an AND, so the
    result is smaller than either count."""
    return {option: len(ids_for(agency, [option], direction) or ())
            for option in options_for(agency)}


DIRECTIONS = ["was missing", "was assessable"]


def ids_for(agency, selection, direction=None):
    """Appraisal ids matching the selection, or None for no filter.

    None means "do not filter" and is returned when nothing is selected or
    the agency has no index, so a caller treats None as a no-op rather than
    as an empty result.

    For G-BA the selected categories are combined with AND, in whichever
    direction is asked for: every ticked category was missing, or every
    ticked category was assessable. Both searches are useful and neither is
    the obvious default, so the direction is the caller's to choose. For HAS
    the two options are mutually exclusive states, so they combine with OR
    and direction does not apply.
    """
    picked = [s for s in (selection or []) if s]
    if not picked:
        return None

    if agency == "G-BA":
        if not _GBA:
            return None
        wanted = [p for p in picked if p in GBA_OPTIONS]
        if not wanted:
            return None
        want_missing = str(direction or DIRECTIONS[0]) != DIRECTIONS[1]
        return {key for key, missing in _GBA.items()
                if all(bool(missing.get(name)) == want_missing
                       for name in wanted)}

    if agency == "HAS":
        if not _HAS:
            return None
        keep = set()
        for key, graded in _HAS.items():
            if "Has a recorded SMR grade" in picked and graded:
                keep.add(key)
            if "No grade recorded" in picked and not graded:
                keep.add(key)
        return keep

    return None


def label_for(agency):
    if agency == "G-BA":
        return "Endpoint evidence — tick the categories you care about"
    return "SMR grade recorded — tick to narrow (none ticked = all)"


def direction_label():
    return "...where that evidence"


def help_for(agency):
    if agency == "G-BA":
        return ("Which endpoint categories the G-BA could or could not "
                "assess. Ticking several narrows rather than widens: every "
                "ticked category must be in the chosen state. 'Was missing' "
                "on all four gives the 468 appraisals where nothing could be "
                "assessed; 'was assessable' on all four gives the 204 where "
                "everything could. A category counts as missing only when no "
                "subgroup produced usable data. This is about evidence "
                "quality, not about whether a benefit was granted — use "
                "Outcomes for that.")
    return ("Whether HAS recorded an SMR grade at all. This is NOT the same "
            "as 'SMR insuffisant', which is a real verdict on 2,026 "
            "appraisals and is kept by both options here.")
