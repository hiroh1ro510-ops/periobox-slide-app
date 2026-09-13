"""Route ``study_type_class`` to the right slide-2/slide-3 renderer.

Falls back to the legacy renderers (``app.add_research_design_slide`` /
``app.add_research_results_slide``) when no dedicated template exists yet.
"""

from typing import Any, Callable, Dict, Optional

from . import design_comparative, result_comparative


COMPARATIVE_CLASSES = {
    "rct",
    "prospective_cohort",
    "retrospective_comparative",
    "mechanism",
}


def get_design_renderer(study_type_class: str) -> Optional[Callable]:
    if study_type_class in COMPARATIVE_CLASSES:
        return design_comparative.render
    return None


def get_result_renderer(study_type_class: str) -> Optional[Callable]:
    if study_type_class in COMPARATIVE_CLASSES:
        return result_comparative.render
    return None


def resolve_class(summary: Dict[str, Any]) -> str:
    slide2 = summary.get("slide2", {}) if isinstance(summary, dict) else {}
    return str(slide2.get("study_type_class", "")).strip()
