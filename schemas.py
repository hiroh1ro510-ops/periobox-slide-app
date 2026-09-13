"""JSON schemas and enums for the PerioBox slide pipeline.

The pipeline now uses a 3-layer schema for the design (slide 2) and result
(slide 3) sections:

  1. Common core   - required for every paper regardless of study design.
  2. Pattern class - one of STUDY_TYPE_CLASSES, decided by the LLM and
                     verified by a second LLM pass.
  3. Pattern-specific extensions - only the fields relevant to the class
                     are populated (groups, comparison_matrix, pico, ...).

The old flat ``slide2 / slide3`` fields are kept alongside the new
``study_type_class`` / ``pattern_extensions`` so existing renderers keep
working while the new templates are being built (Phase 2).
"""

from __future__ import annotations

STUDY_TYPE_CLASSES = [
    "rct",
    "prospective_cohort",
    "retrospective_comparative",
    "cross_sectional",
    "systematic_review",
    "narrative_review",
    "mechanism",
    "case_report",
]

STUDY_TYPE_LABEL_JA = {
    "rct": "ランダム化比較試験",
    "prospective_cohort": "前向きコホート研究",
    "retrospective_comparative": "後ろ向き比較研究",
    "cross_sectional": "横断研究",
    "systematic_review": "システマティックレビュー",
    "narrative_review": "ナラティブレビュー",
    "mechanism": "機序・組織学的研究",
    "case_report": "症例報告",
}


FAITHFUL_SUMMARY_JSON_SHAPE = r"""
{
  "paper": {
    "title_en": "英語タイトル",
    "citation_short": "筆頭著者 et al. ジャーナル名. 年",
    "article_type": "論文タイプ"
  },
  "faithful_summary": {
    "introduction_ja": "Introductionを約400字で忠実に要約",
    "clinical_question_source": "本文から導ける臨床疑問",
    "methods_ja": "Methodsの忠実な日本語要約",
    "participants_ja": "対象",
    "interventions_or_groups_ja": ["群・介入・曝露1", "群・介入・曝露2"],
    "outcomes_ja": ["評価項目1", "評価項目2"],
    "results_ja": "Resultsの忠実な日本語要約",
    "key_numeric_data_ja": ["数値1", "数値2", "数値3"],
    "statistical_notes_ja": ["統計1", "統計2"],
    "discussion_ja": "Discussionの忠実な日本語要約",
    "limitations_ja": ["限界1", "限界2"],
    "figures_tables_ja": "図表の忠実な日本語要約",

    "study_type_class": "rct | prospective_cohort | retrospective_comparative | cross_sectional | systematic_review | narrative_review | mechanism | case_report",
    "study_type_label_ja": "多施設・後ろ向き研究 のような短いサブタイトルピル用文言",
    "study_type_evidence": {
      "source_quote": "分類の根拠になった原文の英語1文",
      "source_page": 3
    },

    "design_common": {
      "subjects_short": "対象を1〜2行で",
      "methods_short": "方法を3〜5行で",
      "evaluation_items": ["評価項目ピル用短語1", "評価項目ピル用短語2"],
      "figure_refs": [
        {"page": 3, "label": "Figure 2", "caption_ja": "図の内容", "source_quote": "英語Figure legend抜粋"}
      ]
    },

    "result_common": {
      "main_conclusion_ja": "1〜3文の主要結論",
      "detail_note_ja": "『さらに詳しく』欄の補足解説1〜2文",
      "figure_refs": [
        {"page": 5, "label": "Figure 4", "caption_ja": "図の内容", "source_quote": "英語Figure legend抜粋"}
      ]
    },

    "pattern_extensions": {
      "groups": [
        {
          "name": "群名",
          "n": "n数（不明なら空文字）",
          "intervention_or_exposure": "介入・曝露内容",
          "source_quote": "群記述の英語原文1文",
          "source_page": 3
        }
      ],
      "comparison_matrix": {
        "row_axis": "行の軸（例: 群）",
        "column_axis": "列の軸（例: 指標）",
        "rows": ["行ラベル1", "行ラベル2"],
        "columns": ["列ラベル1", "列ラベル2"],
        "cells": [["値11", "値12"], ["値21", "値22"]],
        "significance_note": "有意差の記述",
        "source_quote": "対応するResultsの英語原文1文",
        "source_page": 5
      },
      "pico": {
        "population": "P",
        "intervention": "I",
        "comparison": "C",
        "outcome": "O",
        "included_studies_n": "採用研究数",
        "search_flow": {"identified": "n", "screened": "n", "included": "n"},
        "source_quote": "英語原文1文",
        "source_page": 2
      },
      "classification_tree": {
        "title": "分類名",
        "nodes": [
          {"label": "分岐条件", "children": [{"label": "サブ分岐", "children": []}]}
        ],
        "source_quote": "英語原文1文",
        "source_page": 2
      },
      "case_report": {
        "patient_profile": "年齢・性・主訴など",
        "procedure_steps": ["術式ステップ1", "術式ステップ2"],
        "requirements": ["適応要件1"],
        "advantages": ["利点1"],
        "disadvantages": ["欠点1"],
        "source_quote": "英語原文1文",
        "source_page": 4
      }
    }
  }
}
""".strip()


SLIDE_SUMMARY_JSON_SHAPE = r"""
{
  "paper": {
    "title_en": "英語論文タイトル",
    "citation": "筆頭著者 et al. ジャーナル名. 年",
    "article_type": "RCT / Systematic Review / Retrospective study / Case report 等"
  },
  "slide1": {
    "clinical_question_jp": "日本語の問い",
    "key_message_jp": "問いへの結論を1文で",
    "background_heading": "論文の背景と概要",
    "background_overview": "Introductionを中心に、臨床上の課題・既存知見・未解決点・本研究目的を約400字で要約"
  },
  "slide2": {
    "design_summary": "一般的な研究デザイン名",
    "study_type_class": "rct | prospective_cohort | retrospective_comparative | cross_sectional | systematic_review | narrative_review | mechanism | case_report",
    "study_type_label": "多施設・後ろ向き研究 のようなサブタイトルピル用文言",
    "participants": "対象患者・対象資料",
    "selection_criteria": ["選定条件1", "選定条件2", "除外条件1"],
    "methods": ["方法1", "方法2", "方法3"],
    "groups_or_comparison": ["群分け・比較1", "群分け・比較2"],
    "outcomes": ["評価項目1（略称）", "評価項目2（略称）"],
    "design_schema": {
      "title": "中段図表の短いタイトル",
      "steps": ["対象/処置の流れ1", "流れ2"],
      "groups": [{"name": "群1", "description": "短い説明"}],
      "labels": ["図中ラベル1"]
    },
    "pattern_extensions": {
      "groups_detail": [
        {"name": "群名", "n": "n数", "intervention_or_exposure": "介入・曝露"}
      ],
      "pico": {"population": "P", "intervention": "I", "comparison": "C", "outcome": "O", "included_studies_n": "n"},
      "classification_tree": {"title": "分類名", "nodes": [{"label": "分岐条件", "children": []}]},
      "case_report": {"patient_profile": "...", "procedure_steps": ["..."]}
    },
    "figure_refs": [{"page": 3, "label": "Figure 2", "caption_ja": "..."}]
  },
  "slide3": {
    "main_results": ["統計的有意差の有無を含む主な結果1", "主な結果2"],
    "key_numbers": ["重要な数値1"],
    "statistical_notes": ["統計的有意差・P値1"],
    "evidence_schema": {
      "title": "主要結果の根拠を示す短いタイトル",
      "items": [
        {"label": "比較軸・評価項目", "value": "数値または所見", "interpretation": "臨床的な読み取り"}
      ]
    },
    "more_detail": "さらに詳しく：図表や補足結果の要約を1〜2文",
    "pattern_extensions": {
      "comparison_matrix": {
        "row_axis": "行の軸",
        "column_axis": "列の軸",
        "rows": ["群1", "群2"],
        "columns": ["指標1", "指標2"],
        "cells": [["値11", "値12"], ["値21", "値22"]],
        "significance_note": "有意差記述"
      },
      "case_report": {"requirements": ["..."], "advantages": ["..."], "disadvantages": ["..."]}
    },
    "figure_refs": [{"page": 5, "label": "Figure 4", "caption_ja": "..."}]
  },
  "slide4": {
    "selection_points": ["論文選定のポイント1", "論文選定のポイント2"],
    "doctor_comment": "抄読ドクターからひとこと。5〜8文",
    "limitations": ["限界1", "限界2"]
  }
}
""".strip()


STUDY_TYPE_VERIFIER_JSON_SHAPE = r"""
{
  "study_type_class": "rct | prospective_cohort | retrospective_comparative | cross_sectional | systematic_review | narrative_review | mechanism | case_report",
  "study_type_label_ja": "サブタイトルピル用文言（例: 多施設・後ろ向き研究）",
  "verdict": "correct | corrected | uncertain",
  "reason_ja": "判定理由を日本語で1〜2文",
  "source_quote": "分類の根拠になった原文の英語1文",
  "source_page": 3
}
""".strip()
