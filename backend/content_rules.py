from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Mapping


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


# V10 : les reformulations de questions (ex-recouvrements B01/B09 et B06B/B12)
# sont désormais écrites directement dans la feuille 50_Questions de l'Excel
# de paramétrage. Il n'y a plus de dictionnaire codé en dur ici : modifier une
# question dans 50_Questions suffit à changer le comportement de l'outil,
# sans toucher au code. Ce dict est conservé vide pour compatibilité API.
QUESTION_OVERRIDES: Dict[str, Dict[str, Any]] = {}


FALLBACK_GLOSSARY: Dict[str, Dict[str, Any]] = {}

# Le glossaire métier est fourni exclusivement par 40_Glossaire.
# Aucun terme, synonyme ni définition de secours n'est défini dans Python.



def normalize_questions_dataframe(df):
    """Applique les libellés canoniques tout en conservant les identifiants."""
    if df is None or getattr(df, "empty", True):
        return df
    out = df.copy()
    if "id_question" not in out.columns:
        return out
    for idx, row in out.iterrows():
        qid = clean(row.get("id_question"))
        override = QUESTION_OVERRIDES.get(qid)
        if not override:
            continue
        for key, value in override.items():
            if key in out.columns:
                out.at[idx, key] = value
    return out


def _split_aliases(value: Any) -> List[str]:
    text = clean(value)
    if not text:
        return []
    return [clean(item) for item in re.split(r"[;,|]", text) if clean(item)]


def glossary_entries_from_dataframe(df) -> List[Dict[str, Any]]:
    entries: Dict[str, Dict[str, Any]] = {}
    if df is not None and not getattr(df, "empty", True):
        for _, row in df.iterrows():
            term = clean(row.get("terme"))
            definition = clean(row.get("definition") or row.get("definition_excel"))
            if not term or not definition:
                continue
            aliases = []
            aliases.extend(_split_aliases(row.get("mots_cles")))
            aliases.extend(_split_aliases(row.get("synonymes")))
            entries[term.casefold()] = {
                "term": term,
                "definition": definition,
                "source": clean(row.get("source")),
                "practice": clean(row.get("explication_tpe_pme")),
                "category": clean(row.get("categorie")),
                "aliases": sorted({a for a in aliases if a and a.casefold() != term.casefold()}, key=len, reverse=True),
            }
    return sorted(entries.values(), key=lambda item: item["term"].casefold())


def glossary_lookup(entries: Iterable[Mapping[str, Any]]) -> Dict[str, Dict[str, Any]]:
    lookup: Dict[str, Dict[str, Any]] = {}
    for raw in entries or []:
        term = clean(raw.get("term") or raw.get("terme"))
        definition = clean(raw.get("definition") or raw.get("definition_excel"))
        if not term or not definition:
            continue
        entry = {
            "term": term,
            "definition": definition,
            "source": clean(raw.get("source") or raw.get("source_excel")),
            "practice": clean(raw.get("practice") or raw.get("explication_tpe_pme")),
            "category": clean(raw.get("category") or raw.get("categorie")),
            "aliases": list(raw.get("aliases") or []),
        }
        for alias in [term, *entry["aliases"]]:
            alias = clean(alias)
            if alias:
                lookup.setdefault(alias.casefold(), entry)
    return lookup
