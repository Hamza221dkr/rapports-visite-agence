"""
Module d'intégration OpenAI GPT — génère des analyses enrichies pour les rapports.
"""


def get_gpt_fn(api_key: str):
    """Retourne une fonction GPT si la clé est valide, sinon None."""
    if not api_key or not api_key.strip().startswith("sk-"):
        return None
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key.strip())
    except ImportError:
        return None

    def gpt_fn(mode: str, *args):
        if mode == "individual":
            return _analyze_individual(client, *args)
        elif mode == "consolidated":
            return _analyze_consolidated(client, *args)
        return None

    return gpt_fn


# ── Prompt système commun ───────────────────────────────────────────────────

SYSTEM_PROMPT = """Tu es un expert bancaire senior spécialisé dans l'audit et le contrôle des agences.
Tu rédiges des comptes rendus de visite officiels pour une Direction Régionale.

RÈGLES ABSOLUES — à respecter sans exception :
- Rédige uniquement en français, texte courant, style professionnel et direct.
- INTERDIT : markdown, astérisques, dièses, tirets de liste, numérotation, sous-titres.
- INTERDIT : toute phrase d'introduction du type "Voici l'appréciation" ou "En conclusion,".
- INTERDIT : répéter les données brutes (notes chiffrées, listes de critères).
- Commence directement par le contenu, sans formule introductive.
- Chaque paragraphe est séparé par un saut de ligne. Aucun autre formatage."""


# ── Helpers ─────────────────────────────────────────────────────────────────

def _build_agency_summary(agency_name, thematiques, note_globale, notation_globale) -> str:
    THEME_MAP = {
        "FONCTIONNEMENT": "Fonctionnement agence",
        "QUALITE":        "Qualité de service",
        "ENVIRONNEMENT":  "Sécurité & Risques opérationnels",
        "ORGANISATION":   "Organisation du travail",
        "PERFORMANCES":   "Performance Commerciale",
        "MANAGEMENT":     "Management & Leadership",
    }
    lines = [f"Agence : {agency_name}"]
    if note_globale:
        lines.append(f"Appréciation globale : {notation_globale} ({note_globale:.2f}/3)")

    for theme_raw, data in thematiques.items():
        key   = next((k for k in THEME_MAP if k in theme_raw.upper()), theme_raw)
        label = THEME_MAP.get(key, theme_raw)
        note  = data.get("note")
        lines.append(f"\n{label} — {f'{note:.2f}/3' if note else 'non noté'}")
        for q in data["questions"]:
            if q["scoring"] in ("À améliorer", "Bon") or q["obs"]:
                lines.append(f"  [{q['scoring']}] {q['question'][:90]}")
                if q["obs"] and q["scoring"] == "À améliorer":
                    lines.append(f"    Observation : {q['obs'][:150]}")
    return "\n".join(lines)


def _call(client, system: str, user: str, max_tokens: int) -> str | None:
    try:
        r = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system},
                {"role": "user",   "content": user},
            ],
            max_tokens=max_tokens,
            temperature=0.5,
        )
        return r.choices[0].message.content.strip()
    except Exception:
        return None


# ── Rapport individuel ───────────────────────────────────────────────────────

def _analyze_individual(client, agency_name, thematiques, note_globale, notation_globale) -> str | None:
    summary = _build_agency_summary(agency_name, thematiques, note_globale, notation_globale)

    user_prompt = f"""Données de la visite de l'agence {agency_name} :

{summary}

Rédige une appréciation générale en 3 paragraphes :
- Paragraphe 1 : bilan global du fonctionnement de l'agence, niveau de maîtrise des thématiques.
- Paragraphe 2 : points forts constatés et axes d'amélioration prioritaires avec recommandations concrètes.
- Paragraphe 3 : message de conclusion engageant pour l'équipe, avec les prochaines attentes.

Longueur : 150 à 220 mots. Texte continu, aucun formatage."""

    return _call(client, SYSTEM_PROMPT, user_prompt, max_tokens=400)


# ── Rapport consolidé ────────────────────────────────────────────────────────

def _analyze_consolidated(client, all_agencies: list) -> str | None:
    summaries = "\n\n---\n\n".join(
        _build_agency_summary(ag["name"], ag["thematiques"], ag.get("note_globale"), ag.get("notation_globale"))
        for ag in all_agencies
    )
    avg = sum(ag.get("note_globale") or 0 for ag in all_agencies) / max(len(all_agencies), 1)
    names = ", ".join(ag["name"] for ag in all_agencies)

    user_prompt = f"""Données de la tournée de visite — {len(all_agencies)} agences : {names}
Note moyenne de zone : {avg:.2f}/3

{summaries}

Rédige une appréciation générale de zone en 4 paragraphes :
- Paragraphe 1 : bilan de la tournée, note moyenne, tendance générale de la zone.
- Paragraphe 2 : agences performantes et agences nécessitant un accompagnement prioritaire.
- Paragraphe 3 : problématiques transversales communes, vigilances partagées.
- Paragraphe 4 : recommandations stratégiques pour la Direction Régionale et les entités support, objectifs pour les prochaines semaines.

Longueur : 200 à 280 mots. Texte continu, aucun formatage."""

    return _call(client, SYSTEM_PROMPT, user_prompt, max_tokens=550)
