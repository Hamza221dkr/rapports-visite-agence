"""
Module d'intégration OpenAI GPT — génère des analyses enrichies pour les rapports.
"""

import os

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


def _build_agency_summary(agency_name, thematiques, note_globale, notation_globale) -> str:
    """Construit un résumé textuel de l'agence pour le prompt GPT."""
    lines = [f"Agence : {agency_name}",
             f"Note globale : {note_globale:.2f}/3 ({notation_globale})" if note_globale else ""]

    THEME_MAP_SHORT = {
        "FONCTIONNEMENT": "Fonctionnement agence", "QUALITE": "Qualité de service",
        "ENVIRONNEMENT": "Sécurité & Risques", "ORGANISATION": "Organisation commerciale",
        "PERFORMANCES": "Performance", "MANAGEMENT": "Management",
    }

    for theme_raw, data in thematiques.items():
        key = next((k for k in THEME_MAP_SHORT if k in theme_raw.upper()), theme_raw)
        label = THEME_MAP_SHORT.get(key, theme_raw)
        note = data.get("note")
        note_str = f"{note:.2f}" if note else "—"
        lines.append(f"\n[{label} — {note_str}/3]")
        for q in data["questions"]:
            if q["scoring"] in ("À améliorer", "Bon") or q["obs"]:
                lines.append(f"  • [{q['scoring']}] {q['question'][:80]}")
                if q["obs"] and q["scoring"] == "À améliorer":
                    lines.append(f"    → {q['obs'][:120]}")
    return "\n".join(lines)


def _analyze_individual(client, agency_name, thematiques, note_globale, notation_globale) -> str:
    """Génère une conclusion enrichie pour un rapport individuel."""
    summary = _build_agency_summary(agency_name, thematiques, note_globale, notation_globale)

    prompt = f"""Tu es un expert bancaire chargé de rédiger des comptes rendus de visite d'agences.
Voici les données de visite de l'agence {agency_name} :

{summary}

Rédige une appréciation générale professionnelle et synthétique (3 à 4 paragraphes) qui :
1. Donne une appréciation globale du niveau de fonctionnement de l'agence
2. Met en valeur les points forts observés
3. Souligne les axes d'amélioration prioritaires avec des recommandations concrètes
4. Conclut avec une perspective motivante pour les équipes

Style : professionnel, bienveillant mais direct. En français. Sans bullet points, uniquement du texte rédigé."""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=600,
            temperature=0.7,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return None


def _analyze_consolidated(client, all_agencies: list) -> str:
    """Génère une synthèse consolidée pour le rapport de zone."""
    summaries = []
    for ag in all_agencies:
        summaries.append(_build_agency_summary(
            ag["name"], ag["thematiques"], ag.get("note_globale"), ag.get("notation_globale")
        ))

    combined = "\n\n---\n\n".join(summaries)
    avg = sum(ag.get("note_globale") or 0 for ag in all_agencies) / max(len(all_agencies), 1)

    prompt = f"""Tu es un Responsable de Zone bancaire qui rédige le rapport de synthèse après une tournée de visite de {len(all_agencies)} agences.

Voici les données de toutes les agences visitées :

{combined}

Note moyenne de zone : {avg:.2f}/3

Rédige une appréciation générale de zone (4 à 5 paragraphes) qui :
1. Présente le bilan global de la tournée avec la note moyenne de zone
2. Identifie les agences performantes et celles qui nécessitent un suivi prioritaire
3. Met en exergue les problématiques transversales communes à plusieurs agences
4. Formule des recommandations stratégiques adressées à la Direction Régionale et aux entités support
5. Fixe des objectifs et un cadre de suivi pour les prochaines semaines

Style : managérial, structuré, orienté résultats. En français. Texte rédigé, pas de bullet points."""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=800,
            temperature=0.7,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return None
