SUBJECT = "Absences en Sport"

BODY_TEMPLATE = """Bonjour {prenom},

Je me permets de te contacter pour faire le point sur tes absences en SPORT depuis septembre.
Sauf erreur de ma part, tu affiches {nb} absences injustifiées ({dates}).

Pour rappel, le règlement de scolarité de l'ENSTA autorise au maximum 2 absences non-justifiées, et à partir de la troisième, la note au bloc SPORT sera inférieure à 10/20.
Aurais-tu des documents médicaux ou autres à me transmettre pour justifier ces absences ?

Je reste disponible pour en discuter si tu le souhaites.

Cordialement,

"""


def format_dates(dates: list[str]) -> str:
    if not dates:
        return "dates non précisées"
    if len(dates) == 1:
        return dates[0]
    if len(dates) == 2:
        return f"{dates[0]} et {dates[1]}"
    return ", ".join(dates[:-1]) + f" et {dates[-1]}"
