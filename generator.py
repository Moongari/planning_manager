import random
import pandas as pd

def generate_planning_custom_names(person_names, weeks=12):
    jours = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche']
    horaires_jour = [("09:00", "12:30"), ("13:30", "17:30"), ("10:00", "13:30"), ("14:00", "18:00")]

    planning_data = []
    n_persons = len(person_names)

    for semaine in range(1, weeks + 1):
        for p in range(n_persons):
            person = person_names[p]

            # Choisir aléatoirement 2 jours consécutifs de repos (index 0 à 5)
            repos_start_idx = random.randint(0, 5)
            repos_days = [repos_start_idx, repos_start_idx + 1]

            jours_semaine = []
            for j in range(7):
                if j in repos_days:
                    jours_semaine.append("Repos")
                else:
                    hd, hf = random.choice(horaires_jour)
                    jours_semaine.append(f"{hd} - {hf}")

            planning_data.append([semaine, person] + jours_semaine)

    columns = ['Semaine', 'Personne'] + jours
    df = pd.DataFrame(planning_data, columns=columns)
    return df
