from Etudiants import Eleve, CLASSES
from notes import notes, MATIERES


def demander_classe():
    print("Choisissez la classe :")
    for numero, classe in enumerate(CLASSES, start=1):
        print(f"{numero}. {classe}")
    while True:
        choix = input("Votre choix (1-4) : ").strip()
        if choix in {"1", "2", "3", "4"}:
            return CLASSES[int(choix) - 1]
        print("Choix invalide. Veuillez entrer un numéro entre 1 et 4.")


def demander_notes():
    valeurs = []
    for matiere in MATIERES:
        while True:
            texte = input(f"Note en {matiere} (sur 20) : ").strip().replace(",", ".")
            try:
                valeur = float(texte)
                if 0 <= valeur <= 20:
                    valeurs.append(valeur)
                    break
                print("La note doit être entre 0 et 20.")
            except ValueError:
                print("Veuillez entrer un nombre valide.")
    return valeurs


def main():
    while True:
        print("\n===== GESTION DES ÉLÈVES DU COLLÈGE =====")
        print("1. Inscrire un élève")
        print("2. Ajouter ou modifier les notes")
        print("3. Quitter")
        choix = input("Choisissez une opération : ").strip()

        try:
            if choix == "1":
                nom = input("Nom : ").strip()
                prenom = input("Prénom : ").strip()
                classe = demander_classe()
                Eleve(nom, prenom, classe).inscrire()
            elif choix == "2":
                nom = input("Nom : ").strip()
                prenom = input("Prénom : ").strip()
                classe = demander_classe()
                valeurs = demander_notes()
                notes(nom, prenom, classe, valeurs).enregistrer()
            elif choix == "3":
                print("Fin du programme.")
                break
            else:
                print("Choix invalide. Choisissez 1, 2 ou 3.")
        except (ValueError, FileNotFoundError, PermissionError, OSError) as erreur:
            print(f"Erreur : {erreur}")
            print("Vérifiez que le fichier Excel est fermé dans Excel puis réessayez.")


if __name__ == "__main__":
    main()
