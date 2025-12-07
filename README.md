# Logiciel de Génération de Bilans Psychologiques

## Description
Ce logiciel permet aux psychologues de générer automatiquement des bilans psychologiques au format **Word (.docx)**. Il utilise une interface graphique pour saisir les informations des patients et produit un document structuré et personnalisé.

---

## Fonctionnalités
- **Interface graphique** (PyQt5) pour la saisie des données.
- **Génération automatique** de documents Word avec mise en forme personnalisée.
- **Sections incluses** :
  - Indices cognitifs (ICC, INV, etc.)
  - Capacité verbale
  - Visuo-spatial
  - Raisonnement fluide
  - Mémoire de travail
  - Vitesse de traitement

---

## Prérequis
- Python 3.8 ou supérieur
- Bibliothèques Python requises :
  - `python-docx`
  - `PyQt5`

### Installation des dépendances
```bash
pip install python-docx PyQt5
```

---

## Utilisation

### 1. Lancer l'application
Exécute le fichier principal :
```bash
python bilan.py
```

### 2. Saisie des données
- Remplis les champs de l'interface graphique avec les informations du patient.
- Clique sur le bouton **"Générer le bilan"** pour créer le document Word.

### 3. Récupération du document
Le document généré est sauvegardé dans le dossier `bilans/` sous le nom spécifié.

---

## Structure du Projet
- **`bilan.py`** : Interface graphique (PyQt5) et gestion des entrées utilisateur.
- **`autoBilan.py`** : Fonctions pour la génération et la mise en forme du document Word.

---

## Exemple de Sortie
Le document généré inclut :
- Un en-tête personnalisé avec le nom du patient.
- Des sections détaillées pour chaque test psychologique.
- Des tableaux et des notes standardisées.

---

## Contribution
Les contributions sont les bienvenues ! Pour proposer des améliorations :
1. Fork le projet.
2. Crée une branche pour ta fonctionnalité (`git checkout -b feature/ma-fonctionnalité`).
3. Commit tes modifications (`git commit -m "Ajout de ma fonctionnalité"`).
4. Push vers la branche (`git push origin feature/ma-fonctionnalité`).
5. Ouvre une Pull Request.

---

## Licence
Ce projet est sous licence **MIT**.

---

## Contact
Pour toute question ou suggestion, contacte-moi à tom.loustau@gmail.com.

