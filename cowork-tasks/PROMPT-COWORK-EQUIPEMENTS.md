# 🤖 PROMPT COWORK — Extraire équipements des 22 annonces Airbnb

## Objectif
Récupérer la **liste complète des équipements/amenities** affichés sur chaque fiche Airbnb de Claudine, pour les 22 logements, et les reporter dans un fichier JSON structuré.

## Accès
- **Compte Airbnb hôte** : Claudine est déjà connectée dans Chrome
- **URL profil hôte public** : https://www.airbnb.fr/users/show/252016190
- **Dashboard hôte** (si besoin) : https://www.airbnb.fr/hosting/listings

## Méthode la plus simple
Depuis le **profil hôte public** (lien ci-dessus), tu vois les 22 annonces listées.
Pour chaque annonce :

1. Cliquer sur l'annonce (s'ouvre en page publique voyageur)
2. Scroller jusqu'à la section **« Ce que propose ce logement »** (ou « What this place offers »)
3. Cliquer sur **« Afficher les XX équipements »** (bouton qui ouvre une modale avec la liste complète)
4. Copier toute la liste des équipements visibles, organisés par catégorie (SDB, Cuisine, Chambre, Extérieur, Parking, etc.)
5. Fermer la modale, revenir en arrière vers le profil hôte
6. Passer à l'annonce suivante

## Liste des 22 apparts (slug → titre Airbnb à chercher)

| slug | Titre Airbnb (mot-clé pour retrouver) |
|---|---|
| kitesurf | SPA 100m plage BRASERO design KITESURF |
| hamac | HAMAC - vue mer - 100M PLAGE |
| paddle | PADDLE - Vue Mer - 100M PLAGE |
| surf | SURF Vue Mer 100m plage |
| famille | 3 chambres 100m plage |
| balneo | BALNEO 100M PLAGE PARKING JARDIN BRASERO |
| cocon | Apart SPA (à renommer Cocon Romantique) — 25 rue Maréchal Juin |
| terrasse | Proche Hopitaux CALVE - TERRASSE 100m PLAGE |
| maisonnette | Maisonnette parking moto au coeur des sables |
| facemer | FACE MER, CENTRE DE L'ESPLANADE |
| grandlarge | 45 Esplanade Maritime |
| apolove | Romantique Face Mer - 2 Av Marianne (studio) |
| apollo | SAS Appart 112 - 2 Av Marianne (T3 panoramique) |
| kingston | Face Mer Exceptionnel - 2 Av Marianne RDC |
| reserve | La Réserve - 2 Av Marianne (peut-être pas encore publié) |
| albatros | 10 Av Marianne (Podvin Claudine) |
| miniloveroom | Balneo Face Mer Berck |
| grandeloveroom | Romantique Vue Face Mer - 59 Esplanade Parmentier |
| jeanne | 150m Plage Coeur de Berck - 23 Rue Jeanne |
| evasion | REPOS 150m Plage (à renommer Évasion) - Rue Jeanne |
| rotonde | Rotonde 150m Plage - Rue Jeanne |
| patio | Romantique Chic 150m Plage Patio - Rue Jeanne |

## Format de sortie attendu

Créer un fichier **`C:/Users/claud/Downloads/livret-accueil/cowork-tasks/equipements-airbnb.json`** avec cette structure :

```json
{
  "kitesurf": {
    "airbnb_url": "https://www.airbnb.fr/rooms/XXXXXXX",
    "titre_airbnb": "SPA 100m plage BRASERO design KITESURF",
    "capacite": "4 voyageurs · 1 chambre · 2 lits · 1 salle de bains",
    "equipements": {
      "chambre_linge": ["Linge de lit fourni", "Oreillers supplémentaires"],
      "salle_bains": ["Sèche-cheveux", "Shampoing", "Gel douche"],
      "cuisine": ["Réfrigérateur", "Micro-ondes", "Plaques de cuisson vitrocéramique", "Machine à café Nespresso"],
      "divertissement": ["TV avec Netflix", "Wi-Fi haut débit"],
      "chauffage_clim": ["Chauffage central", "Ventilateur"],
      "securite": ["Détecteur de fumée", "Extincteur"],
      "exterieur": ["Brasero", "Mobilier extérieur"],
      "parking": ["Parking gratuit dans la rue"],
      "services": ["Autorise les animaux", "Adapté aux enfants"]
    },
    "points_forts": ["Brasero", "Spa / Bain à remous", "Vue mer"],
    "equipements_absents": ["Ascenseur", "Lave-vaisselle"]
  },
  "hamac": { ... },
  ...
}
```

**Règles de catégorisation** :
- Respecter les catégories Airbnb telles qu'elles apparaissent (chambre, SDB, cuisine, etc.)
- Lister **TOUS** les équipements cochés, même ceux qui semblent anodins
- Noter aussi les équipements marqués comme **« non disponible »** (barrés) dans `equipements_absents` — utile pour être transparent avec les voyageurs
- Copier le titre Airbnb exact + l'URL complète (`airbnb.fr/rooms/XXXXXXX`)

## En cas de problème
- Si un appart n'est **pas trouvable** (pas publié sur Airbnb) → le noter dans le JSON avec `"status": "pas_sur_airbnb"`
- Si **Booking** est une meilleure source → aller sur `https://admin.booking.com` (Claudine est connectée sur 2 comptes : LMP + SAS Maréchal Invest)
- Si un équipement est ambigu → le copier tel quel, on normalisera ensuite

## Livrable final
1. Fichier JSON complet `equipements-airbnb.json` dans `cowork-tasks/`
2. Fichier texte court `RAPPORT-COWORK.md` qui liste :
   - Nombre d'apparts traités avec succès (X/22)
   - Apparts non trouvés et pourquoi
   - Équipements récurrents qui manquent sur certains apparts (utile pour alerter Claudine)
   - Durée totale

## Important — préférences Claudine

- **Données publiques uniquement** — rien de confidentiel à exposer
- **Ne rien modifier** sur Airbnb (READ-ONLY strict)
- **Respecter le rythme** — si Airbnb met des limites de vitesse, attendre et reprendre tranquillement
- Signaler si certains apparts ont des **équipements manquants vs leurs voisins** (ex : Kitesurf a Netflix mais pas Hamac) — c'est utile pour harmoniser les annonces
