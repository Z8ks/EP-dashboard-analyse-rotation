# \# 📊 Dashboard Analyse Rotation - Retailer

# 

# Dashboard automatique d'analyse des ventes et du stock pour le Retailer.

# 

# \## 🎯 Fonctionnalités

# 

# \- \*\*Analyse des ventes par produit\*\* : Identification des meilleures et pires ventes

# \- \*\*Rotation du stock\*\* : Calcul automatique du coefficient de rotation

# \- \*\*Visualisations graphiques\*\* : Graphiques automatiques intégrés dans Excel

# \- \*\*Alertes automatiques\*\* : Produits en surstockage ou en rupture

# \- \*\*Export professionnel\*\* : Dashboard Excel prêt à présenter

# 

# \## 🚀 Installation

# 

# \### Prérequis

# \- Python 3.7+

# \- Excel 2016+ (Windows recommandé)

# 

# \### Dépendances



\## 📁 Structure des fichiers



02\_Analyse\_Rotation/

├── generer\_dashboard.py # Script principal

├── LANCER\_DASHBOARD.bat # Lanceur rapide

├── Ventes\_Stock\_Retailer.xlsx # Fichier de données (à créer)

└── Dashboard/ # Dossiers de sortie (auto-créé)



\## 💻 Utilisation



\### Méthode 1 : Double-clic (Recommandé)

1\. Double-cliquez sur \*\*`LANCER\_DASHBOARD.bat`\*\*

2\. Le dashboard s'ouvre automatiquement dans Excel



\### Méthode 2 : Ligne de commande





\## 📊 Format du fichier source



Le fichier \*\*`Ventes\_Stock\_Retailer.xlsx`\*\* doit contenir :



| Colonne | Description | Type |

|---------|-------------|------|

| `Code EAN` | Code-barres produit | Texte |

| `Désignation` | Nom du produit | Texte |

| `Quantité Vendue` | Ventes période | Nombre |

| `Stock Actuel` | Stock disponible | Nombre |

| `Prix de Vente` | Prix unitaire | Nombre |



\## 📈 Indicateurs calculés



\- \*\*CA Total\*\* : Chiffre d'affaires par produit

\- \*\*Taux de rotation\*\* : Vitesse d'écoulement du stock

\- \*\*Classement ventes\*\* : Top et Flop produits

\- \*\*Alertes stock\*\* : Surstock / Rupture



\## 🔧 Personnalisation



Modifiez les seuils dans `generer\_dashboard.py` :



Seuils d'alerte rotation

SEUIL\_SURSTOCK = 0.5 # Rotation < 50%

SEUIL\_RUPTURE = 2.0 # Rotation > 200%





\## 📄 Licence



Projet privé - Usage interne uniquement



\## 👤 Auteur



\*\*Z8ks\*\* - Dashboard automatisé pour analyse commerciale



