import pandas as pd
import numpy as np
from datetime import datetime
import os
import sys
import glob

print()
print("=" * 100)
print("DASHBOARD RETAILER - VERSION AMELIOREE AVEC RECOMMANDATIONS DETAILLEES")
print("=" * 100)
print()

# ============================================================================
# [1/12] Import des bibliothèques
# ============================================================================
print("[1/12] Import des bibliotheques...")
print("      OK")
print()

# ============================================================================
# [2/12] Navigation vers Data
# ============================================================================
print("[2/12] Navigation vers Data...")
try:
    os.chdir("../Data")
    print(f"      Dossier: {os.getcwd()}")
    print()
except:
    print("      ERREUR: Impossible d'acceder au dossier Data")
    input("\nAppuyez sur Entree...")
    sys.exit(1)

# ============================================================================
# [3/12] Recherche fichiers
# ============================================================================
print("[3/12] Recherche fichiers...")

stock_files = glob.glob("ExcelStock-*.xlsx")
ventes_files = glob.glob("ExcelVenteHebdo-*.xlsx")
recap_files = glob.glob("Classeur1.xlsx")

if not stock_files:
    print()
    print("=" * 100)
    print("ERREUR!")
    print("=" * 100)
    print()
    print("Aucun fichier ExcelStock-*.xlsx trouve")
    print()
    input("Appuyez sur Entree pour quitter...")
    sys.exit(1)

if not ventes_files:
    print()
    print("=" * 100)
    print("ERREUR!")
    print("=" * 100)
    print()
    print("Aucun fichier ExcelVenteHebdo-*.xlsx trouve")
    print()
    input("Appuyez sur Entree pour quitter...")
    sys.exit(1)

stock_file = sorted(stock_files)[-1]
ventes_file = sorted(ventes_files)[-1]
recap_file = "Classeur1.xlsx" if recap_files else None

print(f"      Stock  : {stock_file}")
print(f"      Ventes : {ventes_file}")
print(f"      RECAP  : {'OUI' if recap_file else 'NON'}")
print()

# ============================================================================
# [4/12] Chargement Stock
# ============================================================================
print("[4/12] Chargement Stock...")
try:
    df_stock = pd.read_excel(stock_file, sheet_name='Stock')
    print(f"      {len(df_stock)} lignes")
    print()
except Exception as e:
    print(f"      ERREUR: {e}")
    input("\nAppuyez sur Entree...")
    sys.exit(1)

# ============================================================================
# [5/12] Chargement Ventes
# ============================================================================
print("[5/12] Chargement Ventes...")
try:
    df_ventes = pd.read_excel(ventes_file, sheet_name='Ventes hebdomadaires')
    print(f"      {len(df_ventes)} lignes")
    print()
except Exception as e:
    print(f"      ERREUR: {e}")
    input("\nAppuyez sur Entree...")
    sys.exit(1)

# ============================================================================
# [6/12] Chargement RECAP
# ============================================================================
print("[6/12] Chargement RECAP...")
if recap_file:
    try:
        df_recap = pd.read_excel(recap_file)
        print(f"      {len(df_recap)} articles RECAP")
        print()
    except Exception as e:
        print(f"      ATTENTION: {e}")
        df_recap = pd.DataFrame()
        print()
else:
    df_recap = pd.DataFrame()
    print("      Pas de fichier RECAP")
    print()

# ============================================================================
# [7/12] Normalisation EAN
# ============================================================================
print("[7/12] Normalisation EAN...")

def normaliser_ean(ean):
    if pd.isna(ean):
        return ""
    ean_str = str(int(float(ean))) if isinstance(ean, (int, float)) else str(ean)
    return ean_str.zfill(13)

df_stock['EAN'] = df_stock['EAN'].apply(normaliser_ean)
df_ventes['EAN'] = df_ventes['EAN'].apply(normaliser_ean)
if len(df_recap) > 0 and 'EAN' in df_recap.columns:
    df_recap['EAN'] = df_recap['EAN'].apply(normaliser_ean)

print("      OK")
print()

# ============================================================================
# [8/12] Traitement dates et choix semaine
# ============================================================================
print("[8/12] Traitement dates...")

df_ventes['Début semaine'] = pd.to_datetime(df_ventes['Début semaine'], format='%d/%m/%Y', errors='coerce')
df_ventes['Fin semaine'] = pd.to_datetime(df_ventes['Fin semaine'], format='%d/%m/%Y', errors='coerce')

derniere_semaine = df_ventes['Début semaine'].max()
print(f"      Dernière semaine détectée dans le fichier : {derniere_semaine.strftime('%d/%m/%Y')}")

choix_date = input(f"Entrez la date de début de semaine à traiter (jj/mm/aaaa) ou appuyez sur Entrée pour utiliser {derniere_semaine.strftime('%d/%m/%Y')} : ").strip()

if choix_date:
    try:
        semaine_choisie = pd.to_datetime(choix_date, format='%d/%m/%Y')
        if semaine_choisie not in df_ventes['Début semaine'].values:
            print(f"      ATTENTION: Date {choix_date} non trouvée, utilisation de la dernière semaine")
            semaine_choisie = derniere_semaine
    except:
        print(f"      ATTENTION: Format de date invalide, utilisation de la dernière semaine")
        semaine_choisie = derniere_semaine
else:
    semaine_choisie = derniere_semaine

df_ventes_semaine = df_ventes[df_ventes['Début semaine'] == semaine_choisie].copy()

print(f"      Semaine utilisée : {semaine_choisie.strftime('%d/%m/%Y')}")
print(f"      Ventes retenues pour la semaine du {semaine_choisie.strftime('%d/%m/%Y')} : {len(df_ventes_semaine)} lignes")
print()

# ============================================================================
# [9/12] Creation dashboard
# ============================================================================
print("[9/12] Creation dashboard...")

# Filtrer Retailer
df_stock_retailer = df_stock[df_stock['Enseigne'] == 'Retailer'].copy()
df_ventes_retailer = df_ventes_semaine[df_ventes_semaine['Libellé Enseigne'] == 'Retailer'].copy()

# Ventes par article
ventes_par_article = df_ventes_retailer.groupby('EAN')['Quantité'].sum().reset_index()
ventes_par_article.columns = ['EAN', 'Ventes_Semaine']

# Stock par article
stock_par_article = df_stock_retailer.groupby(['EAN', 'Code article', 'Libellé article']).agg({
    'Quantité': 'sum'
}).reset_index()
stock_par_article.columns = ['EAN', 'Code_Article', 'Libelle', 'Stock_Total']

# Fusion
dashboard = pd.merge(stock_par_article, ventes_par_article, on='EAN', how='outer').fillna(0)
dashboard['Ventes_Semaine'] = dashboard['Ventes_Semaine'].astype(int)
dashboard['Stock_Total'] = dashboard['Stock_Total'].astype(int)

# Ajouter info RECAP
if len(df_recap) > 0:
    dashboard['Dans_RECAP'] = dashboard['EAN'].isin(df_recap['EAN'])
    
    # Récupérer la marque depuis RECAP
    if 'Marque' in df_recap.columns:
        marque_recap = df_recap[['EAN', 'Marque']].drop_duplicates()
        dashboard = pd.merge(dashboard, marque_recap, on='EAN', how='left')
        dashboard['Marque'] = dashboard['Marque'].fillna('Autre')
    else:
        dashboard['Marque'] = 'Autre'
else:
    dashboard['Dans_RECAP'] = False
    dashboard['Marque'] = 'Autre'

# Détection marque depuis libellé si manquante
marques_connues = ['RIVACASE', 'RIVA CASE', 'KASPERSKY', 'ADATA', 'DELL', 'LENOVO', 'ASUS', 'HP', 'ACER', 'EPSON']

def detecter_marque(libelle, marque_actuelle):
    if pd.notna(marque_actuelle) and marque_actuelle != 'Autre':
        return marque_actuelle
    
    libelle_upper = str(libelle).upper()
    for marque in marques_connues:
        if marque in libelle_upper:
            return marque if marque != 'RIVA CASE' else 'RIVACASE'
    return 'Autre'

dashboard['Marque'] = dashboard.apply(lambda row: detecter_marque(row['Libelle'], row['Marque']), axis=1)

# Ajouter colonnes de statut
def get_mouvement(row):
    return '✅ BOUGE' if row['Ventes_Semaine'] > 0 else '❌ INACTIF'

def get_statut(row):
    if row['Ventes_Semaine'] > 0 and row['Stock_Total'] == 0:
        return '❌ RUPTURE'
    elif row['Ventes_Semaine'] > 0 and row['Stock_Total'] < 5:
        return '🔴 CRITIQUE'
    elif row['Ventes_Semaine'] > 0:
        return '📈 ACTIF'
    elif row['Stock_Total'] > 0:
        return '⏸️  INACTIF'
    else:
        return '⚪ ABSENT'

dashboard['Mouvement'] = dashboard.apply(get_mouvement, axis=1)
dashboard['Statut'] = dashboard.apply(get_statut, axis=1)

# Tri par Marque puis Ventes
dashboard = dashboard.sort_values(['Marque', 'Ventes_Semaine'], ascending=[True, False])

print(f"      {len(dashboard)} articles uniques")
print(f"      Dashboard: {len(dashboard)} lignes")
print()

# ============================================================================
# [10/12] Creation fichier Excel
# ============================================================================
print("[10/12] Creation fichier Excel...")

# Créer dossier Resultats
try:
    os.makedirs("../Resultats", exist_ok=True)
    os.chdir("../Resultats")
except Exception as e:
    print(f"      ERREUR: {e}")
    input("\nAppuyez sur Entree...")
    sys.exit(1)

# Nom fichier
date_str = datetime.now().strftime('%d%m%Y')
output_file = f"DASHBOARD_RETAILER_{date_str}.xlsx"

# Créer Excel
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Onglet SYNTHESE
    synthese_data = {
        'Indicateur': [
            '📊 Total Articles',
            '📈 Articles Actifs (ventes > 0)',
            '🔴 Articles Critiques (ventes > 0, stock < 5)',
            '❌ Articles en Rupture (ventes > 0, stock = 0)',
            '⏸️  Articles Inactifs (ventes = 0, stock > 0)',
            '⚪ Articles Absents (ventes = 0, stock = 0)',
            '💰 Ventes Totales Semaine'
        ],
        'Valeur': [
            len(dashboard),
            len(dashboard[dashboard['Ventes_Semaine'] > 0]),
            len(dashboard[(dashboard['Ventes_Semaine'] > 0) & (dashboard['Stock_Total'] < 5) & (dashboard['Stock_Total'] > 0)]),
            len(dashboard[(dashboard['Ventes_Semaine'] > 0) & (dashboard['Stock_Total'] == 0)]),
            len(dashboard[(dashboard['Ventes_Semaine'] == 0) & (dashboard['Stock_Total'] > 0)]),
            len(dashboard[(dashboard['Ventes_Semaine'] == 0) & (dashboard['Stock_Total'] == 0)]),
            int(dashboard['Ventes_Semaine'].sum())
        ]
    }
    
    df_synthese = pd.DataFrame(synthese_data)
    df_synthese.to_excel(writer, sheet_name='SYNTHESE', index=False)
    
    # Onglet TOUS LES ARTICLES
    dashboard_export = dashboard[['Code_Article', 'Libelle', 'EAN', 'Marque', 'Ventes_Semaine', 
                                   'Stock_Total', 'Dans_RECAP', 'Mouvement', 'Statut']]
    dashboard_export.to_excel(writer, sheet_name='TOUS LES ARTICLES', index=False)
    
    # Onglet RECOMMANDATIONS DETAILS
    recommandations = []
    
    # Articles Critiques
    critiques = dashboard[(dashboard['Ventes_Semaine'] > 0) & (dashboard['Stock_Total'] < 5) & (dashboard['Stock_Total'] > 0)].copy()
    critiques['Categorie'] = '🔴 CRITIQUE'
    critiques['Qte_Recommandee'] = (critiques['Ventes_Semaine'] * 4) - critiques['Stock_Total']
    recommandations.append(critiques)
    
    # Ruptures
    ruptures = dashboard[(dashboard['Ventes_Semaine'] > 0) & (dashboard['Stock_Total'] == 0)].copy()
    ruptures['Categorie'] = '❌ RUPTURE'
    ruptures['Qte_Recommandee'] = ruptures['Ventes_Semaine'] * 4
    recommandations.append(ruptures)
    
    # Inactifs
    inactifs = dashboard[(dashboard['Ventes_Semaine'] == 0) & (dashboard['Stock_Total'] > 0)].copy()
    inactifs['Categorie'] = '⏸️  INACTIF'
    inactifs['Qte_Recommandee'] = 0
    recommandations.append(inactifs)
    
    if recommandations:
        df_recommandations = pd.concat(recommandations, ignore_index=True)
        df_recommandations = df_recommandations[['Categorie', 'Code_Article', 'Libelle', 'EAN', 'Marque',
                                                  'Ventes_Semaine', 'Stock_Total', 'Qte_Recommandee']]
        df_recommandations.to_excel(writer, sheet_name='RECOMMANDATIONS DETAILS', index=False)
    
    # TOP 10 MARQUES
    if 'Marque' in dashboard.columns:
        top_marques = dashboard.groupby('Marque').agg({
            'Ventes_Semaine': 'sum',
            'Stock_Total': 'sum',
            'Code_Article': 'count'
        }).reset_index()
        top_marques.columns = ['Marque', 'Ventes_Hebdo', 'Stock_Total', 'Nb_Articles']
        top_marques = top_marques.nlargest(10, 'Ventes_Hebdo')
        
        # Ajouter rotation
        top_marques['Rotation'] = top_marques.apply(
            lambda row: round(row['Ventes_Hebdo'] / row['Stock_Total'], 2) if row['Stock_Total'] > 0 else 0,
            axis=1
        )
        
        top_marques.to_excel(writer, sheet_name='TOP 10 MARQUES', index=False)

print(f"      OK : {output_file}")
print()

# ============================================================================
# [11/12] Application couleurs (optionnel)
# ============================================================================
print("[11/12] Application des styles...")
try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font
    
    wb = load_workbook(output_file)
    
    # Formater l'onglet TOUS LES ARTICLES
    ws = wb['TOUS LES ARTICLES']
    
    # En-têtes
    header_fill = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Colorier selon statut
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        statut = row[8].value  # Colonne Statut
        
        if statut and '🔴' in str(statut):
            fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
        elif statut and '❌' in str(statut):
            fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
        elif statut and '⏸️' in str(statut):
            fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
        elif statut and '📈' in str(statut):
            fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
        else:
            fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
        
        for cell in row:
            cell.fill = fill
    
    wb.save(output_file)
    print("      OK")
    print()
except Exception as e:
    print(f"      ATTENTION: {e}")
    print()

# ============================================================================
# [12/12] Statistiques finales
# ============================================================================
print("[12/12] Statistiques finales...")

stats = dashboard['Statut'].value_counts()
print()
for statut, count in stats.items():
    print(f"      {statut} : {count} articles")

print()
print("=" * 100)
print("GENERATION TERMINEE !")
print("=" * 100)
print()
print(f"Fichier genere : {output_file}")
print(f"Dossier        : {os.getcwd()}")
print()

input("Appuyez sur Entree pour quitter...")
