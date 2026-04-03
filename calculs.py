"""
Moteur de calcul du Plan Financier Hotel 4*
Reproduit la logique des formules Excel (7. RECAP, amortissements, indicateurs)
"""

import numpy as np
import pandas as pd
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
import calendar


# ─── Helpers ───────────────────────────────────────────────────────────────────

def mois_range(date_ouverture: date, nb_mois: int) -> list[date]:
    """Genere la liste des premiers jours de chaque mois a partir de l'ouverture."""
    return [date_ouverture + relativedelta(months=i) for i in range(nb_mois)]


def jours_dans_mois(d: date) -> int:
    return calendar.monthrange(d.year, d.month)[1]


def annee_exploitation(d: date, date_ouverture: date) -> int:
    """Renvoie le numero d'annee d'exploitation (1, 2, 3, 4+)."""
    delta = relativedelta(d, date_ouverture)
    return delta.years + 1


# ─── Parametres par defaut ─────────────────────────────────────────────────────

def params_defaut():
    return {
        # General
        "nom_hotel": "Chateau d'Argenteau",
        "nb_chambres": 70,
        "personnes_par_chambre": 2,
        "date_ouverture": date(2029, 7, 1),
        "nb_mois_projection": 204,  # 17 ans
        "nb_couverts_brasserie": 80,
        "nb_salles_mariage": 1,
        "nb_invites_mariage": 125,
        "nb_invites_seminaire": 50,

        # Taux d'occupation hebergement par annee (1->4+)
        "taux_occ": [0.42, 0.475, 0.55, 0.62],

        # Taux d'occupation brasserie par annee (1->4+)
        "taux_occ_brasserie": [0.35, 0.45, 0.525, 0.60],

        # Saisonnalite (jan->dec) - ponderation du taux d'occupation
        "saisonnalite": [0.626, 0.657, 0.86, 1.017, 1.126, 1.22,
                         1.283, 1.33, 1.173, 1.064, 0.782, 0.86],

        # Segmentation clientele
        "segments": {
            "Loisirs": {"part": 0.50, "prix": 230},
            "Affaires": {"part": 0.25, "prix": 210},
            "Groupes": {"part": 0.05, "prix": 185},
            "MICE": {"part": 0.05, "prix": 185},
            "Evenementiel": {"part": 0.15, "prix": 210},
        },

        # Part OTA par segment (pour calcul commissions)
        "segments_part_ota": {
            "Loisirs": 0.75,
            "Affaires": 0.55,
            "Groupes": 0.20,
            "MICE": 0.0,
            "Evenementiel": 0.0,
        },

        # Hausse prix annuelle (hors inflation)
        "hausse_prix_an": [0.0, 0.025, 0.025, 0.025],

        # Brasserie (structure Excel)
        "brasserie_prix_souper": 75,
        "brasserie_jours_souper": 4,        # jours/semaine
        "brasserie_services_souper": 1,      # services par jour
        "brasserie_prix_diner": 45,
        "brasserie_jours_diner": 4,          # jours/semaine
        "brasserie_services_diner": 1.5,     # services par jour
        "brasserie_part_nourriture": 0.60,   # part revenu lie a nourriture
        "brasserie_part_boissons": 0.40,     # part revenu lie aux boissons

        # Compatibilite UI existante (ancien format)
        "brasserie_ouvert_midi": 4,
        "brasserie_couverts_midi": 40,
        "brasserie_prix_midi": 45,
        "brasserie_ouvert_soir": 4,
        "brasserie_couverts_soir": 50,
        "brasserie_prix_soir": 75,
        "brasserie_part_interne": 0.60,
        "brasserie_part_externe": 0.40,

        # Petit-dejeuner
        "petit_dej_prix": 37.5,
        "petit_dej_taux": 0.85,  # % clients qui prennent pdj

        # Bar (structure Excel simplifiee)
        "bar_prix_moyen": 18,               # EUR/personne/nuitee vendue
        "bar_taux_clients_hotel": 0.40,      # % des clients hotel qui vont au bar
        "bar_conso_moyenne": 18,
        "bar_clients_ext_jour": 10,
        "bar_conso_ext_moyenne": 22,
        "bar_jours_ouvert_semaine": 7,

        # Spa (structure Excel)
        "spa_entree_hotel_prix": 0,          # EUR/entree clients hotel
        "spa_entree_hotel_taux": 0.20,       # ratio par nuitee vendue
        "spa_entree_ext_prix": 55,           # EUR/entree externes
        "spa_entree_ext_nb_mois": 25,        # personnes externes par mois
        "spa_soin_hotel_prix": 120,          # EUR/soin clients hotel
        "spa_soin_hotel_taux": 0.10,         # ratio par nuitee vendue
        "spa_soin_ext_prix": 150,            # EUR/soin externes
        "spa_soin_ext_nb_mois": 15,          # soins externes par mois

        # Compatibilite UI existante
        "spa_capacite_soins_jour": 30,
        "spa_prix_moyen_soin": 55,
        "spa_taux_clients_hotel": 0.25,
        "spa_clients_ext_jour": 5,
        "spa_prix_ext": 65,

        # Salles (mariages + seminaires combines)
        "salle_mariage_prix": 7500,          # EUR/salle/jour
        "salle_mariage_nb_an": 49,           # nombre par an (combine)
        "salle_seminaire_prix": 1750,        # EUR/salle/jour
        "salle_seminaire_nb_an": 49,         # nombre par an

        # Compatibilite UI existante
        "seminaire_nb_an": 50,
        "seminaire_prix_location": 800,
        "seminaire_nb_participants_moy": 25,
        "seminaire_prix_participant": 65,
        "seminaire_duree_moy_jours": 1,
        "mariage_nb_an": 12,
        "mariage_nb_convives_moy": 120,
        "mariage_prix_convive": 95,
        "mariage_prix_location": 2500,

        # Divers (minibar, blanchisserie, extras)
        "divers_prix_nuitee": 3,             # EUR/nuitee vendue
        "divers_taux": 1.0,                  # taux d'utilisation

        # Loyer restaurant (reference externe dans Excel)
        "loyer_restaurant_mensuel": 5860,    # EUR/mois

        # Charges variables (structure Excel)
        "cv_hebergement_par_nuitee": {
            "Linge / Blanchisserie": 5.5,
            "Produits accueil": 3.5,
            "Produits entretien": 2.5,
            "Energie variable": 10.0,
            "Fournitures chambres": 2.5,
        },
        "cv_commission_ota_pct": 0.17,       # % du CA hebergement pondere par part OTA
        "cv_franchise_pct": 0.04,            # % du CA (hors loyer restaurant et spa)

        # Compatibilite ancienne structure
        "cv_hebergement": {
            "Blanchisserie": 8.5,
            "Produits accueil": 3.5,
            "Nettoyage": 12.0,
            "Petit-dejeuner (cout)": 11.0,
            "Commission OTA": 0.15,
            "Divers": 5.0,
        },

        "cv_brasserie_pct": 0.35,            # food cost brasserie = pdj
        "cv_bar_beverage_pct": 0.30,          # beverage cost
        "cv_bar_consommable_unite": 0.20,     # EUR/consommation
        "cv_bar_pct": 0.25,                   # compat

        "cv_spa_soin_cout": 50,               # EUR/soin
        "cv_spa_produits_soin": 5,            # EUR/soin
        "cv_spa_consommable_entree": 1.5,     # EUR/entree spa
        "cv_spa_energie_entree": 3.0,         # EUR/entree spa
        "cv_spa_piscine_entree": 0.5,         # EUR/entree spa
        "cv_spa_pct": 0.20,                   # compat

        "cv_seminaire_pause_participant": 8,  # EUR/participant
        "cv_seminaire_materiel_participant": 3, # EUR/participant
        "cv_seminaire_equipement": 15,        # EUR/seminaire
        "cv_seminaire_energie": 75,           # EUR/seminaire
        "cv_seminaire_nettoyage": 500,        # EUR/seminaire
        "cv_seminaire_pct": 0.30,             # compat

        "cv_mariage_energie": 150,            # par mariage
        "cv_mariage_nettoyage": 1000,         # par mariage
        "cv_mariage_pct": 0.35,               # compat

        "cv_divers_consommable": 1,           # EUR/nuitee

        # ── Charges fixes directes (par service, structure Excel) ──────
        "charges_patronales_pct": 0.35,

        # Personnel par departement/service
        "personnel_hebergement": [
            {"poste": "Housekeeping / Etages", "etp": 7, "cout_brut": 35000},
            {"poste": "Reception / Front Office", "etp": 5, "cout_brut": 37500},
            {"poste": "Night Audit", "etp": 2, "cout_brut": 42500},
        ],
        "personnel_brasserie": [
            {"poste": "Chef", "etp": 1.5, "cout_brut": 67500},
            {"poste": "Cuisinier", "etp": 6, "cout_brut": 55000},
            {"poste": "Serveurs", "etp": 6, "cout_brut": 42500},
        ],
        "personnel_bar": [
            {"poste": "Barman", "etp": 1, "cout_brut": 37500},
        ],
        "personnel_spa": [
            {"poste": "Spa - hote", "etp": 1, "cout_brut": 35000},
        ],
        "personnel_indirect": [
            {"poste": "Direction generale", "etp": 1, "cout_brut": 65000},
            {"poste": "Direction adjointe / Revenue Manager", "etp": 1, "cout_brut": 50000},
            {"poste": "Maintenance / Technique", "etp": 2.33, "cout_brut": 34000},
            {"poste": "Commercial / Ventes / Marketing", "etp": 1, "cout_brut": 40000},
            {"poste": "Administration / Comptabilite", "etp": 1.5, "cout_brut": 34000},
            {"poste": "Extras / Renforts saisonniers", "etp": 4, "cout_brut": 24000},
            {"poste": "Chauffeur / Bagagiste / Concierge", "etp": 2.5, "cout_brut": 42500},
        ],

        # Loyer a repartir par service (%)
        "loyer_mensuel": 120000,
        "loyer_repartition": {
            "hebergement": 0.65,
            "brasserie": 0.15,
            "bar": 0.05,
            "spa": 0.05,
            "salles": 0.10,
        },

        # Autres frais fixes directs (hors personnel, par service)
        "cf_directs_hebergement": {},
        "cf_directs_brasserie": {},
        "cf_directs_spa": {},
        "personnel_evenements": [],
        "cf_directs_evenements": {},

        # ── Charges fixes indirectes (non lies a un service) ─────────
        "charges_fixes_indirectes_par_annee": {
            "Assurances": [35000, 35000, 35000],
            "Precompte immobilier": [50000, 50000, 50000],
            "Energie fixe": [30000, 36000, 40000],
            "Contrats maintenance": [33750, 40500, 45000],
            "Logiciels IT": [18000, 21600, 24000],
            "Marketing Communication": [45000, 32500, 27500],
            "Honoraires comptable juridique": [13500, 16200, 18000],
            "Abonnements cotisations": [3750, 4500, 5000],
            "Frais de gestion": [0, 0, 0],
            "Fournitures bureau": [4500, 5400, 6000],
            "Securite gardiennage": [9000, 10800, 12000],
            "Dechets collecte": [6000, 7200, 8000],
            "Jardin espaces verts": [4500, 5400, 6000],
            "Licences taxes SABAM": [4000, 4000, 4000],
        },

        # Compat ancienne structure
        "charges_fixes_indirectes": {
            "Assurances": 35000,
            "Precompte immobilier": 50000,
            "Energie fixe": 40000,
            "Contrats maintenance": 45000,
            "Logiciels IT": 24000,
            "Marketing Communication": 27500,
            "Honoraires comptable juridique": 18000,
            "Abonnements cotisations": 5000,
            "Fournitures bureau": 6000,
            "Securite gardiennage": 12000,
            "Dechets collecte": 8000,
            "Jardin espaces verts": 6000,
            "Licences taxes SABAM": 4000,
        },

        # Inflation differenciee (structure Excel)
        "inflation_ventes": 0.025,
        "inflation_loyer_restaurant": 0.02,
        "inflation_personnel": 0.025,
        "inflation_loyer": 0.02,
        "inflation_charges_variables": 0.025,
        "inflation_autres_charges": 0.02,
        "inflation_an": 0.025,  # compat

        # ── Investissements & Financement ─────────────────────────────────
        "investissements": [
            {"categorie": "Terrain", "montant": 2951000, "duree_amort": 0},
            {"categorie": "Construction", "montant": 14871850, "duree_amort": 0},
            {"categorie": "Amenagements interieurs", "montant": 1534444, "duree_amort": 10},
            {"categorie": "Mobilier & Equipements", "montant": 2347485, "duree_amort": 5},
            {"categorie": "Branding", "montant": 200000, "duree_amort": 3},
        ],
        "reinvest_pct_an": [0.0, 0.005, 0.01, 0.015, 0.02, 0.02, 0.025],
        "reinvest_categories": ["Amenagements interieurs", "Mobilier & Equipements"],

        "prets": [
            {"nom": "Pret banque non garanti", "montant": 2500000, "taux_annuel": 0.04,
             "duree_ans": 15, "differe_mois": 0, "type": "annuite"},
            {"nom": "Pret banque garanti", "montant": 500000, "taux_annuel": 0.03,
             "duree_ans": 2, "differe_mois": 0, "type": "interet_seul"},
            {"nom": "Pret ROCHER", "montant": 2400000, "taux_annuel": 0.04,
             "duree_ans": 15, "differe_mois": 0, "type": "interet_seul"},
        ],
        "fonds_propres_initial": 1000000,

        # ── Hypotheses fiscales ───────────────────────────────────────────
        "tva_ventes": {
            "Hebergement": 0.06,
            "Brasserie nourriture": 0.12,
            "Brasserie boisson": 0.21,
            "Petit-dejeuner": 0.12,
            "Bar": 0.21,
            "Spa": 0.21,
            "Salles": 0.21,
            "Divers": 0.12,
            "Loyer Restaurant": 0.21,
        },
        "tva_achats": {
            "Charges variables pdj brasserie divers salles": 0.12,
            "Autres": 0.21,
        },
        "tva_periodicite": "Trimestrielle",
        "isoc": 0.25,

        # ── Delais de paiement (mois) ─────────────────────────────────────
        "delais_clients": {
            "hebergement_Loisirs": 0,
            "hebergement_Affaires": 1,
            "hebergement_Groupes": 1,
            "hebergement_MICE": 1,
            "hebergement_Evenementiel": 1,
            "brasserie": 0,
            "bar": 0,
            "spa": 0,
            "salles": 1,
            "divers": 0,
            "loyer_restaurant": 1,
        },
        "delais_fournisseurs": {
            # Charges variables par nature
            "Linge & blanchisserie": 1,
            "Produits & consommables": 1,
            "Energie variable": 1,
            "Nourriture": 1,
            "Boissons": 1,
            "Commissions OTA": 0,
            "Commissions CB & franchise": 0,
            "Soins spa": 1,
            "Evenements (nettoyage, materiel)": 1,
            # Charges fixes par nature
            "Loyer": 0,
            "CF directs departements": 1,
            "Energie fixe": 1,
            "Assurances & taxes": 0,
            "Maintenance & contrats": 1,
            "Marketing": 1,
            "IT & logiciels": 0,
            "Honoraires": 1,
            "Autres CF indirects": 1,
        },
    }


# ─── Calcul du prix moyen pondere (ADR) ─────────────────────────────────────

def calc_prix_moyen_pondere(params, annee_idx):
    """Prix moyen pondere par la segmentation (ADR)."""
    prix = 0
    for seg, data in params["segments"].items():
        prix += data["part"] * data["prix"]
    # Appliquer hausse prix (hors inflation) a partir annee 2
    for i in range(1, min(annee_idx + 1, len(params["hausse_prix_an"]))):
        prix *= (1 + params["hausse_prix_an"][i])
    return prix


# ─── Taux d'occupation ──────────────────────────────────────────────────────

def calc_taux_occupation(params, annee_idx, mois_idx):
    """Taux d'occupation hebergement pour une annee et un mois donnes."""
    taux_base = params["taux_occ"][min(annee_idx, len(params["taux_occ"]) - 1)]
    saisonnalite = params["saisonnalite"][mois_idx]
    return min(taux_base * saisonnalite, 1.0)


def calc_taux_occupation_brasserie(params, annee_idx, mois_idx):
    """Taux d'occupation brasserie."""
    taux_occ_brass = params.get("taux_occ_brasserie", params["taux_occ"])
    taux_base = taux_occ_brass[min(annee_idx, len(taux_occ_brass) - 1)]
    saisonnalite = params["saisonnalite"][mois_idx]
    return min(taux_base * saisonnalite, 1.0)


# ─── Calcul du nombre de mois depuis l'ouverture (pour inflation) ────────────

def _mois_depuis_ouverture(d: date, date_ouverture: date) -> int:
    return (d.year - date_ouverture.year) * 12 + (d.month - date_ouverture.month)


# ─── Revenus mensuels ────────────────────────────────────────────────────────

def calc_revenus_mensuels(params, d: date, annee_idx: int):
    """Calcule tous les revenus pour un mois donne (logique Excel 7.RECAP)."""
    mois_idx = d.month - 1
    jours = jours_dans_mois(d)
    nb_ch = params["nb_chambres"]
    pers_ch = params.get("personnes_par_chambre", 2)
    mois_depuis = _mois_depuis_ouverture(d, params["date_ouverture"])

    # Inflation cumulative par categorie
    infl_ventes = (1 + params.get("inflation_ventes", params["inflation_an"]) / 12) ** mois_depuis
    infl_loyer_rest = (1 + params.get("inflation_loyer_restaurant", 0.02) / 12) ** mois_depuis
    infl_autres = (1 + params.get("inflation_autres_charges", 0.02) / 12) ** mois_depuis

    saisonn = params["saisonnalite"][mois_idx]

    # Taux d'occupation hebergement
    taux_occ = calc_taux_occupation(params, annee_idx, mois_idx)
    # ADR avec hausse prix (hors inflation)
    prix_moyen = calc_prix_moyen_pondere(params, annee_idx)
    # Appliquer inflation ventes a l'ADR
    prix_moyen_infl = prix_moyen * infl_ventes

    # Nuitees (= chambres occupees)
    chambres_dispo = nb_ch * jours
    nuitees = chambres_dispo * taux_occ

    # ── CA Hebergement ──
    ca_hebergement = nuitees * prix_moyen_infl

    # ── CA Brasserie (formule Excel Row 19+20+21) ──
    # Brasserie externe (Row 19): taux_occ_brasserie * saisonnalite *
    #   (prix_souper * jours_souper/7 * services + prix_diner * jours_diner/7 * services)
    #   * jours * nb_couverts
    taux_occ_brass = calc_taux_occupation_brasserie(params, annee_idx, mois_idx)
    nb_couverts = params.get("nb_couverts_brasserie", 80)

    # Utiliser les params brasserie (souper/diner ou midi/soir)
    if "brasserie_prix_souper" in params:
        ca_brass_ext = (taux_occ_brass *
                        (params["brasserie_prix_souper"] * params["brasserie_jours_souper"] / 7 * params["brasserie_services_souper"] +
                         params["brasserie_prix_diner"] * params["brasserie_jours_diner"] / 7 * params["brasserie_services_diner"]) *
                        jours * nb_couverts)
    else:
        # Fallback ancien format
        semaines = jours / 7
        ca_brass_ext = (params["brasserie_couverts_soir"] * params["brasserie_prix_soir"] *
                        params["brasserie_ouvert_soir"] / 7 * 1 +
                        params["brasserie_couverts_midi"] * params["brasserie_prix_midi"] *
                        params["brasserie_ouvert_midi"] / 7 * 1.5) * jours

    # Brasserie interne / petit-dejeuner (Row 20): nuitees * taux_pdj * prix_pdj * pers/chambre
    ca_pdj = nuitees * params["petit_dej_taux"] * params["petit_dej_prix"] * pers_ch

    # Brasserie total (Row 21): (externe + interne) * inflation
    ca_brasserie = (ca_brass_ext + ca_pdj) * infl_ventes

    # ── CA Bar (Row 22) ──
    # Clients hotel : conso_moyenne * taux * nuitees * pers/chambre
    ca_bar_hotel = (params.get("bar_prix_moyen", params["bar_conso_moyenne"]) *
                    params["bar_taux_clients_hotel"] * nuitees * pers_ch)
    # Clients externes : clients_ext/jour * conso_ext * jours_ouverts (sans saisonnalite)
    jours_ouvert_bar = jours * params.get("bar_jours_ouvert_semaine", 7) / 7
    ca_bar_ext = (params.get("bar_clients_ext_jour", 0) *
                  params.get("bar_conso_ext_moyenne", 0) * jours_ouvert_bar)
    ca_bar = (ca_bar_hotel + ca_bar_ext) * infl_ventes

    # ── CA Spa (Row 23) ──
    # (entree_hotel_prix * entree_hotel_taux * nuitees +
    #  soin_hotel_prix * soin_hotel_taux * nuitees +
    #  entree_ext_prix * entree_ext_nb_mois +
    #  soin_ext_prix * soin_ext_nb_mois) * inflation
    if "spa_entree_hotel_prix" in params:
        ca_spa = (params["spa_entree_hotel_prix"] * params["spa_entree_hotel_taux"] * nuitees +
                  params["spa_soin_hotel_prix"] * params["spa_soin_hotel_taux"] * nuitees +
                  params["spa_entree_ext_prix"] * params["spa_entree_ext_nb_mois"] +
                  params["spa_soin_ext_prix"] * params["spa_soin_ext_nb_mois"]) * infl_ventes
    else:
        # Fallback
        ca_spa = (nuitees * params["spa_taux_clients_hotel"] * params["spa_prix_moyen_soin"] +
                  jours * params["spa_clients_ext_jour"] * params["spa_prix_ext"]) * infl_ventes

    # ── CA Salles (Row 24) ──
    ca_seminaires = 0
    ca_mariages = 0
    ca_salles = 0
    # Seminaires : location_salle x nb_seminaires / 12
    if "seminaire_nb_an" in params:
        ca_seminaires = params.get("seminaire_prix_location", 800) * params["seminaire_nb_an"] / 12 * infl_ventes
        ca_salles += ca_seminaires
    # Mariages : location_salle x nb_mariages / 12
    if "mariage_nb_an" in params and params["mariage_nb_an"] > 0:
        ca_mariages = params.get("mariage_prix_location", 2500) * params["mariage_nb_an"] / 12 * infl_ventes
        ca_salles += ca_mariages
    # Location salles du chateau
    ca_salles_chateau = 0
    if params.get("salles_chateau_nb_an", 0) > 0:
        ca_salles_chateau = (params.get("salles_chateau_prix", 1500) *
                             params["salles_chateau_nb_an"] / 12 * infl_ventes)
        ca_salles += ca_salles_chateau

    # ── CA Divers (Row 25): prix * taux * nuitees * inflation ──
    ca_divers = (params.get("divers_prix_nuitee", 3) *
                 params.get("divers_taux", 1.0) * nuitees * infl_ventes)

    # ── Loyer Restaurant gastronomique ──
    ca_loyer_restaurant = params.get("loyer_restaurant_mensuel", 5800) * infl_loyer_rest

    # ── CA TOTAL ──
    ca_total = (ca_hebergement + ca_brasserie + ca_bar + ca_spa +
                ca_salles + ca_divers + ca_loyer_restaurant)

    # CA hebergement ventile par segment (pour delais de paiement)
    hausse_mult = 1.0
    for i in range(1, min(annee_idx + 1, len(params["hausse_prix_an"]))):
        hausse_mult *= (1 + params["hausse_prix_an"][i])
    ca_par_segment = {}
    for seg, data in params["segments"].items():
        ca_par_segment[seg] = nuitees * data["part"] * data["prix"] * hausse_mult * infl_ventes

    return {
        "date": d,
        "nuitees": nuitees,
        "taux_occupation": taux_occ,
        "prix_moyen": prix_moyen_infl,
        "ca_hebergement": ca_hebergement,
        "ca_brasserie": ca_brasserie,
        "ca_pdj": ca_pdj * infl_ventes,  # Pour affichage separe
        "ca_bar": ca_bar,
        "ca_spa": ca_spa,
        "ca_salles": ca_salles,
        "ca_seminaires": ca_seminaires,
        "ca_mariages": ca_mariages,
        "ca_salles_chateau": ca_salles_chateau,
        "ca_divers": ca_divers,
        "ca_loyer_restaurant": ca_loyer_restaurant,
        "ca_total": ca_total,
        # Donnees intermediaires pour charges variables
        "_ca_brass_ext": ca_brass_ext * infl_ventes,
        "_ca_pdj_infl": ca_pdj * infl_ventes,
        "_infl_ventes": infl_ventes,
        "_infl_autres": infl_autres,
        "_saisonnalite": saisonn,
        "_ca_par_segment": ca_par_segment,
    }


# ─── Charges variables ────────────────────────────────────────────────────────

def calc_charges_variables(params, rev: dict, annee_idx: int):
    """Charges variables mensuelles (logique Excel 7.RECAP rows 30-55)."""
    nuitees = rev["nuitees"]
    mois_depuis = _mois_depuis_ouverture(rev["date"], params["date_ouverture"])
    infl_cv = (1 + params.get("inflation_charges_variables", params["inflation_an"]) / 12) ** mois_depuis
    infl_autres = (1 + params.get("inflation_autres_charges", 0.02) / 12) ** mois_depuis

    # ── Hebergement commission OTA (Row 30) ──
    # Pondere par segments: sum(part_segment * part_ota_segment) * CA_hebergement * taux_commission
    if "segments_part_ota" in params:
        poids_ota = sum(
            params["segments"][seg]["part"] * params["segments_part_ota"].get(seg, 0)
            for seg in params["segments"]
        )
        cv_heberg_commission = poids_ota * rev["ca_hebergement"] * params.get("cv_commission_ota_pct", 0.17)
    else:
        # Fallback
        part_ota = params["segments"].get("OTA", {}).get("part", 0.05)
        cv_heberg_commission = rev["ca_hebergement"] * part_ota * params["cv_hebergement"].get("Commission OTA", 0.15)

    # ── Hebergement autres CV (Row 31) ──
    # = sum(cout_par_nuitee) * nuitees + commission_cb + franchise_pct * CA
    if "cv_hebergement_par_nuitee" in params:
        cout_par_nuitee = sum(params["cv_hebergement_par_nuitee"].values())
        # Commission cartes de credit : cout/nuitee * % chambres payees par CB * nuitees
        cv_commission_cb = (params.get("cv_commission_cb_nuitee", 1.5) *
                            params.get("cv_commission_cb_pct_chambres", 0.80) * nuitees * infl_cv)
        # Franchise : % du CA total hebergement
        ca_base_franchise = rev["ca_hebergement"]
        cv_franchise = ca_base_franchise * params.get("cv_franchise_pct", 0.04)
        cv_heberg_autre = (cout_par_nuitee * nuitees * infl_cv + cv_commission_cb + cv_franchise)
    else:
        # Fallback ancien format
        cv_heberg_autre = 0
        for poste, cout in params["cv_hebergement"].items():
            if poste != "Commission OTA":
                cv_heberg_autre += nuitees * cout * infl_cv

    cv_hebergement = cv_heberg_commission + cv_heberg_autre

    # ── Brasserie CV (Row 36) ──
    # Midi & soir : food cost % CA + PDJ : food cost PDJ % CA PDJ
    cv_brasserie_resto = rev.get("_ca_brass_ext", 0) * params["cv_brasserie_pct"]
    cv_pdj = rev.get("_ca_pdj_infl", 0) * params.get("cv_pdj_pct", params["cv_brasserie_pct"])
    cv_brasserie = cv_brasserie_resto + cv_pdj

    # ── Bar CV (Row 41) ──
    # = ca_bar * beverage_cost + nb_conso * consommable
    bar_prix = params.get("bar_prix_moyen", params.get("bar_conso_moyenne", 18))
    nb_conso = rev["ca_bar"] / bar_prix if bar_prix > 0 else 0
    cv_bar = (rev["ca_bar"] * params.get("cv_bar_pct", 0.25) +
              nb_conso * params.get("cv_bar_consommable_unite", 0.20) * infl_cv)

    # ── Spa CV (Row 46) ──
    nb_soins_hotel = rev["nuitees"] * params.get("spa_soin_hotel_taux", 0.10)
    nb_soins_ext = params.get("spa_soin_ext_nb_mois", 15)
    nb_entrees_hotel = rev["nuitees"] * params.get("spa_entree_hotel_taux", 0.20)
    nb_entrees_ext = params.get("spa_entree_ext_nb_mois", 25)

    cv_spa_soins = ((params.get("cv_spa_soin_cout", 50) + params.get("cv_spa_produits_soin", 5)) *
                    (nb_soins_hotel + nb_soins_ext))
    cv_spa_entrees = ((params.get("cv_spa_consommable_entree", 1.5) +
                       params.get("cv_spa_energie_entree", 3.0) +
                       params.get("cv_spa_piscine_entree", 0.5)) *
                      (nb_entrees_hotel + nb_entrees_ext))
    cv_spa = (cv_spa_soins + cv_spa_entrees) * infl_cv

    # ── Seminaires CV ──
    nb_sem = params.get("seminaire_nb_an", 50)
    nb_participants_sem = params.get("seminaire_nb_participants_moy", 25)
    cv_seminaires = ((params.get("cv_seminaire_pause_participant", 8) +
                      params.get("cv_seminaire_materiel_participant", 3)) * nb_participants_sem * nb_sem +
                     nb_sem * (params.get("cv_seminaire_equipement", 15) +
                               params.get("cv_seminaire_energie", 75) +
                               params.get("cv_seminaire_nettoyage", 500))) / 12 * infl_cv

    # ── Mariages CV ──
    nb_mar = params.get("mariage_nb_an", 12)
    cv_mariages = ((params.get("cv_mariage_energie", 150) +
                    params.get("cv_mariage_nettoyage", 1000)) * nb_mar) / 12 * infl_cv

    # ── Salles chateau CV ──
    nb_loc_chateau = params.get("salles_chateau_nb_an", 0)
    cv_salles_chateau = ((params.get("cv_salles_chateau_energie", 100) +
                          params.get("cv_salles_chateau_nettoyage", 500)) * nb_loc_chateau) / 12 * infl_cv

    cv_salles = cv_seminaires + cv_mariages + cv_salles_chateau

    # ── Divers CV (integre dans hebergement) ──
    cv_divers = params.get("cv_divers_consommable", 1) * rev["nuitees"] * infl_cv
    cv_hebergement += cv_divers

    total = cv_hebergement + cv_brasserie + cv_bar + cv_spa + cv_salles

    # ── Detail par nature de fourniture (pour delais de paiement) ──
    # Decompose les charges par type de fournisseur
    cv_par_nuitee = params.get("cv_hebergement_par_nuitee", {})
    _n = rev["nuitees"]
    cv_nat_linge = cv_par_nuitee.get("Linge / Blanchisserie", 0) * _n * infl_cv
    cv_nat_produits = ((cv_par_nuitee.get("Produits accueil", 0) +
                        cv_par_nuitee.get("Produits entretien", 0) +
                        cv_par_nuitee.get("Fournitures chambres", 0)) * _n * infl_cv +
                       cv_spa_entrees * infl_cv / max(infl_cv, 0.01) +  # deja * infl_cv
                       cv_divers +
                       nb_conso * params.get("cv_bar_consommable_unite", 0.20) * infl_cv)
    # Correction: cv_spa_entrees et cv_divers sont deja avec inflation
    cv_nat_produits = ((cv_par_nuitee.get("Produits accueil", 0) +
                        cv_par_nuitee.get("Produits entretien", 0) +
                        cv_par_nuitee.get("Fournitures chambres", 0)) * _n * infl_cv +
                       cv_spa_entrees +
                       cv_divers +
                       nb_conso * params.get("cv_bar_consommable_unite", 0.20) * infl_cv)
    cv_nat_energie = (cv_par_nuitee.get("Energie variable", 0) * _n * infl_cv +
                      (params.get("cv_seminaire_energie", 75) * nb_sem +
                       params.get("cv_mariage_energie", 150) * nb_mar +
                       params.get("cv_salles_chateau_energie", 100) * nb_loc_chateau) / 12 * infl_cv)
    cv_nat_nourriture = cv_brasserie  # food cost brasserie + PDJ
    cv_nat_boissons = rev["ca_bar"] * params.get("cv_bar_pct", 0.25)
    cv_nat_commissions_ota = cv_heberg_commission
    cv_nat_commissions_cb = cv_commission_cb + cv_franchise if "cv_hebergement_par_nuitee" in params else 0
    cv_nat_soins_spa = cv_spa_soins * infl_cv / max(infl_cv, 0.01)
    # Correction: cv_spa_soins n'est pas encore * infl_cv dans la formule, mais cv_spa = (soins+entrees)*infl
    cv_nat_soins_spa = (params.get("cv_spa_soin_cout", 50) + params.get("cv_spa_produits_soin", 5)) * (nb_soins_hotel + nb_soins_ext) * infl_cv
    cv_nat_nettoyage = ((params.get("cv_seminaire_nettoyage", 500) * nb_sem +
                          params.get("cv_mariage_nettoyage", 1000) * nb_mar +
                          params.get("cv_salles_chateau_nettoyage", 500) * nb_loc_chateau) / 12 * infl_cv +
                         (params.get("cv_seminaire_pause_participant", 8) +
                          params.get("cv_seminaire_materiel_participant", 3)) * nb_participants_sem * nb_sem / 12 * infl_cv +
                         params.get("cv_seminaire_equipement", 15) * nb_sem / 12 * infl_cv)

    return {
        "cv_hebergement": cv_hebergement,
        "cv_brasserie": cv_brasserie,
        "cv_pdj": cv_pdj,
        "cv_bar": cv_bar,
        "cv_spa": cv_spa,
        "cv_seminaires": cv_seminaires,
        "cv_mariages": cv_mariages,
        "cv_salles_chateau": cv_salles_chateau,
        "cv_salles": cv_salles,
        "cv_divers": cv_divers,
        "cv_total": total,
        # Detail par nature de fourniture
        "_cv_linge": cv_nat_linge,
        "_cv_produits_consommables": cv_nat_produits,
        "_cv_energie": cv_nat_energie,
        "_cv_nourriture": cv_nat_nourriture,
        "_cv_boissons": cv_nat_boissons,
        "_cv_commissions_ota": cv_nat_commissions_ota,
        "_cv_commissions_cb_franchise": cv_nat_commissions_cb,
        "_cv_soins_spa": cv_nat_soins_spa,
        "_cv_evenements": cv_nat_nettoyage,
    }


# ─── Charges fixes ────────────────────────────────────────────────────────────

def _masse_salariale(personnel_list, charges_patronales_pct):
    """Masse salariale annuelle chargee d'une liste de personnel."""
    total = 0
    for p in personnel_list:
        total += p["cout_brut"] * (1 + charges_patronales_pct) * p["etp"]
    return total


def _charges_personnel_mois(personnel_list, cp, d: date, date_ouverture: date, inflation_rate: float):
    """
    Charges de personnel pour un mois donne.
    Le cout brut annuel encode inclut deja le 13eme mois et pecule vacances.
    Charge P&L = masse annuelle / 12 (lissee, provision mensuelle).
    Cash = salaire de base mensuel (masse / 13.92) chaque mois
           + pecule de vacances en juillet (prorata des mois travailles)
           + 13eme mois en decembre (prorata des mois travailles)
    Systeme belge : l'employeur provisionne 1/12 par mois travaille,
    paie le cumul en juillet (pecule) et decembre (13e mois).
    """
    masse_annuelle = _masse_salariale(personnel_list, cp)
    mois_depuis = _mois_depuis_ouverture(d, date_ouverture)
    infl = (1 + inflation_rate / 12) ** mois_depuis

    # Charge comptable lissee = annuel / 12
    charge_pl = masse_annuelle / 12 * infl

    # Cash : salaire de base + paiements ponctuels en juillet et decembre
    base_mois = masse_annuelle / 13.92
    pecule_annuel = 0.92 * base_mois      # double pecule vacances (part annuelle)
    treizieme_annuel = 1.0 * base_mois     # 13eme mois (part annuelle)

    cash_mois = base_mois  # salaire de base chaque mois

    # Pecule de vacances en juillet : prorata des mois depuis dernier aout (ou ouverture)
    if d.month == 7 and d >= date_ouverture:
        ref_aout = date(d.year - 1, 8, 1)
        debut = max(date_ouverture, ref_aout)
        if debut <= d:
            mois_travailles = min((d.year - debut.year) * 12 + (d.month - debut.month) + 1, 12)
        else:
            mois_travailles = 0
        cash_mois += pecule_annuel * mois_travailles / 12

    # 13eme mois en decembre : prorata des mois depuis dernier janvier (ou ouverture)
    if d.month == 12 and d >= date_ouverture:
        ref_jan = date(d.year, 1, 1)
        debut = max(date_ouverture, ref_jan)
        if debut <= d:
            mois_travailles = min((d.year - debut.year) * 12 + (d.month - debut.month) + 1, 12)
        else:
            mois_travailles = 0
        cash_mois += treizieme_annuel * mois_travailles / 12

    charge_cash = cash_mois * infl

    return charge_pl, charge_cash


def calc_charges_fixes_mensuelles(params, d: date, annee_idx: int):
    """Charges fixes mensuelles (logique Excel 7.RECAP)."""
    date_ouv = params["date_ouverture"]
    mois_depuis = _mois_depuis_ouverture(d, date_ouv)
    cp = params["charges_patronales_pct"]
    infl_pers = params.get("inflation_personnel", params["inflation_an"])
    infl_loyer = params.get("inflation_loyer", 0.02)
    infl_autres_cf = params.get("inflation_autres_charges", 0.02)
    infl_pers_cumul = (1 + infl_pers / 12) ** mois_depuis
    infl_loyer_cumul = (1 + infl_loyer / 12) ** mois_depuis
    infl_autres_cumul = (1 + infl_autres_cf / 12) ** mois_depuis

    # ── Charges fixes directes par service (personnel) ──

    # Personnel hebergement
    cf_pers_heberg_pl, cf_pers_heberg_cash = _charges_personnel_mois(
        params.get("personnel_hebergement", []), cp, d, date_ouv, infl_pers)
    # Personnel brasserie
    cf_pers_brass_pl, cf_pers_brass_cash = _charges_personnel_mois(
        params.get("personnel_brasserie", []), cp, d, date_ouv, infl_pers)
    # Personnel bar
    cf_pers_bar_pl, cf_pers_bar_cash = _charges_personnel_mois(
        params.get("personnel_bar", []), cp, d, date_ouv, infl_pers)
    # Personnel spa
    cf_pers_spa_pl, cf_pers_spa_cash = _charges_personnel_mois(
        params.get("personnel_spa", []), cp, d, date_ouv, infl_pers)
    # Personnel evenements
    cf_pers_events_pl, cf_pers_events_cash = _charges_personnel_mois(
        params.get("personnel_evenements", []), cp, d, date_ouv, infl_pers)

    cf_personnel_direct = cf_pers_heberg_pl + cf_pers_brass_pl + cf_pers_bar_pl + cf_pers_spa_pl + cf_pers_events_pl
    cf_personnel_direct_cash = cf_pers_heberg_cash + cf_pers_brass_cash + cf_pers_bar_cash + cf_pers_spa_cash + cf_pers_events_cash

    # Autres frais fixes directs (hors personnel, pas de loyer ici - loyer dans CF indirects)
    cf_autres_heberg = sum(params.get("cf_directs_hebergement", {}).values()) / 12 * infl_autres_cumul
    cf_autres_brass = sum(params.get("cf_directs_brasserie", {}).values()) / 12 * infl_autres_cumul
    cf_autres_spa = sum(params.get("cf_directs_spa", {}).values()) / 12 * infl_autres_cumul
    cf_autres_events = sum(params.get("cf_directs_evenements", {}).values()) / 12 * infl_autres_cumul
    cf_autres_directs = cf_autres_heberg + cf_autres_brass + cf_autres_spa + cf_autres_events

    # Total charges fixes directes = personnel direct + autres (pas de loyer)
    cf_directs_total = cf_personnel_direct + cf_autres_directs

    # ── Charges fixes indirectes ──

    # Personnel indirect (Row 67)
    cf_pers_indirect_pl, cf_pers_indirect_cash = _charges_personnel_mois(
        params.get("personnel_indirect", []), cp, d, date_ouv, infl_pers)

    # Loyer mensuel (dans CF indirects)
    loyer_base = params.get("loyer_mensuel", 0)
    cf_loyer = loyer_base * infl_loyer_cumul

    # Autres charges fixes indirectes (Row 68)
    # Utiliser les montants par annee si disponibles
    precompte_annees_actives = params.get("precompte_annees_actives", None)
    if "charges_fixes_indirectes_par_annee" in params:
        cf_indirect_dict = params["charges_fixes_indirectes_par_annee"]
        year_idx = min(annee_idx, 2)
        cf_indirectes_an = 0
        for k, v in cf_indirect_dict.items():
            montant = v[year_idx] if isinstance(v, list) else v
            # Precompte immobilier : vérifier si l'année est active
            if k == "Precompte immobilier" and precompte_annees_actives is not None:
                if (annee_idx + 1) not in precompte_annees_actives:
                    montant = 0
            cf_indirectes_an += montant
    else:
        cf_indirectes_an = 0
        for k, v in params.get("charges_fixes_indirectes",
                                params.get("charges_fixes_autres", {})).items():
            montant = v
            if k == "Precompte immobilier" and precompte_annees_actives is not None:
                if (annee_idx + 1) not in precompte_annees_actives:
                    montant = 0
            cf_indirectes_an += montant

    cf_autres_indirects = cf_indirectes_an / 12 * infl_loyer_cumul

    # Detail CF indirectes par nature (pour delais de paiement)
    _cfi_groupes = {
        "Energie fixe": ["Energie fixe"],
        "Assurances & taxes": ["Assurances", "Precompte immobilier", "Licences taxes SABAM"],
        "Maintenance & contrats": ["Contrats maintenance", "Securite gardiennage",
                                    "Dechets collecte", "Jardin espaces verts"],
        "Marketing": ["Marketing Communication"],
        "IT & logiciels": ["Logiciels IT"],
        "Honoraires": ["Honoraires comptable juridique"],
    }
    _cfi_par_nature = {}
    _cfi_affectes = set()
    if "charges_fixes_indirectes_par_annee" in params:
        _cfi_dict = params["charges_fixes_indirectes_par_annee"]
        _yr = min(annee_idx, 2)
        for grp, postes in _cfi_groupes.items():
            tot = 0
            for p_name in postes:
                if p_name in _cfi_dict:
                    v = _cfi_dict[p_name]
                    m = v[_yr] if isinstance(v, list) else v
                    if p_name == "Precompte immobilier" and precompte_annees_actives is not None:
                        if (annee_idx + 1) not in precompte_annees_actives:
                            m = 0
                    tot += m
                    _cfi_affectes.add(p_name)
            _cfi_par_nature[grp] = tot / 12 * infl_loyer_cumul
        # Autres = tout ce qui n'est pas dans les groupes
        _cfi_autres = 0
        for k, v in _cfi_dict.items():
            if k not in _cfi_affectes:
                m = v[_yr] if isinstance(v, list) else v
                _cfi_autres += m
        _cfi_par_nature["Autres CF indirects"] = _cfi_autres / 12 * infl_loyer_cumul
    else:
        for grp in _cfi_groupes:
            _cfi_par_nature[grp] = 0
        _cfi_par_nature["Autres CF indirects"] = cf_autres_indirects

    cf_indirects_total = cf_pers_indirect_pl + cf_loyer + cf_autres_indirects
    cf_indirects_total_cash = cf_pers_indirect_cash + cf_loyer + cf_autres_indirects

    # ── Totaux ──
    cf_personnel = cf_personnel_direct + cf_pers_indirect_pl
    cf_personnel_cash = cf_personnel_direct_cash + cf_pers_indirect_cash

    return {
        "cf_personnel": cf_personnel,
        "cf_personnel_direct": cf_personnel_direct,
        "cf_personnel_indirect": cf_pers_indirect_pl,
        "cf_loyer": cf_loyer,
        "cf_autres": cf_autres_indirects,
        "cf_autres_indirects": cf_autres_indirects,
        "cf_directs_total": cf_directs_total,
        "cf_directs_total_cash": cf_personnel_direct_cash + cf_autres_directs,
        "cf_indirects_total": cf_indirects_total,
        "cf_indirects_total_cash": cf_indirects_total_cash,
        "cf_total": cf_directs_total + cf_indirects_total,
        "cf_total_cash": (cf_personnel_direct_cash + cf_autres_directs) + cf_indirects_total_cash,
        # Detail par service (pour delais de paiement fournisseurs)
        "cf_autres_hebergement": cf_autres_heberg,
        "cf_autres_brasserie": cf_autres_brass,
        "cf_autres_spa": cf_autres_spa,
        "cf_autres_evenements": cf_autres_events,
        # CF directs par service (personnel + autres)
        "cf_directs_hebergement": cf_pers_heberg_pl + cf_autres_heberg,
        "cf_directs_brasserie": cf_pers_brass_pl + cf_autres_brass,
        "cf_directs_bar": cf_pers_bar_pl,
        "cf_directs_spa": cf_pers_spa_pl + cf_autres_spa,
        "cf_directs_evenements": cf_pers_events_pl + cf_autres_events,
        # Detail CF indirectes par nature
        "_cfi_par_nature": _cfi_par_nature,
    }


# ─── Amortissements ───────────────────────────────────────────────────────────

def calc_amortissements_mensuels(params, d: date, date_ouverture: date):
    """Amortissement lineaire mensuel des investissements initiaux."""
    amort_total = 0
    detail = {}

    for inv in params["investissements"]:
        if inv["duree_amort"] <= 0:
            continue
        amort_mensuel = inv["montant"] / (inv["duree_amort"] * 12)
        fin_amort = date_ouverture + relativedelta(years=inv["duree_amort"])
        if d < fin_amort:
            amort_total += amort_mensuel
            detail[inv["categorie"]] = amort_mensuel
        else:
            detail[inv["categorie"]] = 0

    # Reinvestissements
    mois_depuis_ouverture = (d.year - date_ouverture.year) * 12 + (d.month - date_ouverture.month)
    annee_idx = mois_depuis_ouverture // 12

    # Montants de base pour reinvest par categorie (avec leur duree d'amortissement)
    reinvest_cats = params.get("reinvest_categories", ["Amenagements interieurs", "Mobilier & Equipements"])
    reinvest_bases = []  # [(montant, duree_amort), ...]
    for inv in params["investissements"]:
        if inv["categorie"] in reinvest_cats and inv["duree_amort"] > 0:
            reinvest_bases.append((inv["montant"], inv["duree_amort"]))

    reinvest_amort = 0
    reinvest_acq_mois = 0
    pct_list = params.get("reinvest_pct_an", [0.0, 0.005, 0.01, 0.015, 0.02, 0.02, 0.025, 0.025, 0.025, 0.025])

    for a in range(annee_idx + 1):
        # Annees au-dela de la liste : reprendre le dernier %
        pct = pct_list[min(a, len(pct_list) - 1)]
        if pct > 0 and a >= 1:
            acq_annee = 0
            for montant_base, dur_amort in reinvest_bases:
                # Montant reinvesti cette annee = % x montant initial
                montant_reinvesti = pct * montant_base
                acq_annee += montant_reinvesti
                # Amortissement : regle de 3 = montant reinvesti / duree amort originale / 12
                mois_depuis_reinvest = (annee_idx - a) * 12 + (mois_depuis_ouverture % 12)
                if mois_depuis_reinvest < dur_amort * 12:
                    reinvest_amort += montant_reinvesti / (dur_amort * 12)

            # Acquisition reinvest du mois courant (reparti sur 12 mois)
            if a == annee_idx:
                reinvest_acq_mois = acq_annee / 12

    detail["Reinvestissements"] = reinvest_amort
    amort_total += reinvest_amort

    return amort_total, detail, reinvest_acq_mois


# ─── Tableau d'amortissement des prets ─────────────────────────────────────────

def calc_tableau_pret(pret: dict, date_debut: date, nb_mois_max: int = 240):
    """Genere le tableau d'amortissement d'un pret."""
    montant = pret["montant"]
    taux_mensuel = pret["taux_annuel"] / 12
    nb_mensualites = pret["duree_ans"] * 12
    differe = pret.get("differe_mois", 0)
    type_pret = pret.get("type", "annuite")

    if montant == 0 or nb_mensualites == 0:
        return pd.DataFrame()

    rows = []
    capital_restant = montant

    if type_pret == "interet_seul":
        # Pret in fine / interet seul: on ne rembourse que les interets
        for i in range(min(nb_mensualites, nb_mois_max)):
            d = date_debut + relativedelta(months=i)
            interets = capital_restant * taux_mensuel
            capital_rembourse = 0
            paiement = interets

            rows.append({
                "date": d,
                "mois": i + 1,
                "mensualite": paiement,
                "interets": interets,
                "capital": capital_rembourse,
                "capital_restant": capital_restant,
            })
    else:
        # Annuite constante (formule PMT)
        if taux_mensuel > 0:
            mensualite = montant * (taux_mensuel + taux_mensuel / ((1 + taux_mensuel) ** nb_mensualites - 1))
        else:
            mensualite = montant / nb_mensualites

        for i in range(min(nb_mensualites + differe, nb_mois_max)):
            d = date_debut + relativedelta(months=i)
            interets = capital_restant * taux_mensuel

            if i < differe:
                capital_rembourse = 0
                paiement = interets
            else:
                capital_rembourse = mensualite - interets
                paiement = mensualite

            capital_restant -= capital_rembourse

            rows.append({
                "date": d,
                "mois": i + 1,
                "mensualite": paiement,
                "interets": interets,
                "capital": capital_rembourse,
                "capital_restant": max(capital_restant, 0),
            })

            if capital_restant <= 0.01:
                break

    return pd.DataFrame(rows)


def calc_service_dette_mensuel(params, d: date):
    """Calcule le service de la dette total pour un mois donne."""
    total_mensualite = 0
    total_interets = 0
    total_capital = 0

    for pret in params["prets"]:
        if pret["montant"] == 0:
            continue
        df = calc_tableau_pret(pret, params["date_ouverture"])
        row = df[df["date"] == d]
        if not row.empty:
            total_mensualite += row.iloc[0]["mensualite"]
            total_interets += row.iloc[0]["interets"]
            total_capital += row.iloc[0]["capital"]

    return {
        "dette_mensualite": total_mensualite,
        "dette_interets": total_interets,
        "dette_capital": total_capital,
    }


# ─── Projection complete ──────────────────────────────────────────────────────

def projection_complete(params):
    """Genere la projection mensuelle complete (equivalent du 7. RECAP)."""
    dates = mois_range(params["date_ouverture"], params["nb_mois_projection"])

    # Pre-calculer les tableaux de prets (cle par index pour eviter doublons de nom)
    prets_tables = {}
    for idx_pret, pret in enumerate(params["prets"]):
        if pret["montant"] > 0:
            df = calc_tableau_pret(pret, params["date_ouverture"], params["nb_mois_projection"])
            prets_tables[idx_pret] = df

    rows = []
    _pertes_reportees = 0  # Stock de pertes fiscales reportables
    _resultat_annee = 0   # Resultat cumule de l'annee en cours
    isoc_a_payer = {}     # {annee: montant} - ISOC calcule en dec, paye en juin N+1
    tva_buffer = 0       # TVA nette accumulee (a reverser a l'etat)
    tva_a_payer = 0      # TVA du trimestre/mois precedent, payee le mois suivant
    tva_periodicite = params.get("tva_periodicite", "Trimestrielle")

    # Delais de paiement (buffers deque)
    from collections import deque
    delais_cl = params.get("delais_clients", {})
    delais_fn = params.get("delais_fournisseurs", {})
    buf_clients = {k: deque() for k in delais_cl}
    buf_fournisseurs = {k: deque() for k in delais_fn}

    # Provisions sociales (pecule + 13e mois) pour le bilan
    provision_pecule = 0.0
    provision_13e_mois = 0.0

    # Taux TVA
    tva_ventes = params.get("tva_ventes", {})
    tva_heberg = tva_ventes.get("Hebergement", 0.06)
    tva_brass_food = tva_ventes.get("Brasserie nourriture", 0.12)
    tva_brass_boisson = tva_ventes.get("Brasserie boisson", 0.21)
    tva_pdj = tva_ventes.get("Petit-dejeuner", 0.12)
    tva_bar = tva_ventes.get("Bar", 0.21)
    tva_spa = tva_ventes.get("Spa", 0.21)
    tva_salles = tva_ventes.get("Salles", 0.21)
    tva_divers = tva_ventes.get("Divers", 0.12)
    tva_loyer_resto = tva_ventes.get("Loyer Restaurant", 0.21)
    tva_achats_cv = params.get("tva_achats", {}).get("Charges variables pdj brasserie divers salles", 0.12)
    tva_achats_autres = params.get("tva_achats", {}).get("Autres", 0.21)
    part_nourriture = params.get("brasserie_part_nourriture", 0.60)

    for d in dates:
        mois_idx = (d.year - params["date_ouverture"].year) * 12 + (d.month - params["date_ouverture"].month)
        annee_idx = mois_idx // 12

        # Revenus
        rev = calc_revenus_mensuels(params, d, annee_idx)

        # Charges variables
        cv = calc_charges_variables(params, rev, annee_idx)

        # Total charges variables
        total_variables = cv["cv_total"]

        # Charges fixes
        cf = calc_charges_fixes_mensuelles(params, d, annee_idx)

        # Total charges fixes directes
        total_fixes_directes = cf["cf_directs_total"]

        # Marge (Row 63) = CA - variables - fixes directes
        marge = rev["ca_total"] - total_variables - total_fixes_directes

        # EBITDA (Row 72) = marge - charges fixes indirectes
        ebitda = marge - cf["cf_indirects_total"]

        # Amortissements (Row 76)
        amort, amort_detail, reinvest_acq = calc_amortissements_mensuels(params, d, params["date_ouverture"])

        # Resultat d'exploitation (Row 78) = EBITDA - amortissement
        resultat_exploitation = ebitda - amort

        # Service de la dette
        dette_mensualite = 0
        dette_interets = 0
        dette_capital = 0
        for idx_pret, pret in enumerate(params["prets"]):
            if idx_pret in prets_tables:
                df_pret = prets_tables[idx_pret]
                row = df_pret[df_pret["date"] == d]
                if not row.empty:
                    dette_mensualite += row.iloc[0]["mensualite"]
                    dette_interets += row.iloc[0]["interets"]
                    dette_capital += row.iloc[0]["capital"]

        # Subside RW : 1/5 du montant par an au resultat (Chateau), a partir de An 2
        # Lisse mensuellement = montant / 5 / 12
        # Pas d'impact cash (neutre : perception = remboursement)
        subside_rw_resultat = 0
        if annee_idx >= 1:  # A partir de l'annee suivant le lancement
            for pret in params["prets"]:
                if pret.get("subside_rw", False):
                    duree_subside = 5  # 5 ans pour le Chateau
                    annee_subside = annee_idx - 1  # 0-indexed depuis An 2
                    if annee_subside < duree_subside:
                        subside_rw_resultat += pret["montant"] / duree_subside / 12

        # Resultat (Row 82) = resultat exploitation - interets + subside (produit comptable)
        resultat = resultat_exploitation - dette_interets + subside_rw_resultat

        # Resultat cumule annuel (pour ISOC)
        _resultat_annee += resultat

        # Cash flow operationnel (utilise les valeurs cash pour le personnel)
        # Difference cash vs P&L = cf_total_cash - cf_total (P&L)
        cf_cash_delta = cf["cf_total_cash"] - cf["cf_total"]
        cash_flow_operationnel = ebitda - dette_interets - cf_cash_delta

        # ── ISOC : calcule en decembre, paye en juin N+1 ──
        # Les pertes se cumulent et viennent en deduction des benefices futurs
        impot_charge = 0   # Charge comptable du mois
        impot_cash = 0      # Paiement cash du mois

        # En decembre : calculer l'ISOC de l'annee
        if d.month == 12:
            if _resultat_annee > 0:
                # Benefice : deduire les pertes reportees
                _base_imposable = _resultat_annee - _pertes_reportees
                if _base_imposable > 0:
                    impot_charge = _base_imposable * params["isoc"]
                    _pertes_reportees = 0  # Toutes les pertes ont ete absorbees
                else:
                    # Les pertes reportees absorbent tout le benefice
                    impot_charge = 0
                    _pertes_reportees = -_base_imposable  # Reste des pertes non absorbees
            else:
                # Perte de l'annee : ajouter aux pertes reportees
                impot_charge = 0
                _pertes_reportees += abs(_resultat_annee)
            isoc_a_payer[d.year] = impot_charge
            _resultat_annee = 0  # Reset pour l'annee suivante

        # En juin : payer l'ISOC de l'annee precedente
        if d.month == 6 and (d.year - 1) in isoc_a_payer:
            impot_cash = isoc_a_payer.pop(d.year - 1)

        # Resultat net = resultat avant impot - charge impot
        resultat_avant_impot = resultat
        resultat_net = resultat - impot_charge

        # ── TVA : collectee - deductible ──
        # TVA collectee sur ventes
        ca_brass_hors_pdj = rev["ca_brasserie"] - rev.get("ca_pdj", 0)
        tva_collectee = (
            rev["ca_hebergement"] * tva_heberg +
            ca_brass_hors_pdj * (part_nourriture * tva_brass_food + (1 - part_nourriture) * tva_brass_boisson) +
            rev.get("ca_pdj", 0) * tva_pdj +
            rev["ca_bar"] * tva_bar +
            rev["ca_spa"] * tva_spa +
            rev.get("ca_salles", 0) * tva_salles +
            rev.get("ca_divers", 0) * tva_divers +
            rev.get("ca_loyer_restaurant", 0) * tva_loyer_resto
        )

        # TVA deductible sur achats (charges variables + partie des charges fixes)
        cv_food = cv["cv_brasserie"] + cv.get("cv_pdj", 0)
        cv_autres = cv["cv_total"] - cv_food
        tva_deductible = (
            cv_food * tva_achats_cv +
            cv_autres * tva_achats_autres +
            cf.get("cf_autres_indirects", 0) * tva_achats_autres +
            reinvest_acq * tva_achats_autres
        )

        tva_nette_mois = tva_collectee - tva_deductible
        tva_buffer += tva_nette_mois

        # Reversement TVA selon periodicite (paye le mois suivant la cloture)
        # Trimestrielle : Q1 cloture en mars -> paye en avril, etc.
        # Mensuelle : mois M cloture -> paye en M+1
        tva_paiement = 0
        if tva_periodicite == "Mensuelle":
            # Payer la TVA du mois precedent
            tva_paiement = tva_a_payer
            tva_a_payer = tva_buffer
            tva_buffer = 0
        elif tva_periodicite == "Trimestrielle":
            # Payer en avril, juillet, octobre, janvier (mois suivant fin de trimestre)
            if d.month in (1, 4, 7, 10):
                tva_paiement = tva_a_payer
                tva_a_payer = 0
            if d.month in (3, 6, 9, 12):
                # Cloture du trimestre : transferer le buffer vers a_payer
                tva_a_payer = tva_buffer
                tva_buffer = 0

        # ── Delais de paiement (BFR) ──
        # Montants factures clients (HT)
        facture_clients = {}
        for seg in params["segments"]:
            facture_clients[f"hebergement_{seg}"] = rev.get("_ca_par_segment", {}).get(seg, 0)
        facture_clients["brasserie"] = rev["ca_brasserie"]
        facture_clients["bar"] = rev["ca_bar"]
        facture_clients["spa"] = rev["ca_spa"]
        facture_clients["salles"] = rev.get("ca_salles", 0)
        facture_clients["divers"] = rev.get("ca_divers", 0)
        facture_clients["loyer_restaurant"] = rev.get("ca_loyer_restaurant", 0)

        # Montants factures fournisseurs par nature de fourniture (HT, hors personnel)
        facture_fournisseurs = {
            # Charges variables par nature
            "Linge & blanchisserie": cv.get("_cv_linge", 0),
            "Produits & consommables": cv.get("_cv_produits_consommables", 0),
            "Energie variable": cv.get("_cv_energie", 0),
            "Nourriture": cv.get("_cv_nourriture", 0),
            "Boissons": cv.get("_cv_boissons", 0),
            "Commissions OTA": cv.get("_cv_commissions_ota", 0),
            "Commissions CB & franchise": cv.get("_cv_commissions_cb_franchise", 0),
            "Soins spa": cv.get("_cv_soins_spa", 0),
            "Evenements (nettoyage, materiel)": cv.get("_cv_evenements", 0),
            # Charges fixes par nature
            "Loyer": cf["cf_loyer"],
            "CF directs departements": (cf.get("cf_autres_hebergement", 0) +
                                        cf.get("cf_autres_brasserie", 0) +
                                        cf.get("cf_autres_spa", 0) +
                                        cf.get("cf_autres_evenements", 0)),
        }
        # Ajouter le detail CF indirectes par nature
        for grp, montant in cf.get("_cfi_par_nature", {}).items():
            facture_fournisseurs[grp] = montant

        # Appliquer delais et calculer cash encaisse/decaisse
        cash_clients = 0.0
        total_facture_cl = 0.0
        for key, amount in facture_clients.items():
            total_facture_cl += amount
            delay = delais_cl.get(key, 0)
            if key not in buf_clients:
                buf_clients[key] = deque()
            buf_clients[key].append(amount)
            if len(buf_clients[key]) > delay:
                cash_clients += buf_clients[key].popleft()

        cash_fournisseurs = 0.0
        total_facture_fn = 0.0
        for key, amount in facture_fournisseurs.items():
            total_facture_fn += amount
            delay = delais_fn.get(key, 0)
            if key not in buf_fournisseurs:
                buf_fournisseurs[key] = deque()
            buf_fournisseurs[key].append(amount)
            if len(buf_fournisseurs[key]) > delay:
                cash_fournisseurs += buf_fournisseurs[key].popleft()

        delay_adjustment = (cash_clients - total_facture_cl) - (cash_fournisseurs - total_facture_fn)

        # Creances et dettes (bilan)
        creances_clients = sum(sum(q) for q in buf_clients.values())
        dettes_fournisseurs = sum(sum(q) for q in buf_fournisseurs.values())
        dette_tva = tva_a_payer + tva_buffer
        dette_isoc = sum(isoc_a_payer.values())

        # Provisions sociales (pecule + 13e mois)
        all_pers = (params.get("personnel_hebergement", []) +
                    params.get("personnel_brasserie", []) +
                    params.get("personnel_bar", []) +
                    params.get("personnel_spa", []) +
                    params.get("personnel_evenements", []) +
                    params.get("personnel_indirect", []))
        masse_tot = _masse_salariale(all_pers, params["charges_patronales_pct"])
        infl_pers_prov = (1 + params.get("inflation_personnel", params["inflation_an"]) / 12) ** mois_idx
        base_prov = masse_tot / 13.92 * infl_pers_prov
        pecule_mensuel = 0.92 * base_prov / 12
        treizieme_mensuel = 1.0 * base_prov / 12

        if d.month == 7:
            provision_pecule = 0.0
        else:
            provision_pecule += pecule_mensuel

        if d.month == 12:
            provision_13e_mois = 0.0
        else:
            provision_13e_mois += treizieme_mensuel

        dette_sociale = provision_pecule + provision_13e_mois

        # ── Cash flow libre ──
        # Le modele travaille en HT. Chaque mois, l'hotel encaisse la TVA de ses
        # clients et decaisse la TVA a ses fournisseurs : le delta net (tva_nette_mois)
        # est du cash en banque. Ce cash ne sort qu'au paiement trimestriel (tva_paiement).
        # Le delay_adjustment reflète le decalage entre facturation et encaissement/paiement.
        cash_flow = (cash_flow_operationnel + delay_adjustment + tva_nette_mois
                     - tva_paiement - dette_capital - impot_cash - reinvest_acq)

        # Marge brute pour compatibilite
        marge_brute = rev["ca_total"] - cv["cv_total"]

        row_data = {
            "date": d,
            "annee": d.year,
            "mois": d.month,
            "mois_exploitation": mois_idx + 1,
            "annee_exploitation": annee_idx + 1,
            # Revenus
            "nuitees": rev["nuitees"],
            "taux_occupation": rev["taux_occupation"],
            "prix_moyen": rev["prix_moyen"],
            "ca_hebergement": rev["ca_hebergement"],
            "ca_brasserie": rev["ca_brasserie"],
            "ca_pdj": rev.get("ca_pdj", 0),
            "ca_bar": rev["ca_bar"],
            "ca_spa": rev["ca_spa"],
            "ca_salles": rev.get("ca_salles", 0),
            "ca_seminaires": rev["ca_seminaires"],
            "ca_mariages": rev["ca_mariages"],
            "ca_salles_chateau": rev.get("ca_salles_chateau", 0),
            "ca_divers": rev.get("ca_divers", 0),
            "ca_loyer_restaurant": rev.get("ca_loyer_restaurant", 0),
            "ca_total": rev["ca_total"],
            # Charges variables
            "cv_hebergement": cv["cv_hebergement"],
            "cv_brasserie": cv["cv_brasserie"],
            "cv_pdj": cv["cv_pdj"],
            "cv_bar": cv["cv_bar"],
            "cv_spa": cv["cv_spa"],
            "cv_seminaires": cv["cv_seminaires"],
            "cv_mariages": cv["cv_mariages"],
            "cv_salles_chateau": cv.get("cv_salles_chateau", 0),
            "cv_salles": cv.get("cv_salles", 0),
            "cv_divers": cv.get("cv_divers", 0),
            "cv_total": cv["cv_total"],
            # Marges
            "marge_brute": marge_brute,
            "marge_brute_pct": marge_brute / rev["ca_total"] if rev["ca_total"] > 0 else 0,
            "total_variables": total_variables,
            "total_fixes_directes": total_fixes_directes,
            "marge": marge,
            # Charges fixes
            "cf_personnel": cf["cf_personnel"],
            "cf_personnel_direct": cf.get("cf_personnel_direct", 0),
            "cf_personnel_indirect": cf.get("cf_personnel_indirect", 0),
            "cf_loyer": cf["cf_loyer"],
            "cf_autres": cf["cf_autres"],
            "cf_autres_indirects": cf.get("cf_autres_indirects", 0),
            "cf_directs_total": cf.get("cf_directs_total", 0),
            "cf_directs_hebergement": cf.get("cf_directs_hebergement", 0),
            "cf_directs_brasserie": cf.get("cf_directs_brasserie", 0),
            "cf_directs_bar": cf.get("cf_directs_bar", 0),
            "cf_directs_spa": cf.get("cf_directs_spa", 0),
            "cf_directs_evenements": cf.get("cf_directs_evenements", 0),
            "cf_indirects_total": cf.get("cf_indirects_total", 0),
            "cf_total": cf["cf_total"],
            "cf_total_cash": cf.get("cf_total_cash", cf["cf_total"]),
            # Resultats
            "ebitda": ebitda,
            "ebitda_pct": ebitda / rev["ca_total"] if rev["ca_total"] > 0 else 0,
            "amortissement": amort,
            "ebit": resultat_exploitation,
            "resultat_exploitation": resultat_exploitation,
            # Dette
            "dette_mensualite": dette_mensualite,
            "dette_interets": dette_interets,
            "dette_capital": dette_capital,
            # Resultat final
            "resultat_avant_impot": resultat_avant_impot,
            "resultat": resultat,
            "impot": impot_charge,
            "impot_cash": impot_cash,
            "resultat_net": resultat_net,
            # TVA
            "tva_collectee": tva_collectee,
            "tva_deductible": tva_deductible,
            "tva_nette": tva_nette_mois,
            "tva_paiement": tva_paiement,
            # Cash flow
            "cash_flow_operationnel": cash_flow_operationnel,
            "delay_adjustment": delay_adjustment,
            "cash_flow": cash_flow,
            "reinvest_acquisition": reinvest_acq,
            "subside_rw": subside_rw_resultat,
            # Bilan
            "creances_clients": creances_clients,
            "dettes_fournisseurs": dettes_fournisseurs,
            "bfr": creances_clients - dettes_fournisseurs,
            "dette_tva": dette_tva,
            "dette_isoc": dette_isoc,
            "dette_sociale": dette_sociale,
            "provision_pecule": provision_pecule,
            "provision_13e_mois": provision_13e_mois,
        }
        rows.append(row_data)

    df = pd.DataFrame(rows)

    # Cumuls
    df["ca_cumul"] = df["ca_total"].cumsum()
    df["resultat_net_cumul"] = df["resultat_net"].cumsum()
    df["cash_flow_cumul"] = df["cash_flow"].cumsum()

    return df


# ─── Indicateurs annuels ──────────────────────────────────────────────────────

def indicateurs_annuels(df: pd.DataFrame, params: dict, par_calendaire=False) -> pd.DataFrame:
    """Agrege la projection mensuelle en indicateurs annuels.

    Si par_calendaire=True, regroupe par année calendaire au lieu de l'année d'exploitation.
    """
    group_col = "annee" if par_calendaire else "annee_exploitation"
    grouped = df.groupby(group_col)

    rows = []
    for annee, g in grouped:
        ca = g["ca_total"].sum()
        ca_heberg = g["ca_hebergement"].sum()
        ca_brass = g["ca_brasserie"].sum()
        ca_autres = ca - ca_heberg - ca_brass
        cv = g["cv_total"].sum()
        cf = g["cf_total"].sum()
        cf_directs = g["cf_directs_total"].sum() if "cf_directs_total" in g.columns else 0
        cf_indirects = g["cf_indirects_total"].sum() if "cf_indirects_total" in g.columns else 0
        ebitda = g["ebitda"].sum()
        ebit = g["ebit"].sum()
        rn = g["resultat_net"].sum()
        cf_flow = g["cash_flow"].sum()
        nuitees = g["nuitees"].sum()
        amort = g["amortissement"].sum()
        dette_int = g["dette_interets"].sum()
        dette_cap = g["dette_capital"].sum()
        dette_mens = g["dette_mensualite"].sum()

        nb_ch = params["nb_chambres"]
        jours_an = g.apply(lambda r: jours_dans_mois(r["date"]), axis=1).sum()
        chambres_dispo = nb_ch * jours_an

        taux_occ_moyen = nuitees / chambres_dispo if chambres_dispo > 0 else 0
        revpar = ca_heberg / chambres_dispo if chambres_dispo > 0 else 0
        prix_moyen = ca_heberg / nuitees if nuitees > 0 else 0
        goppar = ebitda / chambres_dispo if chambres_dispo > 0 else 0

        # Marge = CA - CV - CF directes
        marge = ca - cv - cf_directs

        annee_cal = int(g["annee"].iloc[0])
        annee_expl = int(g["annee_exploitation"].iloc[0])

        rows.append({
            "Annee exploitation": annee_expl,
            "Annee calendaire": annee_cal,
            "Mois": len(g),
            "CA Total": ca,
            "CA Hebergement": ca_heberg,
            "CA Brasserie": ca_brass,
            "CA Autres": ca_autres,
            "Charges Variables": cv,
            "Marge Brute": ca - cv,
            "Marge Brute %": (ca - cv) / ca * 100 if ca > 0 else 0,
            "Charges Fixes Directes": cf_directs,
            "Charges Fixes Indirectes": cf_indirects,
            "Marge": marge,
            "Charges Fixes": cf,
            "EBITDA": ebitda,
            "EBITDA %": ebitda / ca * 100 if ca > 0 else 0,
            "Amortissement": amort,
            "EBIT": ebit,
            "Interets": dette_int,
            "Resultat": ebit - dette_int,
            "Resultat Net": rn,
            "Cash Flow": cf_flow,
            "Cash Flow Cumul": g["cash_flow_cumul"].iloc[-1],
            "Nuitees": nuitees,
            "Taux Occupation": taux_occ_moyen * 100,
            "Prix Moyen (ADR)": prix_moyen,
            "RevPAR": revpar,
            "GOPPAR": goppar,
            "Service Dette": dette_mens,
            "DSCR": ebitda / dette_mens if dette_mens > 0 else float("inf"),
        })

    return pd.DataFrame(rows)
