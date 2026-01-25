# -*- coding: utf-8 -*-
"""
Created on Sun Nov  9 12:23:12 2025

@author: valer
"""

import os
import requests
import xml.etree.ElementTree as ET
from pathlib import Path
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from scipy.interpolate import interp1d
import numpy as np
import time
from collections import Counter
from scipy.signal import savgol_filter

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert

import lib_vspec as vsp


# TODO : mettre les dates dans graphiques
# TODO : gerer la BR
# ALGO : rectifier le profil
# ALGO : decaler avec vitesse heliocentric
# ALGO : sample au plus grand echx  


def logme (msg, flag_both = True) :
    with open("report.txt", "a") as f:
        f.write(msg+"\n")
        if flag_both :
            print(msg)
    
def is_file_locked(filepath):
    """Renvoie True si le fichier est ouvert/verrouillé par une autre application."""
    if not os.path.exists(filepath):
        return False
    try:
        with open(filepath, 'a'):
            pass
        return False
    except PermissionError:
        return True
    
def get_VOlist_from_object(Be_TargetName, Be_date_d, Be_date_f, Be_lamb_d, Be_HR, MaxRecord=1000) :
    # BeSS_VO est le repertoire de telechargement créé à la racine de l'application
    App_Path = Path(__file__).resolve().parent

    # Prétraitement du nom de cible
    if Be_TargetName.isnumeric():
        Be_TargetName = "HD" + Be_TargetName

    # Remplacement manuel des caractères spéciaux
    Be_TargetName = (
        Be_TargetName
        .replace("+", "%2B")
        .replace("[", "%5B")
        .replace("]", "%5D")
    )
    # Maximum record
    max_rec = 'MAXREC='+str(MaxRecord)
    
    # Formatage des dates
    if len(Be_date_d) == 4:
        Be_date_d += "-01-01"
    if len(Be_date_f) == 4:
        Be_date_f += "-01-01"

    if Be_TargetName == '' :
        #Be_URL = "http://basebe.obspm.fr/cgi-bin/ssapBE_1.0.pl?MAXREC=1000"
        Be_URL = "http://basebe.obspm.fr/cgi-bin/ssapBE_1.0.pl?"+max_rec
    else :
        # Construction de l’URL
        #Be_URL = f"http://basebe.obspm.fr/cgi-bin/ssapBE_1.0.pl?TARGETNAME={Be_TargetName}&MAXREC=1000"
        Be_URL = f"http://basebe.obspm.fr/cgi-bin/ssapBE_1.0.pl?TARGETNAME={Be_TargetName}&"+max_rec

    # Préparation du fichier local
    local_dir = os.path.join(App_Path, "BeSS_VO")
    os.makedirs(local_dir, exist_ok=True)
    local_file = os.path.join(local_dir, "Bess_query_VO.pl")

    if os.path.exists(local_file):
        os.remove(local_file)

    # Ajout des paramètres facultatifs
    if Be_date_d and Be_date_f:
        Be_URL += f"&TIME={Be_date_d}/{Be_date_f}"

    if Be_lamb_d:
        lamb = float(Be_lamb_d) * 1e-10
        Be_URL += f"&BAND={lamb}"

    if Be_HR == 1:
        HR_seuil = 4500.0
        Be_URL += f"&SPECRES={HR_seuil}/"
        

    # Téléchargement du fichier
    try:
        response = requests.get(Be_URL, timeout=15)
        response.raise_for_status()
        with open(local_file, "wb") as f:
            f.write(response.content)
    except Exception as e:
        print(f"Erreur lors du téléchargement : {e}", "warning")
        return

    # Vérification
    if os.path.exists(local_file):
        #print("*** liste telechargée")
        pass
    else:
        print( "warning")
        
def parse_xml_to_table() :
    # Chemin vers le fichier XML qui est dans BeSS_VO à la racine de l'application
    app_path = Path(__file__).resolve().parent
    xml_file = app_path / "BeSS_VO" / "Bess_query_VO.pl"

    if not xml_file.exists():
        print("Fichier XML introuvable :", xml_file)
        return []

    # Charger le XML
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Définir le namespace
    ns = {"bk": "http://www.ivoa.net/xml/VOTable/v1.1"}

    # Récupérer toutes les lignes TR
    rows = root.findall(".//bk:TR", ns)

    results = []

    for row in rows:
        cells = row.findall("bk:TD", ns)
        if not cells or len(cells) < 62:  # au moins 62 colonnes
            continue

        # Extraction des valeurs (VB6 index 0, 28, 61)
        fname = cells[0].text[47:] if cells[0].text else ""
        date_obs = cells[61].text if cells[61].text else ""
        observateur = fname.split('%2F')[0][2:]  # [2:] retire "A_"
        l = float(cells[28].text) * 1e10 if cells[28].text else 0.0
        lamb = f"{l:.2f}"
        obj = cells[4].text[5:]

        results.append({
            "object" : obj,
            "fichiers": fname,
            "observer": observateur,
            "date": date_obs,
            "lamb": lamb,
            "checked": True,
        })

    # Tri (reverse true du plus recent au plus anciens)
    results.sort(key=lambda x: x["date"], reverse=True)

    return results

def download_files(files, save_dir):
    """
    Télécharge les fichiers listés dans `files` depuis `base_url` et les enregistre dans `save_dir`.
    Args:

        files (list of str): noms des fichiers à télécharger.
        save_dir (str or Path): répertoire local où enregistrer les fichiers.
    """
    
    base_url = 'http://basebe.obspm.fr/cgi-bin/extBeSS.pl?fits='
    save_dir = Path(save_dir)
    save_dir.mkdir(parents=True, exist_ok=True)

    for fname in files:
        url = base_url+fname  # construit l'URL complète
        local_file = save_dir / fname

        try:
            #print(f"Téléchargement : {url}")
            response = requests.get(url, timeout=30)
            response.raise_for_status()  # déclenche une erreur si le téléchargement échoue
            with open(local_file, "wb") as f:
                f.write(response.content)
            #print(f"Enregistré dans : {local_file}")
        except Exception as e:
            print(f"Erreur pour {fname}: {e}")
    
def object_list_comp (object_name, mois_courant, year_courant, nb_to_open = 3) :
    # prends tous les spectres du mois et les 3 précedents si ils existent
    # les telecharge
    # les normalises entre lamb_norm1 et lamb_norm2
    # et si la zone n'est pas zone de recouvrement alors modifie
    # affiche en vignettes
    # superpose pour comparaison aka re-echantillonage du dernier spectre
    # sur la grille du premier précédents du mois d'avant ou encore avant
    

    # on force sur lambda H-alpha et spectres HR
    Be_lamb_d = '6563'
    
    # TODO gerer si basse resolution
    Be_HR=1
    
    # année et mois 
    #now = datetime.now()
    #year = now.year
    
    month = mois_courant
    year = year_courant
    if month == 12 :
        year = year+1 
        month = 1
    else :
        month = mois_courant +1
    
        
    datem = f"{year}-{month:02d}-01"
    
    # Convertir en datetime
    date_obj = datetime.strptime(datem, "%Y-%m-%d")
    previous_day = date_obj - timedelta(days=1)
    month = previous_day.month
    date_fin = date_obj.strftime("%Y-%m-%d")
    date_deb = "1901-01-01"
    flag_maxrec = True
    
    # fait la requete
    while flag_maxrec :
        get_VOlist_from_object(object_name, date_deb, date_fin, Be_lamb_d, Be_HR)
        table = parse_xml_to_table()
        flag_maxrec = False
        if len(table)== 1000 :
            flag_maxrec = True
            Be_year = str(int(date_deb[:4])+ 2)
            if Be_year == '1903' :
                Be_year = '2020'
            date_deb = Be_year+'-01-01'
    
    if len(table) == 0 :
        print("Pas de spectre HR")
        Be_HR=0
        # fait la requete
        get_VOlist_from_object(object_name, date_deb, date_fin, Be_lamb_d, Be_HR)
        table = parse_xml_to_table()
    
    if len(table)>= 1000 :
        print("Erreur catch max record")
        #date_deb = "2020-01-01"
        #get_VOlist_from_object(object_name, date_deb, date_fin, Be_lamb_d, Be_HR)
        #table = parse_xml_to_table()

    # mais on ne garde de la table que les spectres du mois courant
    # et les nb_to_open précedent, si ils existent
    prefix = f"{year_courant}-{mois_courant:02d}"  # ex : "2025-11"

    #if object_name == "V442 And" :
        #print(" stop ")
        
    current_month_files = [
        row["fichiers"]
        for row in table
        if row.get("checked") and row.get("date","").startswith(prefix)
        ]
    # Indices de slice : à partir de len(current_month_files), nb_to_open éléments
    if len(current_month_files) != 0 :
        next_nb_to_open_files = [
            row["fichiers"]
            for row in table[len(current_month_files) : len(current_month_files)+ nb_to_open]
            if row.get("checked")
            ]
        file_names = current_month_files + next_nb_to_open_files
        index_to_comp = len(current_month_files)
        if len(next_nb_to_open_files) == 0 :
            Be_HR =0
            get_VOlist_from_object(object_name, date_deb, date_fin, Be_lamb_d, Be_HR)
            table = parse_xml_to_table()
            current_month_files = [
                row["fichiers"]
                for row in table
                if row.get("checked") and row.get("date","").startswith(prefix)
                ]
            if len(current_month_files) != 0 :
                next_nb_to_open_files = [
                    row["fichiers"]
                    for row in table[len(current_month_files) : len(current_month_files)+ nb_to_open]
                    if row.get("checked")
                    ]
                file_names = current_month_files + next_nb_to_open_files
                index_to_comp = len(current_month_files)
                if len(next_nb_to_open_files) == 0 :
                    # spectre unique
                    return file_names, index_to_comp
                else :
                    # les spectres suivants sont surement BR
                    print("Spectre suivants non HR")
                    index_to_comp = -1
                    return file_names, index_to_comp
            
    else :
        # sans doute un spectre BR...
        print("Spectre du mois non HR")
        Be_HR=0
        # fait la requete
        get_VOlist_from_object(object_name, date_deb, date_fin, Be_lamb_d, Be_HR)
        table = parse_xml_to_table()
        
        current_month_files = [
            row["fichiers"]
            for row in table
            if row.get("checked") and row.get("date","").startswith(prefix)
            ]
        # Indices de slice : à partir de len(current_month_files), nb_to_open éléments
        if len(current_month_files) != 0 :
            next_nb_to_open_files = [
                row["fichiers"]
                for row in table[len(current_month_files) : len(current_month_files)+ nb_to_open]
                if row.get("checked")
                ]
            file_names = current_month_files + next_nb_to_open_files
            index_to_comp = len(current_month_files)
                
        else :
            print("Il y a un pb...")
            file_names=[]
            index_to_comp=-1
   
    return file_names, index_to_comp

def object_detect_change (object_name, zone_norm, month_now, year_now, nb_to_open, flag_metrics=False) :
    
    save_dir = Path(__file__).resolve().parent / "BeSS_VO"
    file_names=[]
    

    logme("*********************", False)
    logme(object_name, False)


    # zone de normalisation
    #lamb_norm1 = zone_norm[0]
    #lamb_norm2 = zone_norm[1]
    
        
    # construit la liste avec les fichiers du mois plus les nb_to_open précédents
    file_names, index_to_comp = object_list_comp(object_name, month_now, year_now, nb_to_open )
    
    if len(file_names) == 0 :
       logme( object_name + ' non trouvé')
       
       decision = "Unique"
       delt_ew = 0
       return decision, delt_ew
   
    if index_to_comp == - 1 :
       decision = "HR-BR"
       delt_ew = 0
       return decision, delt_ew
    
    if len(file_names) == 1 :
       decision = "Unique"
       delt_ew = 0 
       return decision, delt_ew
   
    
   
    # premier essai
    try :
        f = file_names[0]
    except:
        time.sleep(5)
        print("Delai...")
    
    # telecharge les fichiers
    download_files(file_names, save_dir)
    
    # premier essai
    try :
        f = file_names[0]
    except:
        time.sleep(5)
        print("Delai...")
    # deuxieme essai
    try :
        f = file_names[0]
    except:
        time.sleep(5)
        print("Delai...")

    
    # il faut ensuite comparer les deux profil 
    # pour file_names[0] et file_names[index_to_comp]
    
    # lecture des deux profils
    lamb_last, pro_last, hdr_last = vsp.read_fits_table(save_dir/file_names[0])
    lamb_comp, pro_comp, hdr_comp = vsp.read_fits_table(save_dir/file_names[index_to_comp])
    lamb_min = np.max([lamb_last[0], lamb_comp[0]])
    lamb_max = np.min([lamb_last[-1], lamb_comp[-1]])

    lamb_ref = np.arange(round(lamb_min+0.5), round(lamb_max-0.5), 0.1, dtype=np.float64)
    
    date_obs_last = hdr_last['DATE-OBS'].split('T')[0]
    date_obs_comp = hdr_comp['DATE-OBS'].split('T')[0]
    BSS_vhel_last = -hdr_last['BSS_RQVH']
    BSS_vhel_comp = -hdr_comp['BSS_RQVH']
    
    # normalise entre lamb_norm1 et lamb_norm2 
    pro_last_norm,_,_ = vsp.profil_norm(pro_last, lamb_last, zone_norm)
    pro_comp_norm,_,_ = vsp.profil_norm(pro_comp, lamb_comp, zone_norm)
    

    # affiche les echantillonage
    echx1 = np.diff(lamb_last)[0]
    echx2 = np.diff(lamb_comp)[0]
    logme("Echx1 : "+ f"{echx1:.2f}"+" ang/pix" )
    logme("Echx2 : "+ f"{echx2:.2f}"+ " ang/pix")
    
    interp_comp = interp1d(lamb_comp, pro_comp_norm, kind='linear', bounds_error=False, fill_value=np.nan)
    pro_comp_norm = interp_comp(lamb_ref)
    interp_last = interp1d(lamb_last, pro_last_norm, kind='linear', bounds_error=False, fill_value=np.nan)
    pro_last_norm = interp_last(lamb_ref)

    # on a deux profils avec meme lamb et normalisés
    # pro_comp_norm et pro_comp_last sur lamb_ref
    # calcul std
    std1,_,_ =  vsp.profil_std(pro_last_norm, lamb_ref, zone_norm)
    std2,_,_ = vsp.profil_std(pro_comp_norm, lamb_ref, zone_norm)

    # faire un crop sur zone
    band = (6540, 6586)
    pro2b, lamb_ref_crop = vsp.profil_crop (pro_comp_norm, lamb_ref, band)
    pro1b, lamb_ref_crop = vsp.profil_crop (pro_last_norm, lamb_ref, band)
    
    # zappe les raies telluriques
    pro2 = vsp.profile_zap_atm(pro2b, lamb_ref_crop)
    pro1 = vsp.profile_zap_atm(pro1b, lamb_ref_crop)
    
    # filtre bruit
    pro2 = savgol_filter(pro2, window_length=31, polyorder=3)
    pro1 = savgol_filter(pro1, window_length=31, polyorder=3)
    
    # Correction vitesse helio
    pro1,_ = vsp.profil_corr_vhel(pro1, lamb_ref_crop, BSS_vhel_last)
    pro2,_ = vsp.profil_corr_vhel(pro2, lamb_ref_crop, BSS_vhel_comp)
    

    # Calcul largeur equivalente ---
    ew1,pro1_ew,_ = vsp.profil_leq(pro1, lamb_ref_crop, lamb_ref_crop[0], lamb_ref_crop[-1], 50)
    ew2,pro2_ew,_ = vsp.profil_leq(pro2, lamb_ref_crop, lamb_ref_crop[0], lamb_ref_crop[-1], 50)
    delta_ew = ew1-ew2
    ew_moy = abs(ew1+ew2)/2
    percent_ew = round(((abs(delta_ew) / (abs(ew1+ew2)*0.5)) *100)+0.5)
    seuil_ew = 15 # seuil en pourcentage de EW
    seuil_ew_forme = 10
    seuil_ew_ME = 5 # seuil en pourcentage de EW
    
    # Calcul variation de forme (corrélation) ---
    corr = np.corrcoef(pro1, pro2)[0,1]
    seuil_corr = 0.92  # si corr < seuil → forme différente was 0.88
    
    # Calcul avec la diff des profils
    
    pro_diff = abs(pro1_ew-pro2_ew)
    pro_diff[pro_diff<0.03]=0
    somme = round((np.sum(pro_diff))+0.5)
      
    # Calcul du chi2
    eps = 1e-12
    # std1 et std2 calculés sur zone de norm
    var = std1**2 + std2**2 + eps 
    
    chi2 = np.sum((pro1 - pro2)**2 / var)
    chi2_red = chi2 / (len(pro1) - 1)
    
    if ew_moy > 16 :
        seuil_diff = 30
        seuil_diff2 = 16 # was 18
    elif ew_moy > 5 :
        seuil_diff = 20
        seuil_diff2 = 10
    else :
        seuil_diff = 15
        seuil_diff2= 7
    
      
    # Décision automatique ---
    if ew_moy  >= 20 :
        variation_ew = abs(percent_ew) >= seuil_ew
        perc_ew =  abs(percent_ew)
        delt_ew = abs(delta_ew)
    
    elif ew_moy <=4 :
        variation_ew= abs(delta_ew)>=0.4
        perc_ew =  abs(percent_ew)
        delt_ew =abs(delta_ew)
    
    else :
        
        variation_ew= abs(delta_ew)>=1
        perc_ew =  abs(percent_ew)
        delt_ew =abs(delta_ew)
    
    variation_ew2 = percent_ew >= seuil_ew_ME
    variation_forme = corr <= seuil_corr
    variation_diff = (abs(somme) >= seuil_diff)
    variation_diff2 = (abs(somme) >= seuil_diff2)
    test_chi2 = chi2_red > 1.5
    
    
    logme(f"Diff ew : {delt_ew :.1f}")
    logme(f"Chi2 : {chi2_red :.1f}")
    
    logme("--------------------------------------------------------------")
    leq_data = f"Leq last : {ew1:.2f},Leq comp : {ew2:.2f}, Var de leq : {delt_ew:.1f}  {perc_ew:.0f} %  → {'Oui' if variation_ew else 'Non'}"
    forme_data =f"Var de forme : corr={corr:.3f} → {'Oui' if variation_forme else 'Non'}"
    diff_data = f"Var de différence : diff={somme:.0f} seuil={seuil_diff:.0f} → {'Oui' if variation_diff else 'Non'}"
    diff2_data = f"Var de différence2 : diff={somme:.0f} seuil={seuil_diff2:.0f}→ {'Oui' if variation_diff2 else 'Non'}"
    chi2_data = f"Test Chi2 : {chi2_red:.1f} → {'Oui' if test_chi2 else 'Non'}"
    logme(leq_data)
    logme(forme_data) 
    logme(diff_data)
    logme(diff2_data)
    logme(chi2_data)
    
  
    
    sens = "EE" if ew2 > ew1 else "DE"
    
    if test_chi2 :
    
        if (variation_ew and variation_diff) or (variation_forme and percent_ew > seuil_ew_forme):
            decision = sens
    
        elif ((variation_ew and variation_diff2) or (variation_ew2 and variation_diff2))  :
        #  elif ((variation_ew and variation_diff2) or (variation_ew2 and variation_diff)):
            decision = "ME"
        
        elif  ((variation_diff2) or (variation_forme)) :
            decision = "SE"
            
        else :
            decision = "-"
    
    else:
        decision = "-"
  
    logme("---------------------")
    logme (object_name + ":  "+ decision)
    logme("---------------------")
    
    
    """
    # Decision
    if variation_ew and variation_diff :
        if ew2 > ew1 :
            print("---------------------")
            print (object_name + ":  EE ")
            decision = "EE"
        else :
            print("---------------------")
            print (object_name + ":  DE ")
            decision = "DE"
    
    
    
    elif variation_forme and percent_ew > seuil_ew_ME:
        print("percent")
        if ew2 > ew1 :
            print("---------------------")
            print (object_name + ":  EE ")
            decision = "EE"
        else :
            print("---------------------")
            print (object_name + ":  DE ")
            decision = "DE"
    
    
    elif variation_ew and not variation_diff and variation_diff2:
        print("---------------------")
        print (object_name + ":  ME ")
        decision ="ME" 
    
    elif percent_ew > seuil_ew_ME and variation_diff :
        print("---------------------")
        print (object_name + ":  ME ")
            
    elif variation_forme  and not variation_ew :
        print("-------------------   --")
        print (object_name + ":  ME ")
        decision ="ME"
    
    elif variation_diff and not variation_ew and not variation_forme:
        print("---------------------")
        print (object_name + ":  ME ")
        decision ="ME" 
    
    else :
        print("---------------------")
        print (object_name + ":  - ")
        decision ="-"
    """
    print(' ')
    
    try :
        with open("tableau.txt", "a") as f:
            f.write(object_name+","+str(round(ew1,2))+","+str(round(ew2,2))
                    +","+str(round(ew_moy,1))
                    +","+str(round(percent_ew,0))
                    +","+str(round(delta_ew,1))
                    + ","+str(round(corr,3))+","+str(round(somme,0))+","+ decision+"\n")
    except:
        pass
    
    if 2 == 1 :
        # trace difference
        plt.plot(lamb_ref_crop, pro_diff)
        plt.xlabel('Longueur d\'onde')
        plt.ylabel('Intensité')
        plt.title('Différence')
        plt.show()
    
    
    # ---- Tracer le spectre
    
    plt.figure()
    plt.margins(y=0.2)
    plt.plot(lamb_ref_crop, pro2_ew, label = date_obs_comp)
    plt.plot(lamb_ref_crop, pro1_ew, label = date_obs_last)
    ymin, ymax = plt.ylim()
    plt.ylim(ymin, ymax) # etait plt.ylim(0, ymax)
    plt.xlabel('Longueur d\'onde')
    plt.ylabel('Intensité')
    title = object_name + " : " + decision
    plt.title(title)
    #subtitle = f"{ew_moy:.2f}" + "  "+f"{delta_ew:.1f}" + "  "+f"{percent_ew:.0f}"+ "  " + f"{corr:.2f}"+"  "+f"{somme:.2f}"
    #plt.suptitle (subtitle)
    
    
    # mode debug pour impression criteres
    
    if flag_metrics :
        ax = plt.gca()   # axe courant
        # Décalage vertical sous l’axe X (en coordonnées d’axes)
        y0 = -0.3   # première ligne
        dy = 0.1    # espacement entre lignes
        
        texts = [leq_data, forme_data, diff_data, diff2_data, chi2_data]
        
        for i, txt in enumerate(texts):
            ax.text(
                0, y0 - i * dy,
                txt,
                transform=ax.transAxes,
                ha="left",
                va="top"
            )
        
        # Laisser de la place en bas pour le texte
        plt.subplots_adjust(bottom=0.35)
    
    fn = object_name+'.png'
    plt.legend()
    plt.savefig(save_dir/fn, bbox_inches="tight")
    plt.show()
    plt.close()

    
    
    return decision, delta_ew
    
def object_list_from_dates (object_name, date_deb, date_fin, lamb_raie, flag_HR, MaxRecord=1000) :
    # Variables d'entrée 
    Be_TargetName = object_name
    Be_date_d = date_deb
    Be_date_f = date_fin
    Be_lamb_d = lamb_raie
    Be_HR = flag_HR
    flag_maxrec = True

    save_dir = Path(__file__).resolve().parent / "BeSS_VO"
    
    while flag_maxrec :
        get_VOlist_from_object(Be_TargetName, Be_date_d, Be_date_f, Be_lamb_d, Be_HR, MaxRecord)
        table = parse_xml_to_table()
        flag_maxrec = False
        if len(table)== 1000 :
            flag_maxrec = True
            Be_year = str(int(date_deb[:4])+ 2)
            if Be_year == '1903' :
                Be_year = '2020'
            Be_date_d = Be_year+'-01-01'
            
        
    # telecharge dans le répertoire les fichiers si checked est true 
    file_names = [row["fichiers"] for row in table if row.get("checked")]
    save_dir = Path(__file__).resolve().parent / "BeSS_VO"
    download_files(file_names, save_dir)
    print("fichiers téléchargés")

    return file_names

def get_all_spectres_between_dates (Be_date_d, Be_date_f):
    Be_TargetName = ''
    Be_HR = 0 
    Be_lamb_d = 6563.0
    
    get_VOlist_from_object(Be_TargetName, Be_date_d, Be_date_f, Be_lamb_d, Be_HR)
    table= parse_xml_to_table()
    
    object_list = list(set([row["object"] for row in table]))
    observer_list = list(set([row["observer"] for row in table]))
    nb_spc = str(len(table))
    count_per_observer = Counter(s["observer"] for s in table)
    observer_list = [[obs, count] for obs, count in count_per_observer.items()]
    
    return object_list, observer_list, nb_spc


def object_composer (object_name, month_now, year_now, flag_thumb) :
    # Download tous les spectres entre date_deb et date_fin pour un objet
    save_dir = Path(__file__).resolve().parent / "BeSS_VO"
    
    # comparaison pour un objet
    #object_name= 'OT Gem'
    cols = 6
    nb_max_pro = 5 *cols # maximum 5 lignes
    
    # récupère le mois et l'année
    #month_now = now.month
    #year_now = now.year

    #date_fin = '2025-12-01'
    date_fin = f"{year_now}-{month_now:02d}-01"
    date_deb = '1901-01-01'
    lamb_raie = '6563'
    flag_HR = 1
    
    zone_norm = (6610.0, 6620.0)
    lamb_min = 6540
    lamb_max=6586
    
    file_names = object_list_from_dates(object_name, date_deb, date_fin, lamb_raie, flag_HR) 
        
    if len(file_names) == 0 :
        print("Pas de spectres trouvés")
        exit()
        
    if len(file_names) == 1000 :
        print("Erreur catch maxrecord")
        #date_deb = '2020-01-01'
        #file_names = object_list_from_dates(object_name, date_deb, date_fin, lamb_raie, flag_HR)
    
    if len(file_names) >= nb_max_pro :
        file_names = file_names[:nb_max_pro]
        
        
    # tableau de profils
    hdrs = []
    profils = []
    lambs = []
    
    for f in file_names:
        filename = save_dir/f
        # Ouvrir le fichier FITS
        lamb, pro, hdr = vsp.read_fits_table(filename)
        # vitesse helio
        BSS_vhel = -hdr['BSS_RQVH']
        # formatte les profils
        pro,_,_ = vsp.profil_norm (pro, lamb, zone_norm)
        pro,_ = vsp.profil_corr_vhel (pro, lamb, BSS_vhel)
        hdrs.append(hdr)
        profils.append(pro)
        lambs.append(lamb)

    all_max = max(pro.max() for pro in profils) * 1.1
    all_min = min(pro.min() for pro in profils) * 0.9
    
    if flag_thumb == True :

        # creation graphique avec vignettes
        n = len(profils)
        rows = (n + cols - 1) // cols
    
        fig, axes = plt.subplots(rows, cols, figsize=(3*cols, 3*rows),
                             sharex=False, sharey=False)
        
        # force axes sous forme de tableau 2D
        axes = np.array(axes).reshape(rows, cols)
        # titre nom de l'objet
        fig.suptitle(object_name, fontsize = 16 , fontweight='bold')
        
        for i, profile in enumerate(profils):
            
            date_obs = hdrs[i]['DATE-OBS'].split('T')[0]
            
            r = i // cols
            c = i % cols
            ax = axes[r, c]
        
            ax.plot(lambs[i],profile, linewidth=1, label= date_obs)
            
            # ticks vers l'intérieur sur tous les côtés et ticks actifs en haut/droite
            ax.tick_params(axis='x', direction='in', top=True, bottom=True)
            ax.tick_params(axis='x', labelbottom=True, labeltop=False,
                       pad=-15)   # pad négatif pour les faire entrer dans la zone
            ax.xaxis.set_ticks_position('both')   # affiche ticks en bas ET en haut
            # pas de labels en Y
            ax.set_ylabel("")
            # retirer labels Y
            ax.tick_params(axis='y', labelleft=False, labelright=False)
            
            ax.tick_params(top=True, right=True)  # active ticks top/right
            
            ax.set_xlim(lamb_min, lamb_max)
            #ax.set_ylim(0, all_max)
            ax.set_ylim(all_min, all_max)
            
            # legend
            ax.legend(frameon=False,handlelength=1) # trait court
        
        # supprimer les cases vides
        for j in range(n, rows * cols):
            fig.delaxes(axes[j // cols, j % cols])
        
        
        # enlever les marges autour de la grille
        #plt.subplots_adjust(wspace=0, hspace=0, left=0, right=1, top=1, bottom=0)
        plt.subplots_adjust(wspace=0.0, hspace=0.0,
                        left=0.02, right=0.98, top=0.98, bottom=0.02)
        fn = object_name+"_t.png"
        plt.savefig(save_dir/fn, bbox_inches="tight")
        plt.show()
        
    else : # diagramme 2D 
        
        offset = 0.8 * (all_max-all_min)   # marge pour éviter toute superposition
        
        
        if len(profils) >= 10 :
            profils = profils[:10]
        
        fig, ax = plt.subplots(figsize=(2, len(profils)*0.5))

        
        for i in range(len(profils)-1, -1, -1) :
            if i == 0 :
                color='orange'
            else :
                color='C0'
                
            j = len(profils)-1 - i # offset pour aller du haut vers le bas
            
            date_obs = hdrs[i]['DATE-OBS'].split('T')[0]
            ax.plot(lambs[i], profils[i] + j * offset, linewidth=1, label= date_obs, color=color)
            # Position de l’annotation (à droite du profil)
            x_text = 6540
            # indice du x le plus proche
            idx = np.abs(lamb - x_text).argmin()
            # valeur de y sur le profil, avec offset vertical
            if idx >= len(profils[i]) :
                idx = len(profils[i])-1
            y_text = profils[i][idx] + j * offset              
        
            ax.text(
                x_text, y_text+0.1, date_obs,
                ha='left', va='bottom',
                fontsize=7,
                color='black'
            )
            
        ax.set_xlim(lamb_min, lamb_max)
        # pas de labels en Y
        ax.set_ylabel("")
        # retirer labels Y
        ax.tick_params(axis='y',left=False, labelleft=False, labelright=False)
        
        # enlever les marges autour de la grille
        fontsize_pt = 11           # taille du texte en points
        fig.suptitle(object_name, fontsize = fontsize_pt , fontweight='bold')
        
        fontsize_inch = fontsize_pt / 72  # hauteur approximative en inches
        fig_height_inch = fig.get_figheight()
        top_margin = 1 - fontsize_inch / fig_height_inch - 0.02  # petite marge supplémentaire
        fig.subplots_adjust(top=top_margin)
        fig.subplots_adjust(wspace=0.0, hspace=0.0,
                        left=0.02, right=0.98, top=top_margin, bottom=0.02)
        # Légende
        #ax.legend(loc="upper right")

        fn = object_name+"_s.png"
        fig.savefig(save_dir/fn, bbox_inches="tight")
        plt.show()

def create_monthly_word (month_name, year_name, nb_stars, nb_spectra,liste_observers,liste_EE, liste_ME, liste_DE, liste_SE, liste_novar, liste_unique, flag_with_novar=False) :
    save_dir = Path(__file__).resolve().parent / "BeSS_VO"
    
    template_path = "Template Bess report.docx"
   
    images_per_row = 3
    display_width = Inches(2)  # adapté au format A4
    #display_height = Inches(1.5)
    cols_table = 6           # nombre de colonnes pour les noms
    
    nom_report = "BeSS report "+month_name+" "+year_name+".docx"
    date_report = " "+month_name+" "+year_name
    output_docx = save_dir / nom_report
    
    nb_obs = str(len(liste_observers))
    observers = [row[0] for row in liste_observers]
    nbobs = [str(row[1]) for row in liste_observers]
    # tri par ordre decroissant
    observers, nbobs = zip(*sorted(zip(observers, nbobs), key=lambda x: int(x[1]), reverse=True))
    observers = list(observers)
    nbobs = list(nbobs)
    
    metrics = nb_stars +"     " + nb_spectra + "     " + nb_obs
    
    
    # --- Créer le document Word ---
    doc = Document(template_path)
    #doc.add_heading("BeSS Monthly Report "+ month_name +" "+str(year), level=0)
    
    style = doc.styles['Table Style']  # par exemple 'Normal', 'Titre', etc.
    style.font.size = Pt(9)  # 14 points
                    
    table = doc.tables[0]
    
    cell = table.cell(0, 0) #li, co
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            if '<<date>>' in run.text:
                run.text = run.text.replace('<<date>>', date_report)
    
    
    cell = table.cell(1, 1)
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            if '<<metrics>>' in run.text:
                run.text = run.text.replace('<<metrics>>', metrics)
                run.font.size = Pt(48)
        
    cell = table.cell(2, 0)
    # parcourir les paragraphes de la cellule
    for paragraph in cell.paragraphs:
        if '<<observers>>' in paragraph.text:

            # Nettoyage : vider tous les runs
            for run in paragraph.runs:
                run.text = ""

            # --- Construire le tableau juste après avoir effacé la cellule ---
            n_rows = len(observers)

            if n_rows == 0:
                cell.paragraphs[0].add_run("(no observers)")
                break

            # On crée un paragraphe pour accueillir le tableau
            container_p = cell.add_paragraph()

            # --- Créer le tableau dans la cellule ---
            inner_table = cell.add_table(rows=n_rows + 1, cols=2)
            inner_table.style = "Table Style"   # ton style Word existant

            # en-têtes
            headers = ["Observer", "nb"]
            header_color = "DBE5F1"
            for col, header in enumerate(headers):
                hdr_cell = inner_table.cell(0, col)
                p_hdr = hdr_cell.paragraphs[0]
                run_hdr = p_hdr.add_run(header)
                run_hdr.bold = True
                p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER

                shading_elm = OxmlElement("w:shd")
                shading_elm.set(qn("w:fill"), header_color)
                hdr_cell._tc.get_or_add_tcPr().append(shading_elm)

            # --- Remplissage ---
            for i in range(n_rows-1):
                inner_table.cell(i + 1, 0).text = observers[i]
                inner_table.cell(i + 1, 1).text = nbobs[i]

            # --- Centrer le texte partout ---
            for row in inner_table.rows:
                for c in row.cells:
                    for p in c.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        break
            
    
    cell = table.cell(2, 1)
    # parcourir les paragraphes de la cellule
    for paragraph in cell.paragraphs:
        # Si la balise est dans le paragraphe
        if '<<events>>' in paragraph.text:
            # On parcourt les runs pour trouver la balise exacte (permet de conserver le style)
            for run in paragraph.runs:
                if '<<events>>' in run.text:
                    # découper autour de la balise (before / placeholder / after)
                    before, _, after = run.text.partition('<<events>>')
    
                    # garder le texte avant dans le run courant
                    run.text = before
    
                    # --- Construire le tableau DANS la cellule ---
                    n_rows = max(len(liste_EE), len(liste_ME), len(liste_DE))
                    if n_rows == 0:
                        # Aucun data → on affiche un message et on sort
                        paragraph.add_run(" (no events) ")
                        break
    
                    # créer le tableau dans la cellule (ajouté à la fin du contenu de la cellule)
                    inner_table = cell.add_table(rows=n_rows + 1, cols=3)
                    inner_table.style = "Table Style"  # ou autre style disponible dans ton template
    
                    # en-têtes
                    headers = ["EE", "ME", "DE"]
                    header_color = "DBE5F1"
                    for col, header in enumerate(headers):
                        hdr_cell = inner_table.cell(0, col)
                        p_hdr = hdr_cell.paragraphs[0]
                        run_hdr = p_hdr.add_run(header)
                        run_hdr.bold = True
                        p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
                        shading_elm = OxmlElement("w:shd")
                        shading_elm.set(qn("w:fill"), header_color)
                        hdr_cell._tc.get_or_add_tcPr().append(shading_elm)
    
                    # remplir les lignes
                    for i in range(n_rows):
                        inner_table.cell(i + 1, 0).text = liste_EE[i] if i < len(liste_EE) else ""
                        inner_table.cell(i + 1, 1).text = liste_ME[i] if i < len(liste_ME) else ""
                        inner_table.cell(i + 1, 2).text = liste_DE[i] if i < len(liste_DE) else ""
    
                    # centrer le texte dans le tableau
                    for row in inner_table.rows:
                        for c in row.cells:
                            for p in c.paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
                    # --- Remettre le texte qui était après la balise (s'il existe) ---
                    if after:
                        # Ajouter un nouveau run après le tableau (le texte après)
                        paragraph.add_run(after)
    
                    # On a remplacé la balise -> on peut sortir
                    break
            # Si on a fait le remplacement on peut sortir de la boucle des paragraphes
            break
    
    
    for p in doc.paragraphs:       
        if "<<DATA>>" in p.text:
            # --- Découper le texte autour de la balise ---
            before, _, after = p.text.partition("<<DATA>>")
    
            # On nettoie le paragraphe courant et remet le texte "avant"
            for run in p.runs:
                run.text = ""
            if before.strip():
                p.add_run(before)
                
            # --- Tableau avec les noms des objets ---
            new_p = p.insert_paragraph_before("Object with no variations")
            new_p.style = "Heading 2"
            
            table = doc.add_table(rows=1, cols=cols_table)  # Word veut au moins 1 ligne
            table.style = "Table Style" 
            # Déplacer le tableau juste avant le paragraphe contenant la balise
            p._p.addprevious(table._tbl)

            
            # --- Remplir le tableau ---
            for i, name in enumerate(liste_novar):
                if i % cols_table == 0 and i != 0:
                    table.add_row()
                row_idx = i // cols_table
                col_idx = i % cols_table
                table.cell(row_idx, col_idx).text = name
        
            # Centrer le texte dans le tableau
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            p.insert_paragraph_before() # espace après le tableau
            
            if liste_unique :       
                # --- Tableau avec les noms des objets ---
                new_p = p.insert_paragraph_before("Object with only one HR spectrum ")
                new_p.style = "Heading 2"
                
                table = doc.add_table(rows=1, cols=cols_table)  # Word veut au moins 1 ligne
                table.style = "Table Style"
                # Déplacer le tableau juste avant le paragraphe contenant la balise
                p._p.addprevious(table._tbl)
                
                # --- Remplir le tableau ---
                for i, name in enumerate(liste_unique):
                    if i % cols_table == 0 and i != 0:
                        table.add_row()
                    row_idx = i // cols_table
                    col_idx = i % cols_table
                    table.cell(row_idx, col_idx).text = name
                
                # Centrer le texte dans le tableau
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        
            p.insert_paragraph_before() # espace après le tableau
        
            # --- Charger toutes les images PNG EE ---
            new_p = p.insert_paragraph_before("Emission Event")
            new_p.style = "Heading 2"
            new_p.paragraph_format.keep_with_next = True
            
            png_names = [n+'.png' for n in liste_EE]
            png_files = [save_dir/n for n in png_names]
            
            # --- Construire le tableau d'images EE ---
            if png_files:
                table = doc.add_table(rows=1, cols=images_per_row)  # Word veut au moins 1 ligne
    
                # Déplacer le tableau juste avant le paragraphe contenant la balise
                p._p.addprevious(table._tbl)
                
                row_cells = table.rows[0].cells
            
                for i, png_file in enumerate(png_files):
                    # Déterminer la ligne et la colonne
                    row_idx = i // images_per_row
                    col_idx = i % images_per_row
            
                    # Ajouter une nouvelle ligne si nécessaire
                    if row_idx >= len(table.rows):
                        row_cells = table.add_row().cells
                    else:
                        row_cells = table.rows[row_idx].cells
            
                    # Ajouter l'image dans la cellule
                    cell = row_cells[col_idx]
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run()
                    run.add_picture(str(png_file), width=display_width)
            
                    # Centrer le contenu de la cellule
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    

                    
            
            p.insert_paragraph_before() # espace après le tableau
            
            # --- Charger toutes les images PNG DE ---
            new_p = p.insert_paragraph_before("Decreasing Event")
            new_p.style = "Heading 2"
            new_p.paragraph_format.keep_with_next = True
            
            png_names = [n+'.png' for n in liste_DE]
            png_files = [save_dir/n for n in png_names]
            
            # --- Construire le tableau d'images DE ---
            if png_files:
                table = doc.add_table(rows=1, cols=images_per_row)  # Word veut au moins 1 ligne
                # Déplacer le tableau juste avant le paragraphe contenant la balise
                p._p.addprevious(table._tbl)
                
                row_cells = table.rows[0].cells
            
                for i, png_file in enumerate(png_files):
                    # Déterminer la ligne et la colonne
                    row_idx = i // images_per_row
                    col_idx = i % images_per_row
            
                    # Ajouter une nouvelle ligne si nécessaire
                    if row_idx >= len(table.rows):
                        row_cells = table.add_row().cells
                    else:
                        row_cells = table.rows[row_idx].cells
            
                    # Ajouter l'image dans la cellule
                    cell = row_cells[col_idx]
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run()
                    run.add_picture(str(png_file), width=display_width)
            
                    # Centrer le contenu de la cellule
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            p.insert_paragraph_before() # espace après le tableau
            
            # --- Charger toutes les images PNG ME ---
            new_p = p.insert_paragraph_before("Moderate Event")
            new_p.style = "Heading 2"
            new_p.paragraph_format.keep_with_next = True
            
            png_names = [n+'.png' for n in liste_ME]
            png_files = [save_dir/n for n in png_names]
            
            # --- Construire le tableau d'images ME ---
            if png_files:
                table = doc.add_table(rows=1, cols=images_per_row)  # Word veut au moins 1 ligne
                # Déplacer le tableau juste avant le paragraphe contenant la balise
                p._p.addprevious(table._tbl)
                
                row_cells = table.rows[0].cells
            
                for i, png_file in enumerate(png_files):
                    # Déterminer la ligne et la colonne
                    row_idx = i // images_per_row
                    col_idx = i % images_per_row
            
                    # Ajouter une nouvelle ligne si nécessaire
                    if row_idx >= len(table.rows):
                        row_cells = table.add_row().cells
                    else:
                        row_cells = table.rows[row_idx].cells
            
                    # Ajouter l'image dans la cellule
                    cell = row_cells[col_idx]
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run()
                    run.add_picture(str(png_file), width=display_width)
            
                    # Centrer le contenu de la cellule
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                     
            
            
            p.insert_paragraph_before() # espace après le tableau
            
        
            # --- Charger toutes les images PNG SE ---
            new_p = p.insert_paragraph_before("Shape Event")
            new_p.style = "Heading 2"
            new_p.paragraph_format.keep_with_next = True
            
            #png_names = [n+'.png' for n in liste_novar]
            png_names = [n+'.png' for n in liste_SE]
            png_files = [save_dir/n for n in png_names]
            
            # --- Construire le tableau d'images novar ---
            if png_files:
                table = doc.add_table(rows=1, cols=images_per_row)  # Word veut au moins 1 ligne
                # Déplacer le tableau juste avant le paragraphe contenant la balise
                p._p.addprevious(table._tbl)
                
                row_cells = table.rows[0].cells
            
                for i, png_file in enumerate(png_files):
                    # Déterminer la ligne et la colonne
                    row_idx = i // images_per_row
                    col_idx = i % images_per_row
            
                    # Ajouter une nouvelle ligne si nécessaire
                    if row_idx >= len(table.rows):
                        row_cells = table.add_row().cells
                    else:
                        row_cells = table.rows[row_idx].cells
            
                    # Ajouter l'image dans la cellule
                    cell = row_cells[col_idx]
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run()
                    run.add_picture(str(png_file), width=display_width)
            
                    # Centrer le contenu de la cellule
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
           
                p.insert_paragraph_before() # espace après le tableau
                
            if flag_with_novar :
                # --- Charger toutes les images PNG novar ---
                new_p = p.insert_paragraph_before("No variations")
                new_p.style = "Heading 2"
                new_p.paragraph_format.keep_with_next = True
                
                #png_names = [n+'.png' for n in liste_novar]
                png_names = [n+'.png' for n in liste_novar]
                png_files = [save_dir/n for n in png_names]
                
                # --- Construire le tableau d'images novar ---
                if png_files:
                    table = doc.add_table(rows=1, cols=images_per_row)  # Word veut au moins 1 ligne
                    # Déplacer le tableau juste avant le paragraphe contenant la balise
                    p._p.addprevious(table._tbl)
                    
                    row_cells = table.rows[0].cells
                
                    for i, png_file in enumerate(png_files):
                        # Déterminer la ligne et la colonne
                        row_idx = i // images_per_row
                        col_idx = i % images_per_row
                
                        # Ajouter une nouvelle ligne si nécessaire
                        if row_idx >= len(table.rows):
                            row_cells = table.add_row().cells
                        else:
                            row_cells = table.rows[row_idx].cells
                
                        # Ajouter l'image dans la cellule
                        cell = row_cells[col_idx]
                        paragraph = cell.paragraphs[0]
                        run = paragraph.add_run()
                        run.add_picture(str(png_file), width=display_width)
                
                        # Centrer le contenu de la cellule
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
               
                    p.insert_paragraph_before() # espace après le tableau
                
           
            # ---- insertion time serie des EE

            new_p = p.insert_paragraph_before("Emission Event time serie")
            new_p.style = "Heading 2"
            new_p.paragraph_format.keep_with_next = True
            
            png_names = [n+'_s.png' for n in liste_EE]
            png_files = [save_dir/n for n in png_names]
            
            # --- Construire le tableau d'images EE time serie ---
            if png_files:
                table = doc.add_table(rows=1, cols=images_per_row)  # Word veut au moins 1 ligne
                # Déplacer le tableau juste avant le paragraphe contenant la balise
                p._p.addprevious(table._tbl)
                images_per_row = 3
                display_width = Inches(2)  # adapté au format A4
                
                row_cells = table.rows[0].cells
            
                for i, png_file in enumerate(png_files):
                    # Déterminer la ligne et la colonne
                    row_idx = i // images_per_row
                    col_idx = i % images_per_row
            
                    # Ajouter une nouvelle ligne si nécessaire
                    if row_idx >= len(table.rows):
                        row_cells = table.add_row().cells
                    else:
                        row_cells = table.rows[row_idx].cells
            
                    # Ajouter l'image dans la cellule
                    cell = row_cells[col_idx]
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run()
                    run.add_picture(str(png_file), width=display_width)
            
                    # Centrer le contenu de la cellule
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
           
   
           
            # --- Réinsérer le texte après la balise ---
            if after.strip():
                p.add_run(after)
    
    
    # --- Sauvegarder ---
    if is_file_locked(output_docx):
        print(f"⚠️  Le fichier '{output_docx}' est actuellement ouvert dans Word.")
        input("👉  Fermez le fichier, puis appuyez sur Entrée pour continuer...")
                
    doc.save(output_docx)
    convert(output_docx)
    print(f"✅ Rapport Word et pdf généré : {output_docx}")
    
# --------------------------------------------------------------------
# --------------------------------------------------------------------
# ---- MAIN - interrogation BeSS
# --------------------------------------------------------------------

print("Lancement de la requête")

#flag = 0 # detection changement par comparaison avec mois précédent pour un objet
#flag = 1 # spectres d'un objet entre deux dates with thumbnails or multiple serie imaging
#flag = 2 # Rapport mensuel automatique, pour tous les spectres de tous les objets entre deux dates detection changement


flag = 2
flag_with_novar = False # imprime aussi les no var
flag_metrics = False # imprime les criteres

# rapport du mois 
mois = 12

now= datetime(2025,mois, 10) # rapport du mois


if flag == 0 :
    # comparaison pour un objet
    object_name= 'HD 37541'
    nb_to_open = 3
    zone_norm = (6610.0, 6620.0)
    
    # récupère le mois et l'année
    month_now = now.month
    year_now = now.year
    
    decision, delta_ew = object_detect_change(object_name,  zone_norm, month_now, year_now,nb_to_open, flag_metrics)
    print(object_name, decision)


elif flag == 1 :
    # Download tous les spectres entre date_deb et date_fin pour un objet
    save_dir = Path(__file__).resolve().parent / "BeSS_VO"
    
    # comparaison pour un objet
    object_name= 'V413 Aur'
    
   
    
    # récupère le mois et l'année
    month_now = now.month
    year_now = now.year
    year_name = str(now.year)
    
  
    
    object_composer(object_name, month_now, year_now, flag_thumb=False)
    

elif flag == 2 :
    
    with open("tableau.txt", "w") as f: #debug file to test classifications
        f.write("\n")
    
    
    # gestion de la date
    months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
    ]
    
    # récupère le mois et l'année
    month_now = now.month
    year_now = now.year
    year_name = str(now.year)
    
    month_name = months[month_now-1]
    
    # année et mois courant
    if month_now == 12 or month_now == 1 :
        Be_date_f = now.replace(day=31)
        Be_date_d = now.replace(day=1)
    else :  
        Be_date_f = now.replace(day=1)
        Be_date_d = (Be_date_f - timedelta(days=1)).replace(day=1)
    date_fin = Be_date_f.strftime("%Y-%m-%d")
    date_deb = Be_date_d.strftime("%Y-%m-%d")
    
    # on lance la requete
    object_list, observer_list, nb_spectra = get_all_spectres_between_dates (date_deb, date_fin)
    nb_stars = str(len(object_list))
    print("Nombre objets : " + nb_stars)
    print("Nombre spectres : " + nb_spectra)
    print("Nombre observateurs : " + str(len(observer_list)))
    print('')
    
    # liste decision
    liste_EE = []
    liste_ME = []
    liste_DE = []
    liste_SE = []
    liste_novar = []
    liste_unique = []
    liste_EE_dew = []
    liste_DE_dew = []
    i = 1 
    
    
    
    for o in object_list:
        
        print("************")
        print(str(i)+" - "+ o)
        object_name= o
        nb_to_open = 3
        zone_norm = (6610.0, 6620.0)
        
        decision, dew = object_detect_change(object_name, zone_norm, month_now, year_now, nb_to_open, flag_metrics)
        #input("Appuie sur Entrée pour continuer...")
        if decision == "EE" :
            liste_EE.append(o)
            liste_EE_dew.append(abs(dew))
        elif decision == "ME" :
            liste_ME.append(o)
        elif decision == "DE" :
            liste_DE.append(o)
            liste_DE_dew.append(abs(dew))
        elif decision == "SE" :
            liste_SE.append(o)
        elif decision == "Unique" :
            liste_unique.append(o)
        else :
            if not decision == "HR-BR":
                liste_novar.append(o)
        i +=1
   
    # trie les liste EE par difference de EW
    liste_EE = [ee for ee, _ in sorted(
                zip(liste_EE, liste_EE_dew),
                key=lambda x: x[1],
                reverse=True
                )]
    
    liste_DE = [de for de, _ in sorted(
                zip(liste_DE, liste_DE_dew),
                key=lambda x: x[1],
                reverse=True
                )]
    
    print("time series")
    
    for o in liste_EE :
        print(o)
        
        object_composer(o, month_now, year_now, flag_thumb=False)
    
    print("Generation rapport word")
    # automatic BeSS monthly report with evolutions classification
    create_monthly_word(month_name, year_name, nb_stars, nb_spectra,observer_list,liste_EE, liste_ME, liste_DE, liste_SE, liste_novar, liste_unique, flag_with_novar)
    

