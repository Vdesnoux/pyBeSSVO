# -*- coding: utf-8 -*-
"""
Created on Sun Nov 16 17:40:13 2025

@author: valerie DESNOUX

********************************************************************

Spectral processing lib from Visual Spec

********************************************************************
"""


from astropy.io import fits
import numpy as np



def read_fits_table (file_name) :
    
    # Lecture fichier fits table 
    #----------------------------------------------------
    
    # Ouvrir le fichier FITS
    with fits.open(file_name, memmap=False) as hdul:
        hdr=  hdul[0].header
        data = hdul[1].data
        lamb = data['wave']
        pro = data['flux']

    return lamb, pro, hdr



def profil_leq (pro, lamb, lamb1, lamb2, nbpts) :
        
    # Calcul de largeur equivalente entre lamb1 et lamb2
    #----------------------------------------------------
    
    # nombre de points en debut et fin de la moyenne pour calculer le continuum
    n_points = nbpts
    
    # Extraction du sous-spectre
    mask = (lamb >= lamb1) & (lamb <= lamb2)
    if not np.any(mask):
        raise ValueError("Aucun point dans l'intervalle lamb1-lamb2.")
    
    lamb_sub = lamb[mask]
    pro_sub = pro[mask]

    # Calcul du continuum linéaire
    flux_start = np.mean(pro_sub[:n_points])
    lamb_start = np.mean(lamb_sub[:n_points])
    
    flux_end = np.mean(pro_sub[-n_points:])
    lamb_end = np.mean(lamb_sub[-n_points:])
    
    a = (flux_end - flux_start) / (lamb_end - lamb_start)
    b = flux_start - a * lamb_start
    continuum = a * lamb_sub + b

    # Normalisation par le continuum
    flux_norm = pro_sub / continuum

    # Calcul de la largeur équivalente
    f = 1 - flux_norm  # absorption
    dl = np.diff(lamb_sub)
    f_avg = (f[:-1] + f[1:]) / 2
    ew = np.sum(f_avg * dl)

    return ew,flux_norm, lamb_sub



def profil_norm (pro, lamb, lamb_norm1, lamb_norm2) :
    
    # Normalisation du profil entre lamb1 et lamb2
    #----------------------------------------------------
    
    # verif si la range de norm est dans les deux spectres
    if lamb_norm1 < lamb[0] :
        lamb_norm2 = lamb_norm2+ (lamb[0]-lamb_norm1)
        lamb_norm1 = lamb[0]
    if lamb_norm2 > lamb[-1] :
        lamb_norm1 = lamb_norm1+ (lamb[-1]-lamb_norm2)
        lamb_norm2 = lamb[-1]
        
    # on normalise
    mask = (lamb >= lamb_norm1) & (lamb <= lamb_norm2)
    if not np.any(mask):
        raise ValueError("Aucun point dans l’intervalle de normalisation.")
    mean_norm = np.nanmean(pro[mask])
    pro_norm = pro / mean_norm
    return pro_norm



def profil_crop (pro, lamb, zone) :
    
    # Decoupe le profil entre lamb1 et lamb2
    #----------------------------------------------------
    
    lamb1 = zone[0]
    lamb2= zone[1]
    mask = (lamb >= lamb1) & (lamb <= lamb2)
    if not np.any(mask):
        raise ValueError("Aucun point dans l’intervalle demandé.")
    
    lamb_crop =lamb[mask] 
    pro_crop =pro[mask]
    
    return pro_crop, lamb_crop



def profile_zap_atm (pro, lamb) :
    
    # Remplace les raies telluriques par une droite
    # Valide autour de H-alpha
    #----------------------------------------------------
    
    
    halpha = 6563
    lines_atm = np.array([6523.850,6530.598,6532.359,6532.359,6534.0,6536.726,6542.313, 6543.912, 6545.781,  6547.705, 6548.627, 6552.632, 
                          6553.785,6557.171, 6564.196, 6568.806,6572.072, 6574.847, 6580.794,  6583.6, 6586.559, 6594.375])
    lamb_min = lamb[0]
    lamb_max = lamb[-1]
    echx = np.diff(lamb)[0]
    range_echx = 0.5 # demi largeur en angstrom de raie tellurique
    nbpas = int(range_echx // echx) +1
    
    if lamb[0] > halpha or lamb[-1] < halpha :
        # le profil n'est pas dans la région de H-alpha
        print(" Le profil ne contient pas H-alpha")
    else :
        # garde les raies qui sont dans le profil
        mask = (lines_atm >= lamb_min) & (lines_atm <= lamb_max)
        mylines_atm = lines_atm [mask]
        # tableau des index des lamb les plus proches des valeurs théorique
        idx = [np.abs(lamb - val).argmin() for val in mylines_atm]
        pro_modif = pro.copy()

        for i in idx:
            i1 = max(0, i - nbpas)
            i2 = min(len(pro), i + nbpas + 1)
            x = np.arange(i1, i2 + 1)
            # Calcul de la pente et de l'ordonnée à l'origine
            a = (pro_modif[i2] - pro_modif[i1]) / (i2 - i1)
            b = pro_modif[i1] - a * i1
            
            # Calcul des valeurs de la droite
            pro_modif[i1:i2 + 1] = a * x + b
            
        return pro_modif
    


def profil_corr_vhel(pro, lamb, vhel):
    
    # Corrige le profil de la vitesse heliocentrque vhel
    #----------------------------------------------------
    
    c = 299792.458  # km/s
    
    facteur = 1 + vhel/c

    # nouvelle grille déplacée
    lamb_shift = lamb * facteur

    # interpolation du flux décalé sur la grille originale
    pro_corr = np.interp(lamb, lamb_shift, pro)

    return pro_corr, lamb
    
