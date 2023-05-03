'''
Code par FERREIRA RODRIGUES Rodrigo
      et ABIBOU M'hammed Salman

Fait en mai 2023

But : calculer le TRI d'un projet fourni au format .xlsx

'''

## import_librairies

import pandas as pd
import math as m
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

## init_var_glob

epsilon = 0.0001
nb_itmax = 30

## lecture_donnees

#pip install pandas si la librairie n'est pas encore installée

#load the xlsx file 
data = pd.read_excel("Data/Projet_eche5.xlsx",skiprows=2, usecols="B:H") #le fichier excel doit contenir qu'une seule ligne contenant les données des flux

# Read the values of the file in the dataframe
#data = pd.DataFrame(data)

#Print data
print("Le contenu des données : \n", data, "\n")

#The value of the sell is the last column of the second line
valRevente = data.iloc[0, (len(data.axes[1]) - 1)]
print("La valeur de revente est : ",valRevente,"\n")

#The value of I0
I0 = data.iloc[0, 0]
print("L'investissement initial est de : ",-I0,"\n")

#The values of B from 1 to n
B = data.iloc[0, 1:(len(data.axes[1]) - 1)]
print("Valeurs des Bk : \n",B,"\n")

#n, the number of years
n = len(data.axes[1]) - 2
print("Le nombre d'années est : ", n, "\n")

#tau
tau = 0.01
 
## Function calcul_VAN

def calcul_VAN(data, n, tau, valRevente = 0):
    sum = 0
    if(tau > 0):
        for i in range(1, n):
            sum += data.iloc[0, i]/pow(1 + tau, i)
        sum += (data.iloc[0, n] + valRevente)/pow(1 + tau, n)
    return I0 + sum

## Function calcul_echeance_moy

def calcul_echeance_moy(data, n, tau, valRevente = 0):
    sumFlux = 0
    sumFluxAct = 0
    for i in range(1, n):
        sumFlux += data[i]
        sumFluxAct += data[i]/pow(1+tau, i)
    sumFlux += data[n] + valRevente
    sumFluxAct += (data[n] + valRevente)/pow(1+tau, n)
    return m.log((sumFlux) /sumFluxAct) / m.log(1 + tau)

## Function calcul_tri_aux

def calcul_tri_aux(I0, B, d, valRevente = 0):
    return pow(((B + valRevente) / -I0), 1 / d) - 1

## Function calcul_tri

def calcul_tri(data, n, tau0, valRevente = 0):
    tri = 0 # Initialisation de la variable tri à 0, qui sera utilisée pour stocker la valeur du TRI.
    if(tau0 > 0): # Vérification si le taux d'actualisation initial (tau0) est supérieur à 0
        tau = tau0
        i = 0 # Initialisation du compteur i à 0, qui sera utilisé pour suivre le nombre d'itérations.
        arret = False # Initialisation de la variable arret à False, qui sera utilisée pour déterminer si l'itération doit être arrêtée.
        somme = sum(B) # Calcul de la somme des valeurs des flux (B) du projet.
        while(not arret):
            i+=1
            d = calcul_echeance_moy(data, n, tau, valRevente)
            tau = calcul_tri_aux(I0, somme, d, valRevente)
            VAN = calcul_VAN(data, n, tau, valRevente)
            print(str(i) + ". VAN(" + str(tau) + ") = " + str(VAN)) # Affichage de la valeur actuelle nette à chaque itération (pour voir la convergence)
            if(((VAN <= epsilon) and (VAN >= -epsilon)) or i >= nb_itmax): # Vérification si la valeur actuelle nette (VAN) est proche de zéro (dans la plage définie par epsilon) ou si le nombre d'itérations dépasse la limite définie (nb_itmax)
                tri = tau
                arret = True
    return tri

## Application numérique

print("VAN(" + str(tau) + ") =", end=" ")
print(calcul_VAN(data, n, tau, valRevente), "\n")

print("d_moy(" + str(tau) + ") =", end=" ")
d0 = calcul_echeance_moy(data, n, tau, valRevente)
print(d0,"\n")

tri_aux = calcul_tri_aux(I0, sum(B), d0, valRevente)
print("tri_aux(" + str(d0) + ") = " + str(tri_aux) 
      + "\n")

tri = calcul_tri(data, n, tau, valRevente)
print("\n")
print("Le taux de rendement interne (TRI) du projet est de : " + str(tri), "\n")

print("VAN(" + str(tri) + ") = " + str(calcul_VAN(data, n, tri, valRevente)), "\n")