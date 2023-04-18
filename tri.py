## import_librairies

import pandas as pd
import math as m

## init_var_glob

epsilon = 0.0001
nb_itmax = 30

## lecture_donnees

#pip install pandas si la librairie n'est pas encore installée

#load the xlsx file 
data = pd.read_excel("maths-fi/Data/Projet_eche5.xlsx",skiprows=2) #le fichier excel doit contenir qu'une seule ligne contenant les données des flux

# Read the values of the file in the dataframe
data = pd.DataFrame(data, columns=[
                    0, 1, 2, 3, 4, 5, 'V'])

#Print data
print("Le contenu des données : \n", data)

## Function calcul_VAN

def VAN(data, n, tau, valRevente):
    sum = 0
    if(tau > 0):
        for i in range(1, n+1):
            sum += data[i]/pow(1+tau, i)
    return data[0] + sum + valRevente

## Function calcul_echeance_moy

def echenceMoy(data, n, tau, valRevente):
    sumFlux = 0
    sumFluxAct = 0
    for i in range(1, n+1):
        sumFlux += data[i]
        sumFluxAct += data[i]/pow(1+tau, i)
    return m.log(sumFlux/sumFluxAct) / m.log(1 + tau)

## Function calcul_tri_aux

def triAux(I0, B, d):
    return pow((B / I0), 1 / d) - 1
