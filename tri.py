## import_librairies

import pandas as pd
import math as m

## init_var_glob

epsilon = 0.0001
nb_itmax = 30

## lecture_donnees

#pip install pandas si la librairie n'est pas encore installée

#load the xlsx file 
data = pd.read_excel("Data/Projet_eche5.xlsx") #le fichier excel doit contenir qu'une seule ligne contenant les données des flux

# Read the values of the file in the dataframe
data = pd.DataFrame(data)

#Print data
print("Le contenu des données : \n", data)

#The value of the sell is the last column of the second line
valRevente = data.iloc[2, (len(data.axes[1]) - 1)]
print(valRevente)

#The values of B with B[0] = I
B = data.iloc[2, 1:(len(data.axes[1]) - 1)]

#n, the number of years
n = len(data.axes[1]) - 3

#tau
tau = 0.01

## Function calcul_VAN

def calcul_VAN(data, n, tau, valRevente):
    sum = 0
    if(tau > 0):
        for i in range(1, n+1):
            sum += data[i]/pow(1+tau, i)
    return data[0] + sum + valRevente

## Function calcul_echeance_moy

def calcul_echeance_moy(data, n, tau, valRevente):
    sumFlux = 0
    sumFluxAct = 0
    for i in range(1, n+1):
        sumFlux += data[i]
        sumFluxAct += data[i]/pow(1+tau, i)
    return m.log(sumFlux/sumFluxAct) / m.log(1 + tau)

## Function calcul_tri_aux

def calcul_tri_aux(I0, B, d):
    return pow((B / -I0), 1 / d) - 1

## Function calcul_tri

def calcul_tri(data, n, tau0, valRevente):
    tri = 0
    if(tau0 > 0):
        tau = tau0
        i = 0
        arret = False
        somme = sum(data[1:n+1])
        while(not arret):
            i+=1
            d = calcul_echeance_moy(data, n, tau, valRevente)
            tau = calcul_tri_aux(data[0], somme, d)
            VAN = calcul_VAN(data, n, tau, valRevente)
            
            if(((VAN <= epsilon) and (VAN >= -epsilon)) or i >= nb_itmax):
                tri = tau
                arret = True
                print(i)
    return tri

print("VAN(" + str(tau) + ") =", end=" ")
print(calcul_VAN(B, n, tau, valRevente))

print("d_moy(" + str(tau) + ") =", end=" ")
print(calcul_echeance_moy(B, n, tau, valRevente))

#d, une date 
d = 4
print("tri_aux(" + str(d) + ") =" + str(calcul_tri_aux(B[0], sum(B[1:n+1]), d)))

tri = calcul_tri(B, n, tau, valRevente)
print("tri = " + str(tri))

#tri = 0.41196
print("VAN(" + str(tri) + ") = " + str(calcul_VAN(B, n, tri, valRevente)))