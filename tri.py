## import_librairies

import pandas as pd
import math as m

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

def calcul_VAN(data, n, tau, valRevente):
    sum = 0
    if(tau > 0):
        for i in range(1, n+1):
            sum += data.iloc[0, i]/pow(1 + tau, i)
    return I0 + sum + valRevente

## Function calcul_echeance_moy

def calcul_echeance_moy(data, n, tau, valRevente):
    sumFlux = 0
    sumFluxAct = 0
    for i in range(1, n+1):
        sumFlux += data[i]
        sumFluxAct += data[i]/pow(1+tau, i)
    return m.log((sumFlux + valRevente) /sumFluxAct) / m.log(1 + tau)

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
        somme = sum(B)
        while(not arret):
            i+=1
            d = calcul_echeance_moy(data, n, tau, valRevente)
            tau = calcul_tri_aux(I0, somme, d)
            VAN = calcul_VAN(data, n, tau, valRevente)
            
            if(((VAN <= epsilon) and (VAN >= -epsilon)) or i >= nb_itmax):
                tri = tau
                arret = True
                #print(i)
    return tri

print("VAN(" + str(tau) + ") =", end=" ")
print(calcul_VAN(data, n, tau, valRevente), "\n")

print("d_moy(" + str(tau) + ") =", end=" ")
d0 = calcul_echeance_moy(data, n, tau, valRevente)
print(d0,"\n")

#d, une date 
d = 4
print("tri_aux(" + str(d0) + ") = " + str(calcul_tri_aux(I0, sum(B), d0)) 
      + "\n")
#print("tri_aux = ", calcul_tri_aux(I0, sum(B), d))

tri = calcul_tri(data, n, tau, valRevente)
print("tri = " + str(tri), "\n")

#tri = 0.41196
print("VAN(" + str(tri) + ") = " + str(calcul_VAN(data, n, tri, valRevente)), "\n")