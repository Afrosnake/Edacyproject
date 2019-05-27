# -*- coding :Latin -1 -*

import os
from xlwt import Workbook

compteur_de_passage = 0

response = 0
male =0
female =0
compteur_moyenne =0
#creation fichier
fichierCSV = Workbook ()

# Creation de la feuille numero 1

feuille1 = fichierCSV.add_sheet ('feuille 1')

# ajout des en-etes

feuille1.write(0,0, 'NOM')
feuille1.write(0,1, 'Prenom')
feuille1.write(0,2, 'Adresse')
feuille1.write(0,3, 'Moyenne ')
feuille1.write(0,4, 'Age')
feuille1.write(0,5, 'Region')
feuille1.write(0,6, 'Specialite')
feuille1.write(0,7, 'Sexe')
# Taille des colonnes 1

feuille1.col(0).width = 5000
feuille1.col(1).width = 7000
feuille1.col(2).width = 12000
feuille1.col(3).width = 2000
feuille1.col(4).width = 2000
feuille1.col(5).width = 7000
feuille1.col(6).width = 7000
feuille1.col(7).width = 7000


# Creation de la feuille numero 2
feuille2 = fichierCSV.add_sheet ('feuille 2')

feuille2.write (0,0, "Liste des eleves ayant la moyenne ")
feuille2.write (0,1, "Moyenne")
feuille2.col(0).width = 12000
feuille2.col(1).width = 2000

# creation feuille numéro 3

feuille3 = fichierCSV.add_sheet ('feuille 3')

feuille3.write (0,0, "Etudiant ayant plus de 20 ans ")
feuille3.write (0,1, "AGE")
feuille3.col(0).width = 12000
feuille3.col(1).width = 2000
# création feuille numéro 4
feuille4 = fichierCSV.add_sheet ('feuille 4')

feuille4.write (0,0, "Moyenne de l'ecole ")
feuille4.write (0,1, "pourcentage de filles ")
feuille4.write (0,2, "Pourcentages de garcons ")
feuille4.write (0,3, "La région avec la plus forte moyenne ")

feuille4.col(0).width = 10000
feuille4.col(1).width = 8000
feuille4.col(2).width = 8000
feuille4.col(3).width = 8000

# Automatisation de la saisie

while response is not 'Q':
    compteur_de_passage = compteur_de_passage+1


    print(compteur_de_passage)

    response = input("Est ce le dernier Etudiant sur la liste ? 'Q' = OUI : ")
    nom = input("Entrez le Nom de l'etudiant: ")
    prenom = input("Entrez le Prenom de l'etudiant: ")
    adresse = input("Entrez l'adresse de l'etudiant :  ")
    moyenne = input ("Entrez la moyenne de l'etudiant : ")
    if type(moyenne) != int :
        print("Vous n'avez pas saisi de chiffre ")
        moyenne = input("Entrez la moyenne de l'etudiant: ")


    elif 0 > int(moyenne) >= 20:
        print ("Vous n'avez pas saisi de moyenne comprise entre 0 et 20")
    else :
        pass

    # elif int(moyenne) > 20 :
    #     print ("La moyenne saisie depasse 20 ")
    #     moyenne = input("Entrez la moyenne de l'etudiant: ")
    # elif int(moyenne) <= 0 :
    #     print("la moyenne saisie est inferieure ou egale à ZERO ")
    #     moyenne = input("Entrez la moyenne de l'etudiant: ")

    compteur_moyenne = compteur_moyenne + int(moyenne)
    age =  input("Entrez l'age de l'etudiant: ")
    region = input("Entrez la region de l'etudiant: ")
    specialite = input("Entrez la spécialité de l'etudiant: ")

    sexe = input("Entrez le sexe de l'etudiant; 'M' pour un homme et 'F' pour un femme  : ")
    sexeinput = str(sexe)
    if sexeinput == 'M':
        print ("L'Etudiant ", prenom,'  ',nom ,"est un HOMME" )
        male +=1
    elif sexeinput == 'F':
        print("L'Etudiant ", prenom, '  ', nom, "est une FEMME")
        female +=1

    else :
        sexe = input("Entrez le sexe de l'etudiant; 'M' pour un homme et 'F' pour une femme  : ")


    # remplissage des lignes feuille1

    ligne_compteur_de_passage = feuille1.row(compteur_de_passage)
    ligne_compteur_de_passage.write(0,nom)
    ligne_compteur_de_passage.write(1,prenom)
    ligne_compteur_de_passage.write(2,adresse)
    ligne_compteur_de_passage.write(3,moyenne)
    ligne_compteur_de_passage.write(4,age)
    ligne_compteur_de_passage.write(5,region)
    ligne_compteur_de_passage.write(6,specialite)
    ligne_compteur_de_passage.write(7,sexe)
    ##remplissage des lignes feuille2
    ligne_compteur_de_passage2 = feuille2.row(compteur_de_passage)

    if int(moyenne) >= 10:


        ligne_compteur_de_passage2.write(0, prenom +' ' +nom)
        ligne_compteur_de_passage2.write(1, moyenne)
    else :
        pass






    #remplissage des lignes feuille3

    ligne_compteur_de_passage3 = feuille3.row(compteur_de_passage)

    if int(age) >= 20 :

        ligne_compteur_de_passage3.write(0, prenom +' ' +nom)
        ligne_compteur_de_passage3.write(1, age)
    else :
        pass

    # remplissage feuille4
    ligne_compteur_de_passage4 = feuille4.row(compteur_de_passage)
    moyenne_ecole = compteur_moyenne / (compteur_de_passage )
    ligne_compteur_de_passage4.write (0, moyenne_ecole)


#pourcentage filles et garcons
pourcentage_fille = (female / compteur_de_passage ) * 100
ligne_compteur_de_passage4.write(1, pourcentage_fille)

pourcentage_garcon = (male / compteur_de_passage ) * 100
ligne_compteur_de_passage4.write(2, pourcentage_garcon)


# Sauvegarde du fichier excel

fichierCSV.save('Fichier_EdacyPythonData.xls')

os. system (" pause ")
