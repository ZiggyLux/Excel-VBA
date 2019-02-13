Attribute VB_Name = "FileSysUtil"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. TableTailleFixe.excelMacro.bas
' Dernière MAJ........ 04/10/18
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
' Brève description... Fonctions utiles pour la gestion de fichier
'
' Emplacement.........
'-----------------------------------------------------------------------------
' Options
Option Explicit

' VERIFIER QU'UN FICHIER EXISTE
' Paramètres:
' strNomFichier ....... Chemin du fihcier à écrire
' Valeur retournée .... Vrai si le fihcier existe
Public Function EstFichierPresent(strNomFichier As String) As Boolean
    EstFichierPresent = Dir(strNomFichier) <> ""
End Function

