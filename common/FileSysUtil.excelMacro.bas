Attribute VB_Name = "FileSysUtil"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. TableTailleFixe.excelMacro.bas
' Derni�re MAJ........ 04/10/18
' Auteur.............. Marc C�sarini
' Remarque............ VBA source file
' Br�ve description... Fonctions utiles pour la gestion de fichier
'
' Emplacement.........
'-----------------------------------------------------------------------------
' Options
Option Explicit

' VERIFIER QU'UN FICHIER EXISTE
' Param�tres:
' strNomFichier ....... Chemin du fihcier � �crire
' Valeur retourn�e .... Vrai si le fihcier existe
Public Function EstFichierPresent(strNomFichier As String) As Boolean
    EstFichierPresent = Dir(strNomFichier) <> ""
End Function

