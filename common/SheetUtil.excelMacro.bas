Attribute VB_Name = "SheetUtil"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. SheetUtil.excelMacro.bas
' Derni�re MAJ........ 11/09/18
' Auteur.............. Marc C�sarini
' Remarque............ VBA source file
' Br�ve description... Fonctions utiles pour les classeurs
'
' Emplacement.........
'-----------------------------------------------------------------------------
' Options
Option Explicit

Private Function normalizeSheetname(str As String) As String
    normalizeSheetname = LTrim(RTrim(UCase(str)))
End Function

' Retourne le TypeName d'une feuille � partir:
'   strFeuille ................ Nom de la feuille
'   wb ........................ classeur
'
'   "Worksheet" ............... Pour une feuille de calcul
'   "Chart" ................... Pour une feuille de type graphe
'   "DialogSheet" ............. Pour une feuille bo�te de dialogue
'   "" ........................ Si la feuille n'a pas �t� trouv�e
'
Public Function getSheetTypeByName( _
    strFeuille As String, _
    wb As Workbook) As String
    
    Dim strType As String
    Dim i As Integer
    Dim strFeuilleNorm As String
    
    ' Normalisation du nom de feuille pass� en param�tre
    strFeuilleNorm = normalizeSheetname(strFeuille)
    
    strType = ""
    For i = 1 To wb.Sheets.Count
        If normalizeSheetname(wb.Sheets(i).Name) = strFeuilleNorm Then
            strType = TypeName(wb.Sheets(i))
            Exit For
        End If
    Next
    getSheetTypeByName = strType
End Function
