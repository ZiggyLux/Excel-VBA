Attribute VB_Name = "Common"
'-----------------------------------------------------------------------------
' Application......... VBA tool box
' Version............. 1.0
' Plateforme.......... Win 32
' Source.............. Common.excelMacro.bas
' Dernière MAJ........
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
' Brève description...
'
' Emplacement.........
'-----------------------------------------------------------------------------
'
' Options
'
Option Explicit

' Déclarations des variables
'
' Déclarations des constantes
'

'-----------------------------------------------------------------------------
' Réactive la mise à jour de l'écran à l'exécution
Sub updateScreen()
'-----------------------------------------------------------------------------
    Application.ScreenUpdating = True
End Sub

'-----------------------------------------------------------------------------
' Suspend la mise à jour de l'écran à l'exécution
Sub freezeScreen()
'-----------------------------------------------------------------------------
    Application.ScreenUpdating = False
End Sub

'-----------------------------------------------------------------------------
' Affiche la formule associée à la cellule
Function getFormula(r As Range) As String
'-----------------------------------------------------------------------------
    Application.Volatile
    If r.HasArray Then
        getFormula = "<-- {" & r.FormulaArray & "}"
    Else
        getFormula = "<-- " & r.FormulaArray
    End If
End Function

