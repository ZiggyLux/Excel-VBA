Attribute VB_Name = "Common"
'-----------------------------------------------------------------------------
' Application......... VBA tool box
' Version............. 1.0
' Plateforme.......... Win 32
' Source.............. Common.excelMacro.bas
' Derni�re MAJ........
' Auteur.............. Marc C�sarini
' Remarque............ VBA source file
' Br�ve description...
'
' Emplacement.........
'-----------------------------------------------------------------------------
'
' Options
'
Option Explicit

' D�clarations des variables
'
' D�clarations des constantes
'

'-----------------------------------------------------------------------------
' R�active la mise � jour de l'�cran � l'ex�cution
Sub updateScreen()
'-----------------------------------------------------------------------------
    Application.ScreenUpdating = True
End Sub

'-----------------------------------------------------------------------------
' Suspend la mise � jour de l'�cran � l'ex�cution
Sub freezeScreen()
'-----------------------------------------------------------------------------
    Application.ScreenUpdating = False
End Sub

'-----------------------------------------------------------------------------
' Affiche la formule associ�e � la cellule
Function getFormula(r As Range) As String
'-----------------------------------------------------------------------------
    Application.Volatile
    If r.HasArray Then
        getFormula = "<-- {" & r.FormulaArray & "}"
    Else
        getFormula = "<-- " & r.FormulaArray
    End If
End Function

