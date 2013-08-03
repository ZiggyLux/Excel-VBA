Attribute VB_Name = "CurrencyTable"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. CurrencyTable.excelMacro.bas
' Dernière MAJ........ 01/08/13
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
' Brève description... Permet de faire une recherche pour un code de monnaie
'                      (numérique ou alphabétique) et de retourner une
'                       propriété
'
' Emplacement.........
'-----------------------------------------------------------------------------
'
' Table des monnaies - Structure
'   Code numérique :        3 chiffres (critère tri)
'   Code alphabétique :     3 lettres
'   Description :           Texte
'   Sous-unité :            1 chiffre
'   Livrabilité :           booléen
'   Commentaire :           Texte
' Index des monnaies - Structure
'   Code alphabétique :     3 lettres (critère tri)
'   Code numérique :        3 chiffres
'
' Options
'
Option Explicit
'
' Déclarations des variables
'
' Déclarations des constantes
'
Private Const RNG_REFTAB_CURRENCY As String = "TableCurr!A2:E22"
Private Const RNG_INDEX_CURRENCY_ALPHA As String = "IndexCurrAlpha!A2:B22"
'
' Colonnes de la tables des monnaies
'
Public Const TCOL_CUR_NUM As Integer = 1
Public Const TCOL_CUR_ALPHA As Integer = 2
Public Const TCOL_CUR_DESCR As Integer = 3
Public Const TCOL_CUR_SBUNT As Integer = 4
Public Const TCOL_CUR_DELIV As Integer = 5
Public Const TCOL_CUR_REM As Integer = 6
Private Const TCOL_CUR_IDXALP_NUM As Integer = 2
'
' Fonctions
'
Public Function getCurrPropFromNum(iCurr As Integer, Optional iCol As Integer = TCOL_CUR_ALPHA) As Variant
    Dim vLU As Variant
    Dim rngTable As Range
    
    Set rngTable = Application.Range(RNG_REFTAB_CURRENCY)
    
    vLU = Application.WorksheetFunction.VLookUp(iCurr, rngTable, iCol, False)
    If Application.WorksheetFunction.IsNA(vLU) Then
        getCurrPropFromNum = ""
    Else
        getCurrPropFromNum = vLU
    End If
    
    Set rngTable = Nothing
End Function
Public Function getCurrNumFromAlpha(strCurr As String) As Integer
    
    Dim vLU As Variant
    Dim rngTable As Range
    
    Set rngTable = Application.Range(RNG_INDEX_CURRENCY_ALPHA)
    
    vLU = Application.WorksheetFunction.VLookUp(strCurr, rngTable, TCOL_CUR_IDXALP_NUM, False)
    If Application.WorksheetFunction.IsNA(vLU) Then
        getCurrNumFromAlpha = 0
    Else
        getCurrNumFromAlpha = vLU
    End If
    
    Set rngTable = Nothing
End Function
Public Function getCurrPropFromAlpha(strCurr As String, Optional iCol As Integer = TCOL_CUR_NUM) As Variant
    Dim iCurr As Integer
    iCurr = getCurrNumFromAlpha(strCurr)
    If iCurr = 0 Then
        getCurrPropFromAlpha = ""
    Else
        getCurrPropFromAlpha = getCurrPropFromNum(iCurr, iCol)
    End If
End Function


