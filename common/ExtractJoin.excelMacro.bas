Attribute VB_Name = "ExtractJoin"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Windows
' Source.............. ExtractJoin.excelMacro.bas
' Dernière MAJ........ 20/09/17
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
' Brève description... Extraire des données d'une feuille avec jointure
'
' Emplacement.........
' Dépendance.......... Requiert RangeExtension
'-----------------------------------------------------------------------------
'
' Options
'
Option Explicit
'
' Déclarations des variables
'
Private wshTarTab As Worksheet  ' La feuille de calcul cible
'
' Déclarations des constantes
'
Private Const TARGET_TABNAME = "target" ' Nom de la feuille de calcul cible
Public Const SOURCE_WSHNAM = "source"   ' Nom de la feuille de calcul source
Public Const SOURCE_FIRST_ROW = 3       ' Numéro de la première ligne de données
Public Const SOURCE_COLTAG_ORIGIN = "A" ' Colonne de la première colonne de données

Public Const SOURCE_CELL_ORIGIN = SOURCE_COLTAG_ORIGIN & SOURCE_FIRST_ROW

Public Const RNG_REFTAB_1 As String _
    = "'refpos'!A2:B5"                  ' Référence vers la table de jointure

'   Numéro des colonnes à extraires de la table source
Private Const _
    EMPNUM_Col = 1, _
    EMPPOS_Col = 6, _
    EMPNOM_Col = 3, _
    EMPSEX_Col = 2

'   Valeur à afficher si la jointure est infructueuse
Private Const UNKNOWN_VALUE = "???"
' Construction de l'en-tête de la table cible
Private Sub BuildHeaders()
    With wshTarTab
        .Cells(1, 1).Value = "Numéro"
        .Columns(1).ColumnWidth = 15
        
        .Cells(1, 2).Value = "Pos"
        .Columns(2).ColumnWidth = 9
        
        .Cells(1, 3).Value = "Nom"
        .Columns(3).ColumnWidth = 9
        
        .Cells(1, 4).Value = "Sexe"
        .Columns(4).ColumnWidth = 9
        
        .Cells(1, 5).Value = "Position"
        .Columns(5).ColumnWidth = 50
        
        With .Range("A1:E1").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 3
        End With
    End With
End Sub
' Jointure
Private Function getCol2Descr(strCode As String) As Variant
    Dim vLookup As Variant
    Dim rngTable As Range
    Dim wsFunc As WorksheetFunction: Set wsFunc = Application.WorksheetFunction
        
    Set rngTable = Application.Range(RNG_REFTAB_1)
    
    strCode = RTrim(strCode)
    
    On Error Resume Next
    vLookup = wsFunc.vLookup(strCode, rngTable, 2, False)
    If Err.Number <> 0 Then
        On Error GoTo 0
        getCol2Descr = CVErr(xlErrNA)
    Else
        On Error GoTo 0
        getCol2Descr = vLookup
    End If
    Set rngTable = Nothing
End Function
' Détermination de l'étendue des données source
'   Requiert RangeExtension
Private Function RowsFromHere(rngStartCell As Range) As Integer
    Dim r As Range
    Set r = ExtendingRange.ExtendingRange(rngStartCell, 1, 0)
    RowsFromHere = r.Rows.Count
End Function
' Extraction des lignes, jointures et alimentation de la table cible
Private Sub GenerateTargetLines()
    Dim iLine, nRows As Integer
    Dim r, tabSrc As Range
    Dim varCol2Descr As Variant
    Dim strCol2Descr As String
    nRows = RowsFromHere(Worksheets(SOURCE_WSHNAM).Range(SOURCE_CELL_ORIGIN))
    Set tabSrc = Worksheets(SOURCE_WSHNAM).Range( _
        SOURCE_CELL_ORIGIN & ":" & _
        SOURCE_COLTAG_ORIGIN & CStr((SOURCE_FIRST_ROW - 1) + nRows))
    
    iLine = 0
    For Each r In tabSrc
        If Not (r Is Nothing) And _
            VarType(r.Cells(1, 1)) <> vbEmpty Then
            
                iLine = iLine + 1
                
                wshTarTab.Cells(iLine + 1, 1).Value = _
                    r.Cells(1, EMPNUM_Col).Value
                    
                wshTarTab.Cells(iLine + 1, 2).Value = _
                    r.Cells(1, EMPPOS_Col).Value
                    
                wshTarTab.Cells(iLine + 1, 3).Value = _
                    r.Cells(1, EMPNOM_Col).Value
                    
                wshTarTab.Cells(iLine + 1, 4).Value = _
                    r.Cells(1, EMPSEX_Col).Value
                    
                varCol2Descr = getCol2Descr(r.Cells(1, EMPPOS_Col).Value)
                If WorksheetFunction.IsError(varCol2Descr) Then
                    ' (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).
                    strCol2Descr = UNKNOWN_VALUE
                Else
                    strCol2Descr = varCol2Descr
                End If
                wshTarTab.Cells(iLine + 1, 5).Value = strCol2Descr
        End If
    Next r
    Set tabSrc = Nothing
End Sub
' Procédure principale d'extraction avec jointure
Public Sub ExtractJoin()
    On Error Resume Next
    Worksheets(TARGET_TABNAME).Delete
    On Error GoTo 0
    
    Set wshTarTab = Worksheets.Add(Type:=xlWorksheet)
    With wshTarTab
        .Name = TARGET_TABNAME
    End With
    
    Application.ScreenUpdating = False
    
    BuildHeaders
        
    GenerateTargetLines
      
    Application.ScreenUpdating = True
    
    Set wshTarTab = Nothing
End Sub





