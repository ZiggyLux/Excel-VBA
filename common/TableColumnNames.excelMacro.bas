'-----------------------------------------------------------------------------
' Application......... VBA tool box                                                     
' Version............. 1.0                                                       
' Plateforme.......... Win 32                                                 
' Source.............. TableColumnNames.excelMacro.bas
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
' Mise en forme d'un en-tête de table (plage Selection)
Sub formatTableColumnNames()
'-----------------------------------------------------------------------------
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub
