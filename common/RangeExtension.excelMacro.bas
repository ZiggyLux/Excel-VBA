Attribute VB_Name = "RangeExtension"
'-----------------------------------------------------------------------------
' Application......... Excel VBA
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. RangeExtension.excelMacro.bas
' Dernière MAJ........ 17/06/2012
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
' Brève description... Détection de plage à partir d'une cellule d'origine
'                      en testant la contenance des cellules adjacentes
' Emplacement.........
'-----------------------------------------------------------------------------
'
' Options
Option Explicit

' Déclarations des variables

' Déclarations des constantes

Private Function isEmptyRange(r As Range) As Boolean
    Dim fResult As Boolean
    Dim c As Range
    fResult = True
    For Each c In r
        If Not (c Is Nothing) And VarType(c.Cells(1, 1).Value) <> vbEmpty Then
            fResult = False
            Exit For
        End If
    Next c
    isEmptyRange = fResult
End Function
Public Function walkExtendingRange(r As Range, iMoveRow As Integer, iMoveCol As Integer) As Range
    Dim iExtRow, iExtCol As Integer
    Dim iStepRow, iStepCol As Integer
    Dim fCont As Boolean
    
    ' Initialize step value for detection
    If iMoveRow > 0 Then
        iStepRow = 1
    ElseIf iMoveRow = 0 Then
        iStepRow = 0
    Else
        iStepRow = -1
    End If
    If iMoveCol > 0 Then
        iStepCol = 1
    ElseIf iMoveCol = 0 Then
        iStepCol = 0
    Else
        iStepCol = -1
    End If
    
    iExtRow = 0
    iExtCol = 0
    fCont = True
    
    While (fCont)
        fCont = False
        If iStepRow <> 0 Then
            If iStepCol <> 0 Then
                If Not isEmptyRange( _
                    Range( _
                        r.Cells(1 + iExtRow + iStepRow, 1), _
                        r.Cells(1 + iExtRow + iStepRow, iExtCol + iStepCol))) Then
                        
                    iExtRow = iExtRow + iStepRow
                    fCont = True
                End If
            End If
            If iStepCol = 0 Then
                If Not isEmptyRange( _
                    Range( _
                        r.Cells(1 + iExtRow + iStepRow, 1), _
                        r.Cells(1 + iExtRow + iStepRow, 1))) Then
                        
                    iExtRow = iExtRow + iStepRow
                    fCont = True
                End If
            End If
        End If
        If iStepCol <> 0 Then
            If iStepRow <> 0 Then
                If Not isEmptyRange( _
                    Range( _
                        r.Cells(1, 1 + iExtCol + iStepCol), _
                        r.Cells(1 + iExtRow + iStepRow, 1 + iExtCol + iStepCol))) Then
                        
                    iExtCol = iExtCol + iStepCol
                    fCont = True
                End If
            End If
            If iStepRow = 0 Then
                If Not isEmptyRange( _
                    Range( _
                        r.Cells(1, 1 + iExtCol + iStepCol), _
                        r.Cells(1, 1 + iExtCol + iStepCol))) Then
                        
                    iExtCol = iExtCol + iStepCol
                    fCont = True
                End If
            End If
        End If
    Wend
    Set walkExtendingRange = Range(r.Cells(1, 1), r.Cells(1 + iExtRow, 1 + iExtCol))
End Function
Public Function getRowCntFromCell(rngStartCell As Range, Optional iNS As Integer = 1) As Integer
    Dim r As Range
    Set r = walkExtendingRange(rngStartCell, 1, 0)
    getRowCntFromCell = r.Rows.Count
End Function
Public Function getColCntFromCell(rngStartCell As Range, Optional iWE As Integer = 1) As Integer
    Dim r As Range
    Set r = walkExtendingRange(rngStartCell, 0, 1)
    getColCntFromCell = r.Columns.Count
End Function
Public Sub DetectRangeActiveCell_SE()
    Dim r As Range
    Set r = walkExtendingRange(ActiveCell, 1, 1)
    r.Select
    Set r = Nothing
End Sub

