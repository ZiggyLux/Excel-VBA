Attribute VB_Name = "TypeCellule"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. CellValueType.excelMacro.bas
' Derni�re MAJ........ 30/07/13
' Auteur.............. Marc C�sarini
' Remarque............ VBA source file
' Br�ve description... Permet de d�terminer le type d'une cellule
'
' Emplacement.........
'-----------------------------------------------------------------------------
'
' Options
'
Option Explicit
'
' D�clarations des variables
'
' D�clarations des constantes
'
Public Function getCellValueType(v As Variant) As Integer
    getCellValueType = VarType(v) ' Voir http://msdn.microsoft.com/en-us/library/gg278470.aspx
End Function
Public Function getCellValueTypeAsText(v As Variant) As String
    ' Voir http://msdn.microsoft.com/en-us/library/gg278470.aspx
    Dim str As String
    Select Case VarType(v)
        Case vbEmpty
            str = "vbEmpty"
        Case vbNull
            str = "vbNull"
        Case vbInteger
            str = "vbInteger"
        Case vbLong
            str = "vbLong"
        Case vbSingle
            str = "vbSingle"
        Case vbDouble
            str = "vbDouble"
        Case vbCurrency
            str = "vbCurrency"
        Case vbDate
            str = "vbDate"
        Case vbString
            str = "vbString"
        Case vbObject
            str = "vbObject"
        Case vbError
            str = "vbError"
        Case vbBoolean
            str = "vbBoolean"
        Case vbVariant
            str = "vbVariant"
        Case vbDataObject
            str = "vbDecimal"
        Case vbByte
            str = "vbByte"
'				vbLongLong est disponible sur les plateformes 64 bits
'       Case vbLongLong
'           str = "vbLongLong"
        Case vbUserDefinedType
            str = "vbUserDefinedType"
        Case vbArray
            str = "vbArray"
        Case Else
            str = "?"
    End Select
    getCellValueTypeAsText = str & " (" & VarType(v) & ")"
End Function

