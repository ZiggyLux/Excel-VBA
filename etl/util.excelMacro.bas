Attribute VB_Name = "util"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Windows
' Source.............. etl.util.excelMacro.bas
' Derni�re MAJ........ 17/08/18
' Auteur.............. Marc C�sarini
' Remarque............ VBA source file
' Br�ve description... Fonctions utilitaires pour travaux ETL
'
' Emplacement.........
'-----------------------------------------------------------------------------
'
' Options
'
Option Explicit
'
' Constantes
Const K_VLD_SUB_LIB As String = "Fonction"
Const K_VLD_TYP_LIB As String = "Erreur de param�tre"
Const K_VLD_CHA_LIB_OFP As String = "Nom du fichier de sortie"
Const K_VLD_CHA_LIB_SEP As String = "S�parateur de valeurs"
Const K_VLD_CHA_LIB_IND As String = "Valeur d'indentation"
Const K_VLD_CHA_LIB_PAW As String = "Largeur de paragraphe"
Const K_VLD_LIB_VALEUR As String = "valeur"
Const K_VLD_LIB_TEST As String = "r�gle"
Const K_VLD_LIB_MANQUANT As String = "manquant"
Const K_VLD_LIB_ENTRE_INF As String = "Doit �tre compris entre "
Const K_VLD_LIB_ENTRE_SUP As String = " et "
'
' Fonction utile pour formater une plage de valeur en texte d�limit�
'   List ............. Liste/Plage de valeurs � incorporer dans la liste
'   Sep .............. Carac�re s�parant les valeurs
'   Quote ............ Caract�re entourant les valeurs
Public Function ImplodeRangeToString(List, sep As String, quote As String) _
    As String
    
    Dim elt As Variant
    Dim str As String
    str = ""
    For Each elt In List
        str = IIf(Len(str) = 0, "", str & sep) & quote & elt & quote
    Next elt
    ImplodeRangeToString = str
End Function
'
' Fonction utile pour formater une plage de valeur en texte d�limit�
'   dans un fichier.
'   Valeur retour..... True si la fonction s'est d�roul�e avec succ�s
'   List.............. Liste/Plage de valeurs � incorporer dans le fichier
'   strOutputFilePath  Chemin du fichier � g�n�rer
'   strSeparator...... S�parateur (ne peut �tre vide)
'   strQuote.......... Quote (vide: pas de s�parateur)
'   nIndent........... Indentation de chaque ligne (0: pas d'indentation)
'   nParagraphWidth... Largeur maxi de chaque ligne (0: pas de limite)
'
Public Function ImplodeRangeToFile( _
    List, _
    strOutputFilePath As String, _
    strSeparator As String, _
    strQuote As String, _
    nIndent As Integer, _
    nParagraphWidth As Integer _
) As Boolean
    Dim nCount As Integer
    Dim chSep As String
    Dim cell As Range
    Dim strBuf As String
    Dim strNew As String
    Dim strTest As String
    Dim strMessage As String
    
    Const K_VLD_SUB_LIB_VAL As String = "ImplodeRangeToFile"
        
    ' V�rification des param�tres
    ' . V�rification de la liste de valeurs
    '   List : Pas de v�rification
    '   strOutputFilePath : V�rification que non-vide
    If strOutputFilePath = "" Then
        strMessage = K_VLD_SUB_LIB & " : " & K_VLD_SUB_LIB_VAL & vbCrLf _
            & K_VLD_CHA_LIB_OFP & vbCrLf _
            & vbTab & K_VLD_LIB_VALEUR & " = " & strOutputFilePath & vbCrLf _
            & vbTab & K_VLD_LIB_TEST & " : "
        MsgBox strMessage & K_VLD_LIB_MANQUANT, , K_VLD_TYP_LIB
        GoTo Exit_Error
    End If
    '   strSeparator : V�rification que non-vide
    If strSeparator = "" Then
        strMessage = K_VLD_SUB_LIB & " : " & K_VLD_SUB_LIB_VAL & vbCrLf _
            & K_VLD_CHA_LIB_SEP & vbCrLf _
            & vbTab & K_VLD_LIB_VALEUR & " = " & strSeparator & vbCrLf _
            & vbTab & K_VLD_LIB_TEST & " : "
        MsgBox strMessage & K_VLD_LIB_MANQUANT, , K_VLD_TYP_LIB
        GoTo Exit_Error
    End If
    '   strQuote : Pas de v�rification
    '   nIndent : V�rification compris entre 0 et 100
    If nIndent < 0 Or nIndent > 100 Then
        strMessage = K_VLD_SUB_LIB & " : " & K_VLD_SUB_LIB_VAL & vbCrLf _
            & K_VLD_CHA_LIB_IND & vbCrLf _
            & vbTab & K_VLD_LIB_VALEUR & " = " & nIndent & vbCrLf _
            & vbTab & K_VLD_LIB_TEST & " : "
        MsgBox strMessage & K_VLD_LIB_ENTRE_INF & "0" & K_VLD_LIB_ENTRE_SUP & "100", , K_VLD_TYP_LIB
        GoTo Exit_Error
    End If
    '   nParagraphWidth : V�rification compris entre 0 et 256
    If nParagraphWidth < 0 Or nIndent > 256 Then
        strMessage = K_VLD_SUB_LIB & " : " & K_VLD_SUB_LIB_VAL & vbCrLf _
            & K_VLD_CHA_LIB_PAW & vbCrLf _
            & vbTab & K_VLD_LIB_VALEUR & " = " & nParagraphWidth & vbCrLf _
            & vbTab & K_VLD_LIB_TEST & " : "
        MsgBox strMessage & K_VLD_LIB_ENTRE_INF & "0" & K_VLD_LIB_ENTRE_SUP & "256", , K_VLD_TYP_LIB
        GoTo Exit_Error
    End If
    
    ' Traitement
    Open strOutputFilePath For Output As #1
    
    If nIndent = 0 Then strBuf = "" Else strBuf = Space(nIndent)
    nCount = 0
    For Each cell In List
        nCount = nCount + 1
        strNew = strQuote & cell.Value & strQuote _
            & strSeparator & Space(1)
        strTest = strBuf & strNew
        If nParagraphWidth > 0 And Len(strTest) > nParagraphWidth Then
            ' Vide la cha�ne du tampon
            Print #1, strBuf
            ' R�initialise la cha�ne du tampon
            If nIndent = 0 Then strBuf = "" Else strBuf = Space(nIndent)
        End If
        strBuf = strBuf & strNew
    Next cell
    If nCount > 0 Then
        ' Supprime le dernier s�parateur et espace
        strBuf = Left(strBuf, Len(strBuf) - 2)
        ' Vide la chaine tampon
        Print #1, strBuf
    End If
    Close #1
    
Exit_Success:
    ImplodeRangeToFile = True
    Exit Function
    
Exit_Error:
    ' Point de sortie en cas d'erreur
    ImplodeRangeToFile = False
End Function

