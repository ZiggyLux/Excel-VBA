Attribute VB_Name = "SheetUtil"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. SheetUtil.excelMacro.bas
' Dernière MAJ........ 11/09/18
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
' Brève description... Fonctions utiles pour les classeurs
'
' Emplacement.........
'-----------------------------------------------------------------------------
' Options
Option Explicit

Private Function normalizeSheetname(str As String) As String
    normalizeSheetname = LTrim(RTrim(UCase(str)))
End Function

' OBTENIR LE TYPENAME D'UNE FEUILLE
' Paramètres :
'   strFeuille ................ Nom de la feuille
'   wb ........................ classeur
' Retour :
'   "Worksheet" ............... Pour une feuille de calcul
'   "Chart" ................... Pour une feuille de type graphe
'   "DialogSheet" ............. Pour une feuille boîte de dialogue
'   "" ........................ Si la feuille n'a pas été trouvée
'
Public Function getSheetTypeByName( _
    strFeuille As String, _
    wb As Workbook) As String
    
    Dim strType As String
    Dim i As Integer
    Dim strFeuilleNorm As String
    
    ' Normalisation du nom de feuille passé en paramètre
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

' VERIFIER EXISTENCE FEUILLE AVEC OPTION DE RECREER
' Paramètres :
' strFeuille ................. Nom de la feuille
' fRecreer ................... Faux : si existe ne fait rien
'                              Vrai : si existe, recréer la feuille
' Retour:
' Vrai : la feuille existe
' Faux : la feuille n'existe pas
Public Function VerifierExistenceFeuille( _
    strFeuille As String, _
    fRecreer As Boolean) As Boolean
    
    Const L_TITRE_DIALOGUE As String = "Test Existence Feuille"
    Dim iReponseDialogue As Integer
    Dim wshType As String
    Dim wshNew As Worksheet
    Dim fSuccess As Boolean

    wshType = getSheetTypeByName(strFeuille, Application.ActiveWorkbook)
    
    If Len(wshType) > 0 Then
        If wshType = "Worksheet" Then
            If fRecreer = True Then
                iReponseDialogue = MsgBox( _
                    "La feuille " + strFeuille + " existe déjà." & vbCrLf & _
                        "Voulez-vous la recréer ?", _
                    vbYesNo, _
                    L_TITRE_DIALOGUE)
                If iReponseDialogue = vbYes Then 'Suppression et création
                    ' Suppression de la feuille
                    Sheets(strFeuille).Delete
                    If Len(getSheetTypeByName(strFeuille, Application.ActiveWorkbook)) = 0 Then
                        ' Nouvelle création de la feuille
                        Set wshNew = Worksheets.Add(Type:=xlWorksheet)
                        wshNew.Name = strFeuille
                        Set wshNew = Nothing
                    End If
                End If
            End If
            fSuccess = True
        Else
            MsgBox "Une feuille nommée " & strFeuille & " est déjà utilisée" & vbCrLf & _
                "mais pour une autre utilisation. Veuillez choisir un autre nom, svp !", _
                , L_TITRE_DIALOGUE
            fSuccess = False
        End If
    Else
        iReponseDialogue = MsgBox( _
            "La feuille " + strFeuille + " n'existe pas." & vbCrLf & _
                "Voulez-vous la créer ?", _
            vbYesNo, _
            L_TITRE_DIALOGUE)
        If iReponseDialogue = vbYes Then ' Création de la feuille
            Set wshNew = Worksheets.Add(Type:=xlWorksheet)
            wshNew.Name = strFeuille
            Set wshNew = Nothing
            fSuccess = True
        Else
            fSuccess = False
        End If
    End If
    VerifierExistenceFeuille = fSuccess
End Function


