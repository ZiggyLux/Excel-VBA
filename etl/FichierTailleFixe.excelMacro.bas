Attribute VB_Name = "FichierTailleFixe"
' Attribute VB_Name = "TableTailleFixe"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. TableTailleFixe.excelMacro.bas
' Dernière MAJ........ 13/02/19
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
'                      Utilise la classe logger de Tim Hall
'                      Tim Hall - https://github.com/VBA-tools/VBA-Log
' Brève description... Permet de charger un fichier à taille fixe
'
' Emplacement.........
'-----------------------------------------------------------------------------
' Options
Option Explicit

' Declarations de types
Public Type MyField
    strFieldname As String
    strInternalFieldname As String
    iOffset As Integer
    iLength As Integer
    iUpTo As Integer
    fQueryDefault As Boolean
    strQueryFieldname As String
    strMyDDSFieldname As String
    fBuildRecord As Boolean
    strExternalFieldname As String
    strRulesAndRemarks As String
    strFormat As String
    fPrimaryKey As Boolean
    strRegExp As String
End Type
Const K_MYFIELD_NBCOLS As Integer = 14

' Charge un tableau de structures MyField à partir:
'   rngFieldList .............. Plage de champs
'   varrFLColNames ............ Arrangement des noms de colonnes de la liste de champs
'                               implémenté comme un Variant de type tableau
'   de ........................ Tableau de structures à charger
'                               Redimensionné avec perte des données existantes
Public Sub ConstruitDescEnreg( _
    rngFieldList As Range, _
    varrFLColNames As Variant, _
    de() As MyField)
    
    Dim i As Integer, j As Integer ' Indices de boucles
    Dim fld As MyField ' Structure provisoire des données de définition d'un champ
    
    Dim nFields As Integer  ' Nombre de champs définis pour le fichier
    Dim nCols As Integer ' Nombre de colonnes dans la table de définition des champs
    
    Dim arrKnownFN(1 To K_MYFIELD_NBCOLS) As String
    Dim arrIndexKnownFN(1 To K_MYFIELD_NBCOLS) As Integer
            
    nFields = rngFieldList.Rows.Count ' Récupération du nombre de champs suivant hauteur plage
    nCols = rngFieldList.Columns.Count ' Récupération du nombre de colonnes suivant largeur plage

    For i = 1 To K_MYFIELD_NBCOLS
        arrIndexKnownFN(i) = -1
    Next
    For i = 1 To nCols
        Select Case varrFLColNames(1, i)
            Case "SQLVALIDFIELDNAME"
                arrIndexKnownFN(1) = i
            Case "INTERNALFIELDNAME"
                arrIndexKnownFN(2) = i
            Case "OFFSET"
                arrIndexKnownFN(3) = i
            Case "LENGTH"
                arrIndexKnownFN(4) = i
            Case "UPTO"
                arrIndexKnownFN(5) = i
            Case "QUERYDEFAULT"
                arrIndexKnownFN(6) = i
            Case "QUERYFIELDNAME"
                arrIndexKnownFN(7) = i
            Case "MYDDSFIELDNAME"
                arrIndexKnownFN(8) = i
            Case "BUILDRECORD"
                arrIndexKnownFN(9) = i
            Case "EXTERNALFIELDNAME"
                arrIndexKnownFN(10) = i
            Case "RULESANDREMARKS"
                arrIndexKnownFN(11) = i
            Case "FORMAT"
                arrIndexKnownFN(12) = i
            Case "PRIMARYKEY"
                arrIndexKnownFN(13) = i
            Case "REGEXP"
                arrIndexKnownFN(14) = i
        End Select
    Next
    
    ReDim de(1 To nFields)
    
    For i = 1 To nFields
        For j = 1 To K_MYFIELD_NBCOLS
            If arrIndexKnownFN(j) > 0 Then
                Select Case j
                    Case 1
                        fld.strFieldname = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 2
                        fld.strInternalFieldname = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 3
                        fld.iOffset = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 4
                        fld.iLength = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 5
                        fld.iUpTo = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 6
                        fld.fQueryDefault = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 7
                        fld.strQueryFieldname = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 8
                        fld.strMyDDSFieldname = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 9
                        fld.fBuildRecord = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 10
                        fld.strExternalFieldname = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 11
                        fld.strRulesAndRemarks = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 12
                        fld.strFormat = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 13
                        fld.fPrimaryKey = rngFieldList.Cells(i, arrIndexKnownFN(j))
                    Case 14
                        fld.strRegExp = rngFieldList.Cells(i, arrIndexKnownFN(j))
                End Select
            End If
        Next
        de(i) = fld
    Next i
End Sub

' ECRITURE D'UNE LIGNE D'ENTETE EN LIGNE 1 D'UNE FEUILLE DE CALCUL
' Paramètres:
' fca ................. Feuille de calcul qui va recevoir les données lues
' rngFieldList ........ Plage où sont décrit les champs du fichier
' varrFLColNames ...... Tableau où sont indiqués les noms d'attributs de champ
Public Sub PlacerEnteteNomsChamps( _
    fca As Worksheet, _
    rngFieldList As Range, _
    varrFLColNames As Variant)
    
    Dim i, iLoop As Integer
    Dim str As String
    Dim nFields As Integer
    Dim myDescEnreg() As MyField
    Dim fLastUpdScrVal As Boolean
            
    nFields = rngFieldList.Rows.Count
    
    ' Vidage des cellules de la feuille
    fca.Cells.Clear
    
    ' Chargement de la description de la structure
    ConstruitDescEnreg rngFieldList, varrFLColNames, myDescEnreg
    
    ' Screen Update False est une bonne pratique
    fLastUpdScrVal = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    iLoop = 0
    For i = 1 To nFields
        str = myDescEnreg(i).strMyDDSFieldname
        If str <> "" Then
            iLoop = iLoop + 1
            With fca.Cells(1, iLoop)
                .Value = str
                .Font.Bold = True
            End With
        End If
    Next i
    Application.ScreenUpdating = fLastUpdScrVal
End Sub

' LECTURE D'UN FICHIER DE TAILLE FIXE DANS UNE FEUILLE DE CALCUL
' Paramètres:
' fca ................. Feuille de calcul qui va recevoir les données lues
' strPathname ......... Nom du fichier de structure connue et depuis lequel
'                       seront lues les données
' rngFieldList ........ Plage où sont décrit les champs du fichier
' varrFLColNames ...... Plage où sont indiqués les noms d'attributs de champ
Public Sub LireFichierTailleFixe( _
    fca As Worksheet, _
    strPathname As String, _
    rngFieldList As Range, _
    varrFLColNames As Variant)
    
    Dim nLine As Integer
    Dim strLigne As String
    Dim strDDSFieldname As String
    Dim nFields As Integer
    Dim myDescEnreg() As MyField
    Dim fLastUpdScrVal As Boolean
    
    Dim i, iLoop, iFile As Integer
                
    ' Placement des en-tête
    PlacerEnteteNomsChamps fca, rngFieldList, varrFLColNames
    
    ' Chargement de la description de la structure
    ConstruitDescEnreg rngFieldList, varrFLColNames, myDescEnreg

    iFile = FreeFile
    Open strPathname For Input As #iFile
    
    ' Screen Update False est une bonne pratique
    fLastUpdScrVal = Application.ScreenUpdating
    Application.ScreenUpdating = False
        
    nFields = rngFieldList.Rows.Count
    nLine = 0
    While Not EOF(1)
        nLine = nLine + 1
        Line Input #iFile, strLigne
        iLoop = 0
        For i = 1 To nFields
            strDDSFieldname = myDescEnreg(i).strMyDDSFieldname
           If strDDSFieldname <> "" Then ' Champ atomique
                iLoop = iLoop + 1
                With fca.Cells(nLine + 1, iLoop)
                    .Value = Mid(strLigne, myDescEnreg(i).iOffset, myDescEnreg(i).iLength)
                End With
            End If
        Next i
    Wend
    
    Application.ScreenUpdating = fLastUpdScrVal
    
    Close #iFile
End Sub

Private Sub MyLogCallback(Level As Long, Message As String, From As String)
    Dim LevelValue As String
    Select Case Level
    Case 1
        LevelValue = "Trace"
    Case 2
        LevelValue = "Debug"
    Case 3
        LevelValue = "Info"
    Case 4
        LevelValue = "WARN"
    Case 5
        LevelValue = "ERROR"
    End Select
    Debug.Print Date & ";" & Time & ";" & LevelValue & ";" & _
        IIf(From <> "", From & ";", "") & Message
End Sub

Private Function IsKeyGood(strKey As String, strRegExp As String) As Boolean
    Dim fGoodKey As Boolean
    Dim ref() As Variant

    fGoodKey = True
    If LTrim(strKey) = "" Then
        fGoodKey = False
    Else
        If TypeName(RegExpFind(strKey, strRegExp)) = "String" Then fGoodKey = False
    End If
    IsKeyGood = fGoodKey
End Function

' ECRITURE D'UN FICHIER DE TAILLE FIXE DEPUIS UNE FEUILLE DE CALCUL
' Paramètres:
' fca ................. Feuille de calcul qui contient les données à écrire
' strPathname ......... Nom du fichier de structure connue et vers lequel
'                       seront écrites les données
' rngFieldList ........ Plage où sont décrit les champs du fichier
' varrFLColNames ...... Plage où sont indiqués les noms d'attributs de champ
Public Sub EcrireFichierTailleFixe( _
    fca As Worksheet, _
    strPathname As String, _
    rngFieldList As Range, _
    varrFLColNames As Variant)

    Const K_VLD_SUB_LIB As String = "Procédure"
    Const K_VLD_SUB_LIB_VAL As String = "EcrireFichierTailleFixe"
    Const K_VLD_LIB_VAL As String = "Valeur"
    Const K_VLD_LIB_VALPRE As String = "Valeur précédente"
    Const K_VLD_LIB_TEST As String = "Règle"
    Const K_VLD_LIB_TYP_INPFIL As String = "Donnée en entrée erronée"
    Const K_VLD_LIB_TEST_REGEXP As String = "Reg Exp "
    Const K_VLD_LIB_TEST_CLEPRINONUNI As String = "Clé primaire doit être unique"
    Const K_VLD_LIB_TEST_CLEPRINONTRI As String = "Clé primaire devrait être triée"

    Dim nFields As Integer
    Dim myDescEnreg() As MyField
    
    Dim iFile As Integer, nRecLen As Integer
    Dim nLine As Integer, nCol As Integer, i As Integer, iKey As Integer
    Dim iDiffLen As Integer
    Dim strLastKey As String, strNewKey As String, fWriteRecord As Boolean
    Dim strLigne As String, strField As String, strLeft As String, strRight As String
    
    Dim nStatLignesLues As Integer
    Dim nStatLignesEcrites As Integer
    
    Dim mylog As logger
    Dim fLastUpdScrVal As Boolean
    
    ' Chargement de la description de la structure
    ConstruitDescEnreg rngFieldList, varrFLColNames, myDescEnreg

    ' Calcul de la taille d'un enregistrement
    nFields = rngFieldList.Rows.Count
    nRecLen = 0
    For i = 1 To nFields
        If myDescEnreg(i).fBuildRecord Then
            nRecLen = nRecLen + myDescEnreg(i).iLength
        End If
    Next i

    iFile = FreeFile
    Open strPathname For Output As iFile Len = nRecLen

    ' Screen Update False est une bonne pratique
    fLastUpdScrVal = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Initialisation du logger
    Set mylog = New logger
    mylog.LogEnabled = True
    mylog.LogCallback = "MyLogCallback"
    mylog.LogTrace "Début de la trace", K_VLD_SUB_LIB & " " & K_VLD_SUB_LIB_VAL
    
    strLastKey = ""
    nStatLignesLues = 0
    nStatLignesEcrites = 0
    For nLine = 2 To fca.UsedRange.Rows.Count
        nStatLignesLues = nStatLignesLues + 1
        strLigne = Space(nRecLen)
        nCol = 0
        iKey = 0
        strNewKey = ""
        For i = 1 To nFields
            If myDescEnreg(i).fBuildRecord Then
                nCol = nCol + 1
                If Len(myDescEnreg(i).strFormat) > 0 Then
                    ' A FAIRE : Sécuriser l'appel à Format (au besoin On Error ...)
                    strField = Format(fca.Cells(nLine, nCol), myDescEnreg(i).strFormat)
                Else
                    ' A FAIRE : Sécuriser suivant le type de valeur contenue dans cellule
                    strField = fca.Cells(nLine, nCol)
                End If
                If myDescEnreg(i).fPrimaryKey = True Then
                    iKey = i
                    strNewKey = strField
                End If
                
                If myDescEnreg(i).iOffset > 1 Then
                    strLeft = Left(strLigne, myDescEnreg(i).iOffset - 1)
                Else
                    strLeft = ""
                End If
                If myDescEnreg(i).iUpTo < nRecLen Then
                    strRight = Right(strLigne, nRecLen - myDescEnreg(i).iUpTo)
                Else
                    strRight = ""
                End If
                iDiffLen = myDescEnreg(i).iLength - Len(strField)
                If iDiffLen > 0 Then
                    strField = strField & Space(iDiffLen)
                End If
                strLigne = strLeft & strField & strRight
           End If
        Next i
        
        
        If strNewKey <> strLastKey Then
            If iKey > 0 _
            And Len(myDescEnreg(iKey).strRegExp) > 0 _
            And Len(myDescEnreg(iKey).strFormat) > 0 Then
                With myDescEnreg(iKey)
                    ' On peut vérifier la bonne tête de la clé
                    fWriteRecord = IsKeyGood(strNewKey, .strRegExp)
                    If Not fWriteRecord Then
                        mylog.LogError _
                            K_VLD_LIB_TYP_INPFIL & ":" & .strExternalFieldname & ";" _
                                & K_VLD_LIB_VAL & ":" & strNewKey & ";" _
                                & K_VLD_LIB_TEST & ":" & K_VLD_LIB_TEST_REGEXP _
                                    & .strRegExp, _
                            K_VLD_SUB_LIB & ":" & K_VLD_SUB_LIB_VAL
                    End If
                    If strNewKey < strLastKey Then
                        ' On avertit que le tri naturel n'est pas respecté
                        mylog.LogWarn _
                            K_VLD_LIB_TYP_INPFIL & ":" & .strExternalFieldname & ";" _
                                & K_VLD_LIB_VAL & ":" & strNewKey & ";" _
                                & K_VLD_LIB_TEST & ":" & K_VLD_LIB_TEST_CLEPRINONTRI _
                                    & ";" _
                                & K_VLD_LIB_VALPRE & ":" & strLastKey, _
                            K_VLD_SUB_LIB & ":" & K_VLD_SUB_LIB_VAL
                        
                    End If
                End With
            Else
                fWriteRecord = True
            End If
            If fWriteRecord Then
                Print #iFile, strLigne
                strLastKey = strNewKey
                
                nStatLignesEcrites = nStatLignesEcrites + 1
            End If
        Else ' Cas d'une clé dupliquée
            If iKey > 0 And myDescEnreg(iKey).fPrimaryKey = True Then
                mylog.LogError _
                    K_VLD_LIB_TYP_INPFIL & ":" _
                        & myDescEnreg(iKey).strExternalFieldname & ";" _
                        & K_VLD_LIB_VAL & ":" & strNewKey & ";" _
                        & K_VLD_LIB_TEST & ":" & K_VLD_LIB_TEST_CLEPRINONUNI, _
                    K_VLD_SUB_LIB & ":" & K_VLD_SUB_LIB_VAL
            End If
        End If
    Next nLine

    Application.ScreenUpdating = fLastUpdScrVal
    
    Close #iFile
    mylog.LogInfo "Nbr lignes lues:" & nStatLignesLues & ";" _
            & "Nbr lignes écrites:" & nStatLignesEcrites _
        , K_VLD_SUB_LIB & " " & K_VLD_SUB_LIB_VAL
    mylog.LogTrace "Fin de la trace", K_VLD_SUB_LIB & " " & K_VLD_SUB_LIB_VAL
    Set mylog = Nothing
    
End Sub


