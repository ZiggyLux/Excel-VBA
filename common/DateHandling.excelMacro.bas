Attribute VB_Name = "DateHandling"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Windows
' Source.............. DateHandling.excelMacro.bas
' Dernière MAJ........ 12/09/17
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
' Brève description... Gestion de dates en VBA
'
' Emplacement.........
' Bibliographie.......
'   [A] .............. Excel 2013 Power Programming with VBA (John Walkenbach)
'-----------------------------------------------------------------------------
'
' Options
'
Option Explicit
'
' Déclarations des variables
'
' Déclarations des constantes
'
' Retourne un tableau de valeurs
Public Function getDaysOfWeek1char(Optional strLng As String)
    Dim myArray
    If (Not IsMissing(strLng) And strLng = "en") Then
        myArray = Array("s", "m", "t", "w", "t", "f", "s")
    Else
        myArray = Array("d", "l", "m", "m", "j", "v", "s")
    End If
    getDaysOfWeek1char = myArray
End Function
Public Function getDaysOfWeek3chars(Optional strLng As String)
    Dim myArray
    If (Not IsMissing(strLng) And strLng = "en") Then
        myArray = Array("sun", "mon", "tue", "wed", "thu", "fri", "sat")
    Else
        myArray = Array("dim", "lun", "mar", "mer", "jeu", "ven", "sam")
    End If
    getDaysOfWeek3chars = myArray
End Function
Public Function getDaysOfWeekAllchars(Optional strLng As String)
    Dim myArray
    If (Not IsMissing(strLng) And strLng = "en") Then
        myArray = Array("sunday", "monday", "tuesday", "wednesday", "thursday", "friday", _
            "saturday")
    Else
        myArray = Array("dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", _
            "samedi")
    End If
    getDaysOfWeekAllchars = myArray
End Function
Public Function getDayOfWeek1char(dt As Date, Optional strLng As String)
    Dim aDays: aDays = getDaysOfWeek1char(strLng)
    
    getDayOfWeek1char = aDays(Weekday(dt, vbSunday) - 1)
End Function
Public Function getDayOfWeek3chars(dt As Date, Optional strLng As String)
    Dim aDays: aDays = getDaysOfWeek3chars(strLng)
    
    getDayOfWeek3chars = aDays(Weekday(dt, vbSunday) - 1)
End Function
Public Function getDayOfWeekAllchars(dt As Date, Optional strLng As String)
    Dim aDays: aDays = getDaysOfWeekAllchars(strLng)
    
    getDayOfWeekAllchars = aDays(Weekday(dt, vbSunday) - 1)
End Function
