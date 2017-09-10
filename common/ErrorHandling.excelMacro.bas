Attribute VB_Name = "ErrorHandling"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Windows
' Source.............. ErrorHandling.excelMacro.bas
' Dernière MAJ........ 08/09/17
' Auteur.............. Marc Césarini
' Remarque............ VBA source file
' Brève description... Gestion des erreurs en VBA
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
Public Function returnExcelError(str As String) As Variant
    ' Voir 'A Function that returns an error value' dans [A]
    ' Voir 'Table 2-2: Excel Formula Error Values' dans [A]
    Select Case str
        Case "#DIV/0!"
            ' La formule essaie de diviser par 0 (zéro), ce
            ' qui n'est pas permis. Cette erreur se produit aussi
            ' lorsque la formule essaie de diviser par une cellule
            ' vide.
            returnExcelError = CVErr(xlErrDiv0)
            
        Case "#N/A"
            ' La formule fait référence, directement ou indirectement,
            ' à une cellule qui utilise la ws-fonction NA pour
            ' signaler le fait que la donnée n'est pas disponible.
            ' Une fonction de look-up qui ne peut localiser une
            ' valeur retourne aussi #N/A
            returnExcelError = CVErr(xlErrNA)
            
        Case "#NAME?"
            ' La formule utilise un nom qu'Excel ne reconnaît pas.
            ' Ceci arrive si on supprime un nom qui est utilisé
            ' dans la formule, si des guillements sont mal appariés
            ' dans un texte, si des parenthèses sont mal appariées
            ' pour une fonction sans paramètre ou si on orthographie
            ' mal un nom de fonction ou de plage. Une formule affiche
            ' aussi cette erreur si elle utilise la fonction d'un
            ' module complémentaire qui n'a pas été installé.
            returnExcelError = CVErr(xlErrName)
            
        Case "#NULL!"
            ' La formule utilise une intersection de deux plages qui
            ' n'ont pas d'intersection.
            returnExcelError = CVErr(xlErrNull)
            
        Case "#NUM!"
            ' Une fonction a un problème; par exemple la fonction
            ' SQRT essaie de calculer la racine carré d'un nombre
            ' négatif. Cette erreur apparaît également si une valeur
            ' calculé est trop élevée ou trop faible.
            ' Excel ne gère pas les valeurs non-nulles inférieures à
            ' 1E-307 ou supérieurs à 1E+308 en valeur absolue.
            returnExcelError = CVErr(xlErrNum)
            
        Case "#REF!"
            ' La formule fait référence à une cellule qui n'est pas
            ' valide. Cela peut arriver si une cellule utilisée dans
            ' une formule a été supprimée de la feuille de calcul.
            returnExcelError = CVErr(xlErrRef)
            
        Case "#VALUE!"
            ' La formule inclut un paramètre ou une opérande du
            ' mauvais type. Une opérande est une valeur ou référence
            ' à une cellule qu'une formule utilise pour calculer un
            ' résultat. Cette erreur arrive aussi si une ws-formule
            ' VBA personnalisée contient une erreur.
            returnExcelError = CVErr(xlErrValue)
            
        Case Else
            returnExcelError = CVErr(xlErrValue)
    End Select
End Function
