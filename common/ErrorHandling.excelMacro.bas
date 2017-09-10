Attribute VB_Name = "ErrorHandling"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Windows
' Source.............. ErrorHandling.excelMacro.bas
' Derni�re MAJ........ 08/09/17
' Auteur.............. Marc C�sarini
' Remarque............ VBA source file
' Br�ve description... Gestion des erreurs en VBA
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
' D�clarations des variables
'
' D�clarations des constantes
'
Public Function returnExcelError(str As String) As Variant
    ' Voir 'A Function that returns an error value' dans [A]
    ' Voir 'Table 2-2: Excel Formula Error Values' dans [A]
    Select Case str
        Case "#DIV/0!"
            ' La formule essaie de diviser par 0 (z�ro), ce
            ' qui n'est pas permis. Cette erreur se produit aussi
            ' lorsque la formule essaie de diviser par une cellule
            ' vide.
            returnExcelError = CVErr(xlErrDiv0)
            
        Case "#N/A"
            ' La formule fait r�f�rence, directement ou indirectement,
            ' � une cellule qui utilise la ws-fonction NA pour
            ' signaler le fait que la donn�e n'est pas disponible.
            ' Une fonction de look-up qui ne peut localiser une
            ' valeur retourne aussi #N/A
            returnExcelError = CVErr(xlErrNA)
            
        Case "#NAME?"
            ' La formule utilise un nom qu'Excel ne reconna�t pas.
            ' Ceci arrive si on supprime un nom qui est utilis�
            ' dans la formule, si des guillements sont mal appari�s
            ' dans un texte, si des parenth�ses sont mal appari�es
            ' pour une fonction sans param�tre ou si on orthographie
            ' mal un nom de fonction ou de plage. Une formule affiche
            ' aussi cette erreur si elle utilise la fonction d'un
            ' module compl�mentaire qui n'a pas �t� install�.
            returnExcelError = CVErr(xlErrName)
            
        Case "#NULL!"
            ' La formule utilise une intersection de deux plages qui
            ' n'ont pas d'intersection.
            returnExcelError = CVErr(xlErrNull)
            
        Case "#NUM!"
            ' Une fonction a un probl�me; par exemple la fonction
            ' SQRT essaie de calculer la racine carr� d'un nombre
            ' n�gatif. Cette erreur appara�t �galement si une valeur
            ' calcul� est trop �lev�e ou trop faible.
            ' Excel ne g�re pas les valeurs non-nulles inf�rieures �
            ' 1E-307 ou sup�rieurs � 1E+308 en valeur absolue.
            returnExcelError = CVErr(xlErrNum)
            
        Case "#REF!"
            ' La formule fait r�f�rence � une cellule qui n'est pas
            ' valide. Cela peut arriver si une cellule utilis�e dans
            ' une formule a �t� supprim�e de la feuille de calcul.
            returnExcelError = CVErr(xlErrRef)
            
        Case "#VALUE!"
            ' La formule inclut un param�tre ou une op�rande du
            ' mauvais type. Une op�rande est une valeur ou r�f�rence
            ' � une cellule qu'une formule utilise pour calculer un
            ' r�sultat. Cette erreur arrive aussi si une ws-formule
            ' VBA personnalis�e contient une erreur.
            returnExcelError = CVErr(xlErrValue)
            
        Case Else
            returnExcelError = CVErr(xlErrValue)
    End Select
End Function
