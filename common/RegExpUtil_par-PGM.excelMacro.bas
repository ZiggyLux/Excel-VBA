Attribute VB_Name = "RegExpUtil"
'-----------------------------------------------------------------------------
' Application......... Templates
' Version............. 1.00
' Plateforme.......... Win 32
' Source.............. RegExpUtil_par-PGM.excelMacro.bas
' Derni�re MAJ........ 30/10/18
' Auteur.............. Patrick G. Matthews
'   https://www.experts-exchange.com/articles/1336/Using-Regular-Expressions-in-Visual-Basic-for-Applications-and-Visual-Basic-6.html
' Remarque............ VBA source file
' Br�ve description... Fonctions utiles pour expressions r�guli�res
'
' Emplacement.........
'-----------------------------------------------------------------------------
' Options
Option Explicit

' Comparaison insensible � la casse et ordres de tris suivants param�tres r�gionaux
Option Compare Text
 
Function RegExpFind(LookIn As String, PatternStr As String, Optional Pos, _
    Optional MatchCase As Boolean = True, Optional ReturnType As Long = 0, _
    Optional MultiLine As Boolean = False)
    
    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code
    
    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer
    
    ' This function uses Regular Expressions to parse a string (LookIn), and return matches to a
    ' pattern (PatternStr).  Use Pos to indicate which match you want:
    ' Pos omitted               : function returns a zero-based array of all matches
    ' Pos = 1                   : the first match
    ' Pos = 2                   : the second match
    ' Pos = <positive integer>  : the Nth match
    ' Pos = 0                   : the last match
    ' Pos = -1                  : the last match
    ' Pos = -2                  : the 2nd to last match
    ' Pos = <negative integer>  : the Nth to last match
    ' If Pos is non-numeric, or if the absolute value of Pos is greater than the number of
    ' matches, the function returns an empty string.  If no match is found, the function returns
    ' an empty string.  (Earlier versions of this code used zero for the last match; this is
    ' retained for backward compatibility)
    
    ' If MatchCase is omitted or True (default for RegExp) then the Pattern must match case (and
    ' thus you may have to use [a-zA-Z] instead of just [a-z] or [A-Z]).
    
    ' ReturnType indicates what information you want to return:
    ' ReturnType = 0            : the matched values
    ' ReturnType = 1            : the starting character positions for the matched values
    ' ReturnType = 2            : the lengths of the matched values
    
    ' If MultiLine = False, the ^ and $ match the beginning and end of input, respectively.  If
    ' MultiLine = True, then ^ and $ match the beginning and end of each line (as demarcated by
    ' new line characters) in the input string
    
    ' If you use this function in Excel, you can use range references for any of the arguments.
    ' If you use this in Excel and return the full array, make sure to set up the formula as an
    ' array formula.  If you need the array formula to go down a column, use TRANSPOSE()
    
    ' Note: RegExp counts the character positions for the Match.FirstIndex property as starting
    ' at zero.  Since VB6 and VBA has strings starting at position 1, I have added one to make
    ' the character positions conform to VBA/VB6 expectations
    
    ' Normally as an object variable I would set the RegX variable to Nothing; however, in cases
    ' where a large number of calls to this function are made, making RegX a static variable that
    ' preserves its state in between calls significantly improves performance
    
    Static RegX As Object
    Dim TheMatches As Object
    Dim Answer()
    Dim Counter As Long
    
    ' Evaluate Pos.  If it is there, it must be numeric and converted to Long
    
    If Not IsMissing(Pos) Then
        If Not IsNumeric(Pos) Then
            RegExpFind = ""
            Exit Function
        Else
            Pos = CLng(Pos)
        End If
    End If
    
    ' Evaluate ReturnType
    
    If ReturnType < 0 Or ReturnType > 2 Then
        RegExpFind = ""
        Exit Function
    End If
    
    ' Create instance of RegExp object if needed, and set properties
    
    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = True
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
    End With
        
    ' Test to see if there are any matches
    
    If RegX.Test(LookIn) Then
        
        ' Run RegExp to get the matches, which are returned as a zero-based collection
        
        Set TheMatches = RegX.Execute(LookIn)
        
        ' Test to see if Pos is negative, which indicates the user wants the Nth to last
        ' match.  If it is, then based on the number of matches convert Pos to a positive
        ' number, or zero for the last match
        
        If Not IsMissing(Pos) Then
            If Pos < 0 Then
                If Pos = -1 Then
                    Pos = 0
                Else
                    
                    ' If Abs(Pos) > number of matches, then the Nth to last match does not
                    ' exist.  Return a zero-length string
                    
                    If Abs(Pos) <= TheMatches.Count Then
                        Pos = TheMatches.Count + Pos + 1
                    Else
                        RegExpFind = ""
                        GoTo Cleanup
                    End If
                End If
            End If
        End If
        
        ' If Pos is missing, user wants array of all matches.  Build it and assign it as the
        ' function's return value
        
        If IsMissing(Pos) Then
            ReDim Answer(0 To TheMatches.Count - 1)
            For Counter = 0 To UBound(Answer)
                Select Case ReturnType
                    Case 0: Answer(Counter) = TheMatches(Counter)
                    Case 1: Answer(Counter) = TheMatches(Counter).FirstIndex + 1
                    Case 2: Answer(Counter) = TheMatches(Counter).Length
                End Select
            Next
            RegExpFind = Answer
        
        ' User wanted the Nth match (or last match, if Pos = 0).  Get the Nth value, if possible
        
        Else
            Select Case Pos
                Case 0                          ' Last match
                    Select Case ReturnType
                        Case 0: RegExpFind = TheMatches(TheMatches.Count - 1)
                        Case 1: RegExpFind = TheMatches(TheMatches.Count - 1).FirstIndex + 1
                        Case 2: RegExpFind = TheMatches(TheMatches.Count - 1).Length
                    End Select
                Case 1 To TheMatches.Count      ' Nth match
                    Select Case ReturnType
                        Case 0: RegExpFind = TheMatches(Pos - 1)
                        Case 1: RegExpFind = TheMatches(Pos - 1).FirstIndex + 1
                        Case 2: RegExpFind = TheMatches(Pos - 1).Length
                    End Select
                Case Else                       ' Invalid item number
                    RegExpFind = ""
            End Select
        End If
    
    ' If there are no matches, return empty string
    
    Else
        RegExpFind = ""
    End If
    
Cleanup:
    ' Release object variables
    
    Set TheMatches = Nothing
    
End Function

Function RegExpFindExtended(LookIn As String, PatternStr As String, Optional Pos, _
    Optional MatchCase As Boolean = True, Optional MultiLine As Boolean = False)
    
    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code
    
    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer
    
    ' This function uses Regular Expressions to parse a string (LookIn), and returns a 0-(N-1), 0-2
    ' array of the matched values (position 0 for the 2nd dimension), the starting character
    ' positions (position 1 for the 2nd dimension), and the length of the matched values (position 2
    ' for the 2nd dimension)
    
    ' Use Pos to indicate which match you want:
    ' Pos omitted               : function returns a zero-based array of all matches
    ' Pos = 1                   : the first match
    ' Pos = 2                   : the second match
    ' Pos = <positive integer>  : the Nth match
    ' Pos = 0                   : the last match
    ' Pos = -1                  : the last match
    ' Pos = -2                  : the 2nd to last match
    ' Pos = <negative integer>  : the Nth to last match
    ' If Pos is non-numeric, or if the absolute value of Pos is greater than the number of
    ' matches, the function returns an empty string.  If no match is found, the function returns
    ' an empty string.
    
    ' If MatchCase is omitted or True (default for RegExp) then the Pattern must match case (and
    ' thus you may have to use [a-zA-Z] instead of just [a-z] or [A-Z]).
    
    ' If MultiLine = False, the ^ and $ match the beginning and end of input, respectively.  If
    ' MultiLine = True, then ^ and $ match the beginning and end of each line (as demarcated by
    ' new line characters) in the input string
    
    ' If you use this function in Excel, you can use range references for any of the arguments.
    ' If you use this in Excel and return the full array, make sure to set up the formula as an
    ' array formula.  If you need the array formula to go down a column, use TRANSPOSE()
    
    ' Note: RegExp counts the character positions for the Match.FirstIndex property as starting
    ' at zero.  Since VB6 and VBA has strings starting at position 1, I have added one to make
    ' the character positions conform to VBA/VB6 expectations
    
    ' Normally as an object variable I would set the RegX variable to Nothing; however, in cases
    ' where a large number of calls to this function are made, making RegX a static variable that
    ' preserves its state in between calls significantly improves performance
    
    Static RegX As Object
    Dim TheMatches As Object
    Dim Answer()
    Dim Counter As Long
    
    ' Evaluate Pos.  If it is there, it must be numeric and converted to Long
    
    If Not IsMissing(Pos) Then
        If Not IsNumeric(Pos) Then
            RegExpFindExtended = ""
            Exit Function
        Else
            Pos = CLng(Pos)
        End If
    End If
    
    ' Create instance of RegExp object
    
    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = True
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
    End With
        
    ' Test to see if there are any matches
    
    If RegX.Test(LookIn) Then
        
        ' Run RegExp to get the matches, which are returned as a zero-based collection
        
        Set TheMatches = RegX.Execute(LookIn)
        
        If Not IsMissing(Pos) Then
            If Pos < 0 Then
                If Pos = -1 Then
                    Pos = 0
                Else
                    
                    ' If Abs(Pos) > number of matches, then the Nth to last match does not
                    ' exist.  Return a zero-length string
                    
                    If Abs(Pos) <= TheMatches.Count Then
                        Pos = TheMatches.Count + Pos + 1
                    Else
                        RegExpFindExtended = ""
                        GoTo Cleanup
                    End If
                End If
            End If
        End If
        
        ' If Pos is missing, user wants array of all matches.  Build it and assign it as the
        ' function's return value
        
        If IsMissing(Pos) Then
            ReDim Answer(0 To TheMatches.Count - 1, 0 To 2)
            For Counter = 0 To UBound(Answer)
                Answer(Counter, 0) = TheMatches(Counter)
                Answer(Counter, 1) = TheMatches(Counter).FirstIndex + 1
                Answer(Counter, 2) = TheMatches(Counter).Length
            Next
            RegExpFindExtended = Answer
        
        ' User wanted the Nth match (or last match, if Pos = 0).  Get the Nth value, if possible
        
        Else
            Select Case Pos
                Case 0                          ' Last match
                    ReDim Answer(0 To 0, 0 To 2)
                    Answer(0, 0) = TheMatches(TheMatches.Count - 1)
                    Answer(0, 1) = TheMatches(TheMatches.Count - 1).FirstIndex + 1
                    Answer(0, 2) = TheMatches(TheMatches.Count - 1).Length
                    RegExpFindExtended = Answer
                Case 1 To TheMatches.Count      ' Nth match
                    ReDim Answer(0 To 0, 0 To 2)
                    Answer(0, 0) = TheMatches(Pos - 1)
                    Answer(0, 1) = TheMatches(Pos - 1).FirstIndex + 1
                    Answer(0, 2) = TheMatches(Pos - 1).Length
                    RegExpFindExtended = Answer
                Case Else                       ' Invalid item number
                    RegExpFindExtended = ""
            End Select
        End If
    
    ' If there are no matches, return empty string
    
    Else
        RegExpFindExtended = ""
    End If
    
Cleanup:

    ' Release object variables
    
    Set TheMatches = Nothing
    
End Function

Function RegExpFindSubmatch(LookIn As String, PatternStr As String, Optional MatchPos, _
    Optional SubmatchPos, Optional MatchCase As Boolean = True, _
    Optional MultiLine As Boolean = False)
    
    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code
    
    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer
    
    ' This function uses Regular Expressions to parse a string (LookIn), and return "submatches"
    ' from the various matches to a pattern (PatternStr).  In RegExp, submatches within a pattern
    ' are defined by grouping portions of the pattern within parentheses.
    
    ' Use MatchPos to indicate which match you want:
    ' MatchPos omitted               : function returns results for all matches
    ' MatchPos = 1                   : the first match
    ' MatchPos = 2                   : the second match
    ' MatchPos = <positive integer>  : the Nth match
    ' MatchPos = 0                   : the last match
    ' MatchPos = -1                  : the last match
    ' MatchPos = -2                  : the 2nd to last match
    ' MatchPos = <negative integer>  : the Nth to last match
    
    ' Use SubmatchPos to indicate which match you want:
    ' SubmatchPos omitted               : function returns results for all submatches
    ' SubmatchPos = 1                   : the first submatch
    ' SubmatchPos = 2                   : the second submatch
    ' SubmatchPos = <positive integer>  : the Nth submatch
    ' SubmatchPos = 0                   : the last submatch
    ' SubmatchPos = -1                  : the last submatch
    ' SubmatchPos = -2                  : the 2nd to last submatch
    ' SubmatchPos = <negative integer>  : the Nth to last submatch
    
    ' The return type for this function depends on whether your choice for MatchPos is looking for
    ' a single value or for potentially many.  All arrays returned by this function are zero-based.
    ' When the function returns a 2-D array, the first dimension is for the matches and the second
    ' dimension is for the submatches
    ' MatchPos omitted, SubmatchPos omitted: 2-D array of submatches for each match.  First dimension
    '                                        based on number of matches (0 to N-1), second dimension
    '                                        based on number of submatches (0 to N-1)
    ' MatchPos omitted, SubmatchPos used   : 2-D array (0 to N-1, 0 to 0) of the specified submatch
    '                                        from each match
    ' MatchPos used, SubmatchPos omitted   : 2-D array (0 to 0, 0 to N-1) of the submatches from the
    '                                        specified match
    ' MatchPos used, SubmatchPos used      : String with specified submatch from specified match
    
    ' For any submatch that is not found, the function treats the result as a zero-length string
    
    ' If MatchCase is omitted or True (default for RegExp) then the Pattern must match case (and
    ' thus you may have to use [a-zA-Z] instead of just [a-z] or [A-Z]).
    
    ' If MultiLine = False, the ^ and $ match the beginning and end of input, respectively.  If
    ' MultiLine = True, then ^ and $ match the beginning and end of each line (as demarcated by
    ' new line characters) in the input string
    
    ' If you use this function in Excel, you can use range references for any of the arguments.
    ' If you use this in Excel and return the full array, make sure to set up the formula as an
    ' array formula.  If you need the array formula to go down a column, use TRANSPOSE()
    
    ' Normally as an object variable I would set the RegX variable to Nothing; however, in cases
    ' where a large number of calls to this function are made, making RegX a static variable that
    ' preserves its state in between calls significantly improves performance
    
    Static RegX As Object
    Dim TheMatches As Object
    Dim Mat As Object
    Dim Answer() As String
    Dim Counter As Long
    Dim SubCounter As Long
    
    ' Evaluate MatchPos.  If it is there, it must be numeric and converted to Long
    
    If Not IsMissing(MatchPos) Then
        If Not IsNumeric(MatchPos) Then
            RegExpFindSubmatch = ""
            Exit Function
        Else
            MatchPos = CLng(MatchPos)
        End If
    End If
    
    ' Evaluate SubmatchPos.  If it is there, it must be numeric and converted to Long
    
    If Not IsMissing(SubmatchPos) Then
        If Not IsNumeric(SubmatchPos) Then
            RegExpFindSubmatch = ""
            Exit Function
        Else
            SubmatchPos = CLng(SubmatchPos)
        End If
    End If
    
    ' Create instance of RegExp object
    
    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = True
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
    End With
        
    ' Test to see if there are any matches
    
    If RegX.Test(LookIn) Then
        
        ' Run RegExp to get the matches, which are returned as a zero-based collection
        
        Set TheMatches = RegX.Execute(LookIn)
        
        ' If MatchPos is missing, user either wants array of all the submatches for each match, or an
        ' array of all the specified submatches for each match.  Build it and assign it as the
        ' function's return value
        
        If IsMissing(MatchPos) Then
            
            ' Return value is a 2-D array of all the submatches for each match
            
            If IsMissing(SubmatchPos) Then
                For Counter = 0 To TheMatches.Count - 1
                    Set Mat = TheMatches(Counter)
                    
                    ' To determine how many submatches there are we need to first evaluate a match.  That
                    ' is why we redim the array inside the for/next loop
                    
                    If Counter = 0 Then
                        ReDim Answer(0 To TheMatches.Count - 1, 0 To Mat.Submatches.Count - 1) As String
                    End If
                    
                    ' Loop through the submatches and populate the array.  If the Nth submatch is not
                    ' found, RegExp returns a zero-length string
                    
                    For SubCounter = 0 To UBound(Answer, 2)
                        Answer(Counter, SubCounter) = Mat.Submatches(SubCounter)
                    Next
                Next
            
            ' Return value is a 2-D array of the specified submatch for each match.
            
            Else
                For Counter = 0 To TheMatches.Count - 1
                    Set Mat = TheMatches(Counter)
                    
                    ' To determine how many submatches there are we need to first evaluate a match.  That
                    ' is why we redim the array inside the for/next loop.  If SubmatchPos = 0, then we want
                    ' the last submatch.  In that case reset SubmatchPos so it equals the submatch count.
                    ' Negative number indicates Nth to last; convert that to applicable "positive" position
                    
                    If Counter = 0 Then
                        ReDim Answer(0 To TheMatches.Count - 1, 0 To 0) As String
                        Select Case SubmatchPos
                            Case Is > 0: 'no adjustment needed
                            Case 0, -1: SubmatchPos = Mat.Submatches.Count
                            Case Is < -Mat.Submatches.Count: SubmatchPos = -SubmatchPos
                            Case Else: SubmatchPos = Mat.Submatches.Count + SubmatchPos + 1
                        End Select
                    End If
                    
                    ' Populate array with the submatch value.  If the submatch value is not found, or if
                    ' SubmatchPos > the count of submatches, populate with a zero-length string
                    
                    If SubmatchPos <= Mat.Submatches.Count Then
                        Answer(Counter, 0) = Mat.Submatches(SubmatchPos - 1)
                    Else
                        Answer(Counter, 0) = ""
                    End If
                Next
            End If
            RegExpFindSubmatch = Answer
            
        ' User wanted the info associated with the Nth match (or last match, if MatchPos = 0)
        
        Else
            
            ' If MatchPos = 0 then make MatchPos equal the match count.  If negative (indicates Nth
            ' to last), convert to equivalent position.
            
            Select Case MatchPos
                Case Is > 0: 'no adjustment needed
                Case 0, -1: MatchPos = TheMatches.Count
                Case Is < -TheMatches.Count: MatchPos = -MatchPos
                Case Else: MatchPos = TheMatches.Count + MatchPos + 1
            End Select
            
            ' As long as MatchPos does not exceed the match count, process the Nth match.  If the
            ' match count is exceeded, return a zero-length string
            
            If MatchPos <= TheMatches.Count Then
                Set Mat = TheMatches(MatchPos - 1)
                
                ' User wants a 2-D array of all submatches for the specified match; populate array.  If
                ' a particular submatch is not found, RegExp treats it as a zero-length string
                
                If IsMissing(SubmatchPos) Then
                    ReDim Answer(0 To 0, 0 To Mat.Submatches.Count - 1)
                    For SubCounter = 0 To UBound(Answer, 2)
                        Answer(0, SubCounter) = Mat.Submatches(SubCounter)
                    Next
                    RegExpFindSubmatch = Answer
                
                ' User wants a single value
                
                Else
                    
                    ' If SubmatchPos = 0 then make it equal count of submatches.  If negative, this
                    ' indicates Nth to last; convert to equivalent positive position
                    
                    Select Case SubmatchPos
                        Case Is > 0: 'no adjustment needed
                        Case 0, -1: SubmatchPos = Mat.Submatches.Count
                        Case Is < -Mat.Submatches.Count: SubmatchPos = -SubmatchPos
                        Case Else: SubmatchPos = Mat.Submatches.Count + SubmatchPos + 1
                    End Select
                    
                    ' If SubmatchPos <= count of submatches, then get that submatch for the specified
                    ' match.  If the submatch value is not found, or if SubmathPos exceeds count of
                    ' submatches, return a zero-length string.  In testing, it appeared necessary to
                    ' use CStr to coerce the return to be a zero-length string instead of zero
                    
                    If SubmatchPos <= Mat.Submatches.Count Then
                        RegExpFindSubmatch = CStr(Mat.Submatches(SubmatchPos - 1))
                    Else
                        RegExpFindSubmatch = ""
                    End If
                End If
            Else
                RegExpFindSubmatch = ""
            End If
        End If
    
    ' If there are no matches, return empty string
    
    Else
        RegExpFindSubmatch = ""
    End If
    
Cleanup:
    ' Release object variables
    Set Mat = Nothing
    Set TheMatches = Nothing
    
End Function

Function RegExpReplace(LookIn As String, PatternStr As String, Optional ReplaceWith As String = "", _
    Optional ReplaceAll As Boolean = True, Optional MatchCase As Boolean = True, _
    Optional MultiLine As Boolean = False)
    
    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code
    
    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer
    
    ' This function uses Regular Expressions to parse a string, and replace parts of the string
    ' matching the specified pattern with another string.  The optional argument ReplaceAll
    ' controls whether all instances of the matched string are replaced (True) or just the first
    ' instance (False)
    
    ' If you need to replace the Nth match, or a range of matches, then use RegExpReplaceRange
    ' instead
    
    ' By default, RegExp is case-sensitive in pattern-matching.  To keep this, omit MatchCase or
    ' set it to True
    
    ' If MultiLine = False, the ^ and $ match the beginning and end of input, respectively.  If
    ' MultiLine = True, then ^ and $ match the beginning and end of each line (as demarcated by
    ' new line characters) in the input string
    
    ' If you use this function from Excel, you may substitute range references for all the arguments
    
    ' Normally as an object variable I would set the RegX variable to Nothing; however, in cases
    ' where a large number of calls to this function are made, making RegX a static variable that
    ' preserves its state in between calls significantly improves performance
    
    Static RegX As Object
    
    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = ReplaceAll
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
    End With
    
    RegExpReplace = RegX.Replace(LookIn, ReplaceWith)
    
End Function

Function RegExpReplaceRange(LookIn As String, PatternStr As String, Optional ReplaceWith As String = "", _
    Optional StartAt As Long = 1, Optional EndAt As Long = 0, Optional MatchCase As Boolean = True, _
    Optional MultiLine As Boolean = False)
    
    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code
    
    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer
    
    ' This function uses Regular Expressions to parse a string, and replace parts of the string
    ' matching the specified pattern with another string.  In particular, this function replaces
    ' the specified range of matched values with the designated replacement string.
    
    ' StartAt indicates the start of the range of matches to be replaced.  Thus, 2 indicates
    ' that the second match gets replaced starts the range of matches to be replaced.  Use zero
    ' to specify the last match.
    
    ' EndAt indicates the end of the range of matches to be replaced.  Thus, a 5 would indicate
    ' that the 5th match is the last one to be replaced. Use zero to specify the last match.
    
    ' Thus, if you use StartAt = 2 and EndAt = 5, then the 2nd through 5th matches will be
    ' replaced.
    
    ' By default, RegExp is case-sensitive in pattern-matching.  To keep this, omit MatchCase
    ' or set it to True
    
    ' If MultiLine = False, the ^ and $ match the beginning and end of input, respectively.  If
    ' MultiLine = True, then ^ and $ match the beginning and end of each line (as demarcated by
    ' new line characters) in the input string
    
    ' If you use this function from Excel, you may substitute range references for all the
    ' arguments
    
    ' Note: Match.FirstIndex assumes that the first character position in a string is zero.
    ' This differs from VBA and VB6, which has the first character at position 1
    
    ' Normally as an object variable I would set the RegX variable to Nothing; however, in
    ' cases where a large number of calls to this function are made, making RegX a static
    ' variable that preserves its state in between calls significantly improves performance
    
    Static RegX As Object
    Dim TheMatches As Object
    Dim StartStr As String
    Dim WorkingStr As String
    Dim Counter As Long
    Dim arr() As String
    Dim StrStart As Long
    Dim StrEnd As Long
    
    ' Instantiate RegExp object
    
    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        
        ' First search needs to find all matches
        
        .Global = True
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
        
        ' Run RegExp to find the matches
        
        Set TheMatches = .Execute(LookIn)
        
        ' If there are no matches, no replacement need to happen
    
        If TheMatches.Count = 0 Then
            RegExpReplaceRange = LookIn
            GoTo Cleanup
        End If
        
        ' Reset StartAt and EndAt if necessary based on matches actually found.  Escape if StartAt > EndAt
        ' 0 or -1 indicates last match.  Negative number indicates Nth to last
        
        Select Case StartAt
            Case Is > 0: 'no adjustment needed
            Case 0, -1: StartAt = TheMatches.Count
            Case Is < -TheMatches.Count: StartAt = -StartAt
            Case Else: StartAt = TheMatches.Count + StartAt + 1
        End Select
        
        Select Case EndAt
            Case Is > 0: 'no adjustment needed
            Case 0, -1: EndAt = TheMatches.Count
            Case Is < -TheMatches.Count: EndAt = -EndAt
            Case Else: EndAt = TheMatches.Count + EndAt + 1
        End Select
        
        If StartAt > EndAt Then
            RegExpReplaceRange = LookIn
            GoTo Cleanup
        End If
        
        ' Now create an array for the partial strings.  The elements of the array correspond to...
        ' 0         : text before the 1st match
        ' 1         : the first match
        ' 2 * N - 2 : text between the (N - 1)th and the Nth match (repeat as needed)
        ' 2 * N - 1 : the Nth match (repeat as needed)
        ' X         : text after the last match (X = 2 * number of matches)
        
        ReDim arr(0 To 2 * TheMatches.Count) As String
        
        ' Loop through the matches to populate the array
        
        For Counter = 1 To TheMatches.Count
            
            ' If Counter = 1 then it's the first match, and we need the text before the first match.
            ' If not, then we need the text between the (N - 1)th and the Nth match
            
            If Counter = 1 Then
                arr(0) = Left(LookIn, TheMatches(0).FirstIndex)
            Else
                
                ' Starting character position for text between the (N - 1)th and the Nth match
                
                StrStart = TheMatches(Counter - 2).FirstIndex + TheMatches(Counter - 2).Length + 1
                
                ' Length of text between the (N - 1)th and the Nth match
                
                StrEnd = TheMatches(Counter - 1).FirstIndex - StrStart + 1
                arr(2 * Counter - 2) = Mid(LookIn, StrStart, StrEnd)
            End If
            
            ' Now we process the match.  If the match number is within the replacement range,
            ' then put the replacement value into the array.  If not, put in the match value
            
            If Counter >= StartAt And Counter <= EndAt Then
'                arr(2 * Counter - 1) = ReplaceWith
                arr(2 * Counter - 1) = .Replace(TheMatches(Counter - 1), ReplaceWith)
            Else
                arr(2 * Counter - 1) = TheMatches(Counter - 1)
            End If
            
            ' If Counter = TheMatches.Count then we need to get the text after the last match
            
            If Counter = TheMatches.Count Then
                StrStart = TheMatches(Counter - 1).FirstIndex + TheMatches(Counter - 1).Length + 1
                arr(UBound(arr)) = Mid(LookIn, StrStart)
            End If
        Next
    End With
    
    ' Use Join to concatenate the elements of the array for our answer
    
    RegExpReplaceRange = Join(arr, "")
    
Cleanup:
    
    ' Clear object variables
    
    Set TheMatches = Nothing
    
End Function

Function RegExpReplaceExpression(LookIn As String, PatternStr As String, Expression As String, _
    Optional StartAt As Long = 1, Optional EndAt As Long = 0, _
    Optional MatchCase As Boolean = True, Optional MultiLine As Boolean = False)
    
    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code
    
    ' For more info, please see:
    ' http://www.experts-exchange.com/articles/Programming/Languages/Visual_Basic/Using-Regular-Expressions-in-Visual-Basic-for-Applications-and-Visual-Basic-6.html
    
    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer
    
    ' This function is intended for use only in Excel-based VBA projects, since it relies on the
    ' Excel Application.Evaluate method to process the expression.  The expression must use only
    ' normal arithmetic operators and/or native Excel functions.  Use $& to indicate where the
    ' entire match value should go, or $1 through $9 to use submatches 1 through 9
    
    ' This function uses Regular Expressions to parse a string, and replace parts of the string
    ' matching the specified pattern with another string.  In particular, this function replaces
    ' the specified range of matched values with the designated replacement string.  In a twist,
    ' though
    
    ' StartAt indicates the start of the range of matches to be replaced.  Thus, 2 indicates that
    ' the second match gets replaced.  Use zero to specify the last match.  Negative numbers
    ' indicate the Nth to last: -1 is the last, -2 the 2nd to last, etc
    
    ' EndAt indicates the end of the range of matches to be replaced.  Thus, a 5 would indicate
    ' that the 5th match is the last one to be replaced. Use zero to specify the last match.
    ' Negative numbers indicate the Nth to last: -1 is the last, -2 the 2nd to last, etc
    
    ' Thus, if you use StartAt = 2 and EndAt = 5, then the 2nd through 5th matches will be replaced.
    
    ' By default, RegExp is case-sensitive in pattern-matching.  To keep this, omit MatchCase or
    ' set it to True
    
    ' Note: Match.FirstIndex assumes that the first character position in a string is zero.  This
    ' differs from VBA and VB6, which has the first character at position 1
    
    ' Normally as an object variable I would set the RegX variable to Nothing; however, in cases
    ' where a large number of calls to this function are made, making RegX a static variable that
    ' preserves its state in between calls significantly improves performance
    
    Static RegX As Object
    Dim TheMatches As Object
    Dim StartStr As String
    Dim WorkingStr As String
    Dim Counter As Long
    Dim arr() As String
    Dim StrStart As Long
    Dim StrEnd As Long
    Dim Counter2 As Long
    
    ' Instantiate RegExp object
    
    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        
        ' First search needs to find all matches
        
        .Global = True
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
        
        ' Run RegExp to find the matches
        
        Set TheMatches = .Execute(LookIn)
        
        ' If there are no matches, no replacement need to happen
    
        If TheMatches.Count = 0 Then
            RegExpReplaceExpression = LookIn
            GoTo Cleanup
        End If
        
        ' Reset StartAt and EndAt if necessary based on matches actually found.  Escape if StartAt > EndAt
        
        Select Case StartAt
            Case Is > 0: 'no adjustment needed
            Case 0, -1: StartAt = TheMatches.Count
            Case Is < -TheMatches.Count: StartAt = -StartAt
            Case Else: StartAt = TheMatches.Count + StartAt + 1
        End Select
            
        Select Case EndAt
            Case Is > 0: 'no adjustment needed
            Case 0, -1: EndAt = TheMatches.Count
            Case Is < -TheMatches.Count: EndAt = -EndAt
            Case Else: EndAt = TheMatches.Count + EndAt + 1
        End Select
        
        If StartAt > EndAt Then
            RegExpReplaceExpression = LookIn
            GoTo Cleanup
        End If
        
        ' Now create an array for the partial strings.  The elements of the array correspond to...
        ' 0         : text before the 1st match
        ' 1         : the first match
        ' 2 * N - 2 : text between the (N - 1)th and the Nth match (repeat as needed)
        ' 2 * N - 1 : the Nth match (repeat as needed)
        ' X         : text after the last match (X = 2 * number of matches)
        
        ReDim arr(0 To 2 * TheMatches.Count) As String
        
        ' Loop through the matches to populate the array
        
        For Counter = 1 To TheMatches.Count
            
            ' If Counter = 1 then it's the first match, and we need the text before the first match.
            ' If not, then we need the text between the (N - 1)th and the Nth match
            
            If Counter = 1 Then
                arr(0) = Left(LookIn, TheMatches(0).FirstIndex)
            Else
                
                ' Starting character position for text between the (N - 1)th and the Nth match
                
                StrStart = TheMatches(Counter - 2).FirstIndex + TheMatches(Counter - 2).Length + 1
                
                ' Length of text between the (N - 1)th and the Nth match
                
                StrEnd = TheMatches(Counter - 1).FirstIndex - StrStart + 1
                arr(2 * Counter - 2) = Mid(LookIn, StrStart, StrEnd)
            End If
            
            ' Now we process the match.  If the match number is within the replacement range,
            ' then put the replacement value into an Evaluate expression, and place the result
            ' into the array.  If not, put in the match value
            
            If Counter >= StartAt And Counter <= EndAt Then
            
                ' $& stands in for the entire match
                
                Expression = Replace(Expression, "$&", TheMatches(Counter - 1))
                
                ' Now loop through the Submatches, if applicable, and make replacements
                
                For Counter2 = 1 To TheMatches(Counter - 1).Submatches.Count
                    Expression = Replace(Expression, "$" & Counter2, TheMatches(Counter - 1).Submatches(Counter2 - 1))
                Next
                
                ' Evaluate the expression
                
                arr(2 * Counter - 1) = Evaluate(Expression)
            Else
                arr(2 * Counter - 1) = TheMatches(Counter - 1)
            End If
            
            ' If Counter = TheMatches.Count then we need to get the text after the last match
            
            If Counter = TheMatches.Count Then
                StrStart = TheMatches(Counter - 1).FirstIndex + TheMatches(Counter - 1).Length + 1
                arr(UBound(arr)) = Mid(LookIn, StrStart)
            End If
        Next
    End With
    
    ' Use Join to concatenate the elements of the array for our answer
    
    RegExpReplaceExpression = Join(arr, "")
    
Cleanup:
    
    ' Clear object variables
    
    Set TheMatches = Nothing
    Set RegX = Nothing
    
End Function





