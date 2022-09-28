Attribute VB_Name = "redact_non_alpha_numeric"
Option Explicit

Public Function funcRedactNonAlphaNumeric(ByVal strTarget As String) As String
    funcRedactNonAlphaNumeric = strTarget
    Dim a$, b$, c$, i As Integer
    'The dollar sign forces the variable to return a string type rather than an undeclared variant.
    'This is faster and this procedure needs to be as fast as it can be.
    a$ = strTarget
    For i = 1 To Len(a$)
        b$ = Mid(a$, i, 1)
        If b$ Like "[A-Za-z0-9 ]" Then
            c$ = c$ & b$
        Else
            c$ = c$ & ""
        End If
    Next i
    If c$ <> strTarget Then
        funcRedactNonAlphaNumeric = Trim(c$)
    End If
End Function

Private Sub TestRedactNonAlphaNumeric()
    'Place your cursor in this procedure and click the play button.
    'Make sure you have the Immediate Window showing (Ctrl + G)
    
    Dim strTest As String
    'strTest = "This is a test."
    strTest = "It's a NP-3607"
    Debug.Print "Remove nonalphanumerics from """ & strTest & """ => " & funcRedactNonAlphaNumeric(strTest)

    strTest = "Did you see that the period (.) got removed?"
    Debug.Print "Remove nonalphanumerics from """ & strTest & """ => " & funcRedactNonAlphaNumeric(strTest)

    strTest = "Jenny's Phone# is 867-5309!"
    Debug.Print "Remove nonalphanumerics from """ & strTest & """ => " & funcRedactNonAlphaNumeric(strTest)

End Sub

'===========================================
'=                                         =
'=  OTHER VERSIONS OF THE ABOVE PROCEDURE  =
'=                                         =
'===========================================

' ALTERNATE PROCEDURE NO. 1
' REMOVES NON ALPHANUMERIC CHARACTERS INCLUDING SPACES
' THIS VERSION REDUCES THE LENGTH OF THE ORIGINAL INPUT
' Example: "It's a NP-3607." becomes "ItsaNP3607"
'Public Function funcRedactNonAlphaNumeric(ByVal strTarget As String) As String
'    funcRedactNonAlphaNumeric = strTarget
'    Dim a$, b$, c$, i As Integer
'    'The dollar sign forces the variable to return a string type rather than an undeclared variant.
'    'This is faster and this procedure needs to be as fast as it can be.
'    a$ = strTarget
'    For i = 1 To Len(a$)
'        b$ = Mid(a$, i, 1)
'        If b$ Like "[A-Za-z0-9]" Then
'            c$ = c$ & b$
'        Else
'            c$ = c$ & ""
'        End If
'    Next i
'    If c$ <> strTarget Then
'        funcRedactNonAlphaNumeric = Trim(c$)
'    End If
'End Function

' ALTERNATE PROCEDURE NO. 2
' REMOVES NON ALPHANUMERIC CHARACTERS BUT REPLACES THEM WITH A BLANK SPACE
' THIS VERSION MAINTAINS THE ORIGINAL LENGTH OF THE USER INPUT
' Example: "It's a NP 3607." becomes "It s a NP 3607"
'Public Function funcRedactNonAlphaNumeric(ByVal strTarget As String) As String
'    funcRedactNonAlphaNumeric = strTarget
'    Dim a$, b$, c$, i As Integer
'    'The dollar sign forces the variable to return a string type rather than an undeclared variant.
'    'This is faster and this procedure needs to be as fast as it can be.
'    a$ = strTarget
'    For i = 1 To Len(a$)
'        b$ = Mid(a$, i, 1)
'        If b$ Like "[A-Za-z0-9]" Then
'            c$ = c$ & b$
'        Else
'            c$ = c$ & " "
'        End If
'    Next i
'    If c$ <> strTarget Then
'        funcRedactNonAlphaNumeric = Trim(c$)
'    End If
'End Function

' ORIGINAL PROCEDURE
' REMOVES NON ALPHANUMERIC CHARACTERS BUT DOES NOT REMOVE USER TYPED SPACES
' THIS VERSION REDUCES THE LENGTH OF THE USER INPUT
' Example: "It's a NP 3607." becomes "Its a NP3607"
'Public Function funcRedactNonAlphaNumeric(ByVal strTarget As String) As String
'    funcRedactNonAlphaNumeric = strTarget
'    Dim a$, b$, c$, i As Integer
'    'The dollar sign forces the variable to return a string type rather than an undeclared variant.
'    'This is faster and this procedure needs to be as fast as it can be.
'    a$ = strTarget
'    For i = 1 To Len(a$)
'        b$ = Mid(a$, i, 1)
'        If b$ Like "[A-Za-z0-9 ]" Then
'            c$ = c$ & b$
'        Else
'            c$ = c$ & ""
'        End If
'    Next i
'    If c$ <> strTarget Then
'        funcRedactNonAlphaNumeric = Trim(c$)
'    End If
'End Function

