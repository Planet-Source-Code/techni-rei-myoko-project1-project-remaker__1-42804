Attribute VB_Name = "publicfunctions"
Option Explicit
Public Function chkdir(directory As String, filename As String) As String
    If Right(directory, 1) <> "\" Then chkdir = directory & "\" & filename Else chkdir = directory & filename
End Function

Public Function GetRelativePath(sBase As String, sFile As String)
sBase = LCase(sBase) 'must end with a slash
sFile = LCase(sFile)
    Dim Base() As String, File() As String
    Dim I As Integer, NewTreeStart As Long, sRel As String
    If Left(sBase, 3) <> Left(sFile, 3) Then
        GetRelativePath = sFile
        Exit Function
    End If
    Base = Split(sBase, "\")
    File = Split(sFile, "\")
    While Base(I) = File(I)
        I = I + 1
    Wend
    If I = UBound(Base) Then
        While I <= UBound(File)
            sRel = sRel + File(I) + "\"
            I = I + 1
        Wend
        GetRelativePath = Left(sRel, Len(sRel) - 1)
        Exit Function
    End If
    NewTreeStart = I
    While I < UBound(Base)
        sRel = sRel & "..\"
        I = I + 1
    Wend
    While NewTreeStart <= UBound(File)
        sRel = sRel & File(NewTreeStart) + "\"
        NewTreeStart = NewTreeStart + 1
    Wend
    GetRelativePath = Left(sRel, Len(sRel) - 1)
End Function
Public Function containsword(phrase As String, word As String) As Boolean
    If Replace(phrase, word, Empty) <> phrase Then containsword = True Else containsword = False
End Function
Public Function countwords(phrase As String, word As String) As Long
    'MsgBox Len(phrase) & vbNewLine & Len(Replace(phrase, word, Empty)) & vbNewLine & Len(word) & vbNewLine & phrase & vbNewLine & word
    countwords = (Len(phrase) - Len(Replace(phrase, word, Empty))) / Len(word)
End Function
Public Function chkpath(ByVal basehref As String, ByVal URL As String) As String
'Debug.Print basehref & " " & URL
Const goback As String = "..\"
Const slash As String = "\"
Dim spoth As Long
If Left(URL, 1) = slash Then URL = Right(URL, Len(URL) - 1)
If Right(basehref, 1) = slash And Len(basehref) > 3 Then basehref = Left(basehref, Len(basehref) - 1)
If LCase(URL) <> LCase(basehref) And URL <> Empty And basehref <> Empty Then
If URL Like "?:*" Then 'is absolute
    chkpath = URL
Else
    If containsword(URL, goback) Then 'is relative
        If containsword(Right(basehref, Len(basehref) - 3), slash) = True Then
            For spoth = 1 To countwords(URL, goback)
                If countwords(basehref, slash) > 0 Then
                    URL = Right(URL, Len(URL) - Len(goback))
                    basehref = Left(basehref, InStrRev(basehref, slash) - 1)
                Else
                    URL = Replace(URL, goback, "")
                End If
            Next
        Else
            URL = Replace(URL, goback, "")
        End If
        If Right(basehref, 1) <> slash Then chkpath = basehref & slash & URL Else chkpath = basehref & URL
    Else 'is additive
        If Right(basehref, 1) <> slash Then chkpath = basehref & slash & URL Else chkpath = basehref & URL
    End If
End If
End If
End Function

Public Function fileexists(filename As String) As Boolean
On Error Resume Next
    If Dir(filename, vbNormal + vbHidden + vbSystem) <> Empty And filename <> Empty Then fileexists = True Else filename = False
End Function
