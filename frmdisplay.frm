VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdisplay 
   Caption         =   "Project Re-Maker"
   ClientHeight    =   3120
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   5535
   Icon            =   "frmdisplay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstmain 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Extracted Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Actual Filename"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtmain 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmdisplay.frx":0E42
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuload 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnufilesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuponcompletion 
         Caption         =   "Upon Completion"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuzipwhenddone 
         Caption         =   "&Zip all files"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuzipillegal 
         Caption         =   "&Zip illegal files"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnugotoplanet 
         Caption         =   "Go to &Planetsourcecode"
      End
      Begin VB.Menu mnufilesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmdisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'en/n added to keep strings the same length
Const lookfor As String = "reference object form module class usercontrol propertypage"
Const referen As String = "reference"
Const modulen As String = "module"
Const classen As String = "class"
Const objectn As String = "object"

Private Sub Form_Resize()
If Width > 120 Then lstmain.Width = Width - 120
If Height > 720 Then lstmain.Height = Height - 720
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub

Private Sub mnugotoplanet_Click()
mnugotoplanet.Checked = Not mnugotoplanet.Checked
End Sub
Public Sub enumeratefrxfiles(filename As String)
On Error Resume Next 'Picture         =   "frmmain.frx":0E4E (fourth from last char is a :, 5th is a "
Dim frxarr() As String, tempstr As String, tempfile As Long, temp2 As Long, found As Boolean, upperlimit As Long
upperlimit = 0
tempfile = FreeFile
If fileexists(filename) Then
    Open filename For Input As #tempfile
        Do Until EOF(tempfile)
            Line Input #tempfile, tempstr
            tempstr = Trim(Replace(tempstr, Chr(0), Empty))
            
            If InStr(tempstr, ".") > 0 Then
            If Mid(tempstr, Len(tempstr) - 5, 2) = """:" And InStr(tempstr, ".") > 0 Then
                upperlimit = upperlimit + 1
                ReDim Preserve frxarr(1 To upperlimit)
                
                'MsgBox tempstr & vbNewLine & Asc(Mid(tempstr, 1, 1)) & ", " & Asc(Mid(tempstr, 2, 1)) & ", " & Asc(Mid(tempstr, 3, 1))
                frxarr(upperlimit) = Mid(tempstr, InStr(tempstr, """") + 1, InStrRev(tempstr, """") - InStr(tempstr, """") - 1)
                'Removes doubles
                found = False
                For temp2 = 1 To upperlimit - 1
                    If found = False Then
                        If LCase(frxarr(upperlimit)) = LCase(frxarr(temp2)) Then
                            ReDim Preserve frxarr(1 To upperlimit - 1)
                            upperlimit = upperlimit - 1
                            found = True
                        End If
                    End If
                Next
            End If
            End If
        Loop
    Close #tempfile
End If
For temp2 = 1 To upperlimit
     additem frxarr(temp2), "Pictures", chkdir(Left(filename, InStrRev(filename, "\")), frxarr(temp2))
Next
lstmain.Refresh
End Sub
Private Sub mnuload_Click()
On Error Resume Next
InitOpen "Visual Basic Project Files (*.vbp)" & Chr(0) & "*.vbp", "Please select a Project file to Load"
Dim filename As String, tempfile As Long, tempstr As String, newtempstr As String, tempstr3 As String
Dim patheddirs() As String, spoth As Long, found As Boolean, tempstr4 As String
patheddirs = Split(Environ("path"), ";")
ReDim Preserve patheddirs(LBound(patheddirs) To UBound(patheddirs) + 1)
filename = Open_File(Me.hWnd)
If fileexists(filename) Then
    patheddirs(UBound(patheddirs)) = Left(filename, InStrRev(filename, "\") - 1)
    tempfile = FreeFile
    lstmain.ListItems.Clear
    txtmain = Empty
    mnusave.Enabled = True
    Open filename For Input As #tempfile
        Do Until EOF(tempfile)
            Line Input #tempfile, tempstr
            If InStr(tempstr, "=") > 0 Then
                If containsword(lookfor, LCase(Left(tempstr, InStr(tempstr, "=") - 1))) = True Then
                    Select Case LCase(Left(tempstr, InStr(tempstr, "=") - 1))
                        Case objectn            'Requires stripping of #'s and ;'s
                            'unlike the others,this can be in the system directory too :(
                            found = False
                            For spoth = LBound(patheddirs) To UBound(patheddirs)
                                If found = False Then
                                    If fileexists(chkpath(patheddirs(spoth), Right(tempstr, Len(tempstr) - InStrRev(tempstr, "; ") - 1))) = True Then
                                        found = True
                                        newtempstr = chkpath(patheddirs(spoth), Right(tempstr, Len(tempstr) - InStrRev(tempstr, "; ") - 1))
                                    End If
                                End If
                            Next
                            If found = True Then
                                additem Right(tempstr, Len(tempstr) - InStrRev(tempstr, "; ") - 1), Left(tempstr, InStr(tempstr, "=") - 1), newtempstr
                            End If
                        Case referen            'Requires stripping of #'s
                            newtempstr = getnewfilename(filename, stripref(tempstr))
                            additem stripref(tempstr), Left(tempstr, InStr(tempstr, "=") - 1), newtempstr
                            tempstr3 = Left(tempstr, InStrRev(tempstr, "#") - 1): tempstr3 = Left(tempstr3, InStrRev(tempstr3, "#") - 1)
                            newtempstr = Right(newtempstr, Len(newtempstr) - InStrRev(newtempstr, "\"))
                            newtempstr = tempstr3 & "#" & newtempstr & "#" & Right(tempstr, Len(tempstr) - InStrRev(tempstr, "#"))
                            tempstr = newtempstr
                        Case modulen, classen   'Requires stripping of ;'s
                            newtempstr = Right(tempstr, Len(tempstr) - InStr(tempstr, "; ") - 1)
                            additem newtempstr, Left(tempstr, InStr(tempstr, "=") - 1), getnewfilename(filename, newtempstr)
                            tempstr = Left(tempstr, InStr(tempstr, "; ") + 1) & Right(newtempstr, Len(newtempstr) - InStrRev(newtempstr, "\"))
                        Case Else               'Requires no stripping other than the = which is in all of them
                            newtempstr = getnewfilename(filename, stripname(tempstr))
                            additem Right(tempstr, Len(tempstr) - InStrRev(tempstr, ";")), Left(tempstr, InStr(tempstr, "=") - 1), newtempstr
                            tempstr = Left(tempstr, InStr(tempstr, "=")) & Right(tempstr, Len(tempstr) - InStrRev(tempstr, "\"))
                            
                            tempstr4 = LCase(Left(tempstr, InStr(tempstr, "=") - 1))
                            If tempstr4 = "form" Or tempstr4 = "usercontrol" Then enumeratefrxfiles newtempstr
                    End Select
                End If
            End If
            If countwords(tempstr, "=") = 2 Then tempstr = Right(tempstr, Len(tempstr) - InStr(tempstr, "="))
            If txtmain <> Empty Then txtmain = txtmain & vbNewLine
            txtmain = txtmain & tempstr
        Loop
    Close #tempfile
End If
autosizeall lstmain
lstmain.Refresh
End Sub
Public Function getnewfilename(filename As String, reference As String) As String
Dim temp As String
getnewfilename = chkpath(Left(filename, InStrRev(filename, "\")), reference)
End Function
Public Function stripref(reference As String) As String
Dim temp As String
temp = reference
temp = Left(temp, InStrRev(temp, "#") - 1)
temp = Right(temp, Len(temp) - InStrRev(temp, "#"))
stripref = temp
End Function
Public Function stripname(reference As String) As String
stripname = Right(reference, Len(reference) - InStrRev(reference, "="))
End Function
Public Sub additem(Name, ftype, newname)
If InStr(Name, "=") > 0 Then Name = Right(Name, Len(Name) - InStr(Name, "="))
    lstmain.ListItems.Add , , Name
    lstmain.ListItems.Item(lstmain.ListItems.count).SubItems(1) = ftype
    lstmain.ListItems.Item(lstmain.ListItems.count).SubItems(2) = newname
End Sub

Private Sub mnusave_Click()
On Error Resume Next
Dim newfilename As String, zipfilename As String, tempfile As Long, tempstr As String
InitSave "Visual Basic Project Files (*.vbp)" & Chr(0) & "*.vbp", "Please select a location to save the Project"
newfilename = Save_File(Me.hWnd, "vbp")
If mnuzipwhenddone = True And newfilename <> Empty Then
    InitSave "Zip Files (*.zip)" & Chr(0) & "*.zip", "Please select a location to zip the Project to"
    zipfilename = Save_File(Me.hWnd, "zip")
End If
If newfilename <> Empty Then
    tempfile = FreeFile
    Open newfilename For Output As #tempfile
        Print #tempfile, txtmain.text
    Close #tempfile
End If
newfilename = Left(newfilename, InStrRev(newfilename, "\"))
For tempfile = 1 To lstmain.ListItems.count
    tempstr = lstmain.ListItems(tempfile).SubItems(2)
    tempstr = Right(tempstr, Len(tempstr) - InStrRev(tempstr, "\"))
    CopyFile lstmain.ListItems(tempfile).SubItems(2), newfilename & tempstr, zipfilename
Next
If mnugotoplanet.Checked = True Then Shell "start http://www.planetsourcecode.com/vb/authors/determine_author_type.asp?lngWId=1"
End Sub

Private Sub mnuzipwhenddone_Click()
mnuzipwhenddone.Checked = Not mnuzipwhenddone.Checked
End Sub

Public Sub CopyFile(Source As String, destination As String, Optional ZipFile As String)
'Dont ask why this part is sloppy
'I told you not to ask
'Fine, since you wont leave me a lone I'll tell you
'For some reason the destination gets set to 'false' somewhere and I dont know why
Dim buffer As String
buffer = destination
    If (fileexists(Source) = True) Then
        destination = buffer
        If fileexists(destination) = True Then
        Else
            destination = buffer
            If LCase(Source) <> LCase(destination) Then
                destination = buffer
                FileCopy Source, destination
            End If
        End If
    End If
End Sub
