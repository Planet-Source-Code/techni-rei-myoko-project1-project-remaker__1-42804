Attribute VB_Name = "ListViewHandling"
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Const LVM_FIRST = &H1000
Public Sub autosizeall(lview As ListView)
If lview.ListItems.count > 0 Then
Dim count As Integer
For count = 1 To lview.ColumnHeaders.count
    AutoSizeColumnHeader lview, lview.ColumnHeaders.Item(count)
Next
'autoalign lview
End If
End Sub
Public Sub AutoSizeColumnHeader(lview As ListView, Column As ColumnHeader, Optional ByVal SizeToHeader As Boolean = True)
On Error Resume Next
    Dim l As Long
    If SizeToHeader Then l = -2 Else l = -1
    Call SendMessage(lview.hWnd, LVM_FIRST + 30, Column.Index - 1, l)
End Sub

Public Sub resizecolumnheaders(lview As ListView)
On Error Resume Next
Dim temp As Integer
If lview.ListItems.count > 0 Then
    For temp = 1 To lview.ColumnHeaders.count
        AutoSizeColumnHeader lview, lview.ColumnHeaders.Item(temp)
    Next
End If
End Sub

Public Sub autoalign(lview As ListView)
Dim count As Long, count2 As Long, foundnonnumeric As Boolean
For count = 1 To lview.ColumnHeaders.count
    foundnonnumeric = False
    For count2 = 1 To lview.ListItems.count
        If isnumeric2(getitem(lview, count, count2)) = False Then foundnonnumeric = True
    Next
    If foundnonnumeric = True Then lview.ColumnHeaders.Item(count).Alignment = lvwColumnLeft
    If foundnonnumeric = False Then lview.ColumnHeaders.Item(count).Alignment = lvwColumnRight
Next
lview.Refresh
End Sub
Public Function getitem(lview As ListView, x As Long, y As Long)
    If x = 1 Then
        getitem = lview.ListItems.Item(y).text
    Else
        getitem = lview.ListItems.Item(y).SubItems(x - 1)
    End If
End Function
Public Function isnumeric2(text As String) As Boolean
    isnumeric2 = IsNumeric(Replace(Replace(text, ".", ""), "-", ""))
End Function
