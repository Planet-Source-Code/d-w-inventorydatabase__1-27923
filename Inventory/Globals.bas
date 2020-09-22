Attribute VB_Name = "Globals"
Option Explicit

Public Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Public Declare Function SendMessageStr Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long

Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const ES_NUMBER = &H2000&
Public Const GWL_STYLE = (-16)
Public Const LB_FINDSTRING As Long = &H18F
Public Const LB_FINDSTRINGEXACT As Long = &H1A2
Public Const CB_ERR As Long = (-1)
Public Const LB_ERR As Long = (-1)
Public Const WM_USER As Long = &H400
Public Const CB_FINDSTRING As Long = &H14C
Public Const CB_SHOWDROPDOWN As Long = &H14F

Public Counter As Integer
Public i As Integer
Public DB As Database
Public Tbl As TableDef
Public Fld As Field
Public RS As Recordset
Public Ind As Index
Public OldValue As Variant
Public StartDate As Date
Public EndDate As Date
Public CurrentID As Long
Public Network As Boolean
Public Function KeyCheck(TheKey As Integer) As Integer
Dim Keys As String
Dim Char As String
Keys = "`~!@^&()_+qwertyuiop[]asdfghjkl;'zxcvbnm,QWERTYUIOP{}|ASDFGHJKL: & """" & ZXCVBNM<>?"
Char = Chr(TheKey)
If InStr(Keys, Char) Then
KeyCheck = 0
Else
KeyCheck = TheKey
End If
End Function


Public Sub AnythingGoes(TheBox As TextBox)
Dim TheValue As Long
Dim Align As Long
Dim TheReturn As Long
Align = ES_NUMBER
TheValue = GetWindowLong(TheBox.hwnd, GWL_STYLE)
TheReturn = SetWindowLong(TheBox.hwnd, GWL_STYLE, TheValue And (Not Align))
TheBox.Refresh
End Sub

Sub Main()
If GetSetting("Inventory", "Settings", "Update", CStr(0)) = 1 Then
UpdateNetwork
End If
OpenDB
Load Data
End Sub

Private Function NetFolder() As String
NetFolder = "Z:\Nutrition"
End Function
Private Function NetPath() As String
NetPath = "Z:\Nutrition\DataFile.MDB"
End Function
Private Sub OpenDB()
On Error GoTo Out:
If Dir(NetPath) > "" Then
Set DB = OpenDatabase(NetPath)
Network = True
Else
MsgBox "Network unavailable at this time, changes will be stored on this computer temporarily.", vbInformation
Set DB = OpenDatabase(App.Path + "\DataFile.MDB")
SaveSetting "Inventory", "Settings", "Update", CStr(1)
End If
Exit Sub
Out:
MsgBox "Unable to open database."
End Sub
Public Sub NumbersOnly(TheBox As TextBox)
Dim TheValue As Long
Dim Align As Long
Dim TheReturn As Long
Align = ES_NUMBER
TheValue = GetWindowLong(TheBox.hwnd, GWL_STYLE)
TheReturn = SetWindowLong(TheBox.hwnd, GWL_STYLE, TheValue Or Align)
TheBox.Refresh
End Sub

Public Sub UpdateLocal()
Dim i As Integer
If Dir(NetFolder, vbDirectory) = "" Then Exit Sub
i = 1
Do
DoEvents
If Dir(App.Path & "\DataFile.MDB" & CStr(i)) = "" Then
Exit Do
Else
i = i + 1
    If i = 25 Then
    'only keep 25 old copies of database
    i = 1
    Exit Do
    End If
End If
Loop
Name App.Path & "\DataFile.MDB" As App.Path & "\DataFile.MDB" & CStr(i)
FileCopy NetPath, App.Path & "\DataFile.MDB"
End Sub

Private Sub UpdateNetwork()
Dim i As Integer
If Dir(NetFolder, vbDirectory) = "" Then Exit Sub
i = 1
Do
DoEvents
If Dir(NetPath & CStr(i)) = "" Then
Exit Do
Else
i = i + 1
End If
Loop
If Dir(NetPath) > "" Then
Name NetPath As NetPath & CStr(i)
End If
FileCopy App.Path & "\DataFile.MDB", NetPath
SaveSetting "Inventory", "Settings", "Update", CStr(0)
End Sub
