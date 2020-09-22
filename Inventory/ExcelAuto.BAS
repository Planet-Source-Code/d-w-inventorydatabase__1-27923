Attribute VB_Name = "ExcelAuto"
Option Explicit
Private Declare Function BringWindowToTop Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" _
    (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const APP_CAPTION = "Microsoft Excel"
Private Const APP_EXIT_DELAY = 3000
Private Const SERVER_BUSY_TIMEOUT = 500
Private Const ERR_SERVER_BUSY = -2147418111
Private Const SM_CXForm = 32
Private Const SM_CYForm = 33
Private Const SM_CYCAPTION = 4
Private Const TYPENAME_APPLICATION = "Application"
Private Const TYPENAME_OBJECT = "Object"
Private Const TYPENAME_WORKBOOK = "Workbook"

Public Enum ServerStatus
    ServerNotCreated = 1
    ServerIsBusy = 2
    ServerIsReady = 3
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public ExcelForm As Form
Public ExitFlag As Boolean
Public ExHwnd As Long
Public Sht As Excel.Worksheet
Private AppEx As Excel.Application

Private ExWb As Excel.Workbook
Private ExOriginalRect As RECT
Private ExOriginalStatusBar As Boolean



Public Function QueryUnloadExcel(ExcelForm As Form) As Boolean
If Not QuitExcel Then
QueryUnloadExcel = False
Exit Function
End If
ExcelForm.Visible = False
DoEvents
Sleep APP_EXIT_DELAY
ExitFlag = True
Unload ExcelForm
QueryUnloadExcel = True
End Function





Public Function CloseExcelWorkbook(ExWb As Excel.Workbook) As Boolean
Dim intAnswer As Integer
On Error GoTo ErrorHandler
If ExWb.Saved Then
ExWb.Close
Else
intAnswer = MsgBox("Do you want to save the changes you made to " _
    & ExWb.Name & "?", vbYesNoCancel + vbExclamation, _
    "Microsoft Excel")
If intAnswer = vbYes Then
ExWb.Close SaveChanges:=True
If TypeName(ExWb) = TYPENAME_WORKBOOK Then Exit Function
ElseIf intAnswer = vbNo Then
ExWb.Close SaveChanges:=False
Else
Exit Function
End If
End If
CloseExcelWorkbook = True
ErrorHandler:
End Function
Private Sub DoMsgBox()
Dim strMsg As String
strMsg = "Cannot automate Microsoft Excel at this time. "
strMsg = strMsg & "Please make sure Microsoft Excel is not busy "
strMsg = strMsg & "before you attempt this action."
MsgBox strMsg, vbCritical, APP_CAPTION
End Sub
Public Function GetServerStatus() As Integer
Dim strTypeName As String
Dim strTest As String
strTypeName = TypeName(AppEx)
If strTypeName = TYPENAME_APPLICATION Then
GetServerStatus = ServerIsReady
ElseIf strTypeName = TYPENAME_OBJECT Then
On Error Resume Next
strTest = AppEx.Name
If Err.Number = ERR_SERVER_BUSY Or Err.Number = 0 Then
GetServerStatus = ServerIsBusy
Else
GetServerStatus = ServerNotCreated
End If
Else
GetServerStatus = ServerNotCreated
End If
End Function
Public Function OpenFile(PathAndName As String) As Boolean
Dim lngRetVal As Long
On Error GoTo ErrorHandler
If GetServerStatus = ServerIsBusy Then
DoMsgBox
Exit Function
End If
lngRetVal = OpenExcel
If lngRetVal <> 0 Then Err.Raise lngRetVal
If TypeName(ExWb) = TYPENAME_WORKBOOK Then _
If Not CloseExcelWorkbook(ExWb) Then Exit Function
SetForegroundWindow ExHwnd
Set ExWb = AppEx.Workbooks.Open(PathAndName)
OpenFile = True
Exit Function
ErrorHandler:
MsgBox "Error " & Err.Number & ":" & vbCrLf & _
Err.Description, vbExclamation, APP_CAPTION
End Function
Public Function Initialize() As Boolean
App.OleServerBusyTimeout = SERVER_BUSY_TIMEOUT
App.OleServerBusyRaiseError = True
Dim lngRetVal As Long
lngRetVal = OpenExcel
If lngRetVal = 0 Then
SetForegroundWindow ExHwnd
Else
MsgBox "Error " & lngRetVal & ":" & vbCrLf & _
Error(lngRetVal), vbExclamation, APP_CAPTION
Initialize = False
Exit Function
End If
Initialize = True
End Function
Private Function QuitExcel() As Boolean
Dim ExWb As Excel.Workbook
On Error GoTo ErrorHandler
If GetServerStatus = ServerIsBusy Then
OpenExcel
DoMsgBox
Exit Function
ElseIf GetServerStatus = ServerIsReady Then
OpenExcel
AppEx.ScreenUpdating = False
For Each ExWb In AppEx.Workbooks
If Not CloseExcelWorkbook(ExWb) Then
AppEx.ScreenUpdating = True
Exit Function
End If
Next
AppEx.ScreenUpdating = True
If AppEx.WindowState <> xlNormal Then
AppEx.Visible = True 'orig
'AppEx.Visible = False
AppEx.WindowState = xlNormal
End If
AppEx.Visible = False
AppEx.DisplayStatusBar = ExOriginalStatusBar
With AppEx.CommandBars("Worksheet Menu Bar")
.Controls("&File").Controls("&New...").Enabled = True
End With
SetParent ExHwnd, 0
With ExOriginalRect
MoveWindow ExHwnd, .Left, .Top, .Right - .Left, .Bottom - .Top, True
End With
AppEx.DisplayAlerts = False
AppEx.Quit
End If
Set ExWb = Nothing
Set AppEx = Nothing
QuitExcel = True
ErrorHandler:
End Function


Public Sub SetAppSize(lngHwnd As Long, ExcelForm As Form)
Dim lngX As Long
Dim lngY As Long
Dim lngW As Long
Dim lngH As Long
Dim AppRect As RECT
GetWindowRect ExcelForm.hwnd, AppRect
lngX = -GetSystemMetrics(SM_CXForm)
lngY = -GetSystemMetrics(SM_CYForm)
lngW = AppRect.Right - AppRect.Left - lngX * 2
lngH = AppRect.Bottom - AppRect.Top - lngY * 2
lngY = lngY - GetSystemMetrics(SM_CYCAPTION)
lngH = lngH + GetSystemMetrics(SM_CYCAPTION)

MoveWindow lngHwnd, lngX, lngY, lngW, lngH, True
End Sub

Private Function OpenExcel() As Long
Static blnGetRect As Boolean
Dim AppTemp As Excel.Application
On Error GoTo ErrorHandler
If GetServerStatus = ServerNotCreated Then
Set AppEx = Nothing
Set AppTemp = CreateObject("Excel.Application")
Set AppEx = CreateObject("Excel.Application")
AppTemp.Quit
Set AppTemp = Nothing
AppEx.Caption = "besuretofindthisinstance"
ExHwnd = FindWindow("XLMAIN", AppEx.Caption)
AppEx.Caption = APP_CAPTION
ExOriginalStatusBar = AppEx.DisplayStatusBar
AppEx.DisplayStatusBar = True
With AppEx.CommandBars("Worksheet Menu Bar")
    .Controls("&File").Controls("&New...").Enabled = False
End With
AppEx.Visible = True 'orig
'AppEx.Visible = False
If AppEx.WindowState <> xlNormal Then
AppEx.WindowState = xlNormal
End If
If Not blnGetRect Then
GetWindowRect ExHwnd, ExOriginalRect
blnGetRect = True
End If
BringWindowToTop ExHwnd
SetParent ExHwnd, ExcelForm.hwnd
SetAppSize ExHwnd, ExcelForm
ElseIf GetServerStatus = ServerIsReady Then
If AppEx.Visible = False Then
AppEx.Visible = True 'orig
AppEx.ScreenUpdating = True
End If
End If
OpenExcel = 0
Exit Function
ErrorHandler:
OpenExcel = Err.Number
MsgBox "Error " & Err.Number & ":" & vbCrLf & _
    Error$(Err.Number), vbExclamation, APP_CAPTION
End Function






