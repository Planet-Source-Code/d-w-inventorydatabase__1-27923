VERSION 5.00
Begin VB.Form ExcelApp 
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   Icon            =   "ExcelApp.frx":0000
   LinkTopic       =   "ExcelApp"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "ExcelApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Caption = "Microsoft Excel"
Set ExcelForm = ExcelApp
If Initialize Then
Me.Show
OpenFile App.Path & "\Temp.xls"
End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not QueryUnloadExcel(Me) Then
Cancel = True
Exit Sub
End If
Set ExcelForm = Nothing
End Sub


Private Sub Form_Resize()
If GetServerStatus <> ServerNotCreated Then _
SetAppSize ExHwnd, Me
End Sub



