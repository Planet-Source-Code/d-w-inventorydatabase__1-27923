VERSION 5.00
Begin VB.Form Data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Management"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   3330
   Icon            =   "Data.frx":0000
   LinkTopic       =   "Data"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   3330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   8
      Left            =   1275
      TabIndex        =   7
      Top             =   3525
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   7
      Left            =   1275
      TabIndex        =   6
      Top             =   3135
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   6
      Left            =   1275
      TabIndex        =   5
      Top             =   2730
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   5
      Left            =   1275
      TabIndex        =   4
      Top             =   2340
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   4
      Left            =   1275
      TabIndex        =   3
      Top             =   1950
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   3
      Left            =   1275
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   2
      Left            =   1275
      TabIndex        =   1
      Top             =   1170
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.ComboBox Categories 
      Height          =   315
      Left            =   195
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   1
      Left            =   1275
      TabIndex        =   0
      Top             =   780
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.ComboBox DataName 
      Height          =   315
      Left            =   195
      Style           =   2  'Dropdown List
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   405
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label FieldLabel 
      Height          =   195
      Index           =   7
      Left            =   60
      TabIndex        =   17
      Top             =   3180
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label FieldLabel 
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   16
      Top             =   2790
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label FieldLabel 
      Height          =   195
      Index           =   5
      Left            =   60
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label FieldLabel 
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   14
      Top             =   2010
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label FieldLabel 
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   13
      Top             =   1620
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label FieldLabel 
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   1230
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label FieldLabel 
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label FieldLabel 
      Height          =   195
      Index           =   8
      Left            =   60
      TabIndex        =   10
      Top             =   3570
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuPrint 
         Caption         =   "Print Table"
         Shortcut        =   {F8}
      End
      Begin VB.Menu menuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menuInventory 
      Caption         =   "&Inventory"
      Begin VB.Menu menuNew 
         Caption         =   "I&n and Out"
      End
      Begin VB.Menu menuInvRpt 
         Caption         =   "&Reports"
         Begin VB.Menu menuCurrent 
            Caption         =   "&Current"
         End
         Begin VB.Menu menuPrevious 
            Caption         =   "&Previous..."
         End
         Begin VB.Menu menuReceived 
            Caption         =   "&Received..."
         End
         Begin VB.Menu menuUsage 
            Caption         =   "&Usage..."
         End
      End
      Begin VB.Menu mnuForms 
         Caption         =   "&Forms"
         Begin VB.Menu mnuList 
            Caption         =   "Inventory"
         End
         Begin VB.Menu mnuCheck 
            Caption         =   "Checklist"
         End
      End
   End
   Begin VB.Menu menuData 
      Caption         =   "&Data"
      Begin VB.Menu menuRecord 
         Caption         =   "&Add Product..."
      End
      Begin VB.Menu menuDelRecord 
         Caption         =   "&Delete Record"
      End
   End
   Begin VB.Menu mnuFind 
      Caption         =   "&Search"
      Begin VB.Menu mnuID 
         Caption         =   "By &ID Number"
      End
   End
End
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myExcelFile As New ExcelFile
Dim TheIndex As Integer
Dim TheArray() As Variant
Dim SC As ExlCell
Dim LastValue As String
Dim Loading As Boolean
Private Type ExlCell
nRow As Long
nCol As Long
End Type
Private Sub AddPrevious()
Dim Find As TableDef
Set Find = New TableDef
   For Each Find In DB.TableDefs
      If Find.Name = "Previous_Inventory" Then
      DB.TableDefs.Delete "Previous_Inventory"
      Set Find = Nothing
      End If
   Next
Set Tbl = DB.CreateTableDef("Previous_Inventory")
Set Fld = Tbl.CreateField("Product_ID", dbLong)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Description", dbText, 50)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Total_Items", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Inventory_Date", dbDate)
Tbl.Fields.Append Fld
DB.TableDefs.Append Tbl
LoadTables
Categories.ListIndex = Categories.ListCount - 1
End Sub

Private Sub AddSummary(Name As String)
Dim Find As TableDef
Set Find = New TableDef
   For Each Find In DB.TableDefs
      If Find.Name = Name Then
         DB.TableDefs.Delete Name
         Set Find = Nothing
      End If
   Next
Set Tbl = DB.CreateTableDef(Name)
Set Fld = Tbl.CreateField("Product_ID", dbLong)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Description", dbText, 50)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Used", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Added", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Start_Date", dbDate)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("End_Date", dbDate)
Tbl.Fields.Append Fld
DB.TableDefs.Append Tbl
LoadTables
Categories.ListIndex = Categories.ListCount - 1
End Sub



Private Sub UpdateRecent()
'updates the Recent_Products table with new items
Dim RC As Recordset
Dim RD As Recordset
Dim RE As Recordset
Dim RF As Recordset
Dim TempID As Long
Screen.MousePointer = vbHourglass
CurrentInventory
Set RC = DB.OpenRecordset("SELECT Product_ID FROM Current_Inventory")
Do While Not RC.EOF
DoEvents
TempID = RC.Fields("Product_ID")
Set RD = DB.OpenRecordset("SELECT Product_ID " _
& "FROM Recent_Products WHERE Product_ID " _
& "=" & TempID)
    If RD.EOF Then
    RD.Close
    Set RE = DB.OpenRecordset("Recent_Products")
    Set RF = DB.OpenRecordset("SELECT * " _
    & "FROM All_Products WHERE Product_ID " _
    & "=" & TempID)
        With RE
        .AddNew
        .Fields("Description") = RF.Fields("Description")
        .Fields("Product_ID") = RF.Fields("Product_ID")
        .Fields("Brand") = RF.Fields("Brand")
        .Fields("Supplier") = RF.Fields("Supplier")
        .Fields("Units_Per_Case") = RF.Fields("Units_Per_Case")
        .Fields("Unit_Size") = RF.Fields("Unit_Size")
        .Fields("Order_Data") = RF.Fields("Order_Data")
        .Fields("Price") = RF.Fields("Price")
        .Update
        .Close
        End With
    RF.Close
    Else
    RD.Close
    End If
RC.MoveNext
Loop
RC.Close
Screen.MousePointer = vbDefault
End Sub

Private Sub CurrentInventory()
Dim i As Integer
Dim n As Integer
Dim Des As String
Dim Tot As Integer
Dim RC As Recordset
Dim RD As Recordset
Dim TempID As Long
Set RC = DB.OpenRecordset("SELECT * FROM Transactions " _
  & "ORDER BY Product_ID")
AddCurrent
Set RD = DB.OpenRecordset("Current_Inventory")
For i = 1 To RC.RecordCount
If RC.Fields("Product_ID") = TempID Then GoTo Skip:
TempID = RC.Fields("Product_ID")
Des = RC.Fields("Description")
Tot = RC.Fields("Updated_Inventory")
With RD
.AddNew
.Fields("Product_ID") = TempID
CurrentID = TempID
.Fields("Description") = Des
.Fields("Total_Items") = Tot
.Fields("Inventory_Date") = Date
.Update
End With
Skip:
RC.MoveNext
If RC.EOF Then Exit For
Next
RC.Close
Set RC = Nothing
RD.Close
Set RD = Nothing
LoadTables
Categories.ListIndex = Categories.ListCount - 1
End Sub
Private Function DescrFromID(ID As Long, Inform As Boolean) As String
Dim RC As Recordset
Set RC = DB.OpenRecordset("SELECT Product_ID," _
& "Description FROM All_Products WHERE Product_ID " _
& "=" & ID)
If RC.EOF Then
If Inform Then
MsgBox "Product ID not found in Database."
End If
DescrFromID = " "
Else
DescrFromID = RC.Fields("Description")
End If
RC.Close
End Function

Private Sub FindID()
Dim TempID As String
Dim ID As Long
Dim temp As String
TempID = InputBox("Enter Product ID Number", "CHECK FOR EXISTING ID NUMBER")
If TempID = "" Then Exit Sub
On Error GoTo ID_Err:
ID = CLng(TempID)
On Error GoTo 0
If DescrFromID(ID, True) <> " " Then
Categories = "All_Products"
DataName = DescrFromID(ID, False)
OpenFromID ID
End If
Exit Sub
ID_Err:
MsgBox "Numbers only."
FindID
End Sub

Private Function FindInList(CB As ComboBox, TheItem As String) As Integer
FindInList = SendMessageStr(CB.hwnd, CB_FINDSTRING, -1, TheItem)
End Function

Private Sub OpenFromID(ID As Long)
  Loading = True
  Set RS = DB.OpenRecordset("SELECT * " _
& "FROM " & Categories & " WHERE Product_ID = " & ID)
With RS
       Value(1) = .Fields("Product_ID")
       FieldLabel(1) = "Product ID:"
       Value(2) = .Fields("Brand")
       FieldLabel(2) = "Brand:"
       Value(3) = .Fields("Supplier")
       FieldLabel(3) = "Supplier:"
       Value(4) = .Fields("Units_Per_Case")
       FieldLabel(4) = "Items In Case:"
       Value(5) = .Fields("Unit_Size")
       FieldLabel(5) = "Item Size:"
       Value(6) = .Fields("Order_Data")
       FieldLabel(6) = "Order By:"
       Value(7) = Format(.Fields("Price"), "$##.#0")
       FieldLabel(7) = "Price:"
          For i = 1 To 7
             Value(i).Visible = True
             FieldLabel(i).Visible = True
          Next
          Data.Height = Value(7).Top _
          + (Value(7).Height * 4)
.Close
End With
Loading = False
End Sub


Private Sub SummarizeReceived()
Dim i As Integer
Dim n As Integer
Dim Des As String
Dim Used As Integer
Dim Tot As Integer
Dim RC As Recordset
Dim RD As Recordset
Dim TempID As Long
Set RC = DB.OpenRecordset("SELECT * FROM Transactions " _
  & "ORDER BY Product_ID")
AddSummary "Added"
Set RD = DB.OpenRecordset("Added")

For i = 1 To RC.RecordCount
If TempID <> 0 Then
   If TempID <> RC.Fields("Product_ID") Then
      If Tot > 0 Then
      RC.MovePrevious
      TempID = RC.Fields("Product_ID")
      Des = RC.Fields("Description")
      RD.AddNew
      RD.Fields("Product_ID") = TempID
      RD.Fields("Description") = Des
      RD.Fields("Added") = Tot
      RD.Fields("Start_Date") = StartDate
      RD.Fields("End_Date") = EndDate
      RD.Update
      RC.MoveNext
      Tot = 0
      End If
   End If
End If

If RC.Fields("Date") < StartDate Then GoTo Skip:

TempID = RC.Fields("Product_ID")
Des = RC.Fields("Description")

   If RC.Fields("Date") <= EndDate Then
   Tot = Tot + RC.Fields("Received")
   Else
   RC.MovePrevious
      If Not RC.EOF Then
         If Tot > 0 Then
         TempID = RC.Fields("Product_ID")
         Des = RC.Fields("Description")
         RD.AddNew
         RD.Fields("Product_ID") = TempID
         RD.Fields("Description") = Des
         RD.Fields("Added") = Tot
         RD.Fields("Start_Date") = StartDate
         RD.Fields("End_Date") = EndDate
         RD.Update
         RC.MoveNext
         Tot = 0
         End If
      Else
      RC.MoveFirst
      End If
   End If
Skip:
RC.MoveNext
If RC.EOF Then
RC.MoveLast
   If RC.Fields("Date") > StartDate _
   And RC.Fields("Date") <= EndDate Then
      If Tot > 0 Then
      TempID = RC.Fields("Product_ID")
      Des = RC.Fields("Description")
      RD.AddNew
      RD.Fields("Product_ID") = TempID
      RD.Fields("Description") = Des
      RD.Fields("Added") = Tot
      RD.Fields("Start_Date") = StartDate
      RD.Fields("End_Date") = EndDate
      RD.Update
      End If
   End If
End If
Next
RC.Close
Set RC = Nothing
RD.Close
Set RD = Nothing
LoadTables
Categories.ListIndex = Categories.ListCount - 1
End Sub
Private Sub SummarizeUsage()
Dim i As Integer
Dim n As Integer
Dim Des As String
Dim Used As Integer
Dim Tot As Integer
Dim RC As Recordset
Dim RD As Recordset
Dim TempID As Long
Set RC = DB.OpenRecordset("SELECT * FROM Transactions " _
  & "ORDER BY Product_ID")
AddSummary "Used"
Set RD = DB.OpenRecordset("Used")

For i = 1 To RC.RecordCount
If TempID <> 0 Then
   If TempID <> RC.Fields("Product_ID") Then
      If Tot > 0 Then
      RC.MovePrevious
      TempID = RC.Fields("Product_ID")
      Des = RC.Fields("Description")
      RD.AddNew
      RD.Fields("Product_ID") = TempID
      RD.Fields("Description") = Des
      RD.Fields("Used") = Tot
      RD.Fields("Start_Date") = StartDate
      RD.Fields("End_Date") = EndDate
      RD.Update
      RC.MoveNext
      Tot = 0
      End If
   End If
End If
If RC.Fields("Date") < StartDate Then GoTo Skip:

TempID = RC.Fields("Product_ID")
Des = RC.Fields("Description")

   If RC.Fields("Date") <= EndDate Then
   Tot = Tot + RC.Fields("Removed")
   Else
   RC.MovePrevious
      If Not RC.EOF Then
         If Tot > 0 Then
         TempID = RC.Fields("Product_ID")
         Des = RC.Fields("Description")
         RD.AddNew
         RD.Fields("Product_ID") = TempID
         RD.Fields("Description") = Des
         RD.Fields("Used") = Tot
         RD.Fields("Start_Date") = StartDate
         RD.Fields("End_Date") = EndDate
         RD.Update
         RC.MoveNext
         Tot = 0
         End If
      Else
      RC.MoveFirst
      End If
   End If
Skip:
RC.MoveNext
If RC.EOF Then
   RC.MoveLast
   If RC.Fields("Date") > StartDate _
   And RC.Fields("Date") <= EndDate Then
      If Tot > 0 Then
      TempID = RC.Fields("Product_ID")
      Des = RC.Fields("Description")
      RD.AddNew
      RD.Fields("Product_ID") = TempID
      RD.Fields("Description") = Des
      RD.Fields("Used") = Tot
      RD.Fields("Start_Date") = StartDate
      RD.Fields("End_Date") = EndDate
      RD.Update
      End If
   End If
End If
Next
RC.Close
Set RC = Nothing
RD.Close
Set RD = Nothing
LoadTables
Categories.ListIndex = Categories.ListCount - 1
End Sub


Private Function NextTransaction() As Integer
Dim RC As Recordset
Set RC = DB.OpenRecordset("SELECT Transaction_ID " _
  & "FROM Transactions ORDER BY Transaction_ID")
If Not RC.EOF Then
RC.MoveLast
NextTransaction = RC.Fields("Transaction_ID") + 1
Else
NextTransaction = 1
End If
RC.Close
Set RC = Nothing
End Function

Private Sub PreviousInventory()
Dim i As Integer
Dim n As Integer
Dim Des As String
Dim Tot As Integer
Dim RC As Recordset
Dim RD As Recordset
Dim TempID As Long
Set RC = DB.OpenRecordset("SELECT * FROM Transactions " _
  & "ORDER BY Product_ID")
AddPrevious
Set RD = DB.OpenRecordset("Previous_Inventory")
For i = 1 To RC.RecordCount
If RC.Fields("Product_ID") = TempID Then GoTo Skip:
If RC.Fields("Date") > EndDate Then GoTo Skip:
TempID = RC.Fields("Product_ID")
Des = RC.Fields("Description")
Tot = RC.Fields("Updated_Inventory")
With RD
.AddNew
.Fields("Product_ID") = TempID
.Fields("Description") = Des
.Fields("Total_Items") = Tot
.Fields("Inventory_Date") = EndDate
.Update
End With
Skip:
RC.MoveNext
If RC.EOF Then Exit For
Next
RC.Close
Set RC = Nothing
RD.Close
Set RD = Nothing
LoadTables
Categories = "Previous_Inventory"
DataName_Click
End Sub

Private Sub AddTransaction(Optional ID As Long)
Dim RC As Recordset
Dim temp As String
Dim Descr As String
Dim TheDate As Date
If ID = 0 Then
temp = InputBox("Enter Product ID", "ADD NEW TRANSACTION")
temp = Trim(temp)
    If temp = "" Then Exit Sub
    If Not IsNumeric(temp) Then
    MsgBox "Invalid ID"
    Exit Sub
    End If
Else
temp = ID
End If
TheDate = Date
Loading = True
Set RC = DB.OpenRecordset("SELECT * " _
& "FROM Transactions WHERE " _
& "Date = #" & TheDate & "# AND " _
& "Product_ID = " & Val(temp))
Descr = DescrFromID(Val(temp), True)
    If Descr = " " Then
    Exit Sub
    End If
With RC
   If .RecordCount > 0 Then
   MsgBox "Found entry dated today. Please edit it for this transaction."
   Categories = "Transactions"
   DataName = Descr
  
   Value(1) = .Fields("Date")
   FieldLabel(1) = "Date:"
   Value(2) = .Fields("Product_ID")
   CurrentID = .Fields("Product_ID")
   FieldLabel(2) = "Product ID:"
   Value(3) = .Fields("Units_Per_Case")
   FieldLabel(3) = "Units Per Case:"
   Value(4) = .Fields("Received")
   FieldLabel(4) = "Units Received:"
   Value(5) = .Fields("Removed")
   FieldLabel(5) = "Units Removed:"
   Value(6) = .Fields("Updated_Inventory")
   
   FieldLabel(6) = "Units In Stock:"
   Value(7) = .Fields("Transaction_ID")
   FieldLabel(7) = "Transaction ID:"
   
   .Close
   Loading = False
   Exit Sub
   End If
End With

Set RS = DB.OpenRecordset("Transactions")
With RS
    .AddNew
    .Fields("Date") = Date
    .Fields("Product_ID") = Val(temp)
    .Fields("Description") = Descr
    .Fields("Units_Per_Case") = UnitsFromID(Val(temp))
    .Fields("Received") = 0
    .Fields("Removed") = 0
    .Fields("Updated_Inventory") = OldTotal(Val(temp))
    .Fields("Transaction_ID") = NextTransaction
    .Update
    ClearBoxes
    Value(1) = .Fields("Date")
    FieldLabel(1) = "Date:"
    Value(2) = .Fields("Product_ID")
    CurrentID = .Fields("Product_ID")
    FieldLabel(2) = "Product ID:"
    Value(3) = .Fields("Units_Per_Case")
    FieldLabel(3) = "Units Per Case:"
    Value(4) = .Fields("Received")
    FieldLabel(4) = "Units Received:"
    Value(5) = .Fields("Removed")
    FieldLabel(5) = "Units Removed:"
    Value(6) = .Fields("Updated_Inventory")
    FieldLabel(6) = "Units In Stock:"
    Value(7) = .Fields("Transaction_ID")
    FieldLabel(7) = "Transaction ID:"
   .Close
    Loading = False
End With

LoadRecords
Categories = "Transactions"
DataName = Descr

End Sub
Private Sub AddRecord()
Dim TempID As String
Dim ID As Long
Dim temp As String
TempID = InputBox("Enter Product ID Number", "CHECK FOR EXISTING ID NUMBER")
If TempID = "" Then Exit Sub
On Error GoTo ID_Err:
ID = CLng(TempID)
On Error GoTo 0
If DescrFromID(ID, False) = " " Then
temp = InputBox("Enter Product Description.", "PRODUCT DESCRIPTION FOR ID # " & ID)
temp = Trim(temp)
If temp = "" Then Exit Sub
Set RS = DB.OpenRecordset("All_Products")
With RS
.AddNew
.Fields("Description") = temp
.Fields("Product_ID") = ID
CurrentID = ID
.Fields("Brand") = ">BRAND<"
.Fields("Supplier") = ">SUPPLIER<"
.Fields("Units_Per_Case") = 0
.Fields("Unit_Size") = ">UNITSIZE<"
.Fields("Order_Data") = ">ORDERDATA<"
.Fields("Price") = 0
.Update
.Close
End With
Categories = "All_Products"
DataName = temp
DataName_Click
MsgBox "Please fill in all remaining information."
Else
MsgBox "Item found in database."
Categories = "All_Products"
DataName = DescrFromID(ID, False)
OpenFromID ID
End If
Exit Sub
ID_Err:
MsgBox "Numbers only in ID!"
AddRecord
End Sub




Private Sub AddTransactionTable()
'no menu for this, used only once
Set Tbl = DB.CreateTableDef("Transactions")
Set Fld = Tbl.CreateField("Date", dbDate)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Description", dbText, 50)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Product_ID", dbLong)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Units_Per_Case", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Received", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Removed", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Updated_Inventory", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Transaction_ID", dbInteger)
Tbl.Fields.Append Fld
DB.TableDefs.Append Tbl
Categories.ListIndex = Categories.ListCount - 1
LoadTables
End Sub

Private Sub AddCurrent()
Dim Find As TableDef
Set Find = New TableDef
   For Each Find In DB.TableDefs
      If Find.Name = "Current_Inventory" Then
         DB.TableDefs.Delete "Current_Inventory"
         Set Find = Nothing
      End If
   Next
Set Tbl = DB.CreateTableDef("Current_Inventory")
Set Fld = Tbl.CreateField("Product_ID", dbLong)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Description", dbText, 50)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Total_Items", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Inventory_Date", dbDate)
Tbl.Fields.Append Fld

DB.TableDefs.Append Tbl
LoadTables
Categories.ListIndex = Categories.ListCount - 1
End Sub
Private Sub AddTable()
'no menu for this, used only to build database tables
Dim Find As TableDef
Dim temp As String
temp = InputBox("Enter Name", "ADD NEW CATEGORY")
temp = Trim(temp)
If temp = "" Then Exit Sub
Set Find = New TableDef
   For Each Find In DB.TableDefs
      If Find.Name = temp Then
      MsgBox "A category by that name already exists."
         For i = 0 To Categories.ListCount - 1
         Categories.ListIndex = i
            If temp = Categories.Text Then
            Set Find = Nothing
            Exit For
            End If
         Next
      Exit Sub
      End If
   Next
Set Tbl = DB.CreateTableDef(temp)
Set Fld = Tbl.CreateField("Product ID", dbLong)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Description", dbText, 50)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Brand", dbText, 15)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Supplier", dbText, 15)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Units Per Case", dbInteger)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Unit Size", dbText, 15)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Order Data", dbText, 15)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Price", dbCurrency)
Tbl.Fields.Append Fld
DB.TableDefs.Append Tbl
LoadTables
Categories.ListIndex = Categories.ListCount - 1
End Sub





Private Sub DeleteRecord()
Dim Action As VbMsgBoxResult
Action = MsgBox("Delete the " & Categories & ": """ & DataName & """ " & _
 " and all its fields?", vbYesNo, "DELETE " & UCase(Categories) & "?")
If Action = vbYes Then

Set RS = DB.OpenRecordset("SELECT * " _
& "FROM " & Categories & " ORDER BY Description")

With RS
.Move DataName.ListIndex
   If .Fields("Description") = DataName Then
   .Delete
   End If
.Close
End With
LoadTables
LoadRecords
End If
End Sub

Private Sub DeleteTable()
'no menu for this, used in debugging database
Dim Action As VbMsgBoxResult
If Categories.ListCount = 0 Then Exit Sub
Action = MsgBox("Do you really want to delete the category " & """" & Categories & """" & " and all the data it contains.", vbYesNo, "DELETE TABLE")
If Action = vbYes Then
DB.TableDefs.Delete Categories
LoadTables
End If
End Sub
Private Sub LoadTables()
Dim temp As String
Categories.Clear
DataName.Clear
DataName.Visible = False
ClearBoxes
For i = 0 To DB.TableDefs.Count - 1
temp = DB.TableDefs(i).Name
   If Left(temp, 4) <> "MSys" Then
   Categories.AddItem temp
   End If
Next
   If Categories.ListCount > 0 Then
   Categories.ListIndex = Categories.ListCount - 3
   Categories.Visible = True
   Else
   Categories.Visible = False
   End If
End Sub
Private Sub LoadRecords()
DataName.Clear
Loading = True
ClearBoxes
If Categories.ListCount > 0 Then

Set RS = DB.OpenRecordset("SELECT * " _
& "FROM " & Categories & " ORDER BY Description")

With RS
   If .RecordCount <> 0 Then
      For Counter = 1 To .RecordCount
      DataName.AddItem .Fields("Description")
      .MoveNext
      If .EOF Then Exit For
      Next
   End If
.Close
End With

End If

If DataName.ListCount > 0 Then
DataName.ListIndex = 0
DataName.Visible = True
Else
DataName.Visible = False
Data.Height = Categories.Top + (Categories.Height * 4)
End If
End Sub

Private Function SaveData(Index As Integer) As Boolean
On Error GoTo SaveErr:
Dim Total As Integer
Set RS = DB.OpenRecordset("SELECT * FROM " _
& Categories & " WHERE Product_ID = " & CurrentID)
With RS
If .EOF Then GoTo Skip:
.Edit
   If Trim(Value(Index)) <> "" Then
   Select Case Categories
   Case "All_Products", "Recent_Products"
      Select Case Index
      Case 1
      .Fields("Product_ID") = Val(Trim(Value(1)))
      Case 2
      .Fields("Brand") = Trim(Value(2))
      Case 3
      .Fields("Supplier") = Trim(Value(3))
      Case 4
      .Fields("Units_Per_Case") = Val(Trim(Value(4)))
      Case 5
      .Fields("Unit_Size") = Trim(Value(5))
      Case 6
      .Fields("Order_Data") = Trim(Value(6))
      Case 7
      .Fields("Price") = Val(Trim(Value(7)))
      
      End Select
      OldValue = Value(Index)
  Case "Transactions"
          Select Case Index
          Case 4, 5
          If Index = 5 Then
          Total = Val(Value(6)) - (Val(Value(Index)) - Val(OldValue))
          Else
          Total = Val(Value(6)) + (Val(Value(Index)) - Val(OldValue))
          End If
              If Total < 0 Then
              MsgBox "That value results in a negative inventory."
              Value(Index) = OldValue
              Value(Index).SetFocus
              SaveData = False
              Exit Function
              End If
          .Fields("Received") = Val(Value(4))
          .Fields("Removed") = Val(Value(5))
          .Fields("Updated_Inventory") = Total
          Value(6) = Total
          OldValue = Val(Value(Index))
          FieldLabel(6) = "Units In Stock:"
          Value(6).Visible = True
          FieldLabel(6).Visible = True
          End Select
      End Select
   .Update
   Else
   MsgBox "Blank fields not allowed"
   Value(Index) = OldValue
   Value(Index).SetFocus
   SaveData = False
   Exit Function
   End If
.Close
End With
SaveData = True
Exit Function
Skip:
SaveData = True
Exit Function
SaveErr:
SaveData = False
End Function

Private Sub ClearBoxes()
For i = 1 To 8
Value(i) = ""
Value(i).Visible = False
FieldLabel(i) = ""
FieldLabel(i).Visible = False
Next
End Sub

Private Sub SetBoxes(TheType As Integer)

Select Case TheType
Case 1
Value(1).Locked = True
NumbersOnly Value(1)
Value(2).Locked = False
AnythingGoes Value(2)
Value(2).Locked = False
AnythingGoes Value(3)
Value(3).Locked = False
NumbersOnly Value(4)
Value(4).Locked = False
AnythingGoes Value(5)
Value(5).Locked = False
AnythingGoes Value(6)
Value(6).Locked = False
AnythingGoes Value(7)
Value(7).Locked = False
Case 2
AnythingGoes Value(1)
Value(1).Locked = True
NumbersOnly Value(2)
Value(2).Locked = True
NumbersOnly Value(3)
Value(3).Locked = True
NumbersOnly Value(4)
Value(4).Locked = False
NumbersOnly Value(5)
Value(5).Locked = False
NumbersOnly Value(6)
Value(6).Locked = True
NumbersOnly Value(7)
Value(7).Locked = True
Case 3
NumbersOnly Value(1)
Value(1).Locked = True
AnythingGoes Value(2)
Value(2).Locked = True
NumbersOnly Value(3)
Value(3).Locked = True
AnythingGoes Value(4)
Value(4).Locked = True
Case 4
NumbersOnly Value(1)
Value(1).Locked = True
AnythingGoes Value(2)
Value(2).Locked = True
NumbersOnly Value(3)
Value(3).Locked = True
AnythingGoes Value(4)
Value(4).Locked = True
AnythingGoes Value(5)
Value(5).Locked = True
End Select
End Sub


Private Function UnitsFromID(ID As Long) As Integer
Set RS = DB.OpenRecordset("SELECT Product_ID," _
& "Units_Per_Case FROM All_Products WHERE Product_ID " _
& "= " & ID)
If RS.EOF Then
MsgBox "Product ID not found."
UnitsFromID = 0
Else
UnitsFromID = RS.Fields("Units_Per_Case")
End If
RS.Close
End Function

Private Function OldTotal(ID As Long) As Integer
Dim RC As Recordset
Set RC = DB.OpenRecordset("SELECT Date, Product_ID, " _
& "Updated_Inventory FROM Transactions WHERE Product_ID " _
& "= " & ID & " ORDER BY Date DESC")
With RC
If Not .EOF Then
OldTotal = .Fields("Updated_Inventory")
Else
OldTotal = 0
End If
.Close
Set RC = Nothing
End With
End Function

Private Function PreviousItems(ID As Long) As Integer
Dim RC As Recordset
Set RC = DB.OpenRecordset("SELECT Date, Product_ID, " _
& "Updated_Inventory FROM Transactions WHERE Product_ID " _
& "= " & ID & " ORDER BY Date DESC")
With RC
.MoveNext
If Not .EOF Then
PreviousItems = .Fields("Updated_Inventory")
Else
PreviousItems = 0
End If
.Close
Set RC = Nothing
End With
End Function


Private Sub WriteAll(Value As Variant, ByVal nRow As Long, ByVal nCol As Long)
With myExcelFile
Select Case VarType(Value)
Case 0 'empty
.WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, nRow, nCol, ""
Case 1 'null
.WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, nRow, nCol, ""
Case 2 'integer
.WriteValue xlsInteger, xlsFont0, xlsRightAlign, xlsNormal, nRow, nCol, Value
Case 3, 4, 5, 6 'long, single, double
.WriteValue xlsNumber, xlsFont0, xlsRightAlign, xlsNormal, nRow, nCol, Value
Case 7 'date
.WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, nRow, nCol, Value
Case 8 'string
.WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, nRow, nCol, Value
Case 12 'variant
.WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, nRow, nCol, Value
Case 14 'decimal
.WriteValue xlsNumber, xlsFont0, xlsRightAlign, xlsNormal, nRow, nCol, Value
Case Else
'9=object, 10=error, 11=boolean, 17=byte, 36=userdef, 8192=array
End Select
End With
End Sub

Private Sub WriteInventory()
UpdateRecent
On Error GoTo FileError
Dim FileName As String
Dim i As Integer
Dim Row As Long
Row = 3
With myExcelFile
FileName = App.Path & "\Temp.xls"
.CreateFile FileName
.PrintGridLines = True
.SetMargin xlsTopMargin, 1
.SetMargin xlsLeftMargin, 1
.SetMargin xlsRightMargin, 1
.SetMargin xlsBottomMargin, 1
.SetFont "Arial", 10, xlsNoFormat
.SetFont "Arial", 10, xlsBold
.SetFont "Arial", 10, xlsBold + xlsUnderline
.SetFont "Courier", 12, xlsItalic
.SetHeader "INVENTORY"
    'Description
    .SetColumnWidth 1, 1, 26
    'Product_ID
    .SetColumnWidth 2, 2, 8
    'Brand
    .SetColumnWidth 3, 3, 8
    'Supplier
    .SetColumnWidth 4, 4, 7
    'Units_Per_Case
    .SetColumnWidth 5, 5, 4
    'Unit_Size
    .SetColumnWidth 6, 6, 8
    'Cases
    .SetColumnWidth 7, 7, 10
    
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 1, "DESCRIPTION"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 2, "ID #"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 3, "BRAND"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 4, "SUPPL"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 5, "#/Cs"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 6, "SIZE"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 7, "CASES"

Set RS = DB.OpenRecordset("SELECT * " _
  & "FROM Recent_Products ORDER BY Description")
If RS.RecordCount > 0 Then
    For i = 3 To RS.RecordCount + 2
        WriteAll RS.Fields("Description"), Row, 1
        WriteAll RS.Fields("Product_ID"), Row, 2
        WriteAll RS.Fields("Brand"), Row, 3
        WriteAll RS.Fields("Supplier"), Row, 4
        WriteAll RS.Fields("Units_Per_Case"), Row, 5
        WriteAll RS.Fields("Unit_Size"), Row, 6
    RS.MoveNext
    If RS.EOF Then Exit For
    Row = Row + 1
    Next
    RS.Close
    Set RS = Nothing
 End If
.CloseFile
End With
Exit Sub
FileError:
MsgBox "Error writing to Excel, " & Err.Description
End Sub

Private Sub WriteCheckList()
UpdateRecent
On Error GoTo FileError
Dim FileName As String
Dim i As Integer
Dim Row As Long
Row = 3
With myExcelFile
FileName = App.Path & "\Temp.xls"
.CreateFile FileName
.PrintGridLines = True
.SetMargin xlsTopMargin, 1
.SetMargin xlsLeftMargin, 1
.SetMargin xlsRightMargin, 1
.SetMargin xlsBottomMargin, 1
.SetFont "Arial", 10, xlsNoFormat
.SetFont "Arial", 10, xlsBold
.SetFont "Arial", 10, xlsBold + xlsUnderline
.SetFont "Courier", 12, xlsItalic
.SetHeader "RECIEVING AND USAGE"
    'Description
    .SetColumnWidth 1, 1, 26
    'Product_ID
    .SetColumnWidth 2, 2, 8
    'Brand
    .SetColumnWidth 3, 3, 8
    'Supplier
    .SetColumnWidth 4, 4, 7
    'Units_Per_Case
    .SetColumnWidth 5, 5, 4
    'Unit_Size
    .SetColumnWidth 6, 6, 8
    'Received
    .SetColumnWidth 7, 7, 10
    'Removed
    .SetColumnWidth 8, 8, 10
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 1, "DESCRIPTION"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 2, "ID #"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 3, "BRAND"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 4, "SUPPL"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 5, "#/CS"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 6, "SIZE"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 7, "RECEIVED"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 8, "REMOVED"
Set RS = DB.OpenRecordset("SELECT * " _
  & "FROM Recent_Products ORDER BY Description")
If RS.RecordCount > 0 Then
    For i = 3 To RS.RecordCount + 2
        WriteAll RS.Fields("Description"), Row, 1
        WriteAll RS.Fields("Product_ID"), Row, 2
        WriteAll RS.Fields("Brand"), Row, 3
        WriteAll RS.Fields("Supplier"), Row, 4
        WriteAll RS.Fields("Units_Per_Case"), Row, 5
        WriteAll RS.Fields("Unit_Size"), Row, 6
    RS.MoveNext
    If RS.EOF Then Exit For
    Row = Row + 1
    Next
    RS.Close
    Set RS = Nothing
 End If
.CloseFile
End With
Exit Sub
FileError:
MsgBox "Error writing to Excel, " & Err.Description
End Sub


Private Sub WriteExcelFile()
On Error GoTo FileError
Dim FileName As String
Dim i As Integer
Dim Row As Long
Row = 3
With myExcelFile
FileName = App.Path & "\Temp.xls"
.CreateFile FileName
.PrintGridLines = False
.SetMargin xlsTopMargin, 1
.SetMargin xlsLeftMargin, 1
.SetMargin xlsRightMargin, 1
.SetMargin xlsBottomMargin, 1
.SetFont "Arial", 10, xlsNoFormat
.SetFont "Arial", 10, xlsBold
.SetFont "Arial", 10, xlsBold + xlsUnderline
.SetFont "Courier", 12, xlsItalic
Select Case Categories
    Case "All_Products", "Recent_Products"
    'Description
    .SetColumnWidth 1, 1, 26
    'Product_ID
    .SetColumnWidth 2, 2, 8
    'Brand
    .SetColumnWidth 3, 3, 8
    'Supplier
    .SetColumnWidth 4, 4, 7
    'Units_Per_Case
    .SetColumnWidth 5, 5, 4
    'Unit_Size
    .SetColumnWidth 6, 6, 5
    'Order_Data
    .SetColumnWidth 7, 7, 7
    'Price
    .SetColumnWidth 8, 8, 6
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 1, "DESCRIPTION"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 2, "ID #"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 3, "BRAND"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 4, "SUPPL"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 5, "#/Cs"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 6, "SIZE"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 7, "ORDER"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 8, "PRICE"
Set RS = DB.OpenRecordset("SELECT * " _
  & "FROM " & Categories & " ORDER BY Description")
If RS.RecordCount > 0 Then
    For i = 3 To RS.RecordCount + 2
        WriteAll RS.Fields("Description"), Row, 1
        WriteAll RS.Fields("Product_ID"), Row, 2
        WriteAll RS.Fields("Brand"), Row, 3
        WriteAll RS.Fields("Supplier"), Row, 4
        WriteAll RS.Fields("Units_Per_Case"), Row, 5
        WriteAll RS.Fields("Unit_Size"), Row, 6
        WriteAll RS.Fields("Order_Data"), Row, 7
        WriteAll RS.Fields("Price"), Row, 8
    RS.MoveNext
    If RS.EOF Then Exit For
    Row = Row + 1
    Next
    RS.Close
    Set RS = Nothing
 End If
 '**********************************************
    Case "Transactions"
    'Description
    .SetColumnWidth 1, 1, 26
    'Date
    .SetColumnWidth 2, 2, 10
    'Product_ID
    .SetColumnWidth 3, 3, 10
    'Units_Per_Case
    .SetColumnWidth 4, 4, 5
    'Received
    .SetColumnWidth 5, 5, 5
    'Removed
    .SetColumnWidth 6, 6, 5
    'Updated_Inventory
    .SetColumnWidth 7, 7, 5
    'Transaction_ID
    .SetColumnWidth 8, 8, 8
    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 1, "DESCRIPTION"
    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 2, "DATE"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 3, "ID #"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 4, "#/Cs"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 5, "IN"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 6, "OUT"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 7, "TOT"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 8, "TRAN ID#"
Set RS = DB.OpenRecordset("SELECT * " _
  & "FROM " & Categories & " ORDER BY Description")
If RS.RecordCount <> 0 Then
    For i = 3 To RS.RecordCount + 2
        WriteAll RS.Fields("Description"), Row, 1
        WriteAll RS.Fields("Date"), Row, 2
        WriteAll RS.Fields("Product_ID"), Row, 3
        WriteAll RS.Fields("Units_Per_Case"), Row, 4
        WriteAll RS.Fields("Received"), Row, 5
        WriteAll RS.Fields("Removed"), Row, 6
        WriteAll RS.Fields("Updated_Inventory"), Row, 7
        WriteAll RS.Fields("Transaction_ID"), Row, 8
    RS.MoveNext
    If RS.EOF Then Exit For
    Row = Row + 1
    Next
    RS.Close
    Set RS = Nothing
 End If
'********************************************************************
    Case "Current_Inventory", "Previous_Inventory"
    'Description
    .SetColumnWidth 1, 1, 26
    'Product_ID
    .SetColumnWidth 2, 2, 18
    'Total_Items
    .SetColumnWidth 3, 3, 10
    'Inventory_Date
    .SetColumnWidth 4, 4, 18
    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 1, "DESCRIPTION"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 2, "ID #"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 3, "TOTAL"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 4, "DATE"
Set RS = DB.OpenRecordset("SELECT * " _
  & "FROM " & Categories & " ORDER BY Description")
If RS.RecordCount <> 0 Then
    For i = 3 To RS.RecordCount + 2
        WriteAll RS.Fields("Description"), Row, 1
        WriteAll RS.Fields("Product_ID"), Row, 2
        WriteAll RS.Fields("Total_Items"), Row, 3
        WriteAll RS.Fields("Inventory_Date"), Row, 4
    RS.MoveNext
    If RS.EOF Then Exit For
    Row = Row + 1
    Next
    RS.Close
    Set RS = Nothing
 End If
 '***************************************************************
    Case Else
    'Description
    .SetColumnWidth 1, 1, 26
    'Product_ID
    .SetColumnWidth 2, 2, 14
    'Used or Added
    .SetColumnWidth 3, 3, 5
    'Start_Date
    .SetColumnWidth 4, 4, 14
    'End_Date
    .SetColumnWidth 5, 5, 14
    .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, 1, 1, "DESCRIPTION"
    .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 2, "ID #"
        Select Case Left(Categories, 3)
        Case "Use"
        .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 3, "USED"
        Case "Add"
        .WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 1, 3, "ADDED"
        End Select
    .WriteValue xlsText, xlsFont0, xlsRightAlign, xlsNormal, 1, 4, "START"
    .WriteValue xlsText, xlsFont0, xlsRightAlign, xlsNormal, 1, 5, "END"
Set RS = DB.OpenRecordset("SELECT * " _
  & "FROM " & Categories & " ORDER BY Description")
If RS.RecordCount <> 0 Then
    For i = 3 To RS.RecordCount + 2
        WriteAll RS.Fields("Description"), Row, 1
        WriteAll RS.Fields("Product_ID"), Row, 2
            Select Case Left(Categories, 3)
            Case "Use"
            WriteAll RS.Fields("Used"), Row, 3
            Case "Add"
            WriteAll RS.Fields("Added"), Row, 3
            End Select
        WriteAll RS.Fields("Start_Date"), Row, 4
        WriteAll RS.Fields("End_Date"), Row, 5
    RS.MoveNext
    If RS.EOF Then Exit For
    Row = Row + 1
    Next
    RS.Close
    Set RS = Nothing
 End If
End Select
.CloseFile
End With
Exit Sub
FileError:
MsgBox "Error writing to Excel, " & Err.Description
End Sub

Private Sub CopyRecords()
'alternate method of automating Excel, not used
Dim i As Single
Dim Recs As Integer
Dim nRow As Long
Dim nCol As Long
Dim Fd As Field
If RS.EOF And RS.BOF Then Exit Sub
RS.MoveLast
ReDim TheArray(RS.RecordCount + 1, RS.Fields.Count)
nCol = 0
    For Each Fd In RS.Fields
    TheArray(0, nCol) = Fd.Name
    nCol = nCol + 1
    Next
   
RS.MoveFirst
Recs = RS.RecordCount
For nRow = 1 To RS.RecordCount
    For nCol = 0 To RS.Fields.Count - 1
    TheArray(nRow, nCol) = RS.Fields(nCol).Value
        If IsNull(TheArray(nRow, nCol)) Then
        TheArray(nRow, nCol) = ""
        End If
    Next
RS.MoveNext
Next
Sht.Range(Sht.Cells(SC.nRow, SC.nCol), _
  Sht.Cells(SC.nRow + RS.RecordCount + 1, _
  SC.nCol + RS.Fields.Count)).Value = TheArray
End Sub


Private Sub Categories_Click()
LoadRecords
End Sub



Private Sub DataName_KeyDown(KeyCode As Integer, Shift As Integer)
Dim BoxwHND As Long
Dim r As Long
    If KeyCode = 13 Then
        Const WM_USER = &H400
        Const CB_SHOWDROPDOWN = WM_USER + 15
        DataName.SetFocus
        BoxwHND = GetFocus()
        r = SendMessage(BoxwHND, CB_SHOWDROPDOWN, 0, 0)
        KeyCode = 0
    End If
End Sub

Private Sub DropList()
SendMessageStr Categories.hwnd, CB_SHOWDROPDOWN, True, 0&
End Sub

Private Sub menuCurrent_Click()
CurrentInventory
End Sub

Private Sub menuDelRecord_Click()
DeleteRecord
End Sub




Private Sub menuExit_Click()
End
End Sub

Private Sub menuNew_Click()
AddTransaction
End Sub


Private Sub menuPrevious_Click()
SingleDate.Show 1
If EndDate <> -1 Then
PreviousInventory
End If
End Sub

Private Sub menuPrint_Click()
WriteExcelFile
Load ExcelApp
End Sub

Private Sub menuReceived_Click()
Calendar.Show 1
If StartDate <> -1 And EndDate <> -1 Then
SummarizeReceived
End If
End Sub

Private Sub menuRecord_Click()
AddRecord
End Sub




Private Sub DataName_Click()
Loading = True
ClearBoxes
If Categories.ListCount = 0 Then Exit Sub

Select Case Categories
Case "All_Products", "Recent_Products"
SetBoxes 1
Case "Transactions"
SetBoxes 2
Case "Current_Inventory", "Previous_Inventory"
SetBoxes 3
Case Else
SetBoxes 4
End Select
Set RS = DB.OpenRecordset("SELECT * " _
& "FROM " & Categories & " ORDER BY Description")
With RS
If .RecordCount <> 0 Then
   .Move DataName.ListIndex
    If .Fields("Description") = DataName Then
       Select Case Categories
       Case "All_Products", "Recent_Products"
       Value(1) = .Fields("Product_ID")
       CurrentID = .Fields("Product_ID")
       FieldLabel(1) = "Product ID:"
       Value(2) = .Fields("Brand")
       FieldLabel(2) = "Brand:"
       Value(3) = .Fields("Supplier")
       FieldLabel(3) = "Supplier:"
       Value(4) = .Fields("Units_Per_Case")
       FieldLabel(4) = "Items In Case:"
       Value(5) = .Fields("Unit_Size")
       FieldLabel(5) = "Item Size:"
       Value(6) = .Fields("Order_Data")
       FieldLabel(6) = "Order By:"
       Value(7) = Format(.Fields("Price"), "$##.#0")
       FieldLabel(7) = "Price:"
          For i = 1 To 7
             Value(i).Visible = True
             FieldLabel(i).Visible = True
          Next
          Data.Height = Value(7).Top _
          + (Value(7).Height * 4)
      Case "Transactions"
       Value(1) = .Fields("Date")
       FieldLabel(1) = "Date:"
       Value(2) = .Fields("Product_ID")
       CurrentID = .Fields("Product_ID")
       FieldLabel(2) = "Product ID:"
       Value(3) = .Fields("Units_Per_Case")
       FieldLabel(3) = "Units Per Case:"
       Value(4) = .Fields("Received")
       FieldLabel(4) = "Units Received:"
       Value(5) = .Fields("Removed")
       FieldLabel(5) = "Units Removed:"
       Value(6) = .Fields("Updated_Inventory")
       
       FieldLabel(6) = "Units In Stock:"
       Value(7) = .Fields("Transaction_ID")
       FieldLabel(7) = "Transaction ID:"
          For i = 1 To 7
          Value(i).Visible = True
          FieldLabel(i).Visible = True
          Next
          Data.Height = Value(7).Top _
          + (Value(7).Height * 4)
       Case "Current_Inventory", "Previous_Inventory"
       Value(1) = .Fields("Product_ID")
       CurrentID = .Fields("Product_ID")
       FieldLabel(1) = "Product ID:"
       Value(2) = .Fields("Total_Items")
       FieldLabel(2) = "Total Items:"
       Value(3) = .Fields("Inventory_Date")
       FieldLabel(3) = "Inventory Date:"
          For i = 1 To 3
          Value(i).Visible = True
          FieldLabel(i).Visible = True
          Next
          Data.Height = Value(3).Top _
          + (Value(3).Height * 4)
       Case Else
       Value(1) = .Fields("Product_ID")
       CurrentID = .Fields("Product_ID")
       FieldLabel(1) = "Product ID:"
       Select Case Left(Categories, 3)
       Case "Use"
       Value(2) = .Fields("Used")
       FieldLabel(2) = "Items Used:"
       Case "Add"
       Value(2) = .Fields("Added")
       FieldLabel(2) = "Items Added:"
       End Select
       Value(3) = .Fields("Start_Date")
       FieldLabel(3) = "Start Date:"
       Value(4) = .Fields("End_Date")
       FieldLabel(4) = "End Date:"
          For i = 1 To 4
          Value(i).Visible = True
          FieldLabel(i).Visible = True
          Next
          Data.Height = Value(4).Top _
          + (Value(4).Height * 4)
       End Select
       
    End If
End If
.Close
End With
Loading = False
End Sub


Public Function IsFile(FileString As String) As Boolean
Dim FileNumber As Integer
On Error Resume Next
FileNumber = FreeFile()
Open FileString For Input As #FileNumber
If Err Then
IsFile = False
Exit Function
End If
IsFile = True
Close #FileNumber
End Function
Private Sub Form_Load()
LoadTables
Me.Show
DropList
End Sub











Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
DB.Close
If Network Then UpdateLocal
Set Fld = Nothing
Set RS = Nothing
Set Tbl = Nothing
Set DB = Nothing
Set Ind = Nothing
Unload ExcelApp
End Sub



Private Sub menuUsage_Click()
Calendar.Show 1
If StartDate <> -1 And EndDate <> -1 Then
SummarizeUsage
End If
End Sub

Private Sub mnuCheck_Click()
WriteCheckList
Load ExcelApp
End Sub

Private Sub mnuID_Click()
FindID
End Sub


Private Sub mnuList_Click()
WriteInventory
Load ExcelApp
End Sub

Private Sub Value_Change(Index As Integer)
If Loading Then Exit Sub
Loading = False
Select Case Categories
Case "All_Products", "Recent_Products"
    Select Case Index
    Case 1
    Case 2, 3, 4, 5, 6
    SaveData Index
    Case 7
    If Not IsNumeric(Value(Index)) Then
    Value(Index) = LastValue
    Value(Index).SelStart = Len(Value(Index))
    Else
    SaveData Index
    End If
End Select
Case "Transactions"
    Select Case Index
    Case 1
    Case 2
    Case 3
    Case 4
    SaveData Index
    Case 5
    SaveData Index
    Case 6
    Case 7
    End Select
End Select
LastValue = Value(Index)
End Sub

Private Sub Value_Click(Index As Integer)
Value(Index).SelStart = 0
Value(Index).SelLength = Len(Value(Index))
End Sub

Private Sub Value_DblClick(Index As Integer)
Select Case Categories
Case "All_Products", "Recent_Products"
    Select Case Index
    Case 1
    AddTransaction Value(Index)
    Case Else
    Value(Index).SelLength = 0
    Value(Index).SelStart = Len(Value(Index))
    End Select
Case "Transactions"
    Select Case Index
    Case 2
    AddTransaction Value(Index)
    Case Else
    Value(Index).SelLength = 0
    Value(Index).SelStart = Len(Value(Index))
    End Select
End Select
End Sub

Private Sub Value_GotFocus(Index As Integer)
LastValue = Value(Index)
If Index = 4 Or Index = 5 Then
OldValue = Value(Index)
End If
TheIndex = Index
Value(Index).SelStart = 0
Value(Index).SelLength = Len(Value(Index))
End Sub

Private Sub Value_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
Value(Index).SelStart = 0
Value(Index).SelLength = Len(Value(Index))
End If

End Sub


