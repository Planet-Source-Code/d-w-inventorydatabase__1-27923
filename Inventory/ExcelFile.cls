VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Enum ValueTypes
    xlsInteger = 0
    xlsNumber = 1
    xlsText = 2
End Enum

Public Enum CellAlignment
    xlsGeneralAlign = 0
    xlsLeftAlign = 1
    xlsCenterAlign = 2
    xlsRightAlign = 3
    xlsFillCell = 4
    xlsLeftBorder = 8
    xlsRightBorder = 16
    xlsTopBorder = 32
    xlsBottomBorder = 64
    xlsShaded = 128
End Enum

Public Enum CellFont
    'used by rgbAttr2
    'bits 0-5 cell formatting, not font
    'bits 6-7 font number(includes formatting)
    xlsFont0 = 0
    xlsFont1 = 64
    xlsFont2 = 128
    xlsFont3 = 192
End Enum

Public Enum CellHiddenLocked
    'used by rgbAttr1
    'bits 0-5 must be zero
    'bit 6 locked/unlocked
    'bit 7 hidden/not hidden
    xlsNormal = 0
    xlsLocked = 64
    xlsHidden = 128
End Enum
'layout variables
Public Enum MarginTypes
   xlsLeftMargin = 38
   xlsRightMargin = 39
   xlsTopMargin = 40
   xlsBottomMargin = 41
End Enum

Public Enum FontFormatting
   'add enums together: xlsBold + xlsUnderline
   xlsNoFormat = 0
   xlsBold = 1
   xlsItalic = 2
   xlsUnderline = 4
   xlsStrikeout = 8
End Enum

Private Type FONT_RECORD
   opcode As Integer  '49
   length As Integer  '5+len(fontname)
   FontHeight As Integer
   'bit0 bold, bit1 italic, bit2 underline, bit3 strikeout, bit4-7 reserved
   FontAttributes1 As Byte
   FontAttributes2 As Byte  'reserved - always 0
   FontNameLength As Byte
End Type

Private Type HEADER_FOOTER_RECORD
   opcode As Integer  '20 Header, 21 Footer
   length As Integer  '1+len(text)
   TextLength As Byte
End Type

Private Type PROTECT_SPREADSHEET_RECORD
   opcode As Integer  '18
   length As Integer  '2
   Protect As Integer
End Type

Private Type FORMAT_COUNT_RECORD
   opcode As Integer  '1f
   length As Integer '2
   Count As Integer
End Type

Private Type FORMAT_RECORD
   opcode As Integer  '1e
   length As Integer  '1+len(format)
   FormatLenght As Byte 'len(format)
End Type '+ followed by the Format-Picture

Private Type COLWIDTH_RECORD
   opcode As Integer  '36
   length As Integer  '4
   col1 As Byte       'first column
   col2 As Byte       'last column
   ColumnWidth As Integer   '1/256th of a character
End Type

Private Type BEG_FILE_RECORD
  opcode As Integer
  length As Integer
  version As Integer
  ftype As Integer
End Type

Private Type END_FILE_RECORD
  opcode As Integer
  length As Integer
End Type

Private Type PRINT_GRIDLINES_RECORD
  opcode As Integer
  length As Integer
  PrintFlag As Integer
End Type

Private Type tInteger
  opcode As Integer
  length As Integer
  Row As Integer     'unsigned integers
  Col As Integer
  'rgbAttr1 handles whether cell is hidden and/or locked
  rgbAttr1 As Byte
  'rgbAttr2 handles the Font# and Formatting assigned to this cell
  rgbAttr2 As Byte
  'rgbAttr3 handles the Cell Alignment/borders/shading
  rgbAttr3 As Byte
  intValue As Integer  'the actual integer value
End Type

'Number record
Private Type tNumber
  opcode As Integer
  length As Integer
  Row As Integer
  Col As Integer
  rgbAttr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  NumberValue As Double  '8 Bytes
End Type

'text record
Private Type tText
  opcode As Integer
  length As Integer
  Row As Integer
  Col As Integer
  rgbAttr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  TextLength As Byte
End Type

Private Type MARGIN_RECORD_LAYOUT
  opcode As Integer
  length As Integer
  MarginValue As Double  '8 bytes
End Type

Private Type HPAGE_BREAK_RECORD
  opcode As Integer
  length As Integer
  NumPageBreaks As Integer
End Type

Private BEG_FILE_MARKER As BEG_FILE_RECORD
Private END_FILE_MARKER As END_FILE_RECORD
Private HORIZ_PAGE_BREAK As HPAGE_BREAK_RECORD
Private FileNumber As Integer
Private HorizPageBreakRows() As Integer
Private NumHorizPageBreaks As Integer
Private TheRow As Integer
Private TheCol As Integer
Private TheLen As Integer
Private a As Integer
Public Sub SampleCall()
On Error GoTo FileError
Dim myExcelFile As New ExcelFile
Dim FileName As String
Dim d As Date
With myExcelFile
FileName = App.Path & "\test.xls"
.CreateFile FileName

.PrintGridLines = False

.SetMargin xlsTopMargin, 1.5
.SetMargin xlsLeftMargin, 1.5
.SetMargin xlsRightMargin, 1.5
.SetMargin xlsBottomMargin, 1.5

.InsertHorizPageBreak 10
.InsertHorizPageBreak 20

.SetFont "Arial", 10, xlsNoFormat
.SetFont "Arial", 10, xlsBold
.SetFont "Arial", 10, xlsBold + xlsUnderline
.SetFont "Courier", 12, xlsItalic

.SetColumnWidth 1, 5, 18

.SetHeader "This is the header"
.SetFooter "This is the footer"

'Use format #3 - "#,##0"
.WriteValue xlsInteger, xlsFont0, xlsLeftAlign, xlsNormal, 6, 1, 2000, 3

'write a cell with a shaded number with a bottom border
.WriteValue xlsNumber, xlsFont1, xlsRightAlign + xlsBottomBorder + xlsShaded, xlsNormal, 7, 1, 12123.456, 4

'write a normal left aligned string using font2 (bold & underline)
.WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 8, 1, "This is a test string"

'write a locked cell,not be able to be overwritten
'set sheet PROTECTION on before it will take effect!!!
.WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsLocked, 9, 1, "This cell is locked"

'fill a cell with "F"'s
.WriteValue xlsText, xlsFont0, xlsFillCell, xlsNormal, 10, 1, "F"

'write a hidden formulae cell
'Text, Number, Integer values cannot be hidden
.WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsHidden, 11, 1, "If this were a formula it would be hidden!"

'date cell
d = "2001/01/15"
.WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, 15, 1, d, 12

.ProtectSpreadsheet = False

.CloseFile
End With
Exit Sub
FileError:
MsgBox "Error in SampleCall, " & Err.Description
End Sub


Public Sub CreateFile(ByVal FileName As String)
On Error GoTo Write_Error
    FileNumber = FreeFile
    Open FileName For Binary As #FileNumber
    Put #FileNumber, , BEG_FILE_MARKER  'must always be written first
    WriteDefaultFormats
    'create the Horizontal Page Break array
    ReDim HorizPageBreakRows(0)
    NumHorizPageBreaks = 0
Exit Sub
Write_Error:
MsgBox "Error in CreateFile, " & Err.Description
End Sub

Public Sub CloseFile()
On Error GoTo Write_Error
Dim lLoop1 As Long
Dim lLoop2 As Long
Dim lTemp As Long
Dim X As Integer
If FileNumber = 0 Then Exit Sub
'write the horizontal page breaks if necessary
If NumHorizPageBreaks > 0 Then
'the Horizontal Page Break array must be in sorted order.
For lLoop1 = UBound(HorizPageBreakRows) To LBound(HorizPageBreakRows) Step -1
    For lLoop2 = LBound(HorizPageBreakRows) + 1 To lLoop1
        If HorizPageBreakRows(lLoop2 - 1) > HorizPageBreakRows(lLoop2) Then
            lTemp = HorizPageBreakRows(lLoop2 - 1)
            HorizPageBreakRows(lLoop2 - 1) = HorizPageBreakRows(lLoop2)
            HorizPageBreakRows(lLoop2) = lTemp
        End If
    Next lLoop2
Next lLoop1
'write the Horizontal Page Break Record
    With HORIZ_PAGE_BREAK
        .opcode = 27
        .length = 2 + (NumHorizPageBreaks * 2)
        .NumPageBreaks = NumHorizPageBreaks
    End With
    Put #FileNumber, , HORIZ_PAGE_BREAK
    'now write the actual page break values
    For X = 1 To UBound(HorizPageBreakRows)
        Put #FileNumber, , MKI(HorizPageBreakRows(X))
    Next
End If
Put #FileNumber, , END_FILE_MARKER
Close #FileNumber
Exit Sub
Write_Error:
MsgBox "Error in CloseFile,  " & Err.Description
End Sub


Private Sub Class_Initialize()
    With BEG_FILE_MARKER
        .opcode = 9
        .length = 4
        .version = 2
        .ftype = 10
    End With
    With END_FILE_MARKER  'end of file marker
        .opcode = 10
    End With
End Sub


Public Sub InsertHorizPageBreak(TheRow As Long)
On Error GoTo Page_Break_Error

If TheRow > 32767 Then
TheRow = CInt(TheRow - 65536)
Else
TheRow = CInt(TheRow) - 1
End If
        
NumHorizPageBreaks = NumHorizPageBreaks + 1
ReDim Preserve HorizPageBreakRows(NumHorizPageBreaks)
HorizPageBreakRows(NumHorizPageBreaks) = TheRow
Exit Sub
Page_Break_Error:
MsgBox "Error in InsertHorizPageBreak, " & Err.Description
End Sub



Public Sub WriteValue(ValueType As ValueTypes, _
CellFontUsed As CellFont, Alignment As CellAlignment, _
HiddenLocked As CellHiddenLocked, _
TheRow As Long, TheCol As Long, Value As Variant, _
Optional CellFormat As Long = 0)

On Error GoTo Write_Error
Dim St As String
Dim NUMBER_RECORD As tNumber
Dim INTEGER_RECORD As tInteger
Dim TEXT_RECORD As tText
Dim b As Byte
If TheRow > 32767 Then
TheRow = CInt(TheRow - 65536)
Else
TheRow = CInt(TheRow) - 1
End If
If TheCol > 32767 Then
TheCol = CInt(TheCol - 65536)
Else
TheCol = CInt(TheCol) - 1
End If
Select Case ValueType
Case ValueTypes.xlsInteger

With INTEGER_RECORD
.opcode = 2
.length = 9
.Row = TheRow
.Col = TheCol
.rgbAttr1 = CByte(HiddenLocked)
.rgbAttr2 = CByte(CellFontUsed + CellFormat)
.rgbAttr3 = CByte(Alignment)
.intValue = CInt(Value)
End With
Put #FileNumber, , INTEGER_RECORD
    
Case ValueTypes.xlsNumber
With NUMBER_RECORD
.opcode = 3
.length = 15
.Row = TheRow
.Col = TheCol
.rgbAttr1 = CByte(HiddenLocked)
.rgbAttr2 = CByte(CellFontUsed + CellFormat)
.rgbAttr3 = CByte(Alignment)
.NumberValue = CDbl(Value)
End With
Put #FileNumber, , NUMBER_RECORD
      
Case ValueTypes.xlsText
St = CStr(Value)
TheLen = Len(St)
With TEXT_RECORD
.opcode = 4
.length = 10
.TextLength = TheLen
.length = 8 + TheLen
.Row = TheRow
.Col = TheCol
.rgbAttr1 = CByte(HiddenLocked)
.rgbAttr2 = CByte(CellFontUsed + CellFormat)
.rgbAttr3 = CByte(Alignment)
Put #FileNumber, , TEXT_RECORD

   For a = 1 To TheLen
   b = Asc(Mid(St, a, 1))
   Put #FileNumber, , b
   Next
End With
End Select
Exit Sub
Write_Error:
MsgBox "Error number in WriteValue, " & Err.Description
End Sub


Public Sub SetMargin(Margin As MarginTypes, MarginValue As Double)
On Error GoTo Write_Error
    'write the spreadsheet's layout information (in inches)
    Dim MarginRecord As MARGIN_RECORD_LAYOUT
    With MarginRecord
      .opcode = Margin
      .length = 8
      .MarginValue = MarginValue 'in inches
    End With
    Put #FileNumber, , MarginRecord
Exit Sub
Write_Error:
MsgBox "Error in SetMargin, " & Err.Description
End Sub


Public Sub SetColumnWidth(FirstColumn As Byte, LastColumn As Byte, WidthValue As Integer)
On Error GoTo Write_Error
Dim COLWIDTH As COLWIDTH_RECORD
With COLWIDTH
.opcode = 36
.length = 4
.col1 = FirstColumn - 1
.col2 = LastColumn - 1
.ColumnWidth = WidthValue * 256  'values are specified as 1/256 of a character
End With
Put #FileNumber, , COLWIDTH
Exit Sub
Write_Error:
MsgBox "Error in SetColumnWidth, " & Err.Description
End Sub


Public Sub SetFont(FontName As String, FontHeight As Integer, FontFormat As FontFormatting)
On Error GoTo Write_Error
Dim b As Byte
Dim FONTNAME_RECORD As FONT_RECORD
TheLen = Len(FontName)
With FONTNAME_RECORD
.opcode = 49
.length = 5 + TheLen
.FontHeight = FontHeight * 20
.FontAttributes1 = CByte(FontFormat) 'bold/underline etc...
.FontAttributes2 = CByte(0) 'reserved-always zero!!
.FontNameLength = CByte(Len(FontName))
End With
Put #FileNumber, , FONTNAME_RECORD
    For a = 1 To TheLen
    b = Asc(Mid(FontName, a, 1))
    Put #FileNumber, , b
    Next
Exit Sub
Write_Error:
MsgBox "Error in SetFont, " & Err.Description
End Sub


Public Sub SetHeader(HeaderText As String)
On Error GoTo Write_Error
Dim b As Byte
Dim HEADER_RECORD As HEADER_FOOTER_RECORD
TheLen = Len(HeaderText)
With HEADER_RECORD
.opcode = 20
.length = 1 + TheLen
.TextLength = CByte(Len(HeaderText))
End With
Put #FileNumber, , HEADER_RECORD
  'Then the actual Header text
   For a = 1 To TheLen
   b = Asc(Mid(HeaderText, a, 1))
   Put #FileNumber, , b
   Next
Exit Sub
Write_Error:
MsgBox "Error in SetHeader, " & Err.Description
End Sub



Public Sub SetFooter(FooterText As String)

On Error GoTo Write_Error
Dim b As Byte
Dim FOOTER_RECORD As HEADER_FOOTER_RECORD
TheLen = Len(FooterText)
With FOOTER_RECORD
.opcode = 21
.length = 1 + TheLen
.TextLength = CByte(Len(FooterText))
End With
Put #FileNumber, , FOOTER_RECORD
   For a = 1 To TheLen
   b = Asc(Mid(FooterText, a, 1))
   Put #FileNumber, , b
   Next
Exit Sub
Write_Error:
MsgBox "Error in SetFooter, " & Err.Description
End Sub



Public Property Let PrintGridLines(ByVal newvalue As Boolean)

On Error GoTo Write_Error

    Dim GRIDLINES_RECORD As PRINT_GRIDLINES_RECORD
    
    With GRIDLINES_RECORD
      .opcode = 43
      .length = 2
      If newvalue = True Then
        .PrintFlag = 1
      Else
        .PrintFlag = 0
      End If
      
    End With
    Put #FileNumber, , GRIDLINES_RECORD

Exit Property
Write_Error:
MsgBox "Error in PrintGridLines, " & Err.Description
End Property




Public Property Let ProtectSpreadsheet(ByVal newvalue As Boolean)
On Error GoTo Write_Error
Dim PROTECT_RECORD As PROTECT_SPREADSHEET_RECORD
With PROTECT_RECORD
.opcode = 18
.length = 2
If newvalue = True Then
.Protect = 1
Else
.Protect = 0
End If
End With
Put #FileNumber, , PROTECT_RECORD
Exit Property
Write_Error:
MsgBox "Error in ProtectSpreadsheet, " & Err.Description
End Property


Public Sub WriteDefaultFormats()
On Error GoTo Format_Error
Dim cFORMAT_COUNT_RECORD As FORMAT_COUNT_RECORD
Dim cFORMAT_RECORD As FORMAT_RECORD
Dim lIndex As Long
Dim iFormat(0 To 23) As String
Dim TheLen As Long
Dim Quote As String
Dim b As Byte
Dim a As Long
Quote = Chr(34)
iFormat(0) = "General"
iFormat(1) = "0"
iFormat(2) = "0.00"
iFormat(3) = "#,##0"
iFormat(4) = "#,##0.00"
iFormat(5) = "#,##0\ " & Quote & "$" & Quote & ";\-#,##0\ " & Quote & "$" & Quote
iFormat(6) = "#,##0\ " & Quote & "$" & Quote & ";[Red]\-#,##0\ " & Quote & "$" & Quote
iFormat(7) = "#,##0.00\ " & Quote & "$" & Quote & ";\-#,##0.00\ " & Quote & "$" & Quote
iFormat(8) = "#,##0.00\ " & Quote & "$" & Quote & ";[Red]\-#,##0.00\ " & Quote & "$" & Quote
iFormat(9) = "0%"
iFormat(10) = "0.00%"
iFormat(11) = "0.00E+00"
iFormat(12) = "dd/mm/yy"
iFormat(13) = "dd/\ mmm\ yy"
iFormat(14) = "dd/\ mmm"
iFormat(15) = "mmm\ yy"
iFormat(16) = "h:mm\ AM/PM"
iFormat(17) = "h:mm:ss\ AM/PM"
iFormat(18) = "hh:mm"
iFormat(19) = "hh:mm:ss"
iFormat(20) = "dd/mm/yy\ hh:mm"
iFormat(21) = "##0.0E+0"
iFormat(22) = "mm:ss"
iFormat(23) = "@"
With cFORMAT_COUNT_RECORD
.opcode = &H1F
.length = &H2
.Count = CInt(UBound(iFormat))
End With
Put #FileNumber, , cFORMAT_COUNT_RECORD
For lIndex = LBound(iFormat) To UBound(iFormat)
TheLen = Len(iFormat(lIndex))
With cFORMAT_RECORD
.opcode = &H1E
.length = CInt(TheLen + 1)
.FormatLenght = CInt(TheLen)
End With
Put #FileNumber, , cFORMAT_RECORD

   For a = 1 To TheLen
   b = Asc(Mid(iFormat(lIndex), a, 1))
   Put #FileNumber, , b
   Next

Next
Exit Sub
Format_Error:
MsgBox "Error in WriteDefaultFormats, " & Err.Description
End Sub


Private Function MKI(X As Integer) As String
On Error GoTo Function_Error
Dim temp As String
'used for writing integer array values to file
temp = Space(2)
CopyMemory ByVal temp, X, 2
MKI = temp
Exit Function
Function_Error:
MsgBox "Error in MKI Function, " & Err.Description
End Function



