VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub blacklistItemsCheck()
Dim i As Integer, j As Integer, ErrorLog As Integer
ErrorLog = 0
Dim bl As Worksheet, bom As Worksheet
Set bl = Worksheets("Blacklist")
Set bom = Worksheets("YourBOM")
Dim col_iacpn As Integer
bom.Select
col_iacpn = InputBox("Please input the column number of IAC P/N in your BOM:", , 1)
Call clearColumn(col_iacpn, bom.Name, , False) 'clear interior color
For j = 2 To BottomRow(4, bl.Name, 2)
    For i = 2 To BottomRow(col_iacpn, bom.Name, 2)
        If Trim(bl.Cells(j, 4) <> "") Then
            If Trim(bl.Cells(j, 4)) = Trim(bom.Cells(i, col_iacpn)) Then
                bom.Cells(i, col_iacpn).Interior.Color = 255
                ErrorLog = ErrorLog + 1
            End If
        End If
    Next i
Next j
If ErrorLog = 0 Then
    MsgBox "Ok, No Blacklist Items Included!"
Else
    MsgBox "" & ErrorLog & " Blacklist Items Found in YourBOM (Marked with RED BG)!"
End If
End Sub
Function BottomRow(Optional col As Integer, Optional sheetName As String = "specialValue", Optional defaultValueIfFalse As String = "DefaultValue", Optional UsedRangeMode As Boolean = False)
Dim x65535 As String
x65535 = Chr(64 + col) & 65535
If sheetName = "specialValue" Then
    sheetName = ActiveSheet.Name
End If
BottomRow = Sheets(sheetName).Range(x65535).End(xlUp).Row
If defaultValueIfFalse <> "DefaultValue" And BottomRow = 1 And Cells(1, col) = "" Then
    BottomRow = Val(defaultValueIfFalse)
End If
If UsedRangeMode = True Then
    BottomRow = Sheets(sheetName).UsedRange.Rows.Count
End If
End Function
Function clearColumn(tarCol As Integer, tarSheetName As String, Optional startRow As Integer = 1, Optional clearData As Boolean = True)
Dim i As Integer
For i = startRow To BottomRow(tarCol, tarSheetName, , True)
    If clearData = True Then
        Sheets(tarSheetName).Cells(i, tarCol) = ""
    End If
    With Sheets(tarSheetName).Cells(i, tarCol).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Next i

End Function


