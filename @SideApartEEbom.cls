VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub delNotFound()
Dim sb As Worksheet
Set sb = Sheets("SideBOM")
sb.Select
Application.ScreenUpdating = False
Dim i As Integer
For i = BottomRow(8, sb.Name, 2) To 2 Step -1
    If sb.Cells(i, 6) = "" Then
        Rows(i).Delete
    End If
Next i
Application.ScreenUpdating = True
End Sub
Sub SideApartEEbom()
Dim ee As Worksheet, loc As Worksheet, otp As Worksheet, Desk As Worksheet
Set ee = Sheets("BOM")
Set loc = Sheets("SideLocation")
Set otp = Sheets("SideBOM")
Set Desk = Sheets("Desktop")
Desk.Unprotect Password:="ac"
'[ee fomat check]--------------------------------------------------------------------
If MsgBox("Please make sure BOM is Right... Continue?", vbYesNo, "Columns' Title with * MUST Have Data") = vbNo Then
    GoTo theend
End If
Application.ScreenUpdating = False
'clear old data-->otp
otp.Select
Cells.Select
Selection.Clear
'cpy normal title
Desk.Select
Rows(1).Select 'cpy title-->ee,otp
Selection.Copy
ee.Select
Cells(1, 1).Select
ActiveSheet.Paste
'_________
otp.Select
Cells(1, 1).Select
ActiveSheet.Paste

'[1][Trim Data]-------------------------------------------------------------------------
Call TrimData(6, 9, ee.Name)
'[2][Repeat Item check]----------------------------------------------------------------
If MarkRepeatItemWithinColumn(6, ee.Name) = True Then
    MsgBox "[BOM ERROR] Column [IAC P/N*] Contains Repeat IAC PN!(Marked With RED Background)"
    GoTo theend
End If
'[3][*col blank cell check]----------------------------------------------------------------
Call clearColumn(6, ee.Name, 2, False) 'clear BG color
Call clearColumn(8, ee.Name, 2, False) 'clear BG color
Call clearColumn(9, ee.Name, 2, False) 'clear BG color
Dim bot1 As Integer, bot2 As Integer, bot3 As Integer, maxBot As Integer
bot1 = BottomRow(6, ee.Name)
'bot2 = BottomRow(8, ee.Name)
bot3 = BottomRow(9, ee.Name)
maxBot = WorksheetFunction.Max(bot1, bot3)

If bot1 <> maxBot Or bot3 <> maxBot Then
   Application.ScreenUpdating = True
   ee.Select
   If bot1 <> maxBot Then
        Cells(maxBot, 6).Select
    ElseIf bot3 <> maxbox Then
        Cells(maxBot, 9).Select
    End If
    Call MsgBox("[Warning] Last cell cannot be empty!", , "Sheet[BOM]")
    GoTo theend
Else
    '1
    If Counter("NULL", 8, , ee.Name) > 0 Then
        If MsgBox("[Warning] Cell Value of column [Location/Reference*] cannot be empty! Auto-Fill(if column Location is the style of FULLBOM)?", vbYesNo) = vbYes Then
            For i = 2 To BottomRow(6, ee.Name, 2) 'bot based on IACPN
                If ee.Cells(i, 8) = "" Then
                    ee.Cells(i, 8) = ee.Cells(i - 1, 8)
                    ee.Cells(i, 8).Interior.Color = 15773696 'bright blue
                End If
            Next i
        Else
            GoTo theend
        End If
    End If
    '2
    If Counter("NULL", 6, , ee.Name) > 0 Then
        ee.Select
        Columns(6).Select
        Call MsgBox("[Warning] Cell Value of column [IAC P/N*] cannot be empty!", , "Sheet[BOM]")
        GoTo theend
    End If
    '3
    If Counter("NULL", 9, , ee.Name) > 0 Then
        ee.Select
        Columns(9).Select
        Call MsgBox("[Warning] Cell Value of column [IAC Description] cannot be empty!", , "Sheet[BOM]")
    End If
    Application.ScreenUpdating = False
End If
'-----------------------------------------------------[main]--------------------------------------------------------------------
'[seperate locations]--------
Call clearColumn(1, Desk.Name, 2) '¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X¡X
For j = 2 To BottomRow(1, loc.Name, 2)
    Pbot = BottomRow(1, Desk.Name, 2) + 1
    Call LocationsSeparate(loc.Cells(j, 1), Desk.Cells(Pbot, 1), Desk.Name) 'save to default sheet "LocationsSeparate"
Next j
loc.Select
Cells.Select
Selection.Clear
Call copyColumn(1, 1, Desk.Name, loc.Name)
'[side locations repetence check]__________________________
If MarkRepeatItemWithinColumn(1, loc.Name) = True Then
    MsgBox "[ERROR!] SideLocations  Contains Repeat Item(s)(Marked With RED Background)!"
    GoTo theend
End If
'__________________________________________
'loc.Cells(1, 1) = "SideLocation"
'loc.Cells(1, 2) = "[temp]FoundInRow"
'[save found row num]-----------------------------------
Dim teamNum As Integer
For i = 2 To BottomRow(1, loc.Name)
    teamNum = 0
    For j = 2 To BottomRow(8, ee.Name)
        If InStr("," & ee.Cells(j, 8) & ",", "," & loc.Cells(i, 1) & ",") Then
            loc.Cells(i, 3 + teamNum) = j
            teamNum = teamNum + 1
        End If
    Next j
Next i
'----
For i = 2 To BottomRow(1, loc.Name, 2)
    For j = 3 To BottomCol(Val(i), loc.Name)
        loc.Cells(i, 2) = loc.Cells(i, 2) & loc.Cells(i, j) & ","
    Next j
Next i
'[merge location]-------------------------------------------------------------
Dim TBot As Integer
TBot = BottomRow(1, loc.Name, 2)
For i = 2 To TBot - 1
    For j = i + 1 To TBot
        If loc.Cells(i, 2) <> "" Then
            If loc.Cells(i, 2) = loc.Cells(j, 2) Then
                loc.Cells(i, 1) = loc.Cells(i, 1) & "," & loc.Cells(j, 1) 'merge locations
                loc.Rows(j).Clear
            End If
        Else 'mark
            loc.Cells(i, 2).Interior.Color = 49407 'orange
        End If
    Next j
Next i
'[DEL BLANK ROWS]----------
For i = BottomRow(1, loc.Name, 2) To 2 Step -1
    If loc.Cells(i, 1) = "" Then
        loc.Rows(i).Delete
    End If
Next i
'[otp eeBOM]--------------------------------------------------------------------
Dim r As Integer
r = 2
For i = 2 To BottomRow(1, loc.Name, 2)
    otp.Cells(r, 8) = loc.Cells(i, 1) 'cpy location
    If loc.Cells(i, 2) <> "" Then
        For j = 3 To BottomCol(Val(i), loc.Name)
            otp.Cells(r, 2) = ee.Cells(loc.Cells(i, j), 2) 'free colum--data not defined
            otp.Cells(r, 3) = ee.Cells(loc.Cells(i, j), 3)
            otp.Cells(r, 4) = ee.Cells(loc.Cells(i, j), 4)
            'IAC PN & DES
            otp.Cells(r, 6) = ee.Cells(loc.Cells(i, j), 6)
            otp.Cells(r, 9) = ee.Cells(loc.Cells(i, j), 9)
            r = r + 1
        Next j
        
    Else
        otp.Cells(r, 8).Interior.Color = 49407 'orange
        r = r + 1
    End If
Next i
'[bom check]
If MarkRepeatItemWithinColumn(6, otp.Name) = True Then
    Application.ScreenUpdating = True
    otp.Select
    Columns(6).Select
    Call MsgBox("[ERROR!] Column Component* Contains Repeat Item(s)(Marked With RED Background)!", , "[outputEE]")
    GoTo theend
End If
'[Data Generate]_____________________________________________________________
'[1][Item No Generate]
Call ItemNoGenerate(8, otp.Name)
'[2][Priority Generate]
Call PriorityGenerate_F(5, 8, otp.Name)
'[3][Qty Generate]
Call QtyGenerate(7, 8, otp.Name)
For i = 2 To BottomRow(7, otp.Name)
    If otp.Cells(i, 7) = "" Then
        otp.Cells(i, 7) = otp.Cells(i - 1, 7)
    End If
Next i
'[clear footprint of loc.]---------------------------
'side location
loc.Select
Cells.Select
Selection.Clear
loc.Cells(1, 1) = "SideLocations"
Call copyColumn(1, 1, Desk.Name, loc.Name)
'desk
Call clearColumn(1, Desk.Name, 2)
'_______________________________________________________________________________
otp.Select
Dim t As Integer
t = Counter("NULL", 6, , otp.Name)
If BottomRow(6, otp.Name) < BottomRow(8, otp.Name) Then
   t = t + BottomRow(8, otp.Name) - BottomRow(6, otp.Name)
End If
If t > 0 Then
    MsgBox "[Warning]" & t & " Item(s) cannot find IAC PN(Marked with brown BG)!"
Else
    MsgBox "Done!"
End If
theend:
Application.ScreenUpdating = True
Desk.Protect Password:="ac"
End Sub
Function TrimData(srcCol As Integer, desCol As Integer, sheetName As String)
'
' TrimData ¥¨¶°
' Trim Data from worksheet column i to j
'
'[Trim data of column"srcCol" to column"desCol" -on sheet1]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'srcCol = 1     '''''''''''''''''''''''''''''''''''change these value for particular using
'desCol = 8     '''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim col As Integer, row As Integer
    Dim Bottom As Integer
    Dim strCol As String, actSheet As String
    TrimData = 0
    For col = srcCol To desCol Step 1
        strCol = chr(64 + col) & 65535
        Bottom = Sheets(sheetName).Range(strCol).End(xlUp).row
        '__________________________________________________________
        For row = 1 To Bottom Step 1
           ' If Left(Cells(row, col).value, 1) = " " Or Right(Cells(row, col).value, 1) = " " Then
               Sheets(sheetName).Cells(row, col) = Trim(Sheets(sheetName).Cells(row, col))
               TrimData = TrimData + 1
            'End If
        Next row
    Next col
End Function
Function MarkRepeatItemWithinColumn(col As Integer, sheetName As String)
'mark repeat items within one column with red BG;return a flag of repeat item exist/not
'Col:the column to deal with
'sheetName:default value is active sheet; point to particular sheet to deal with
Dim i As Integer, j As Integer
Dim bot As Integer, x65535 As String, activeSheetName As String
'_______________________________________________________________________________'
x65535 = chr(64 + col) & 65535
bot = Sheets(sheetName).Range(x65535).End(xlUp).row
MarkRepeatItemWithinColumn = False
For i = 2 To bot - 1
    For j = i + 1 To bot
       If Sheets(sheetName).Cells(i, col) <> "" Then
            If Sheets(sheetName).Cells(i, col) = Sheets(sheetName).Cells(j, col) Then
                Sheets(sheetName).Cells(j, col).Interior.Color = 255
                Sheets(sheetName).Cells(i, col).Interior.Color = 255
                MarkRepeatItemWithinColumn = True
            End If
        End If
    Next j
Next i
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
Function BottomRow(Optional col As Integer, Optional sheetName As String = "specialValue", Optional defaultValueIfFalse As String = "DefaultValue", Optional UsedRangeMode As Boolean = False)
Dim x65535 As String
x65535 = chr(64 + col) & 65535
If sheetName = "specialValue" Then
    sheetName = ActiveSheet.Name
End If
BottomRow = Sheets(sheetName).Range(x65535).End(xlUp).row
If defaultValueIfFalse <> "DefaultValue" And BottomRow = 1 And Cells(1, col) = "" Then
    BottomRow = Val(defaultValueIfFalse)
End If
If UsedRangeMode = True Then
    BottomRow = Sheets(sheetName).UsedRange.Rows.Count
End If
End Function
Function Counter(Index As String, targetCol As Integer, Optional startRow As Integer = 2, Optional targetSheetName As String = "specialValue", Optional TheRestNum As Boolean = False) As Integer
Dim i As Integer, botRow As Integer
Counter = 0
botRow = BottomRow(targetCol, targetSheetName)
If targetSheetName = "specialValue" Then
    targetSheetName = ActiveSheet.Name
End If
If UCase(Index) = "NULL" Then
    For i = startRow To botRow
        If Sheets(targetSheetName).Cells(i, targetCol) = "" Then
            Counter = Counter + 1
        End If
    Next i
Else
    For i = startRow To botRow
        If Sheets(targetSheetName).Cells(i, targetCol) = Index Then
            Counter = Counter + 1
        End If
    Next i
End If
If TheRestNum = True Then
    Counter = botRow - startRow - Counter + 1
End If
End Function
'==================================================================================================
Function LocationsSeparate(ByVal targetCell As String, ByVal savHeadCell As Range, Optional savSheet As String = "LocationsSeparate")
'[Core algorithm]:Separate data from string (locations in one cell) by checking the commas
'[Whwere used]:Checking locations in a column or so if they are unique;
'The separated data will saved to column 1 of created sheet "LocationsSeparate"(will creat it if not exist).
'____________________col_____________________________________
'|'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'row'''''''''''''''''A,B,C,D,E ''''''''''''''''''''''''''''''
'|'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'|'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'____________________"LocationsSeparate"_____________________
'|'''''''''''''''''''A         ''''''''''''''''''''''''''''''
'|'''''''''''''''''''B         ''''''''''''''''''''''''''''''
'|'''''''''''''''''''C         ''''''''''''''''''''''''''''''
'|'''''''''''''''''''D         ''''''''''''''''''''''''''''''
'|'''''''''''''''''''E         ''''''''''''''''''''''''''''''
'____________________________________________________________
  Dim src As Integer, des As Integer
  Dim str As String, chr As String
  Dim savRow As Integer, savCol As Integer
  Dim i As Integer, j As Integer
    '_________________________________________________________________
  'Call existenceOfSheetsCheck(savSheet, True)
  'Sheets(savSheet).Select  'saving-sheet active
  'savSheet = ActiveSheet.Name->NO NEED!
  savRow = savHeadCell.row 'to concreat parameters of this function
  savCol = savHeadCell.Column
  '__________________________________________________________________
  str = targetCell & ","
  src = 0
  des = 0
  j = 0
  For i = 1 To Len(str) Step 1
    chr = Mid(str, i, 1)
    If (Asc(chr) = 44) Then
       des = i 'Renew des
      'Save separated data
       Sheets(savSheet).Cells(savRow + j, savCol) = Mid(str, src + 1, des - src - 1)
       src = i 'Renew src
       j = j + 1
    End If
  Next
End Function
Function copyColumn(srcCol As Integer, tarCol As Integer, Optional srcSheetName As String = "specialValue", Optional desSheetName As String = "specialValue", Optional srcStartRow As Integer = 2, Optional desColStartRow As Integer = 2, Optional includeBlankCells As Boolean = False)
'Default mode: copy the data  of the srcCol that not null from 2nd row to  bottom to  tarCol from 2nd row to bottom
Dim tempBot As Integer, i As Integer

If srcSheetName = "specialValue" Then
    srcSheetName = ActiveSheet.Name
End If
If desSheetName = "specialValue" Then
    desSheetName = ActiveSheet.Name
End If
tempBot = desColStartRow
'__________________________________________________
If includeBlankCells = False Then
    For i = srcStartRow To BottomRow(srcCol, srcSheetName, Val(srcStartRow))
            If Sheets(srcSheetName).Cells(i, srcCol) <> "" Then
                Sheets(desSheetName).Cells(tempBot, tarCol) = Sheets(srcSheetName).Cells(i, srcCol)
                tempBot = tempBot + 1
            End If
    Next i
Else
    For i = srcStartRow To BottomRow(srcCol, srcSheetName, Val(srcStartRow))
        Sheets(desSheetName).Cells(tempBot, tarCol) = Sheets(srcSheetName).Cells(i, srcCol)
        tempBot = tempBot + 1
    Next i
End If
End Function
'===============================================================
Function ItemNoGenerate(col_Location As Integer, tarSheetName As String)
'DES:Creat Item No basing on Col-Location. Exp:EE BOM
'Item No defalut location: col 1
Dim i As Integer, tarSheet As Worksheet
Set tarSheet = Sheets(tarSheetName)
Call clearColumn(1, tarSheet.Name, 2)
tarSheet.Cells(2, 1) = 10
For i = 3 To realBottomRow(8, tarSheet.Name)
    If Trim(tarSheet.Cells(i, col_Location)) <> "" Then
        tarSheet.Cells(i, 1) = tarSheet.Cells(i - 1, 1) + 10
    Else
         tarSheet.Cells(i, 1) = tarSheet.Cells(i - 1, 1)
    End If
Next i
For i = 2 To realBottomRow(8, tarSheet.Name)
    If tarSheet.Cells(i, 1) < 100 Then
        tarSheet.Cells(i, 1) = "'" & "00" & tarSheet.Cells(i, 1)
    ElseIf tarSheet.Cells(i, 1) < 1000 Then
        tarSheet.Cells(i, 1) = "'" & "0" & tarSheet.Cells(i, 1)
    Else
        tarSheet.Cells(i, 1) = "'" & tarSheet.Cells(i, 1)
    End If
Next i
'tarSheet.Cells(1, 1) = "Item No"
End Function
Function PriorityGenerate_F(col_Pri As Integer, col_Location As Integer, sheetName As String, Optional methord As Integer = 1)
'Des: basing on item_No (col 1)
Dim i As Integer
Dim tarSheet As Worksheet
Set tarSheet = Sheets(sheetName)
If methord = 1 Then
    For i = 2 To realBottomRow(8, sheetName)
        If tarSheet.Cells(i, 1) <> tarSheet.Cells(i - 1, 1) And tarSheet.Cells(i, 1) = tarSheet.Cells(i + 1, 1) Then
            tarSheet.Cells(i, col_Pri) = "'1"
        Else
            tarSheet.Cells(i, col_Pri) = ""
        End If
    Next i
'[2]Base on REFDESIG
Else
    For i = 2 To realBottomRow(5, ActiveSheet.Name)
        If tarSheet.Cells(i, 8) <> "" And tarSheet.Cells(i + 1, 8) = "" Then
            If Val(tarSheet.Cells(i, col_Pri)) <> 1 Then
                tarSheet.Cells(i, col_Pri).Interior.Color = 255
            End If
        Else
            If tarSheet.Cells(i, col_Pri) <> "" Then
                tarSheet.Cells(i, col_Pri).Interior.Color = 255
            End If
        End If
    Next i
End If
End Function
Function QtyGenerate(colQty As Integer, colRef As Integer, tarSheetName As String)
Dim i As Integer, t As Integer, str As String, Qty As Integer
Dim tarSheet As Worksheet
Set tarSheet = Sheets(tarSheetName)
'[qty Generate] & [Qty check]
For i = 2 To BottomRow(colRef, tarSheet.Name)
    Qty = 0
    If Trim(tarSheet.Cells(i, colRef)) <> "" Then
        str = tarSheet.Cells(i, colRef) & ","
        For t = 1 To Len(str)
            If Asc(Mid(str, t, 1)) = 44 Then
                Qty = Qty + 1
            End If
        Next t
    End If
    'MsgBox "" & qty & ""
    If Qty <> tarSheet.Cells(i, colQty) Then
        If Qty = 0 Then
            tarSheet.Cells(i, colQty).Clear
        Else
            tarSheet.Cells(i, colQty) = Qty
            tarSheet.Cells(i, colQty).Interior.Color = 15773696 'bright blue
        End If
    End If
    '[format normalize]
    tarSheet.Cells(i, colQty) = "'" & tarSheet.Cells(i, colQty)
Next i
End Function
Function BottomCol(Optional row As Integer = 1, Optional sheetName As String = "specialValue", Optional UsedRangeMode As Boolean = False)
If sheetName = "specialValue" Then
    sheetName = ActiveSheet.Name
End If
If UsedRangeMode = False Then
    Dim IVx As String
    IVx = "IV" & row
    BottomCol = Sheets(sheetName).Range(IVx).End(xlToLeft).Column
Else
    BottomCol = Sheets(sheetName).UsedRange.Columns.Count
End If
End Function
Function realBottomRow(firstXCols As Integer, sheetName As String) As Integer
Dim x65535 As String, i As Integer
realBottomRow = 1
For i = 1 To firstXCols
    x65535 = chr(64 + i) & 65535
    realBottomRow = WorksheetFunction.Max(Sheets(sheetName).Range(x65535).End(xlUp).row, realBottomRow)
Next i
'ActiveSheet.UsedRange.Rows.Count
End Function



