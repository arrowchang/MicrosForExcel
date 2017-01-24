VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub aotoFillup1main2loc()
'designed for 1col=mainCol 2col=location and need to fill the blank cell
Dim a As Worksheet, b As Worksheet
Set a = Sheets("SheetA")
Set b = Sheets("SheetB")
Dim i As Integer
For i = 2 To BottomRow(2, a.Name, 2) '+ 1
    If a.Cells(i, 2) = "" Then
        a.Cells(i, 2) = a.Cells(i - 1, 2)
    End If
Next i

For i = 2 To BottomRow(2, b.Name, 2) '+ 1
    If b.Cells(i, 2) = "" Then
        b.Cells(i, 2) = b.Cells(i - 1, 2)
    End If
Next i

End Sub
Sub A_B()
Dim a As Worksheet, b As Worksheet
Set a = Sheets("DifferA")
Set b = Sheets("DifferB")
Dim i As Integer, j As Integer
For i = 2 To BottomRow(1, a.Name)
    For j = 2 To BottomRow(1, b.Name)
        If a.Cells(i, 1) = b.Cells(j, 1) Then
            a.Cells(i, 3) = Val(a.Cells(i, 2)) - Val(b.Cells(j, 2))
            b.Cells(j, 1) = ""
            b.Cells(j, 2) = ""
            Exit For
        End If
    Next j
Next i
End Sub
Sub tt()
Dim i As Long
temp = inColumn(1, 1, ActiveSheet.Name)
MsgBox "" & temp & ""
End Sub
Sub t()
'MsgBox "" & Cells(9, 8).Font.ThemeColor & ""
Dim a() As String
ReDim a(10) As String
a = Split(Cells(2, 1), ",")
Call sortarray(a)
For i = 0 To UBound(a)
MsgBox "" & a(i) & ""
Next i
End Sub
Function sortarray(ByRef a() As String)
Dim i As Long, j As Long, bot As Long, t As String
bot = UBound(a)
For i = 0 To bot - 1
    For j = i + 1 To bot
        If StrComp(a(i), a(j), vbTextCompare) = 1 Then  'reverse
            t = a(i)
            a(i) = a(j)
            a(j) = t
        End If
    Next j
Next i
End Function
Sub analyzeDifference()
Dim a As Worksheet, b As Worksheet, d1 As Worksheet, d2 As Worksheet
Set a = Sheets("SheetA")
Set b = Sheets("SheetB")
Set d1 = Sheets("DifferA")
Set d2 = Sheets("DifferB")
Dim pnCol As Integer, locCol As Integer, cmtCol As Integer, i As Integer
pnCol = 1
locCol = 2
cmtCol = 3
Dim del As Integer, add As Integer, aa As Integer, bb As Integer

For i = 2 To BottomRow(Val(pnCol), d1.Name, 2)
    aa = Counter(d1.Cells(i, locCol), 2, 2, a.Name)
    bb = Counter(d1.Cells(i, locCol), 2, 2, b.Name)
    del = Counter(d1.Cells(i, locCol), 2, 2, d1.Name)
    add = Counter(d1.Cells(i, locCol), 2, 2, d2.Name)
    
    If del = aa Then 'the group on a should del
        If bb = 0 Then ' b has no these loc
            d1.Cells(i, cmtCol) = "DEL Parts"
        ElseIf add >= del Then
                d1.Cells(i, cmtCol) = "ReplaceAll(& +" & add - del & ")"
        End If
    ElseIf del < aa Then
        If add = 0 Then
            d1.Cells(i, cmtCol) = "DEL 2nd"
        Else
            If add = del Then
                d1.Cells(i, cmtCol) = "Replace"
            ElseIf add > del Then
                d1.Cells(i, cmtCol) = "Del 2nd(& add " & add & ")"
            Else
                d1.Cells(i, cmtCol) = "Replace/Del 2nd"
            End If
        End If
    End If
Next i
'+++++++++++++++++++++++++++++
For i = 2 To BottomRow(Val(pnCol), d2.Name, 2)
    aa = Counter(d2.Cells(i, locCol), 2, 2, a.Name)
    bb = Counter(d2.Cells(i, locCol), 2, 2, b.Name)
    del = Counter(d2.Cells(i, locCol), 2, 2, d1.Name)
    add = Counter(d2.Cells(i, locCol), 2, 2, d2.Name)
    
    If add = bb Then
        If aa = 0 Then
            d2.Cells(i, cmtCol) = "ADD Parts"
        ElseIf add >= del Then
                d2.Cells(i, cmtCol) = "ReplaceAll(& +" & add - del & ")"
        End If
    ElseIf add < bb Then
        If del = 0 Then
            d2.Cells(i, cmtCol) = "ADD 2nd"
        Else
            If add = del Then
                d2.Cells(i, cmtCol) = "Replace"
            ElseIf del > add Then
                d2.Cells(i, cmtCol) = "Add 2nd(+ " & add & ")"
            Else
                d2.Cells(i, cmtCol) = "Replace " & del & " &+" & add - del & " src"
            End If
        End If
    End If
Next i

End Sub
Function Counter(Index As String, targetCol As Integer, Optional startRow As Integer = 2, Optional targetSheetName As String = "specialValue", Optional TheRestNum As Boolean = False) As Integer
Dim i As Integer, botRow As Integer
Counter = 0
botRow = Sheets(targetSheetName).Range(chr(64 + targetCol) & 65535).End(xlUp).row
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


Sub whatIsDifferentBetween2BOMs() '*if include repeat item(s), the comparing result is index order dependent
Dim i As Long, j As Long, k As Long
'___________________________________________________________
Dim a As Worksheet, b As Worksheet, t As Worksheet
Set a = Sheets("SheetA")
Set b = Sheets("SheetB")
If MsgBox("[Tips:]Data Formats & Order should be the same! Continue?", vbYesNo) = vbNo Then
    GoTo endFunction
End If
a.Select 'easy judge
Application.ScreenUpdating = False
Call clearAllFormats(a)
Call clearAllFormats(b)
'___________________________________________________________
Dim botCol(2) As Long, mainCol As Long, locationsCol As Long, tempStr1 As Boolean, tempStr2 As Boolean

mainCol = InputBox("Please type in the Index Column No. ", , 6)
locationsCol = InputBox("Please type in the Locations Column No.", , 8)
botCol(0) = realBottomCol(5, a.Name)
botCol(1) = realBottomCol(5, b.Name)
'____________________________________________________________
'[Used Range check(check if {row bottom,maincol} is empty)]
If BottomRow(mainCol, a.Name) < realBottomRow(Val(botCol(0)), a.Name) Then
   MsgBox "[Sheet " & a.Name & " ]: Please add items of column " & mainCol & " or to delete the redundant data in the bottom row(s)."
    GoTo endFunction
End If
If BottomRow(mainCol, b.Name) < realBottomRow(Val(botCol(1)), b.Name) Then
   MsgBox "[Sheet " & b.Name & " ]: Please add items of column " & mainCol & " or to delete the redundant data in the bottom row(s)."
   GoTo endFunction
End If
'_____________________________________________________________
'[Trim data]
Call TrimData(1, botCol(0), a.Name)
Call TrimData(1, botCol(1), b.Name)
'===================================
'[deal with mainCol=locationsCol]
Dim arr() As String
If mainCol = locationsCol Then
    For i = 2 To BottomRow(mainCol, a.Name, 2)
        'ReDim arr(Len(a.Cells(i, mainCol)) / 3 + 2) As String
        If a.Cells(i, mainCol) <> "" Then
            arr = Split(a.Cells(i, mainCol), ",")
            Call sortarray(arr)
            a.Cells(i, mainCol) = arr(0)
            For j = 1 To UBound(arr)
                a.Cells(i, mainCol) = a.Cells(i, mainCol) & "," & arr(j)
            Next j
        End If
    Next i
    '--
    For i = 2 To BottomRow(mainCol, b.Name, 2)
        If a.Cells(i, mainCol) <> "" Then
            arr = Split(b.Cells(i, mainCol), ",")
            Call sortarray(arr)
            b.Cells(i, mainCol) = arr(0)
            For j = 1 To UBound(arr)
                b.Cells(i, mainCol) = b.Cells(i, mainCol) & "," & arr(j)
            Next j
        End If
    Next i
End If
'_____________________________________________________________
'[Compare]
Dim foundRow As Long
For i = 2 To BottomRow(mainCol, a.Name)
    If a.Cells(i, mainCol) <> "" Then
        foundRow = inColumn(a.Cells(i, mainCol), mainCol, b.Name, , , True)
        If foundRow > 0 Then 'incolumn true
            For j = 1 To WorksheetFunction.Max(BottomCol(i, a.Name), BottomCol(foundRow, b.Name))
                If j = locationsCol Then 'string compare
                    Call MarkDifferenceBetween2STRINGs(a.Cells(i, locationsCol), b.Cells(foundRow, locationsCol), False)
                    Call MarkDifferenceBetween2STRINGs(b.Cells(foundRow, locationsCol), a.Cells(i, locationsCol), True)
                Else 'value compare
                    If a.Cells(i, j) <> b.Cells(foundRow, j) Then
                        With a.Cells(i, j).Characters.Font
                        .Strikethrough = True
                        .Color = -16776961
                        End With
                        With b.Cells(foundRow, j).Characters.Font
                        .FontStyle = "粗體"
                        .Size = 12
                        .ThemeColor = xlThemeColorLight2
                        End With
                    End If
                End If
            Next j
            b.Cells(foundRow, botCol(1) + 10) = "OFF" 'next loop compare will not compare this row
        Else
            With a.Rows(i).Characters.Font
            .Strikethrough = True
            .Color = -16776961
            End With
        End If
    End If
Next i
'____________________________________
For i = 2 To BottomRow(mainCol, b.Name)
    If b.Cells(i, botCol(1) + 10) <> "OFF" And b.Cells(i, mainCol) <> "" Then 'hadn't been compared
            With b.Rows(i).Characters.Font
                .FontStyle = "粗體"
                .Size = 12
                .ThemeColor = xlThemeColorLight2
            End With
    End If
Next i
b.Columns(botCol(1) + 10).Clear
'==========================================================
'[Package difference to new sheet]________________________________________________________
Application.ScreenUpdating = True
Call existenceOfSheetsCheck("DifferA", True)
Call existenceOfSheetsCheck("DifferB", True)
Sheets("DifferA").Select
Cells.Select
Selection.Clear
Sheets("DifferB").Select
Cells.Select
Selection.Clear
'_____________________________________________
Dim tempBot As Long
tempBot = 2
'Sheets("DifferA").Cells(1, 1) = "mainColDiffer(SheetA)"
'Sheets("DifferB").Cells(1, 1) = "mainColDiffer(SheetB)"
'A.Select
Dim Mark As Boolean '
Dim BC As Long 'Bottomcol
BC = realBottomCol(10, a.Name)
For i = 2 To BottomRow(mainCol, a.Name, 2)
    Mark = False
    For j = 1 To BC
        If a.Cells(i, j).Font.Strikethrough = True Then
            a.Cells(i, j).Interior.Color = 5296274 'light green
            Mark = True
            'Exit For
        End If
    Next j
    '[location Col mark deal]
    For j = 1 To Len(a.Cells(i, locationsCol).Value)
        If a.Cells(i, locationsCol).Characters(Start:=j, Length:=1).Font.Strikethrough = True Then
            Mark = True
            a.Cells(i, locationsCol).Interior.Color = 5296274 'light green '
            Exit For
        End If
    Next j
    If Mark = True Then
        a.Select
        Rows(i).Select
        Selection.Copy
        Sheets("DifferA").Select
        Cells(tempBot, 1).Select
        ActiveSheet.Paste
        If a.Cells(i, mainCol) = "" Then
            Sheets("DifferA").Cells(tempBot, 1) = "[NULL]"
        End If
        tempBot = tempBot + 1
    End If
Next i
'______________________________________________
tempBot = 2
BC = realBottomCol(10, b.Name)

For i = 2 To BottomRow(mainCol, b.Name, 2)
    Mark = False
    For j = 1 To BC
        If b.Cells(i, j).Font.ThemeColor = xlThemeColorLight2 Then
            Mark = True
            b.Cells(i, j).Interior.Color = 65535 '
            'Exit For
        End If
    Next j
    '[location Col mark deal]
    For j = 1 To Len(b.Cells(i, locationsCol).Value)
        If b.Cells(i, locationsCol).Characters(Start:=j, Length:=1).Font.ThemeColor = xlThemeColorLight2 Then
            Mark = True
            b.Cells(i, locationsCol).Interior.Color = 65535 '
            Exit For
        End If
    Next j
    If Mark = True Then
        b.Select
        Rows(i).Select
        Selection.Copy
        Sheets("DifferB").Select
        Cells(tempBot, 1).Select
        ActiveSheet.Paste
            If b.Cells(i, mainCol) = "" Then
                Sheets("DifferB").Cells(tempBot, 1) = "[NULL]"
            End If
            tempBot = tempBot + 1
    End If
Next i
'_______________________________________________
If realBottomRow(realBottomCol(3, "DifferA"), "DifferA") = 1 And realBottomRow(realBottomCol(3, "DifferB"), "DifferB") = 1 Then
    MsgBox "ALL THE SAME!"
End If
'msg repeat item
tempStr1 = MarkRepeatItemWithinColumn(mainCol, a.Name, False)
tempStr2 = MarkRepeatItemWithinColumn(mainCol, b.Name, False)
If tempStr1 = True Or tempStr2 = True Then
    MsgBox "SheetA or SheetB contains repeat items! (Marked with RED BG)"
    'GoTo endFunction
End If
endFunction:
End Sub
Function clearAllFormats(targetSheet As Worksheet)

Call existenceOfSheetsCheck("tempUseForClearAllFormats", True)
targetSheet.Select
Cells.Select
Selection.Copy
Sheets("tempUseForClearAllFormats").Select
Cells.Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False 'clear all formats
Selection.Copy
targetSheet.Select
'Cells.Select
ActiveSheet.Paste 'data return
Sheets("tempUseForClearAllFormats").Select
ActiveWindow.SelectedSheets.Visible = False
'Sheets("tempUseForClearAllFormats").Hide
End Function
Function MarkRepeatItemWithinColumn(col As Long, sheetName As String, Optional Mark As Boolean = True) As Boolean
'mark repeat items within one column with red BG;return a flag of repeat item exist/not
'Col:the column to deal with
'sheetName:default value is active sheet; point to particular sheet to deal with
Dim i As Long, j As Long
Dim bot As Long, x65535 As String
'_______________________________________________________________________________'
x65535 = chr(64 + col) & 65535
bot = Sheets(sheetName).Range(x65535).End(xlUp).row
MarkRepeatItemWithinColumn = False
If Mark = False Then
    For i = 2 To bot - 1
    If Sheets(sheetName).Cells(i, col) <> "" Then
        For j = i + 1 To bot
            If Sheets(sheetName).Cells(i, col) = Sheets(sheetName).Cells(j, col) Then
                MarkRepeatItemWithinColumn = True
                GoTo theend
            End If
        Next j
    End If
    Next i
GoTo theend
End If
For i = 2 To bot - 1
    If Sheets(sheetName).Cells(i, col) <> "" Then
        For j = i + 1 To bot
            If Sheets(sheetName).Cells(i, col) = Sheets(sheetName).Cells(j, col) Then
                Sheets(sheetName).Cells(j, col).Interior.Color = 255
                Sheets(sheetName).Cells(i, col).Interior.Color = 255
                MarkRepeatItemWithinColumn = True
            End If
        Next j
    End If
Next i
theend:
End Function
Function realBottomCol(firstXRows As Long, sheetName As String)
'des:return the bottom Column as long by searching the first "firstXRows" rows
Dim IVx As String
realBottomCol = 1
For i = 1 To firstXRows
    IVx = "IV" & i
    realBottomCol = WorksheetFunction.Max(realBottomCol, Sheets(sheetName).Range(IVx).End(xlToLeft).Column)
Next i
End Function
Function realBottomRow(firstXCols As Long, sheetName As String) As Long
Dim x65535 As String, i As Long
realBottomRow = 1
For i = 1 To firstXCols
    x65535 = chr(64 + i) & 65535
    realBottomRow = WorksheetFunction.Max(Sheets(sheetName).Range(x65535).End(xlUp).row, realBottomRow)
Next i
'ActiveSheet.UsedRange.Rows.Count
End Function
Function BottomRow(Optional col As Long, Optional sheetName As String = "specialValue", Optional defaultValueIfFalse As String = "DefaultValue", Optional UsedRangeMode As Boolean = False)
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
Function BottomCol(Optional row As Long = 1, Optional sheetName As String = "specialValue", Optional UsedRangeMode As Boolean = False)
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
Function TrimData(srcCol As Long, desCol As Long, Optional sheetName As String = "SpecilValue")
'
' TrimData 巨集
' Trim Data from worksheet column i to j
'
'[Trim data of column"srcCol" to column"desCol" -on sheet1]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'srcCol = 1     '''''''''''''''''''''''''''''''''''change these value for particular using
'desCol = 8     '''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim col As Long, row As Long
    Dim Bottom As Long
    Dim strCol As String, actSheet As String
    TrimData = 0
    actSheet = ActiveSheet.Name
    If sheetName <> "SpecilValue" Then
       Sheets(sheetName).Select
    End If
    For col = srcCol To desCol Step 1
        strCol = chr(64 + col) & 65535
        Bottom = ActiveSheet.Range(strCol).End(xlUp).row
        '__________________________________________________________
        For row = 1 To Bottom Step 1
            If Left(Cells(row, col).Value, 1) = " " Or Right(Cells(row, col).Value, 1) = " " Then
               Cells(row, col) = Trim(Cells(row, col))
               TrimData = TrimData + 1
            End If
        Next row
    Next col
    Sheets(actSheet).Select
End Function
Function inColumn(indexStr As String, tarCol As Long, tarSheetName As String, Optional markAllMetWithColor As Boolean = False, Optional startRowOfTarCol As Long = 1, Optional searchOneTime As Boolean = False)
  'find the loaction of indexStr if it's in column tarCol of sheet tarSheetName
  'searchOneTime=true : the found tarRow will not join the next serching
  Dim Bottom As Long
  Dim i As Long
  Dim x65535 As String
If Len(indexStr) = 0 Then
    inColumn = 0 'default value
    GoTo theend
End If
  '____________________________________________________________________________________
  inColumn = 0
  x65535 = chr(64 + tarCol) & 65535
  Bottom = Sheets(tarSheetName).Range(x65535).End(xlUp).row
  If searchOneTime = False Then
        For i = startRowOfTarCol To Bottom
            If indexStr = Sheets(tarSheetName).Cells(i, tarCol) Then
                inColumn = i
            Exit For
            End If
        Next i
Else
    For i = startRowOfTarCol To Bottom
        If Sheets(tarSheetName).Cells(i, tarCol).Interior.ThemeColor <> xlThemeColorDark1 Then
            If indexStr = Sheets(tarSheetName).Cells(i, tarCol) Then
                inColumn = i
                Sheets(tarSheetName).Cells(i, tarCol).Interior.ThemeColor = xlThemeColorDark1 'mark
                Exit For
            End If
        End If
    Next i
End If
    
  '_______________________________________________________________________________________
  If markAllMetWithColor = True Then 'mark all the met-value
     For i = 1 To Bottom
        If indexStr = Sheets(tarSheetName).Cells(i, tarCol) Then
            Sheets(tarSheetName).Cells(i, tarCol).Interior.Color = 255
        End If
    Next i
  End If
theend:
End Function
Function MarkDifferenceBetween2STRINGs(cell1 As Range, cell2 As Range, markStyle As Boolean)
' Check difference between two strings(cells in sheet1 and sheet2), whoes form should be"x,x,x,x,x,x" (Data Separated by commas).
' Then marked the "differnce" in both cells
'==================================================================================================================
  Dim i As Long, j As Long, t As Long
  Dim src() As Long, des() As Long 'save the src and des of the difference string
  Dim tempSrc As Long, tempDes As Long
  
  Dim num(0 To 1) As String 'num of commas
  Dim found As Boolean 'mark of found
  Dim str(0 To 1) As String, chr As String
  'Dim cell1 As Range, cell2 As Range
  
  Dim arr(0 To 100, 0 To 2, 0 To 1) As String 'save src,des,separated data
  Dim AB As Long 'just to control the loop times(twice)
  '=================================================================
  str(0) = cell1 & ","  '
  str(1) = cell2 & ","
'Abstract single data and save it to arr()
 For AB = 0 To 1
        tempSrc = 0
        tempDes = 0
        num(AB) = 0
        'Separate single data
        For i = 1 To Len(str(AB)) Step 1
                chr = Mid(str(AB), i, 1)
                If Asc(chr) = 44 Then  'Meet comma or String-end
                    'Renew tempDes
                    tempDes = i
                    arr(num(AB), 2, AB) = Mid(str(AB), tempSrc + 1, tempDes - tempSrc - 1) 'save separated value
                    arr(num(AB), 0, AB) = tempSrc 'save tempSrc
                    arr(num(AB), 1, AB) = tempDes
                    tempSrc = i
                    num(AB) = num(AB) + 1
                End If
        Next i
Next AB
'=================================================================
'[Record index of difference(src,des)]
ReDim src(0 To Len(str(0))) As Long, des(0 To Len(str(0))) As Long
        t = 0
        For i = 0 To num(0) - 1 Step 1
                found = False
                For j = 0 To num(1) - 1 Step 1
                        If arr(i, 2, 0) = arr(j, 2, 1) Then
                            found = True
                            Exit For
                        End If
                Next j
                If found = False Then 'found part of string with blue color
                    src(t) = arr(i, 0, 0) + 1
                    des(t) = arr(i, 1, 0)
                    t = t + 1
                End If
        Next i
'[Mark difference basing on index]
        '____________(Mark Difference)_________________________________________
        t = 0
        If markStyle = True Then
        While src(t) > 0
               With cell1.Characters(Start:=src(t), Length:=des(t) - src(t)).Font
                        '.Name = "Arial"
                        .FontStyle = "粗體"
                        .Size = 12
                        .ThemeColor = xlThemeColorLight2
                        .TintAndShade = 0.399975585192419
                        .ThemeFont = xlThemeFontNone
              End With
              t = t + 1
        Wend
        Else
        While src(t) > 0
               With cell1.Characters(Start:=src(t), Length:=des(t) - src(t)).Font
                        .Strikethrough = True
                        .Color = -16776961
            End With
        t = t + 1
        Wend
End If

End Function
Function existenceOfSheetsCheck(sheetName As String, Optional add As Boolean = False) As Boolean
'[Function]:Existence check of sheet "sheetName";if not exist and add=true,add a sheet named "sheetName"
    Dim Ws As Worksheet
    existenceOfSheetsCheck = True
    On Error Resume Next
    Set Ws = Sheets(sheetName)
    If Ws Is Nothing Then
       existenceOfSheetsCheck = False
       If add = True Then
          Sheets.add.Name = sheetName
          MsgBox "Add a sheet (" & sheetName & ")!", , "Function: existenceOfSheetsCheck()"
       Else
          MsgBox "Not found sheet (" & sheetName & ")!", , "Function: existenceOfSheetsCheck()"
       End If
    Else
       Sheets(sheetName).Visible = True
    End If
    
End Function

