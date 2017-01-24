VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub PLMtoEE()
'auto copy related data of PLM-downloaded BOM to sheet[OriginalDataSideA/B]
Dim srcSheet As Worksheet, desSheet As Worksheet
Set srcSheet = Sheets("PLMBOM1")
Set desSheet = Sheets("OriginalDataSideA")

End Sub
Sub MergeEEBOM()
'
'Merge EE BOM side A & side B into one BOM. To use this procedure, excel file should contains sheets "Side A" and "Side B"
'EEBOM formula¡G
'Item Plant   PART TYPE       PARENT_NUMBER   Priority    Component     Qty   REFDESIG    Component Description
'0010     SUB-ASSEMBLY    111MBV122015    1           6019A0538201    1   U121        IC-DETECTOR,S-80840CNMC-B8ZT2G,GP,SOT23,TAP,5PIN,x,SII
'0010     SUB-ASSEMBLY    111MBV122015                6019A0722201    1               IC-VDETAC,NCP301LSN40T1G,GP,TSOP-5,TAP,5PIN,VF,ONSEMI
'==========================================================================================================================================
'Head 'G':for Group dealing
Dim GsearchingResults As String 'Record searching results of every item within a group in the type of bit :"1010101..."
Dim Growtop As Integer, Gfoundrow As Integer 'Growtop: the top row of a group; Gfoundrow: the normal row of a group
'___________________________
'Head 'P' for 'Public' using
Dim Ptemp As String, PflagMsgBox As String
Dim Pbot As Integer 'bot of temp using
'___________________________
'Default for public using
Dim i As Integer, j As Integer
Dim Bot1 As Integer, Bot2 As Integer 'bottom of sheet A, bottom of sheet B
Dim a As String, b As String 'Point to sheet name
Dim reverseAB As Boolean 'Code flow select
Dim errorLog(1 To 10) As String, iERRORLOG As Integer
'+++++++++++++++++++++++++++++++++++
Dim Flag_repeatCheck As Boolean
Flag_repeatCheck = False '[Test]

'===========================================================================================================================================
errorLog(1) = "DESKTOP CHECK...OK!"
errorLog(4) = "REPEAT ITEM CHECK...OK!"
errorLog(5) = "REPEAT LOCATION CHECK...OK!"
errorLog(6) = "GROUP UNION CHECK...OK!"
errorLog(7) = "REDUNDANT SOURCE CHECK...OK!"
errorLog(9) = "REPEAT ITEM OF MERGED BOM CHECK...OK!"
iERRORLOG = 0
'______________________________________________________________________________________
PflagMsgBox = vbNo 'MsgBox("Show Procedure?" & chr(13) & " [Tips:If you don't want to analyse the ERROR or WARNING that may occur or it's the FIRST TIME you run this procedure, please press No]", vbYesNo)

'(All sheets' existence check)_____________________________________________________
Call existenceOfSheetsCheck("AnalyseSideA", True)
Call existenceOfSheetsCheck("AnalyseSideB", True)
Call existenceOfSheetsCheck("MergedBOM", True)
Sheets("MergedBOM").Cells.Clear
Ptemp = existenceOfSheetsCheck("Desktop")
If Ptemp = False Then
   MsgBox "Please Add sheet'Desktop'!"
   GoTo endFunction
Else
   If Sheets("Desktop").Cells(12, 1) <> "ArrowChang" Then
   MsgBox "Sheet'Desktop'NOT RIGHT'!"
   iERRORLOG = 1
   errorLog(1) = "Desktop has been modified illegally!" '+++++++++++++++++++++++++++
   GoTo endFunction
   End If
End If
Call existenceOfSheetsCheck("OriginalDataSideA", True)
Call existenceOfSheetsCheck("OriginalDataSideB", True)
Call existenceOfSheetsCheck("ErrorLog", True)
'===================================================================================
a = "AnalyseSideA"
b = "AnalyseSideB"
'__________________________________(Copy original data to sheet Analyseside)________
'Application.ScreenUpdating = False
'++++++++++++++++++
Sheets(a).Select '+
'++++++++++++++++++
For i = 0 To 1
    Sheets("OriginalDataSide" & chr(65 + i)).Cells.Copy ActiveSheet.range("a1")
    Cells.Select
    Selection.ClearFormats
    '++++++++++++++++++
    Sheets(b).Select '+
    '++++++++++++++++++
Next i
'_____________________(Trim)________________________________________________________
iERRORLOG = 2
errorLog(2) = "" & TrimData(2, 9, a) + TrimData(2, 9, b) & " cells have been TRIMED." '++++++++++++++++++++
'________________________(Shift Priority to top row)________________________________
iERRORLOG = 3
errorLog(3) = "" & shiftLocation(0, a) + shiftLocation(0, b) & " groups have shifted 'Priority 1' to their top row." '+++++++++++++++++++++
'===================================================================================
Application.ScreenUpdating = True

'PflagMsgBox = MsgBox("Checking Repeat Item?", vbYesNo)
'If PflagMsgBox = vbNo Then
'GoTo jumpRepeatItemCheck
'End If
'___________________________________(Check repeat item[IACPN] 3-1)__________________
If Flag_repeatCheck = True Then
Ptemp = MarkRepeatItemWithinColumn(6, a) & MarkRepeatItemWithinColumn(6, b)
iERRORLOG = 4
If InStr(Ptemp, "True") > 0 Then 'temp change
  MsgBox "BOM includes repeating items! [Location with RED BG]" & chr(13) & "Please correct them!", , "ERROR"
  errorLog(4) = "Sheet " & a & " or/and " & b & " includes REPEAT ITEMS!" '+++++++++++++++++++++++++++++
   GoTo endFunction
End If
Application.ScreenUpdating = False
End If
'_________________________________(Check repeat item[location] 3-2)_________________
'[1. List separate locations to sheet "LocationsSeparate"]
With Sheets("LocationsSeparate")
     .Cells.Clear
     .Cells(1, 1) = "Location"
     .Cells(1, 2) = "BelongsToWhichSheet"
End With
For i = 0 To 1
  For j = 2 To Sheets(a).range("H65535").End(xlUp).row
      Pbot = Sheets("LocationsSeparate").range("A65535").End(xlUp).row + 1
      Sheets("LocationsSeparate").Cells(Pbot, 2) = a 'mark where the locations belong to
      Call LocationsSeparate(Sheets(a).Cells(j, 8), Cells(Pbot, 1)) 'save to default sheet "LocationsSeparate"
  Next j
  '++++++++++++++++++++
  a = "AnalyseSideB" '+
  b = "AnalyseSideA" '+
  '++++++++++++++++++++
Next i
'[2.Check repeat location]___________________________________________________________________
If Flag_repeatCheck = True Then
Ptemp = MarkRepeatItemWithinColumn(1, "LocationsSeparate")
iERRORLOG = 5
If InStr(Ptemp, "True") > 0 Then
   Sheets("LocationsSeparate").Select
   MsgBox "Location(s) repeat! [Location with RED BG]" & chr(13) & "Please refer to sheet 'LocationsSeparate'!", , "ERROR"
   errorLog(5) = "Location(s) REPEAT!" '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   GoTo endFunction
End If
'===================================================================================
jumpRepeatItemCheck:
End If
'________________________________________(Search A/B)_______________________________
Application.ScreenUpdating = True
reverseAB = True 'only alow below loop one time
reverseABtoCheck:
Sheets(b).Select
Bot1 = Sheets(a).range("F65535").End(xlUp).row
Bot2 = ActiveSheet.range("F65535").End(xlUp).row
If PflagMsgBox = vbYes Then
   MsgBox "BotB=" & Bot1 & ",BotA=" & Bot2 & ""
End If
For j = 2 To Bot2 Step 1
    If Cells(j, 8) <> "" Then 'Based on a new group to deal with
       '[find Group-Pbot-row]-----------------------
       Pbot = edgeRow("BOT", Cells(j, 8), Bot2, b) '-
       'MsgBox "j,Pbot"& j & Pbot & ""
       '-------------------------------------------
       GsearchingResults = "" 'set initial value
       Growtop = 1 'default value, means not found
       For i = j To Pbot Step 1 'Group search
           Cells(i, 6).Select '[easy test]
           Gfoundrow = inColumn(Cells(i, 6), 6, a) 'inColumn() returns the row No of sheet"A",default value is "0"
           If PflagMsgBox = vbYes Then
              MsgBox "incolumn=" & Gfoundrow & ""
           End If
           '[find Group-row-top]--------------------------------
           Gfoundrow = edgeRow("top", Cells(Gfoundrow, 8), , a) '-
           '----------------------------------------------------
           If PflagMsgBox = vbYes Then
              MsgBox "top=:" & Gfoundrow & ""
           End If
           '========================[1.Rowtop recheck]==================================
           iERRORLOG = 6
           If Growtop > 1 And Gfoundrow > 1 And Growtop <> Gfoundrow Then 'found the items belong to same group of B not in the same group of A
              Sheets(a).Cells(Gfoundrow, 8).Interior.Color = 255 '[mark]
              Cells(j, 8).Interior.Color = 255
              Cells(i, 6).Interior.Color = 255
              Cells(jtb, 6).Interior.Color = 255
              MsgBox "One-group source of sheet[B] belong to multi group of sheet[A]!" & chr(13) & "[Location BG marked with RED]"
              errorLog(6) = "One-group source of sheet[B] belong to multi group of sheet[A]!" '++++++++++++++++++++++++++
              GoTo endFunction
           End If
           
           If Gfoundrow > 1 Then 'Gfoundrow is the leagle value
              jtb = i
              Growtop = Gfoundrow 'save the legal value as the Growtop
           End If
           '========================[2.record searching result of every item===========
           If Gfoundrow > 1 Then '[found group in sheet A]
              GsearchingResults = GsearchingResults & "1"
           Else '[not found group in sheet A]
              GsearchingResults = GsearchingResults & "0"
           End If
       Next i
       '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
       If PflagMsgBox = vbYes Then
          MsgBox "Searching Results=" & GsearchingResults & ""
       End If
'_______________________________(RedundantIACPN Marked)________________________________
       If InStr(GsearchingResults, "1") > 0 And InStr(GsearchingResults, "0") > 0 Then 'This PN not found its conterpart but in this group other PN had found_
         iERRORLOG = 7
         For i = 1 To Len(GsearchingResults)
             If Mid(GsearchingResults, i, 1) = "0" Then
               Cells(j + i - 1, 6).Interior.Color = 49407 'Orange
               MsgBox "This IAC PN Not Found in the Group of Side A! May need to be deleted", , "Warning:"
               errorLog(7) = "[Warning]: BOM (may) includes REDUNDANT source" '+++++++++++
             End If
         Next
       End If
       If reverseAB = False Then 'only allow the 2nd time to action
'=======================================(Action)========================================
          'If InStr(GsearchingResults, "1") > 0 Then '[add location]
           If InStr(GsearchingResults, "1") > 0 And Sheets(a).Cells(Growtop, 8) <> Cells(j, 8) Then '[add location]
             src = Len(Sheets(a).Cells(Growtop, 8)) + 1 'before cell change its value
             Sheets(a).Cells(Growtop, 8) = Sheets(a).Cells(Growtop, 8) & "," & Cells(j, 8) 'Location
             Sheets(a).Cells(Growtop, 7) = Val(Sheets(a).Cells(Growtop, 7)) + Val(Cells(j, 7)) 'Qty
             '[Mark change]
             des = Len(Sheets(a).Cells(Growtop, 8)) 'after cell change its value
             Sheets(a).Cells(Growtop, 8).Characters(Start:=src, Length:=des).Font.Color = 12611584 'Blue
          Else 'A not contain, [all move]
             For i = 1 To Len(GsearchingResults)
                 For k = 1 To 9
                     Sheets(a).Cells(Bot1 + i, k) = Cells(j + i - 1, k)
                 Next k
             Next i
             Bot1 = Bot1 + (i - 1) 'renew bot1
          End If
       End If
    End If
Interruption:
Next j
'========================================================================================
If reverseAB = True Then
   reverseAB = False
   '+++++++++++++++++++++'reverse A&B:Now A B recovered to the normal value
   a = "AnalyseSideA" '++
   b = "AnalyseSideB" '++
   '+++++++++++++++++++++
   GoTo reverseABtoCheck
End If
'______________________________(Update Qty.)____________________________________________
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
iERRORLOG = 8
errorLog(8) = QtyCheck(7, 8, a, "True") & " , " & QtyCheck(7, 8, b, "True")  '++++++++++++
errorLog(8) = "Qty. check results: " & errorLog(8) & "" '++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'______________________________(copy to sheet "MergedBOM")_____________________________
Sheets("AnalyseSideA").Cells.Copy Sheets("MergedBOM").range("a1")
Sheets("MergedBOM").Select
'______________________________(Check repeating item 2-2)______________________________
If Flag_repeatCheck = True Then
Ptemp = MarkRepeatItemWithinColumn(6, "MergedBOM")
iERRORLOG = 9
If InStr(Ptemp, "True") > 0 Then
   MsgBox "MergedBOM includes repeating items! [Location with RED BG]" & chr(13) & "Please correct them!", , "ERROR"
   errorLog(9) = "Sheet 'MergedBOM' includes repeating items!" '++++++++++++++++++++++++++++++++++++++++++++++++++++++
   GoTo endFunction
End If
End If
'_______________________________________________________________________________________
MsgBox "Merge EEBOM DONE!Please refer to ErrorLog for reference!", , "The End"
endFunction:
'________________________________(Vew Error Log)_______________________________________
Sheets("ErrorLog").Select
Cells.Clear
Cells(1, 1) = "ErrorLog"
Cells(1, 1).Interior.Color = 65535 'yellow
Cells(1, 2).Interior.Color = 65535
For i = 1 To iERRORLOG
    With Cells(i + 1, 1)
         .Value = "ERRORLOG(" & i & ")"
         .Interior.Color = 12611584 'Blue
         .Font.ThemeColor = xlThemeColorDark1 'white
    End With
    With Cells(i + 1, 2)
         .Value = errorLog(i)
         .Interior.Color = 12611584 'Blue
         .Font.ThemeColor = xlThemeColorDark1 'white
    End With
Next i
Application.ScreenUpdating = True
End Sub
Function ErrorLogGenerator(saveCell As range, msg As String)
Call existenceOfSheetsCheck("ErrorLog", True)
Sheets("ErrorLog").Cells(saveCell.row, saveCell.Column) = msg
End Function
Function existenceOfSheetsCheck(sheetName As String, Optional ADD As Boolean = False) As Boolean
'[Function]:Existence check of sheet "sheetName";if not exist and add=true,add a sheet named "sheetName"
    Dim Ws As Worksheet
    existenceOfSheetsCheck = True
    On Error Resume Next
    Set Ws = Sheets(sheetName)
    If Ws Is Nothing Then
       existenceOfSheetsCheck = False
       If ADD = True Then
          Sheets.ADD.Name = sheetName
          MsgBox "Add a sheet (" & sheetName & ")!", , "Function: existenceOfSheetsCheck()"
       Else
          MsgBox "Not found sheet (" & sheetName & ")!", , "Function: existenceOfSheetsCheck()"
       End If
    Else
       Sheets(sheetName).Visible = True
    End If
    
End Function
Function TrimData(srcCol As Integer, desCol As Integer, Optional sheetName As String = "SpecilValue")
'
' Trim Data from worksheet column i to j
'[Trim data of column"srcCol" to column"desCol" -on sheet1]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'srcCol = 1     '''''''''''''''''''''''''''''''''''change these value for particular using
'desCol = 8     '''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim col As Integer, row As Integer
    Dim Bottom As Integer
    Dim strCol As String, actSheet As String
    TrimData = 0
    actSheet = ActiveSheet.Name
    If sheetName <> "SpecilValue" Then
       Sheets(sheetName).Select
    End If
    For col = srcCol To desCol Step 1
        strCol = chr(64 + col) & 65535
        Bottom = ActiveSheet.range(strCol).End(xlUp).row
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
Function shiftLocation(Strategy As Integer, Optional sheetName As String = "specialValue")
'In some BOM, it's location marked Priority 1 is not the first row of this group. We neeed to change this type
'A:Item No
'E:Priority
'H:Location

Dim i As Integer, j As Integer, k As Integer
Dim bot As Integer
Dim Item As String
tempsave = ActiveSheet.Name
If sheetName <> "specialValue" Then
   Sheets(sheetName).Select
End If

shiftLocation = 0 'record the number of modified items
bot = range("A65535").End(xlUp).row
Select Case Strategy
Case 0 'only move the location and priority to the top row of the group
     Item = Cells(1, 1)
     For i = 2 To bot
         If Cells(i, 1) <> Item Then
            'Cells(i, 1).Select '[test]
            'MsgBox ""
            Item = Cells(i, 1)
            If Cells(i, 8) = "" Then 'location is NULL,find the row which has location(Priority 1) in this group
               shiftLocation = shiftLocation + 1
               botRow = edgeRow("bot", Cells(i, 8), , sheetName) + 1
              ' Cells(botrow, 1).Select '[test]
               ' MsgBox "botrow=" & botrow & ""
               If Cells(botRow, 1) = Item Then 'recheck if the found row in the group
                  '[Action]
                  Cells(i, 5) = "1"
                 ' Cells(i, 5).Interior.color = 65535
                  Cells(i, 8) = Cells(botRow, 8)
                  Cells(i, 8).Interior.Color = 65535
                  Cells(botRow, 5) = ""
                  Cells(botRow, 8) = ""
               Else
                  Cells(i, 1).Select
                  Cells(i, 1).Interior.Color = 255
                  MsgBox "Error! this group NO Location!"
               End If
            End If
         End If
     Next
Case 1

Case Else

End Select
Sheets(tempsave).Select
End Function
Function MarkRepeatItemWithinColumn(col As Integer, Optional sheetName As String = "specialValue")
'mark repeat items within one column with red BG;return a flag of repeat item exist/not
'Col:the column to deal with
'sheetName:default value is active sheet; point to particular sheet to deal with
Dim i As Integer, j As Integer
Dim bot As Integer, Mark As Boolean, x65535 As String, activeSheetName As String
'_______________________________________________________________________________'
activeSheetName = ActiveSheet.Name 'save the active name because may change the active sheet for easy coding
x65535 = chr(64 + col) & 65535
If sheetName <> "specialValue" Then
   Sheets(sheetName).Select
End If
bot = ActiveSheet.range(x65535).End(xlUp).row
'____________________________________________________________'
Mark = False
For i = 2 To bot - 1
    For j = i + 1 To bot
        If Cells(i, col) = Cells(j, col) Then
           Cells(j, col).Interior.Color = 255
           Cells(i, col).Interior.Color = 255
           Mark = True
        End If
    Next j
Next i
'____________________________________________________________'
MarkRepeatItemWithinColumn = Mark
Sheets(activeSheetName).Select 'recover active sheet
End Function
'==================================================================================================
Function LocationsSeparate(ByVal targetCell As String, ByVal savHeadCell As range, Optional savSheet As String = "LocationsSeparate")
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
  
  activeSheetName = ActiveSheet.Name 'to protect present active sheet
  '_________________________________________________________________
  Call existenceOfSheetsCheck(savSheet, True)
  Sheets(savSheet).Select  'saving-sheet active
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
       Cells(savRow + j, savCol) = Mid(str, src + 1, des - src - 1)
       src = i 'Renew src
       j = j + 1
    End If
  Next
  Sheets(activeSheetName).Select 'return to pre-workingsheet
End Function
Function edgeRow(TopOrBot As String, ByVal startCell As range, Optional range As Integer = 0, Optional sheetName As String = "specialValue")
'To find the top or bottom row of a group. givenCol As Integer, givenRow As Integer
'For example, if topOrBot="Top",GivenRow=10,givenCol=8,range=2
'This function will return the value of nearest row,in column 8, above givenRow 10 (but below Range) whose value is not NULL
Dim givenRow As Integer, givenCol As Integer
Dim row As Integer
Dim Mark As Boolean
Dim x65535 As String, TB As String
'(Get row&column)____________________________________________________________________
givenRow = startCell.row
givenCol = startCell.Column
'(Parameters check)___________________________________________________________________
Mark = True
TB = UCase(TopOrBot)
Select Case TB
Case "TOP"
     If range = 0 Then 'default value
        range = 1 '"Top" case default value
     End If
     '_____________________________________
     If givenRow < range Then
        MsgBox "Parameter:" & chr(13) & "givenRow(" & givenRow & ") < Range(" & range & ")", , "Function edgeRow(,  , givenRow As Integer, Range As Integer,)"
        Mark = False
     End If
Case "BOT"
     If range = 0 Then 'default value
        x65535 = chr(64 + givenCol) & 65535
        'Range = ActiveSheet.Range(x65535).End(xlUp).row '"Bot" case default value,[BUG]
        range = WorksheetFunction.Max(ActiveSheet.range("A65535").End(xlUp).row, ActiveSheet.range("B65535").End(xlUp).row, _
               ActiveSheet.range(x65535).End(xlUp).row, ActiveSheet.range("E65535").End(xlUp).row, ActiveSheet.range("F65535").End(xlUp).row, _
               ActiveSheet.range("G65535").End(xlUp).row, ActiveSheet.range("H65535").End(xlUp).row)
     End If
     '_____________________________________
     If givenRow > range Then
        MsgBox "Parameter:" & chr(13) & "givenRow(" & givenRow & ") > Range(" & range & ")", , "Function edgeRow(,  , givenRow As Integer, Range As Integer,)"
        Mark = False
     End If
Case Else
     MsgBox "Parameter TopOrBot(" & TopOrBot & ") Not RIGHT!", , "Function edgeRow(TopOrBot As String,  ,  ,  ,  )"
     Mark = False
End Select
If Mark = False Then
   edgeRow = "False"
   GoTo endFunction
End If
'(Save Present Parameters)________________________________________________________
row = givenRow
activeSheetName = ActiveSheet.Name 'save the active name because may change the active sheet for easy coding
If sheetName <> "specialValue" Then
   Sheets(sheetName).Select
End If
'(Find the top/bot row)___________________________________________________________
Select Case TB
  Case "TOP"
      While Cells(row, givenCol).Value = "" And row > range
        row = row - 1
      Wend
  Case "BOT"
      While Cells(row + 1, givenCol).Value = "" And row < range
        row = row + 1
      Wend
End Select
edgeRow = row
'(edge-cell deal;if the searching meet edge of range)_____________________________
'Cells(edgeRow, givenCol).Select '[test only]
If row = range Then 'loop to the edge, this row may be not NULL(TOP case)
  If Cells(range, givenCol).Value = "" And TB = "TOP" Then
      edgeRow = "False"
      MsgBox "From row " & givenRow & " to " & range & " Not Found 'edgeRow'!", , "Function edgeRow(TopOrBot As String,  givenCol As Integer, givenRow As Integer, Range As Integer, Optional sheetName As String = 'specialValue')"
  End If
End If
'_________________________________________________________________________________
Sheets(activeSheetName).Select 'recover active sheet
endFunction:
End Function
Function inColumn(ByVal str As String, ByVal col As Integer, Optional ByVal sheetName As String = "specialValue", Optional markWithColor As Boolean = False)
  'find the loaction of str if it's in column Col of sheet sheetName
  Dim Bottom As Integer
  Dim row As Integer
  Dim x65535 As String
  Dim actSheet As String
  
  actSheet = ActiveSheet.Name
  '____________________________________________________________________________________
  inColumn = 1 'default value
  If sheetName <> "specialValue" Then
     Sheets(sheetName).Select
  End If
  '____________________________________________________________________________________
  x65535 = chr(64 + col) & 65535
  Bottom = ActiveSheet.range(x65535).End(xlUp).row
  For row = 1 To Bottom
      If str = Cells(row, col) Then
         inColumn = row
         Exit For
      End If
  Next row
  '_______________________________________________________________________________________
  If markWithColor = True Then 'mark all the met-value
     For row = 1 To Bottom
      If str = Cells(row, col) Then
         Cells(row, col).Interior.Color = 255
      End If
  Next row
  End If
  Sheets(actSheet).Select
End Function
Function QtyCheck(colQty As Integer, colRef As Integer, Optional sheetName As String = "specialValue", Optional emptyRefDeal As String = "DefaultValue")
'Item title better contains Qty & Ref Des
'===============================================================================================================
'     1(A)             2(B)         3(C)     4(D)       5(E)            6(F)            7(G)       8(H)
'===============================================================================================================
'Item Number | Item Description  |  Qty  | Ref Des |  Mfr. Name |  Mfr. Part Number  | IAC PN | IAC Description
'===========================================================================================================
  Dim i As Integer, t As Integer
  Dim Bottom As Integer, delta As Integer, Mark As Integer
  Dim str As String
  QtyCheck = "" & sheetName & "...Error Exist"
  activeSheetName = ActiveSheet.Name 'save the active name because may change the active sheet for easy coding
'_______________________________________________________________________________
  strCol = chr(64 + colRef) & 65535
  If sheetName <> "specialValue" Then
   Sheets(sheetName).Select
  End If
  Bottom = ActiveSheet.range(strCol).End(xlUp).row
  If Bottom < 2 Then
     Bottom = 2
  End If
'_______________________________________________________________________________
  Mark = 0
  For i = 2 To Bottom Step 1
     delta = 0
     If Cells(i, colRef) <> "" Then
        str = Cells(i, colRef) & "," '
        For t = 1 To Len(str) Step 1
            If Asc(Mid(str, t, 1)) = 44 Then  'Meet comma or String-end
            delta = delta + 1
            End If
        Next
        'MsgBox "" & delta & ""
        '[Compara delta with old Qty in column C]
        If delta <> Val(Cells(i, colQty)) And delta > 0 Then
            Mark = 1
            MsgBox "Found Item of Qty Not Right and Modified (Marked With Light Blue)", , "Warming"
            Cells(i, colQty) = delta
            Cells(i, colQty).Interior.Color = 15773696 'Bright Blue
        End If
     End If
  Next i
  '____________________________________________________________________________________
  If emptyRefDeal <> "DefaultValue" Then
     For i = 2 To Bottom Step 1
         If Trim(Cells(i, colRef)) = "" Then
            Cells(i, colQty) = "'" & Cells(i - 1, colQty)
         End If
     Next i
  End If
  '____________________________________________________________________________________
  If Mark = 0 Then
     'MsgBox ("All Qty. Is Right!")
     QtyCheck = "" & sheetName & " All Right"
  End If
  Sheets(activeSheetName).Select 'recover active sheet
  
End Function

