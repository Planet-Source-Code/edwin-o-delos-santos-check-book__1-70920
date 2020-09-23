Attribute VB_Name = "modPROCEDURE"
Option Explicit

'** Add_Item
'** autoAlignCol
'** BindDatasource
'** errorMsg
'** Delete_Record
'** DisableX
'** Drag_It
'** FillListView
'** FlatHeader
'** Gradient (form) <<syntax: Gradient me,100,150,200,1>>
'** InsertColumn
'** Insert_Fields
'** Listview_Total
'** ListView_Search
'** loadForm
'** LvwReplaceData
'** Print_Details
'** Print_Headings
'** Print_Total
'** PrintValue
'** RemoveDuplicate (listbox) <<syntax: RemoveDuplicate list1>>
'** ShowFldsLabel
'** SortListView <<syntax: SortListView ListView1, ColumnHeader >>
'** TextBox_Visible
'** TextLocked;*TxtLocked
'** UnloadAllForms
'** WriteData
'===========================================================
Public Sub FlatHeader(ByRef Lvw As ListView)
'// API
Dim r As Long
Dim style As Long
Dim hHeader As Long
hHeader = SendMessageLong(Lvw.hWnd, LVM_GETHEADER, 0, ByVal 0&)
style = GetWindowLong(hHeader, GWL_STYLE)
style = style Xor HDS_BUTTONS
If style Then
r = SetWindowLong(hHeader, GWL_STYLE, style)

End If
End Sub
Public Sub UnloadAllForms()
Dim Form As Form
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
End Sub

Public Sub RemoveDuplicate(ByRef lstBox As ListBox)
Dim i As Long
Dim j As Long
Dim currItem As String
For i = 0 To lstBox.ListCount - 1
   currItem = lstBox.List(i)
   For j = 0 To lstBox.ListCount - 1
      If j <> i Then
'       If currItem <> "" Then 'added by edwin
         If lstBox.List(j) = currItem Then
              lstBox.RemoveItem (j)
         End If
'       End If 'edwin
      End If
   Next j
Next i

End Sub

Public Sub Gradient(TheObject As Object, ByVal Redval As Long, ByVal Greenval As Long, _
    ByVal Blueval As Long, ByVal Direction As Integer)
    Dim Step As Integer, Reps As Integer, FillTop As Integer
    Dim FillLeft As Integer, FillRight As Integer, FillBottom As Integer
    If Direction < 1 Or Direction > 4 Then Direction = 1
    FillTop = 0
    FillLeft = 0
    If Direction < 3 Then
        Step = (TheObject.Height / 100)
        If Direction = 2 Then FillTop = TheObject.Height - Step
        FillBottom = FillTop + Step
        FillRight = TheObject.Width
    Else
        Step = (TheObject.Width / 100)
        If Direction = 4 Then FillLeft = TheObject.Width - Step
        FillRight = FillLeft + Step
        FillBottom = TheObject.Height
    End If
    For Reps = 1 To 100
        If Direction = 2 And Reps = 100 Then FillTop = 0
        If Direction = 4 And Reps = 100 Then FillLeft = 0
        Redval = Redval - 3
        Greenval = Greenval - 3
        Blueval = Blueval - 3
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), _
            BF
        If Direction < 3 Then
            If Direction = 1 Then
                FillTop = FillBottom
            Else
                FillTop = FillTop - Step
            End If
            FillBottom = FillTop + Step
        Else
            If Direction = 3 Then
                FillLeft = FillRight
            Else
                FillLeft = FillLeft - Step
            End If
            FillRight = FillLeft + Step
        End If
    Next Reps
End Sub




Public Sub TextBox_Visible(ByRef frm As Form, ByVal rs As Recordset)
  Dim i As Integer
  Dim numba As Integer
  i = 0
      For i = 0 To frm.txtEntry.UBound             'make all textbox visible
               frm.txtEntry(i).Visible = False
               frm.lblFLDi(i).Visible = False
          Next i
  i = 0
  numba = (rs.Fields.Count - 1)
      For i = 0 To numba                           'make number of textbox available
               frm.txtEntry(i).Visible = True
               frm.lblFLDi(i).Visible = True
          Next i
End Sub


Public Sub DisableX(TheForm As Form)
    '** Description:
    '** Disable X in upper right corner of the form
    Dim lngMenu As Long
    lngMenu = GetSystemMenu(TheForm.hWnd, False)
    DeleteMenu lngMenu, 6, MF_BYPOSITION
End Sub

Public Sub Print_Total(ByRef srcRS As Recordset, ByVal tabStart As Integer, ByRef lst As ListBox)
'// i prefer to use Tab() function rather that currentX/Y
Dim X As ListItem
Dim fldName As String
Dim strvalue As String
Dim currentTab As Integer
Dim iCount As Integer
Dim i As Integer
ReDim dblTotal(srcRS.Fields.Count - 1)
'//initialize value
For i = 0 To (srcRS.Fields.Count - 1)
   dblTotal(i) = 0
   Next i
   i = i + 1
iCount = 0
With srcRS
  .MoveFirst
While Not .EOF = True
iCount = iCount + 1
currentTab = tabStart
  For i = 0 To lst.ListCount - 1
     If lst.Selected(i) = True Then
       fldName = Empty
       strvalue = Empty
       fldName = printIndex(i)
       strvalue = srcRS.Fields(fldName)
       If srcRS.Fields(fldName).Type = 6 Or srcRS.Fields(fldName).Type = 5 Then
               dblTotal(i) = dblTotal(i) + Val(strvalue)   'array/declared public
               If iCount = .RecordCount Then
                  Call PrintValue(currentTab, dblTotal(i), 12) 'print total if EOF reached
               End If
           currentTab = currentTab + 15
        ElseIf srcRS.Fields(fldName).Type = 7 Then
           currentTab = currentTab + 15
        ElseIf srcRS.Fields(fldName).Type = 3 Then
           currentTab = currentTab + 13
        Else
            currentTab = currentTab + 35
        End If 'isnumeric
      End If  'selected = true
    Next i
         currentTab = tabStart
         i = i + 1
.MoveNext

Wend
End With
End Sub

Public Sub Print_Details(ByRef srcRS As Recordset, ByVal tabStart As Integer, ByRef lst As ListBox)
Dim strvalue As String
Dim fldName As String
Dim currentTab As Integer
Dim i As Integer
On Error Resume Next
currentTab = tabStart
  For i = 0 To lst.ListCount - 1
     If lst.Selected(i) = True Then
       fldName = Empty
       strvalue = Empty
       fldName = printIndex(i)
       strvalue = srcRS.Fields(fldName) '//srcRS.Fields(i) *replaced
       If srcRS.Fields(fldName).Type = 6 Or srcRS.Fields(fldName).Type = 5 Then
            Call PrintValue(currentTab, strvalue, 12)
            currentTab = currentTab + 15
        ElseIf srcRS.Fields(fldName).Type = 7 Then
            Printer.Print Tab(currentTab); strvalue;
           currentTab = currentTab + 15
        ElseIf srcRS.Fields(fldName).Type = 3 Then
            Printer.Print Tab(currentTab); strvalue;
            currentTab = currentTab + 13
        Else
            Printer.Print Tab(currentTab); Mid(strvalue, 1, 25);
            currentTab = currentTab + 35
        End If 'isnumeric
            End If  ' selected = true
    Next i
         currentTab = tabStart
End Sub

'//procedure to align value to the right: as in  31220.00
'//                                                200.00
'//original coding by myself
'//no modify within this proc
'---------------------------------------------------------
Public Sub PrintValue(ByRef iTab As Integer, ByVal srcValue As String, maxLEN As Integer)
'//Remarks: maxlen must be equal to maxlen with Print_Headings
'//currLen(15) declared public
 Dim intLEN As Integer, currtab As Integer
 Dim strvalue As Double
 Dim i As Integer
  intLEN = Len(Format(srcValue, "#,###,##0.00"))
  For i = 0 To maxLEN
       currLen(i) = i
          If currLen(i) = intLEN Then
            currtab = iTab + (maxLEN - intLEN)
            GoSub printerP
          End If
      Next i
        i = i + 1
'//sub
printerP:
  If Val(srcValue) > 0 Then
   Printer.Print Tab(currtab); Format(srcValue, "#,###,##0.00");
  Else
     Printer.Print Tab(currtab); "--";
  End If
End Sub

Public Sub Print_Headings(ByRef srcRS As Recordset, _
                         ByVal tabStart As Integer, _
                         ByRef lst As ListBox, _
                         ByVal maxLEN As Integer)
'//Remarks: maxlen must be equal to maxlen with PrintValue
Dim strvalue As String
Dim currentTab As Integer
Dim fldName As String  'hanlde feidlName
Dim heading As String
currentTab = tabStart
Dim hdng As String
On Error Resume Next
'// recordset required
Dim i As Integer
  For i = 0 To lst.ListCount - 1
     If lst.Selected(i) = True Then
       fldName = Empty
       heading = Empty
       fldName = printIndex(i)
       'heading = Mid(printIndex(i), 1, maxLEN) 'use by integer/double/date
       heading = printIndex(i)
       heading = SplitString(heading)
       strvalue = Empty
       'strvalue = SplitString(srcRS.Fields(fldName))
       If srcRS.Fields(fldName).Type = 6 Or srcRS.Fields(fldName).Type = 5 _
          Or srcRS.Fields(fldName).Type = 7 Then      '//If IsNumeric(strvalue) Then *replaced
            Printer.Print Tab(currentTab); Mid(heading, 1, maxLEN);
            currentTab = currentTab + 15
       ElseIf srcRS.Fields(fldName).Type = 3 Then
            Printer.Print Tab(currentTab); Mid(heading, 1, maxLEN);
            currentTab = currentTab + 13
       Else
           Printer.Print Tab(currentTab); Mid(heading, 1, 25);
            currentTab = currentTab + 35
        End If
            End If
    Next i
         currentTab = tabStart
End Sub

Public Sub Insert_Fields(ByRef ctlLIST As Control, ByRef sRecordSource As Recordset)
    Dim X As String
    Dim i As Integer
    Dim sNumOfFields As Integer
    '// initialize value
    sNumOfFields = (sRecordSource.Fields.Count - 1)
    ctlLIST.Clear
    On Error Resume Next
         For i = 0 To sNumOfFields
             X = sRecordSource.Fields.Item(i).Name
             ctlLIST.AddItem X
             Next i
        i = i + 1
End Sub

Public Sub Drag_It(ByVal lngHwnd As Long)
Dim lngReturn As Long
    lngReturn = ReleaseCapture()
    lngReturn = SendMessage(lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, CLng(0))
End Sub

Public Sub loadForm(ByRef srcForm As Form, Optional ByVal Max As Boolean = False)
    srcForm.Show
    If Max Then
       srcForm.WindowState = vbMaximized
    Else
      srcForm.WindowState = vbNormal
    End If
    srcForm.SetFocus
End Sub
Public Sub LvwReplaceData(ByRef frm As Form, _
                      ByRef rs As Recordset, _
                      ByRef lv As ListView, _
                      Optional ByVal numOfFlds As Integer = 0)
Dim i As Integer
Dim NOF As Integer  'number of fields
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (rs.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 1 To NOF
  lv.SelectedItem.ListSubItems(i).Text = frm.txtEntry(i).Text
  Next i
End Sub

Public Sub Listview_Total(ByRef Lvw As ListView, ByRef srcRS As Recordset)
   With srcRS
     If .RecordCount < 1 Then
          Exit Sub
     End If
   End With
Dim rec_Count As Long
Dim isCurr As Boolean    'flag for currency or double
Dim X As ListItem
Dim strvalue As String
Dim iCount As Long  'to determine the last record
Dim i As Integer
Dim NumOfFields As Integer
On Error Resume Next
'//initialize value
isCurr = False
NumOfFields = (srcRS.Fields.Count - 1)
ReDim dblTotal(NumOfFields) As Double    'number of elements
For i = 0 To NumOfFields
   dblTotal(i) = 0
   Next i
   i = i + 1
iCount = 0
With srcRS
   rec_Count = CStr(srcRS.RecordCount)
   Set X = Lvw.ListItems.Add(, , "(" & rec_Count & ")" & "Record")
       X.Bold = True
       X.ForeColor = vbBlue
  .MoveFirst
While Not .EOF = True
iCount = iCount + 1
  For i = 1 To NumOfFields
       strvalue = Empty
      If Not IsNull(srcRS.Fields(i)) Then
             If srcRS.Fields(i).Type = 6 Or srcRS.Fields(i).Type = 5 Then
                 strvalue = toMoney(srcRS.Fields(i))
                 isCurr = True
'             ElseIf srcRS.Fields(i).Type = 3 Then
'                 strvalue = toNumber(srcRS.Fields(i))
'                 isCurr = False
'             ElseIf srcRS.Fields(i).Type = 202 Or srcRS.Fields(i).Type = 203 Then
'                 If IsNumeric(srcRS.Fields(i).Value) Then
'                     strvalue = toNumber(srcRS.Fields(i))
'                     isCurr = False
'                  End If
             Else
                 strvalue = ""
             End If
             '// dblTotal(0),dblTotal(1),dblTotal(2) and so on ...
             If isCurr = True Then
                 dblTotal(i) = dblTotal(i) + Val(strvalue)
             End If
            If iCount = .RecordCount Then
                   If dblTotal(i) > 0 Then
                      If isCurr = True Then
                        With X.ListSubItems.Add(, , Format(dblTotal(i), "standard"))
                           X.ListSubItems(i).Bold = True
                           X.ListSubItems(i).ForeColor = vbRed
                        End With
                      Else
                        With X.ListSubItems.Add(, , toNumber(dblTotal(i)))
                           X.ListSubItems(i).Bold = True
                           X.ListSubItems(i).ForeColor = vbRed
                        End With
                      End If
                    Else  '//dbltotal() = 0
                      With X.ListSubItems.Add(, , " - ")
                      End With
                    End If
           End If   '//icount
       Else
        If iCount = .RecordCount Then
          With X.ListSubItems.Add(, , " - ") 'if null string
          End With
        End If
       End If 'not isnull
    Next i
.MoveNext
Wend
End With
Set X = Nothing
End Sub


Public Sub BindDatasource(ByRef frm As Form, _
                          ByRef srcRS As Recordset, _
                          ByRef lv As ListView, _
                          Optional ByVal findFirst As Boolean = True, _
                          Optional ByVal numOfFlds As Integer = 0)
'//findFIRST - optional/false when use for next,previous,last,first
   If srcRS Is Nothing Then Exit Sub
With srcRS
  If .RecordCount = 0 Then
      Exit Sub
   End If
End With
If findFirst = False Then
 If srcRS.RecordCount > 0 Then
   If srcRS.EOF = True Then
       MsgBox "EOF reached.", vbInformation, "Bind Data!"
       srcRS.MoveLast
   ElseIf srcRS.BOF = True Then
       MsgBox "BOF reached.", vbInformation, "Bind Data!"
       srcRS.MoveFirst
   End If
 Else  'recordcount =0
     Exit Sub
 End If
End If
Dim abPos As Boolean   'absolutePosition
Dim i As Integer
Dim strFind As String
Dim strMatch As String
Dim NOF As Integer 'Number Of Feilds
'//
If srcRS.RecordCount = 0 Then Exit Sub
'// initialized
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 0 To NOF
   frm.txtEntry(i) = Empty
   Next i
If IsNumeric(TrimSpaces(CStr(lv.SelectedItem.Text))) Then
    strFind = TrimSpaces(CStr(lv.SelectedItem.Text))
    abPos = False
Else
    strFind = lv.SelectedItem.Index
    abPos = True
End If
If findFirst = True Then
 With srcRS
 .MoveFirst
   Do Until srcRS.EOF
   If abPos = False Then
        lv.MousePointer = vbHourglass
       strMatch = TrimSpaces(CStr(toNumber(srcRS.Fields(0))))
     Else
       lv.MousePointer = vbHourglass
       'slower//i use only on alpha type// so you can show the value one
       'row even if there is duplicate reference for viewing record
       'remember that reference must be a unique key
       strMatch = srcRS.Bookmark '// .AbsolutePosition
   End If
   If strMatch = strFind Then
         lv.MousePointer = vbDefault

         GoTo iFound
   Else
     .MoveNext
   End If
   Loop
 End With
 lv.MousePointer = vbDefault
End If 'findFirst
iFound:
With srcRS
         If srcRS.EOF = True Or srcRS.BOF = True Then Exit Sub
         For i = 0 To NOF
          If Not IsNull(srcRS.Fields(i)) Then
             frm.txtEntry(i) = FormatRS(srcRS.Fields(i))
              If srcRS.Fields(i).Type = 6 Or srcRS.Fields(i).Type = 5 Then
                frm.txtEntry(i).Alignment = 1
                 If Val(frm.txtEntry(i)) = 0 Then
                   frm.txtEntry(i).ForeColor = &HD38545
                 ElseIf Val(frm.txtEntry(i)) < 0 Then
                   frm.txtEntry(i).ForeColor = vbRed      ' if the value is negative
                 Else
                   frm.txtEntry(i).ForeColor = vbBlack
                End If
             Else                                          'string value and non-zero value
                 frm.txtEntry(i).ForeColor = vbBlack
            End If
          Else
              frm.txtEntry(i) = Empty
          End If
         Next i
    '//end of Search
End With

End Sub

Public Sub autoAlignCol(ByVal lv As ListView)
Dim col As Long
For col = 0 To lv.ColumnHeaders.Count - 1
    SendMessage lv.hWnd, LVM_SETCOLUMNWIDTH, col, LVSCW_AUTOSIZE_USEHEADER
Next col
End Sub



Public Sub Add_Item(ByRef recset As Recordset, _
                     ByRef fld As String, _
                     ByRef ctl As Control, _
                     Optional TrapDup As Boolean = False)
Dim uBnd As Long   'upperbound indeces/number of elements
Dim icnt As Integer 'count number of records
Dim txt1 As String
Dim i As Integer
'initialize
If recset.RecordCount = 0 Then Exit Sub
uBnd = (recset.RecordCount - 1)
ReDim item_added(uBnd)
For i = 0 To uBnd     'recset.RecordCount - 1
    item_added(i) = Empty
    Next i
txt1 = "!@#$%^&*()"
icnt = 0
On Error Resume Next
If recset.RecordCount = 0 Then Exit Sub
If recset.RecordCount > 0 Then
    recset.MoveFirst
    ctl.Clear
   While Not recset.EOF
      If Not IsNull(recset.Fields(fld)) Then
         If recset.Fields(fld) <> txt1 Then
'**           If TrapDup = True Then
'**              Dim X As Boolean
'**              X = alReady_Added(recset, recset.Fields(fld))
'**              If X = False Then
'**                ctl.AddItem recset.Fields(fld) 'add only one record to listbox
'**              End If
'**                item_added(icnt) = recset.Fields(fld)  'continue add to array
'**            Else  '//trapDup = false
                If recset.Fields(fld) <> "" Then
                   ctl.AddItem recset.Fields(fld)
                End If
'**            End If
         End If
        If Not IsNull(recset.Fields(fld)) Then
           txt1 = recset.Fields(fld)
        End If
      End If  'not isnull
        icnt = icnt + 1
        recset.MoveNext
   Wend
End If
End Sub

Public Sub SortListView(ByVal Lvw As MSComctlLib.ListView, _
                        ByVal colHdr As MSComctlLib.ColumnHeader)
'//Sort/ReSort ListView by the clicked column
'//<< syntax >>  SortListView ListView1, ColumnHeader

'//Sort by clicked ListView Column
'--set the sortkey to the column header's index - 1
Lvw.SortKey = colHdr.Index - 1
Lvw.Sorted = True

'--toggle the sort order between ascending & descending
Lvw.SortOrder = 1 Xor Lvw.SortOrder
End Sub
Public Sub ListView_Search(ByRef Lvw As ListView, _
                           ByVal sFind As String, _
                           Optional ByVal valSetting = 1)
Rem valSeeting :>> 0 = lvwtext ; 1 = lvwsubitem
'//input exact string ...
Dim itmFound As ListItem
If valSetting = 0 Then
  Set itmFound = Lvw.FindItem(sFind, 0, 1, 1)
Else
  Set itmFound = Lvw.FindItem(sFind, 1, 1, 1)
End If
  If Not itmFound Is Nothing Then
    itmFound.EnsureVisible
    itmFound.Selected = True
    Lvw.SetFocus
 End If
End Sub
Public Sub InsertColumn(ByRef lv As ListView, _
                        ByVal sRecordSource As Recordset, _
                        Optional ByVal sNumFields As Integer, _
                        Optional CH_clear As Boolean = True)
    With sRecordSource
     If .RecordCount = 0 Then
          Exit Sub
     ElseIf .EOF = True Or .BOF = True Then
          Exit Sub
     End If
   End With
    Dim X As String
    Dim i As Integer
    Dim idx As Integer 'index use to align column right
    Dim sNumOfColumn As Integer  'number of fields
    Dim clmHead As ColumnHeader
If sNumFields > 0 Then
   sNumOfColumn = sNumFields
Else
    sNumOfColumn = (sRecordSource.Fields.Count - 1)
End If
    '// initialize value
    If CH_clear = True Then
       lv.ColumnHeaders.Clear
    End If
    On Error Resume Next
         For i = 0 To sNumOfColumn
            X = SplitString(sRecordSource.Fields.Item(i).Name)
             Set clmHead = lv.ColumnHeaders.Add(, , X)
             '// align column data to right if it is currency or double
             '// no modify below this commented lines
             If sRecordSource.Fields.Item(i).Type = 6 Or _
                sRecordSource.Fields.Item(i).Type = 5 Then
                 idx = i + 1
                    lv.ColumnHeaders(idx).Alignment = lvwColumnRight
             End If
         Next i
End Sub
Public Sub FillListView(ByRef sListView As ListView, _
                        ByRef sRecSource As Recordset, _
                        ByVal sIcoNdx As Byte, _
                        Optional ByVal wchAlign As Integer)
'//set details
   With sRecSource
     If .RecordCount = 0 Then
          Exit Sub
     ElseIf .EOF = True Or .BOF = True Then
          Exit Sub
     End If
   End With
    Dim X As ListItem
    Dim i As Byte
    Dim sFieldsNum As Integer
    On Error Resume Next
    '//initialize
    sListView.ListItems.Clear
    sFieldsNum = (sRecSource.Fields.Count - 1)
    sRecSource.MoveFirst
    Do While Not sRecSource.EOF
         Set X = sListView.ListItems.Add(, , sRecSource.Fields(0), sIcoNdx, sIcoNdx)
         For i = 1 To sFieldsNum
               If Not IsNull(sRecSource.Fields(CInt(i))) Then
                  X.SubItems(i) = FormatRS(sRecSource.Fields(CInt(i)))
               End If
               If i = wchAlign Then
                 SendMessage sListView.hWnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE_USEHEADER
               End If
              Next i
        sRecSource.MoveNext
    Loop
End Sub
'<< ediwn delos santos>>
Public Sub Delete_Record(ByRef srcRS As Recordset, ByRef lvName As ListView)
Dim abPos As Boolean
Dim itemStr As Variant
Dim ans As Integer
Dim strMatch As String
Dim toDelete As String
'// INTIALIZE
toDelete = ""
strMatch = ""
If srcRS.RecordCount = 0 Then Exit Sub
If srcRS.EOF = True Or srcRS.BOF = True Then Exit Sub
itemStr = lvName.SelectedItem.Text
If IsNumeric(TrimSpaces(CStr(lvName.SelectedItem.Text))) Then
   toDelete = TrimSpaces(CStr(lvName.SelectedItem.Text))
   abPos = False
Else
    toDelete = lvName.SelectedItem.Index
    abPos = True
End If
ans = MsgBox("Are you Sure you want to delete selected item#:" & "( " & itemStr & ")" & "?", vbYesNo, "Delete")
If ans = vbYes Then
  With srcRS
     If .RecordCount = 0 Then Exit Sub
    .MoveFirst
     While Not .EOF
     If abPos = False Then
        lvName.MousePointer = vbHourglass
        strMatch = TrimSpaces(CStr(toNumber(srcRS.Fields(0))))
     Else
       lvName.MousePointer = vbHourglass
       'slower//i use only on alpha type// so you can show the value one
       'row even if there is duplicate reference for viewing record
       'remember that reference must be a unique key
       strMatch = srcRS.AbsolutePosition
   End If
        '// if record found
        If toDelete = strMatch Then
            '//delete current record
            srcRS.Delete
            lvName.ListItems.Remove lvName.SelectedItem.Index
            lvName.SetFocus
            lvName.MousePointer = vbDefault
            Exit Sub
        Else
          .MoveNext '//if record not found
       End If
     Wend
  End With  'rsprod
  
ElseIf ans = vbNo Then
  MsgBox "Deletion Cancelled!", , "Delete!"
End If
  lvName.MousePointer = vbDefault
End Sub

'<< edwin delos santos>>
Public Sub TxtLocked(ByRef frm As Form, ByRef idxList As ListBox)
    Dim i As Integer
    Dim idx As Integer
    On Error Resume Next
    For i = 0 To idxList.ListCount - 1
           idx = Val(idxList.List(i))                   'get the value from listbox, you can use array if you want
           frm.txtEntry(idx).Locked = True
         Next i
End Sub
'<< edwin delos sntos>>
Public Sub TextLocked(ByRef frm As Form, ByRef idxList As ListBox)
    Dim i As Integer
    Dim idx As Integer
    On Error Resume Next
    For i = 0 To idxList.ListCount - 1
           idx = Val(idxList.List(i))                   'get the value from listbox, you can use array if you want
           frm.TextEntry(idx).Locked = True
         Next i
End Sub
'<<edwin delos santos>>
Public Sub WriteData(ByRef frm As Form, ByRef srcRS As Recordset, _
                      ByVal newRec As Boolean, _
                      Optional ByVal srcNumFlds As Integer = 0)
'//addnew = true for new record else false > forced
'//srcnumflds = number of fields loaded in textbox  > optional
                'if not all fields are loaded, srcnumflds is equal to text upperbound indeces
                'based on the numbers of textbox showed in the form (see enabled textbox procedures)
If srcRS Is Nothing Then Exit Sub
If srcRS.RecordCount > 0 Then
 If srcRS.EOF = True Or srcRS.BOF = True Then
   'MsgBox "Either EOF or BOF reached.", vbInformation, "Write Data!"
   'Exit Sub
   srcRS.MoveLast
 End If
End If
Dim i As Integer
Dim NOF As Integer 'Number Of Feilds
If srcNumFlds > 0 Then
   NOF = srcNumFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
ReDim entries(NOF) As TextBox
For i = 0 To NOF
    Set entries(i) = frm.txtEntry(i)  'm tired of using frm, set number of elements allowed
    Next i
i = 0
With srcRS
  If newRec = True Then
      .AddNew
  End If
      For i = 0 To NOF
      Select Case srcRS.Fields.Item(i).Type
       Case Is = 3   'integer
           If IsNumeric(entries(i).Text) Then
              srcRS.Fields(i) = toNumber(entries(i).Text)
           End If
      Case Is = 5, 6  'currency or double
           If IsNumeric(entries(i).Text) Then
             srcRS.Fields(i) = toMoney(entries(i).Text)
           End If
       Case Is = 7   'date
           If IsDate(entries(i).Text) Then
               srcRS.Fields(i) = CDate(entries(i).Text)
           Else '//save empty entry
               srcRS.Fields(i) = Null
           End If
       Case Is = 202, 203    'text, memo
             srcRS.Fields(i) = CStr(entries(i).Text)
      End Select
      Next i
      .Update
End With
End Sub


'<<edwin delos santos>>
Public Sub ShowFldsLabel(ByRef frm As Form, _
                         ByRef srcRS As Recordset, _
                         Optional ByVal numOfFlds As Integer = 0, _
                         Optional ByVal strChar As String = " ")
Dim i As Integer
Dim splitCHR As String
Dim NOF As Integer 'Number Of Feilds
'// initialized
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 0 To NOF
      frm.lblFLDi(i) = ""              'caption for each field
   Next i
i = 0
For i = 0 To NOF
        If Not IsNull(srcRS.Fields(i)) Then
              splitCHR = SplitString(srcRS.Fields(i).Name)
              frm.lblFLDi(i) = splitCHR & strChar '" :"
              'frm.lblFLDi(i) = srcRS.Fields(i).Name & " :"
        Else
               frm.lblFLDi(i) = srcRS.Fields(i).Name & strChar '" :"
        End If
        Next i
End Sub

'<< edwin delos santos>
Public Sub errorMsg(ByVal errNUM As ErrObject, _
                    Optional ByVal ModuleName As String, _
                    Optional ByVal OccurIn As String)
 Select Case errNUM
 Case Is = 0
   Exit Sub
 Case Is = 5
   MsgBox "Invalid procedure call or argument", vbCritical, "Warning!"
   Exit Sub
 Case Is = 13
   MsgBox "Data type mismatch!", vbCritical, "Warning!"
   Exit Sub
' Case Is = 3021  'requested operations require a curren record. Current Record has been deleted
' Case Is = 340   'Array doesnot exist
 Case Is = 32755 'Cancelled open
   Exit Sub
' Case Is = 3704  'Operation is not allowed whent the object is close
' Case Is = 9     'Subscrip out of range
 Case Is = 7005
   MsgBox "RowSet not available!", vbInformation, "Warning!"
   Exit Sub
 Case Is = -2147217843
   MsgBox "Not a valid password!", vbInformation, "Enter valid password"
   Exit Sub
 Case Is = -2147217887
   MsgBox "Cannot update (expression)!", vbInformation, "Field not updatable."
   Exit Sub
 Case Is = 3709
   Dim errMsg As String
   errMsg = "The connection cannot be used"
   errMsg = errMsg & Chr(10) & "to perform this operation"
   errMsg = errMsg & Chr(10) & "It is either closed or invalid"
   errMsg = errMsg & Chr(10) & "in this context.!"
   MsgBox errMsg, vbCritical, "Disconnected Recordset"
   Exit Sub
  Case Else
   MsgBox "Error From: " & ModuleName & vbNewLine & _
           "Occur In: " & OccurIn & vbNewLine & _
           "Error Number: " & errNUM.Number & vbNewLine & _
           "Description: " & errNUM.Description, vbCritical, "Application Error"
    'Save the error log (The save error log will be display later on in the program)
    Open App.Path & "\ErrorLog.log" For Append As #1
                Print #1, Format(Date, "MMM-dd-yyyy") & "]~~~~[" & Time & "]~~~~[" & Err.Number & "]~~~~[" & Err.Description & "]~~~~[" & ModuleName & "]~~~~[" & OccurIn
    Close #1
 End Select
End Sub




