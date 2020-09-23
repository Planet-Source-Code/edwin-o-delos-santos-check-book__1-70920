Attribute VB_Name = "modFUNCTION"
Option Explicit
'** alReady_Added
'** FilterNumber
'** FilterString
'** FormatRS
'** FormLoaded  (determine if already load)
'** IsAmount
'** last_Recc
'** print_Init
'** Print_Amt
'** prnCenterText
'** SplitString
'** toMoney
'** toNumber
'** TrimSpaces
'** WordWrap  <syntaxt> text1 = wordwrap(text2)
'=====================================
Public Function WordWrap(ByVal strExpression As String, ByVal lLength As Long) As String
    Dim BufferCrLf() As String, BufferSpace() As String, Buffer As String
    Dim K As Long, J As Long, iCount As Long

    BufferCrLf() = Split(strExpression, vbCrLf)
    For K = LBound(BufferCrLf()) To UBound(BufferCrLf())
        If Len(BufferCrLf(K)) <= lLength Then
            Buffer = Buffer & BufferCrLf(K) & vbCrLf
        Else
            BufferSpace() = Split(BufferCrLf(K), " ")
            For J = 0 To UBound(BufferSpace())
                iCount = iCount + Len(BufferSpace(J)) + 1
                If (iCount <= lLength) Then
                    Buffer = Buffer & BufferSpace(J) & " "
                Else
                    iCount = 0
                    Buffer = Buffer & vbCrLf & BufferSpace(J) & " "
                    iCount = Len(BufferSpace(J)) + 1
                End If
            Next J
            Buffer = Buffer & vbCrLf
        End If
    Next K
    WordWrap = Buffer
End Function
 
 
Public Function FormLoaded(FormName As String) As Boolean
Dim i As Integer, fnamelc As String
fnamelc = LCase$(FormName)
FormLoaded = False
For i = 0 To Forms.Count - 1
  If LCase$(Forms(i).Name) = fnamelc Then
    FormLoaded = True
    Exit Function
  End If
Next
End Function

Public Function FilterString(ByVal Text As String, _
                             ByVal validChars As String) As String
    Dim i As Long, result As String
    For i = 1 To Len(Text)
        If InStr(validChars, Mid$(Text, i, 1)) Then
            result = result & Mid$(Text, i, 1)
        End If
    Next
    FilterString = result
End Function

Public Function FilterNumber(ByRef Text As String, Optional ByVal TrimZeros As Boolean = False) As String
    Dim decSep As String, i As Long, result As String
    '// Retrieve the decimal separator symbol.
    decSep = Format$(0.1, ".")
    '// Use FilterString for most of the work.
    result = FilterString(Text, decSep & "-0123456789")
    '// Do the following only if there is a decimal part and the
    '// user requested that nonsignificant digits be trimmed.
    If TrimZeros And InStr(Text, decSep) > 0 Then
        For i = Len(result) To 1 Step -1
            Select Case Mid$(result, i, 1)
                Case decSep
                    result = Left$(result, i - 1)
                    Exit For
                Case "0"
                    result = Left$(result, i - 1)
                Case Else
                    Exit For
            End Select
        Next
    End If
    FilterNumber = result
End Function


Public Function prnCenterText(ByRef txt As String, ByVal b_line As Integer) As String
Dim s_col As Integer
s_col = (b_line - Len(txt)) / 2
Printer.Print Tab(s_col); txt
End Function

Public Function Print_Amt(ByRef iTab As Integer, ByVal amt As Double, maxLEN As Integer) As Double
 Dim intLEN As Integer, currtab As Integer

 intLEN = Len(Trim(Format(amt, "Standard")))
 Dim i As Integer
  intLEN = Len(Format(amt, "#,###,##0.00"))
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
  If Val(amt) > 0 Then
   Printer.Print Tab(currtab); Format(amt, "#,###,##0.00");
  Else
     Printer.Print Tab(currtab); "--";
  End If
End Function
Public Function print_Init(ByRef lst As ListBox) As Boolean
    Dim i As Integer
    Dim ii As Integer
    On Error Resume Next
    '// initialize
     For ii = 0 To 50
         printIndex(ii) = Empty
         Next ii
    For i = 0 To lst.ListCount - 1
           lst.Selected(i) = False
         Next i
    '//
    If lst.ListCount = 0 Then Exit Function
    For i = 0 To lst.ListCount - 1
            lst.Selected(i) = True
            printIndex(i) = lst.Text
         Next i
      print_Init = True
End Function

Public Function TrimSpaces(Text As String) As String
    Dim Loop1 As Long, SpaceCheck As String
    Dim FullString As String
    For Loop1 = 1 To Len(Text)
        SpaceCheck = Mid(Text, Loop1, 1)
        If SpaceCheck <> " " Then
            FullString = FullString & SpaceCheck
        End If
    Next Loop1
    TrimSpaces = FullString
End Function

'function to check if record already added to listbox
'//coded by: edwin delos santos
Public Function alReady_Added(ByRef rs As Recordset, ByRef srcStr As String) As Boolean
  Dim i As Integer
On Error Resume Next
For i = 0 To rs.RecordCount - 1
    'find and match record from array; if found added = true
    If srcStr = item_added(i) Then
      alReady_Added = True
      i = 0
      Exit Function
    Else
      alReady_Added = False
    End If
  Next i
End Function

'// convert string to number / return the current value
'* Test   :  1,222
'* result :  1222
Public Function toNumber(ByVal srcNum As String) As Long
Dim numba As Long
If IsNumeric(srcNum) Then
  If Val(srcNum) = 0 Then srcNum = 0
  numba = Val(CLng(srcNum))
  toNumber = numba
  numba = 0
End If
End Function

'Function that will return a current format
'* Test  :  1,222.45
'* result:  1222.45
Public Function toMoney(ByVal srcCurr As String) As Double
 Dim sdbl As Double
 If IsNumeric(srcCurr) Then
   If Val(srcCurr) = 0 Then srcCurr = 0
   sdbl = Val(CDbl(srcCurr))
   toMoney = sdbl
   sdbl = 0
 End If
End Function

Public Function IsAmount(ByVal txt As String) As Boolean
    Dim ch As String
    Dim isamountentry As Boolean
    Dim i As Integer, J As Integer
    isamountentry = False
    
    If Len(LTrim(RTrim(txt))) = 0 Then
        IsAmount = False
        Exit Function
    End If
    J = 0
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch < "0" Or ch > "9" Then
            If ch <> "." Then
                IsAmount = False
                Exit Function
            Else
                J = J + 1
            End If
        End If
    Next i
    If J > 1 Then
        Exit Function
    End If
    IsAmount = True
End Function

Public Function SplitString(ByVal strText As String) As String
'label1 = edwin_delos_santos
'<< syntax >>
'Private Sub Command1_Click()
'   Text1.Text = RemoveAllNonAlphaNumeric(Label1.Caption)
'End Sub
'result:  text1 = "edwin delos santos"
    Dim strResult As String
    Dim i As Integer

    For i = 1 To Len(strText)
        Select Case Asc(Mid$(strText, i, 1))
        Case 48 To 57, 65 To 90, 97 To 122 'a digit or Uppercase Alphabet or Lowercase Alphabet
            strResult = strResult & Mid$(strText, i, 1)
        Case Else 'Reject any other key.
            strResult = strResult & Space(1) 'add space
        End Select
    Next i

ExitHere:
    SplitString = strResult
End Function

Public Function FormatRS(ByVal srcField As Field) As String
    Dim strRet As String
     With srcField
        If srcField.Type = adCurrency Or srcField.Type = adDouble Then
            strRet = Format$(srcField, "#,###,##0.00")
        ElseIf srcField.Type = 7 Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        ElseIf srcField.Type = 3 Then
           If IsNumeric(srcField) Then
             strRet = Format$(srcField, "###,##0")
           End If
        ElseIf srcField.Type = 202 Or srcField.Type = 203 Then
            strRet = CStr(srcField)
        End If
    End With
    FormatRS = strRet
    strRet = vbNullString
End Function

Public Function Last_Recc(ByRef rs As Recordset, _
                          Optional ByVal fld As Long = 0) As Long
If rs Is Nothing Then Exit Function
Dim maxRecc As Long, maxExist As Boolean
Dim LastRec As Long
Dim recExist As Long
'// initialize
  maxRecc = 0
  LastRec = 0
  maxExist = False
With rs
    .Requery
  If .RecordCount = 0 Then
      Last_Recc = 1
      Exit Function
  End If
   LastRec = .RecordCount + 1
  .MoveFirst
  If Not IsNumeric(.Fields(fld)) Then Exit Function    '// determine if the first field is numeric
  While Not .EOF()
     If maxRecc > recExist Then
        maxRecc = maxRecc
        maxExist = True
     End If
     recExist = .Fields(fld)       '//current number encountered
     If recExist > LastRec Then
        Last_Recc = recExist + 1
        If maxExist = False Then
          maxRecc = recExist     '// Rem determine the largest number >
        End If
: Rem debug.print maxrecc
     ElseIf recExist = LastRec Then
       Last_Recc = LastRec + 1
     ElseIf recExist < LastRec Then
       Last_Recc = LastRec
    End If
    .MoveNext
  Wend
  If maxExist = True Then
    Last_Recc = maxRecc + 1
  End If
End With
End Function

