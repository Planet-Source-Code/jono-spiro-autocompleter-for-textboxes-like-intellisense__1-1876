<div align="center">

## Autocompleter for textboxes\-Like IntelliSense


</div>

### Description

'This function can be implemented anywhere to finish off a word in a textbox using a list of words with a custom delimeter. It is fairly complex and difficult to document, so bare with me. It also uses the amazing extract argument function written by another code of the day submitter. I have spent lots of time fine tuning this code and making it as flexible and foolproof as the one used in Internet Explorer 4.0.
 
### More Info
 
'Usage: textComplete textBox object, the words list to search through, the last key hit (use keyCode)

'example: textComplete text1,"www.microsoft.com,www.microwave.com",keyCode

'It is very complex code, I have spent a while beta testing it to make sure no modifications are needed. Documenting this would have taken too long and would be difficult to understand, so please just trust the code, it will fry your brain if you try to pull it apart and understand it...

'Returns nothing

'No Side Effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jono Spiro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jono-spiro.md)
**Level**          |Unknown
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jono-spiro-autocompleter-for-textboxes-like-intellisense__1-1876/archive/master.zip)

### API Declarations

```
'None
```


### Source Code

```
'If you want to test this code, I have written a complex program that not only demonstrates how the code works, but it also allows you to dynamically change the delimeter of the textList and, when adding to the list a new word, if the word uses a character that is already being used as the delimeter, it finds a new delimeter so that you can still add the item. First add 3 text fields, and three labels to the form. Name the fields txtType,txtDelim,txtList.
'add this code to the form:
'THIS IS ALL OPTIONAL
Public lastDelimeter As String
Option Compare Text
Private Sub Form_Load()
 Width = 7860
 Height = 1500
 Label1.Caption = "List to search from:"
 Label1.AutoSize = True
 Label1.Left = 45
 Label1.Top = 135
 Label3.Caption = "Text Delimeter:"
 Label3.AutoSize = True
 Label3.Left = 315
 Label3.Top = 450
 Label2.Caption = "Type text here:"
 Label2.AutoSize = True
 Label2.Left = 315
 Label2.Top = 765
 txtDelim.Left = 1395
 txtType.Left = 1395
 txtList.Left = 1395
 txtDelim.Width = 5505
 txtType.Width = 5505
 txtList.Width = 5505
 txtList.Top = 90
 txtDelim.Top = 405
 txtType.Top = 720
 txtDelim.Height = 285
 txtType.Height = 285
 txtList.Height = 285
 txtDelim.Text = ","
 txtList.Text = "greg,gregory,tom,dick,harry,www.microsoft.com,www.microware.com"
 lastDelimeter = txtDelim.Text
End Sub
Private Sub Form_Resize()
 txtType.Width = ScaleWidth - 1500
 txtList.Width = ScaleWidth - 1500
 txtDelim.Width = ScaleWidth - 1500
 Height = 1500
End Sub
Private Sub txtType_KeyPress(KeyAscii As Integer)
 Dim a As Integer
 If KeyAscii = vbKeyReturn And txtType.Text <> "" And txtList.Text <> "" And InStr(txtType.Text, lastDelimeter) = 0 Then
 txtList.Text = txtList.Text & txtDelim.Text & txtType.Text
 ElseIf KeyAscii = vbKeyReturn And txtType.Text <> "" And InStr(txtType.Text, lastDelimeter) = 0 Then
 txtList.Text = txtType.Text
 ElseIf KeyAscii = vbKeyReturn And InStr(txtType.Text, lastDelimeter) Then
 For a = 255 To 0 Step -1
 If InStr(txtType.Text & lastDelimeter & txtList.Text, Chr(a)) = 0 Then
 Exit For
 ElseIf a = 1 And InStr(txtType.Text & lastDelimeter & txtList.Text, Chr(a)) Then
 MsgBox "Error: there are no unique delimeters left, cannot add to datalist."
 Exit Sub
 End If
 Next a
 txtDelim.Text = Chr(a)
 Dim List As String, b As Integer: b = 0
 For a = Len(txtList.Text) To 1 Step -1
 If Mid$(txtList.Text, a, Len(lastDelimeter)) = lastDelimeter Then
 List = List & a & ","
 b = b + 1
 End If
 Next a
 For a = 1 To b
 txtList.SetFocus
 txtList.SelStart = ExtractArgument(a, List, ",") - 1
 txtList.SelLength = Len(lastDelimeter)
 txtList.SelText = txtDelim.Text
 txtType.SetFocus
 Next a
 lastDelimeter = txtDelim.Text
 txtList.Text = txtList.Text & lastDelimeter & txtType.Text
 ElseIf txtDelim.Text <> lastDelimeter Then
 txtDelim.Text = lastDelimeter
 MsgBox "You can only enter delimeter characters that do not exist in the list."
 End If
End Sub
Private Sub txtType_KeyUp(KeyCode As Integer, Shift As Integer)
 textComplete txtType, txtList.Text, txtDelim.Text, KeyCode
End Sub
Private Sub txtDelim_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
 If InStr(txtList.Text, txtDelim.Text) = 0 Then
 Dim List As String, a As Integer, b As Integer: b = 0
 For a = Len(txtList.Text) To 1 Step -1
 If Mid$(txtList.Text, a, Len(lastDelimeter)) = lastDelimeter Then
 List = List & a & ","
 b = b + 1
 End If
 Next a
 For a = 1 To b
 txtList.SelStart = ExtractArgument(a, List, ",") - 1
 txtList.SelLength = Len(lastDelimeter)
 txtList.SelText = txtDelim.Text
 Next a
 lastDelimeter = txtDelim.Text
 ElseIf txtDelim.Text <> lastDelimeter Then
 txtDelim.Text = lastDelimeter
 MsgBox "You can only enter delimeter characters that do not exist in the list."
 End If
 End If
End Sub
'END OF EXAMPLE CODE
'
'
'THIS IS THE ACTUAL CODE FOR THE FUNCTION FROM HERE ON TO THE BOTTOM
'ALL ABOVE IS OPTIONAL!!
Function textComplete(textBox As textBox, searchList As String, delimeter As String, keyHit As Integer)
 '''''''''''''''''''''''''''''''''''''''''''
 'Place me in the KeyUp script of a textbox'
 'Usage: textComplete textBox object, the words to search through, the last key hit (use keyCode)
 '''''''''''''''''''''''''''''''''''''''''''
 With textBox
 If keyHit <> vbKeyBack And keyHit > 48 Then
 Dim List As String, a As Integer, searchText As String, numDelim As Integer: numDelim = 0
 For a = 1 To Len(searchList)
 If Mid$(searchList, a, 1) = delimeter Then numDelim = numDelim + 1
 Next a
 For a = 1 To numDelim + 1
 searchText = ExtractArgument(a, searchList, delimeter)
 If InStr(searchText, .Text) And (Left$(.Text, 1) = Left$(searchText, 1)) And .Text <> "" Then
 .SelText = ""
 .SelLength = 0
 length = Len(.Text)
 .Text = .Text & Right$(searchText, Len(searchText) - Len(.Text))
 .SelStart = length
 .SelLength = Len(.Text)
 Exit Function
 End If
 Next a
 End If
 End With
End Function
Function ExtractArgument(ArgNum As Integer, srchstr As String, Delim As String) As String
 On Error GoTo Err_ExtractArgument
 Dim ArgCount As Integer
 Dim LastPos As Integer
 Dim Pos As Integer
 Dim Arg As String
 Arg = ""
 LastPos = 1
 If ArgNum = 1 Then Arg = srchstr
 Do While InStr(srchstr, Delim) > 0
 Pos = InStr(LastPos, srchstr, Delim)
 If Pos = 0 Then
 'No More Args found
 If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
 Exit Do
 Else
 ArgCount = ArgCount + 1
 If ArgCount = ArgNum Then
 Arg = Mid(srchstr, LastPos, Pos - LastPos)
 Exit Do
 End If
 End If
 LastPos = Pos + 1
 Loop
 ExtractArgument = Arg
 Exit Function
Err_ExtractArgument:
 MsgBox "Error " & Err & ": " & Error
 Resume Next
End Function
```

