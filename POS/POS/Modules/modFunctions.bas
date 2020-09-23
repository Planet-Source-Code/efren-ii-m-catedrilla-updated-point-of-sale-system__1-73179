Attribute VB_Name = "modFunctions"
Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C


Public Function is_Numeric(ByRef txt As String) As Boolean
If IsNumeric(txt) = False Then
  is_Numeric = False
  MsgBox "The field required a numeric input.Please check it!", vbExclamation
  Exit Function
  Else
  is_Numeric = True
  End If
End Function

Public Function is_Empty(ByRef txt As TextBox) As Boolean
 If txt.Text = "" Then
   is_Empty = False
   MsgBox "The field required is empty.Please check it!", vbExclamation
   Exit Function
   Else
   is_Empty = True
   End If
End Function




Public Function Autocomplete(cbCombo As ComboBox, sKeyAscii As Integer, Optional bUpperCase As Boolean = True) As Integer
    Dim lngFind As Long, intPos As Integer, intLength As Integer
    Dim tStr As String


    With cbCombo


        If sKeyAscii = 8 Then
            If .SelStart = 0 Then Exit Function
            .SelStart = .SelStart - 1
            .SelLength = 32000
            .SelText = ""
        Else
            intPos = .SelStart '// save intial cursor position
            tStr = .Text '// save string


            If bUpperCase = True Then
                .SelText = Chr(sKeyAscii) '// change string. (uppercase only)
            Else
                .SelText = Chr(sKeyAscii) '// change string. (leave case alone)
            End If
        End If
        
        lngFind = SendMessage(.hwnd, CB_FINDSTRING, 0, ByVal .Text) '// Find string in combobox


        If lngFind = -1 Then '// if string not found
            .Text = tStr '// set old string (used for boxes that require charachter monitoring
            .SelStart = intPos '// set cursor position
            .SelLength = (Len(.Text) - intPos) '// set selected length
            Autocomplete = 0 '// return 0 value to KeyAscii
            Exit Function
            
        Else '// If string found
            intPos = .SelStart '// save cursor position
            intLength = Len(.List(lngFind)) - Len(.Text) '// save remaining highlighted text length
            .SelText = .SelText & Right(.List(lngFind), intLength) '// change new text in string
            '.Text = .List(lngFind)'// Use this inst
            '     ead of the above .Seltext line to set th
            '     e text typed to the exact case of the it
            '     em selected in the combo box.
            .SelStart = intPos '// set cursor position
            .SelLength = intLength '// set selected length
        End If
    End With
    
End Function


Public Function PrevUpdate(ByVal sUser As String) As Boolean

execSQL "select * from privileges where users='" & sUser & "' and update_data=1"

If RS.RecordCount = 0 Then
 MsgBox "You don't have the privileges to access this area.", vbInformation
 PrevUpdate = True
 
 Else
 PrevUpdate = False
 
 End If
 
 Set RS = Nothing
 Exit Function
End Function

Public Function PrevPass(ByVal sUser As String) As Boolean

execSQL "select * from privileges where users='" & sUser & "' and change_pass=1"

If RS.RecordCount = 0 Then
 MsgBox "You don't have the privileges to access this area.", vbInformation
 PrevPass = True
 
 Else
 PrevPass = False
 
 End If
 

 Set RS = Nothing
 Exit Function
 
End Function

Public Function PrevPrint(ByVal sUser As String) As Boolean

execSQL "select * from privileges where users='" & sUser & "' and print_reports=1"

If RS.RecordCount = 0 Then
 MsgBox "You don't have the privileges to access this area.", vbInformation
 PrevPrint = True
 
 Else
 PrevPrint = False
 
 End If
 

 Set RS = Nothing
 Exit Function
 
End Function

Public Function PrevDelete(ByVal sUser As String) As Boolean

execSQL "select * from privileges where users='" & sUser & "' and delete_data=1"

If RS.RecordCount = 0 Then
 MsgBox "You don't have the privileges to access this area.", vbInformation
 PrevDelete = True
 
 Else
 PrevDelete = False
 
 End If
 

 Set RS = Nothing
 Exit Function
 
End Function

Public Function PrevAdmin(ByVal sUser As String) As Boolean

execSQL "select * from privileges where users='" & sUser & "' and admin=1"

If RS.RecordCount = 0 Then
 MsgBox "You don't have the privileges to access this area.", vbInformation
 PrevAdmin = True
 
 Else
 PrevAdmin = False
 
 End If

 Set RS = Nothing
 Exit Function
End Function





