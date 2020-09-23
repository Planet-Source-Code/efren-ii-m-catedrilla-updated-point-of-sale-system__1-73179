Attribute VB_Name = "modProcedures"
Option Explicit


Public Sub ClearText(ByRef frm As Form)
Dim CONTROL As CONTROL
For Each CONTROL In frm.Controls
If (TypeOf CONTROL Is TextBox) Then CONTROL = vbNullString
Next CONTROL
  Set CONTROL = Nothing
End Sub

Public Sub HLText(ByRef txt)
On Error Resume Next
  With txt
  .SelStart = 0
  .SelLength = Len(txt.Text)
  End With
End Sub

Public Sub CenterForm(frm As Form)
Dim x As Integer
Dim Y As Integer

    x = (Screen.Height - frm.Height) / 3
    Y = (Screen.Width - frm.Width) / 2
    frm.Move Y, x

End Sub

Public Sub Proper_Case(ByVal sText As Variant)
   sText.Text = StrConv(sText, vbProperCase)
End Sub

