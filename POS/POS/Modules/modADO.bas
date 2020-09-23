Attribute VB_Name = "modADO"

Option Explicit

Public Function MySql(sServer As String, sUser As String, sPass As String, sDB As String)
  On Error GoTo err:
  CN.Open "DRIVER={MySQL ODBC 3.51 Driver};Server=" & sServer & ";UID=" & sUser & ";PWD=" & sPass & ";Database=" & sDB
 
  Exit Function
err:
   If MsgBox("Description: " & err.Description & " Err#: " & err.Number, vbRetryCancel) = vbRetry Then
      Resume
      Else
    
    End If
End Function


Public Function execSQL(ByVal sSQL As String)
  On Error GoTo err:

   With RS
   .CursorLocation = adUseClient
   .Open sSQL, CN, adOpenKeyset, adLockOptimistic
   End With
   
  Exit Function
  
err:
 If MsgBox("Description: " & err.Description & " Err#: " & err.Number, vbRetryCancel) = vbRetry Then
      
      Resume
      
      Else
      
    End If
    End Function
Public Sub closeDB()
Set CN = Nothing
Set RS = Nothing
 CN.Close
End Sub

